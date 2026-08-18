"""
mock_sharepoint.py
Built-in Mock mechanism for local development and testing.
When real Azure credentials are unavailable, the system automatically falls back
to this local simulation mode, ensuring all APIs run out of the box.
"""

import json
import os
import shutil
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, BinaryIO, Dict, List, Optional

# ---------------------------------------------------------------------------
# Configurable paths
# ---------------------------------------------------------------------------
DEFAULT_DB_PATH = Path("./data/campus_sync.db")
DEFAULT_MOCK_FILES_DIR = Path("./data/mock_files")


def _ensure_dirs():
    DEFAULT_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    DEFAULT_MOCK_FILES_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# SQLite schema bootstrap
# ---------------------------------------------------------------------------
INIT_SQL = """
CREATE TABLE IF NOT EXISTS tickets (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    description TEXT,
    category TEXT,
    reporter_email TEXT,
    status TEXT DEFAULT 'open',
    priority TEXT DEFAULT 'medium',
    assigned_to TEXT,
    sharepoint_item_id TEXT,
    created_at TEXT,
    updated_at TEXT,
    metadata TEXT
);

CREATE TABLE IF NOT EXISTS mock_files (
    id TEXT PRIMARY KEY,
    drive_id TEXT,
    name TEXT,
    local_path TEXT,
    mime_type TEXT,
    size INTEGER DEFAULT 0,
    created_at TEXT,
    updated_at TEXT
);
"""


def _get_conn(db_path: Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    _ensure_dirs()
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript(INIT_SQL)
    conn.commit()
    return conn


# ---------------------------------------------------------------------------
# Ticket helpers
# ---------------------------------------------------------------------------
def _row_to_ticket(row: sqlite3.Row) -> Dict[str, Any]:
    ticket = dict(row)
    meta = ticket.pop("metadata", None)
    if meta:
        try:
            ticket.update(json.loads(meta))
        except Exception:
            pass
    return ticket


# ---------------------------------------------------------------------------
# Mock SharePoint Client
# ---------------------------------------------------------------------------
class MockSharePointClient:
    """
    Local mock implementation of SharePoint List CRUD and Document Library
    file operations backed by SQLite and the local filesystem.
    """

    def __init__(
        self,
        db_path: Optional[Path] = None,
        files_dir: Optional[Path] = None,
    ):
        self.db_path = db_path or DEFAULT_DB_PATH
        self.files_dir = files_dir or DEFAULT_MOCK_FILES_DIR
        _ensure_dirs()
        # ensure schema
        conn = _get_conn(self.db_path)
        conn.close()

    # -- Tickets (SharePoint List simulation) -------------------------------

    def create_ticket(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        conn = _get_conn(self.db_path)
        try:
            ticket_id = str(uuid.uuid4())
            now = datetime.utcnow().isoformat() + "Z"
            # Extract known fields
            title = payload.get("title", "Untitled")
            description = payload.get("description", "")
            category = payload.get("category", "general")
            reporter_email = payload.get("reporter_email", "")
            status = payload.get("status", "open")
            priority = payload.get("priority", "medium")
            assigned_to = payload.get("assigned_to", "")

            # Remaining fields go into metadata json
            known = {
                "title",
                "description",
                "category",
                "reporter_email",
                "status",
                "priority",
                "assigned_to",
            }
            metadata = {k: v for k, v in payload.items() if k not in known}

            conn.execute(
                """
                INSERT INTO tickets
                (id, title, description, category, reporter_email, status,
                 priority, assigned_to, created_at, updated_at, metadata)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    ticket_id,
                    title,
                    description,
                    category,
                    reporter_email,
                    status,
                    priority,
                    assigned_to,
                    now,
                    now,
                    json.dumps(metadata) if metadata else None,
                ),
            )
            conn.commit()
            return self.get_ticket(ticket_id)
        finally:
            conn.close()

    def list_tickets(
        self,
        status: Optional[str] = None,
        category: Optional[str] = None,
        reporter_email: Optional[str] = None,
        limit: int = 100,
        offset: int = 0,
    ) -> List[Dict[str, Any]]:
        conn = _get_conn(self.db_path)
        try:
            query = "SELECT * FROM tickets WHERE 1=1"
            params: List[Any] = []
            if status:
                query += " AND status = ?"
                params.append(status)
            if category:
                query += " AND category = ?"
                params.append(category)
            if reporter_email:
                query += " AND reporter_email = ?"
                params.append(reporter_email)
            query += " ORDER BY created_at DESC LIMIT ? OFFSET ?"
            params.extend([limit, offset])

            cur = conn.execute(query, params)
            rows = cur.fetchall()
            return [_row_to_ticket(r) for r in rows]
        finally:
            conn.close()

    def get_ticket(self, ticket_id: str) -> Optional[Dict[str, Any]]:
        conn = _get_conn(self.db_path)
        try:
            cur = conn.execute("SELECT * FROM tickets WHERE id = ?", (ticket_id,))
            row = cur.fetchone()
            return _row_to_ticket(row) if row else None
        finally:
            conn.close()

    def update_ticket(
        self, ticket_id: str, payload: Dict[str, Any]
    ) -> Optional[Dict[str, Any]]:
        conn = _get_conn(self.db_path)
        try:
            # Build dynamic update
            allowed = {"title", "description", "category", "status", "priority", "assigned_to"}
            fields = []
            values: List[Any] = []
            for k, v in payload.items():
                if k in allowed:
                    fields.append(f"{k} = ?")
                    values.append(v)
            if not fields:
                return self.get_ticket(ticket_id)

            now = datetime.utcnow().isoformat() + "Z"
            fields.append("updated_at = ?")
            values.append(now)
            values.append(ticket_id)

            conn.execute(
                f"UPDATE tickets SET {', '.join(fields)} WHERE id = ?",
                values,
            )
            conn.commit()
            return self.get_ticket(ticket_id)
        finally:
            conn.close()

    def delete_ticket(self, ticket_id: str) -> bool:
        conn = _get_conn(self.db_path)
        try:
            cur = conn.execute("DELETE FROM tickets WHERE id = ?", (ticket_id,))
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()

    # -- Document Library simulation ----------------------------------------

    def upload_file(
        self,
        drive_id: str,
        file_name: str,
        content: BinaryIO,
        mime_type: Optional[str] = None,
    ) -> Dict[str, Any]:
        _ensure_dirs()
        file_id = str(uuid.uuid4())
        local_path = self.files_dir / f"{file_id}_{file_name}"
        with open(local_path, "wb") as f:
            shutil.copyfileobj(content, f)

        size = local_path.stat().st_size
        now = datetime.utcnow().isoformat() + "Z"
        conn = _get_conn(self.db_path)
        try:
            conn.execute(
                """
                INSERT INTO mock_files (id, drive_id, name, local_path, mime_type, size, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (file_id, drive_id, file_name, str(local_path), mime_type or "application/octet-stream", size, now, now),
            )
            conn.commit()
        finally:
            conn.close()

        return {
            "id": file_id,
            "name": file_name,
            "size": size,
            "mimeType": mime_type or "application/octet-stream",
            "createdDateTime": now,
            "webUrl": f"file://{local_path.resolve()}",
        }

    def download_file(self, drive_id: str, file_id: str) -> Optional[BinaryIO]:
        conn = _get_conn(self.db_path)
        try:
            cur = conn.execute(
                "SELECT local_path FROM mock_files WHERE id = ? AND drive_id = ?",
                (file_id, drive_id),
            )
            row = cur.fetchone()
            if not row:
                return None
            path = Path(row["local_path"])
            if not path.exists():
                return None
            return open(path, "rb")
        finally:
            conn.close()

    def list_files(self, drive_id: str) -> List[Dict[str, Any]]:
        conn = _get_conn(self.db_path)
        try:
            cur = conn.execute(
                "SELECT * FROM mock_files WHERE drive_id = ? ORDER BY created_at DESC",
                (drive_id,),
            )
            rows = cur.fetchall()
            return [
                {
                    "id": r["id"],
                    "name": r["name"],
                    "size": r["size"],
                    "mimeType": r["mime_type"],
                    "createdDateTime": r["created_at"],
                    "webUrl": f"file://{Path(r['local_path']).resolve()}",
                }
                for r in rows
            ]
        finally:
            conn.close()

    def delete_file(self, drive_id: str, file_id: str) -> bool:
        conn = _get_conn(self.db_path)
        try:
            cur = conn.execute(
                "SELECT local_path FROM mock_files WHERE id = ? AND drive_id = ?",
                (file_id, drive_id),
            )
            row = cur.fetchone()
            if row:
                path = Path(row["local_path"])
                if path.exists():
                    path.unlink()
            cur = conn.execute(
                "DELETE FROM mock_files WHERE id = ? AND drive_id = ?",
                (file_id, drive_id),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()
