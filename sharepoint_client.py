"""
sharepoint_client.py
Encapsulates Microsoft Graph API OAuth2 Client Credentials authentication,
SharePoint List CRUD, and Document Library file upload/download.
Auto-falls back to mock mode when Azure credentials are missing or invalid.
"""

import logging
import os
from io import BytesIO
from typing import Any, BinaryIO, Dict, List, Optional

import requests
from msal import ConfidentialClientApplication

from mock_sharepoint import MockSharePointClient

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"
SCOPE = ["https://graph.microsoft.com/.default"]


class SharePointClientError(Exception):
    """Custom exception for SharePoint client errors."""
    pass


class SharePointClient:
    """
    SharePoint Graph API client with automatic mock fallback.
    """

    def __init__(
        self,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        site_id: Optional[str] = None,
        list_id: Optional[str] = None,
        drive_id: Optional[str] = None,
        force_mock: bool = False,
    ):
        self.tenant_id = tenant_id or os.getenv("TENANT_ID")
        self.client_id = client_id or os.getenv("CLIENT_ID")
        self.client_secret = client_secret or os.getenv("CLIENT_SECRET")
        self.site_id = site_id or os.getenv("SITE_ID")
        self.list_id = list_id or os.getenv("LIST_ID")
        self.drive_id = drive_id or os.getenv("DRIVE_ID")

        self._access_token: Optional[str] = None
        self._mock: Optional[MockSharePointClient] = None

        # Determine whether to use mock mode
        if force_mock or not self._has_real_credentials():
            logger.info("SharePointClient initializing in MOCK mode.")
            self._mock = MockSharePointClient()
        else:
            logger.info("SharePointClient initializing in LIVE mode.")

    # -----------------------------------------------------------------------
    # Internal helpers
    # -----------------------------------------------------------------------

    def _has_real_credentials(self) -> bool:
        return all(
            [
                self.tenant_id,
                self.client_id,
                self.client_secret,
                self.site_id,
            ]
        )

    def _get_token(self) -> str:
        if self._access_token:
            return self._access_token
        if not self.tenant_id or not self.client_id or not self.client_secret:
            raise SharePointClientError("Azure credentials are incomplete.")

        authority = AUTHORITY_TEMPLATE.format(tenant_id=self.tenant_id)
        app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=authority,
        )
        result = app.acquire_token_for_client(scopes=SCOPE)
        if "access_token" not in result:
            error_desc = result.get("error_description", "Unknown error")
            raise SharePointClientError(f"Authentication failed: {error_desc}")
        self._access_token = result["access_token"]
        return self._access_token

    def _graph_request(
        self,
        method: str,
        endpoint: str,
        json_payload: Optional[Dict[str, Any]] = None,
        data: Optional[Any] = None,
        headers: Optional[Dict[str, str]] = None,
        stream: bool = False,
    ) -> requests.Response:
        url = f"{GRAPH_API_BASE}{endpoint}"
        default_headers = {
            "Authorization": f"Bearer {self._get_token()}",
            "Accept": "application/json",
        }
        if headers:
            default_headers.update(headers)

        response = requests.request(
            method=method,
            url=url,
            headers=default_headers,
            json=json_payload,
            data=data,
            stream=stream,
            timeout=60,
        )

        # Basic error handling
        if not response.ok:
            try:
                detail = response.json()
            except Exception:
                detail = response.text
            raise SharePointClientError(
                f"Graph API error {response.status_code}: {detail}"
            )
        return response

    def _is_mock(self) -> bool:
        return self._mock is not None

    # -----------------------------------------------------------------------
    # SharePoint List operations (Tickets)
    # -----------------------------------------------------------------------

    def create_ticket(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        if self._is_mock():
            return self._mock.create_ticket(payload)  # type: ignore[union-attr]

        if not self.list_id:
            raise SharePointClientError("LIST_ID is not configured.")

        endpoint = f"/sites/{self.site_id}/lists/{self.list_id}/items"
        # Graph list item creation requires fields wrapper
        graph_payload = {"fields": payload}
        resp = self._graph_request("POST", endpoint, json_payload=graph_payload)
        return resp.json()

    def list_tickets(
        self,
        status: Optional[str] = None,
        category: Optional[str] = None,
        reporter_email: Optional[str] = None,
        limit: int = 100,
        offset: int = 0,
    ) -> List[Dict[str, Any]]:
        if self._is_mock():
            return self._mock.list_tickets(  # type: ignore[union-attr]
                status=status,
                category=category,
                reporter_email=reporter_email,
                limit=limit,
                offset=offset,
            )

        if not self.list_id:
            raise SharePointClientError("LIST_ID is not configured.")

        # Build OData filter
        filters: List[str] = []
        if status:
            filters.append(f"fields/status eq '{status}'")
        if category:
            filters.append(f"fields/category eq '{category}'")
        if reporter_email:
            filters.append(f"fields/reporterEmail eq '{reporter_email}'")

        filter_str = " and ".join(filters) if filters else None
        endpoint = f"/sites/{self.site_id}/lists/{self.list_id}/items?expand=fields"
        if filter_str:
            endpoint += f"&$filter={filter_str}"
        endpoint += f"&$top={limit}"
        if offset > 0:
            endpoint += f"&$skip={offset}"

        resp = self._graph_request("GET", endpoint)
        data = resp.json()
        return data.get("value", [])

    def get_ticket(self, item_id: str) -> Optional[Dict[str, Any]]:
        if self._is_mock():
            return self._mock.get_ticket(item_id)  # type: ignore[union-attr]

        if not self.list_id:
            raise SharePointClientError("LIST_ID is not configured.")

        endpoint = f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}?expand=fields"
        resp = self._graph_request("GET", endpoint)
        return resp.json()

    def update_ticket(self, item_id: str, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        if self._is_mock():
            return self._mock.update_ticket(item_id, payload)  # type: ignore[union-attr]

        if not self.list_id:
            raise SharePointClientError("LIST_ID is not configured.")

        endpoint = f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}"
        graph_payload = {"fields": payload}
        resp = self._graph_request("PATCH", endpoint, json_payload=graph_payload)
        return resp.json()

    def delete_ticket(self, item_id: str) -> bool:
        if self._is_mock():
            return self._mock.delete_ticket(item_id)  # type: ignore[union-attr]

        if not self.list_id:
            raise SharePointClientError("LIST_ID is not configured.")

        endpoint = f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}"
        resp = self._graph_request("DELETE", endpoint)
        return resp.status_code in (204, 200)

    # -----------------------------------------------------------------------
    # Document Library operations
    # -----------------------------------------------------------------------

    def upload_file(
        self,
        file_name: str,
        content: BinaryIO,
        mime_type: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        if self._is_mock():
            target_drive = drive_id or self.drive_id or "mock_drive"
            return self._mock.upload_file(  # type: ignore[union-attr]
                drive_id=target_drive,
                file_name=file_name,
                content=content,
                mime_type=mime_type,
            )

        target_drive = drive_id or self.drive_id
        if not target_drive:
            raise SharePointClientError("DRIVE_ID is not configured.")

        endpoint = f"/drives/{target_drive}/root:/{file_name}:/content"
        file_bytes = content.read()
        headers = {"Content-Type": mime_type or "application/octet-stream"}
        resp = self._graph_request("PUT", endpoint, data=file_bytes, headers=headers)
        return resp.json()

    def download_file(
        self, file_id: str, drive_id: Optional[str] = None
    ) -> Optional[BinaryIO]:
        if self._is_mock():
            target_drive = drive_id or self.drive_id or "mock_drive"
            return self._mock.download_file(target_drive, file_id)  # type: ignore[union-attr]

        target_drive = drive_id or self.drive_id
        if not target_drive:
            raise SharePointClientError("DRIVE_ID is not configured.")

        endpoint = f"/drives/{target_drive}/items/{file_id}/content"
        resp = self._graph_request("GET", endpoint, stream=True)
        return BytesIO(resp.content)

    def list_files(self, drive_id: Optional[str] = None) -> List[Dict[str, Any]]:
        if self._is_mock():
            target_drive = drive_id or self.drive_id or "mock_drive"
            return self._mock.list_files(target_drive)  # type: ignore[union-attr]

        target_drive = drive_id or self.drive_id
        if not target_drive:
            raise SharePointClientError("DRIVE_ID is not configured.")

        endpoint = f"/drives/{target_drive}/root/children"
        resp = self._graph_request("GET", endpoint)
        data = resp.json()
        return data.get("value", [])

    def delete_file(self, file_id: str, drive_id: Optional[str] = None) -> bool:
        if self._is_mock():
            target_drive = drive_id or self.drive_id or "mock_drive"
            return self._mock.delete_file(target_drive, file_id)  # type: ignore[union-attr]

        target_drive = drive_id or self.drive_id
        if not target_drive:
            raise SharePointClientError("DRIVE_ID is not configured.")

        endpoint = f"/drives/{target_drive}/items/{file_id}"
        resp = self._graph_request("DELETE", endpoint)
        return resp.status_code in (204, 200)
