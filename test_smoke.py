"""
Quick smoke tests for SAS-CampusSync.
Run with: pytest test_smoke.py -v
"""

import os
import sys

# Ensure mock mode for tests
os.environ["MOCK_MODE"] = "true"

from fastapi.testclient import TestClient

# Import after setting env var
from main import app

client = TestClient(app)


def test_health_check():
    with TestClient(app) as c:
        response = c.get("/health")
        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "ok"
        assert "timestamp" in data


def test_create_ticket():
    with TestClient(app) as c:
        payload = {
            "title": "Test projector issue",
            "description": "No HDMI signal",
            "category": "hardware",
            "reporter_email": "student@campus.edu",
            "priority": "high",
        }
        response = c.post("/tickets", json=payload)
        assert response.status_code == 201
        data = response.json()
        assert data["title"] == payload["title"]
        assert data["status"] == "open"
        assert "id" in data
        return data["id"]


def test_list_tickets():
    with TestClient(app) as c:
        response = c.get("/tickets")
        assert response.status_code == 200
        data = response.json()
        assert isinstance(data, list)


def test_get_ticket():
    with TestClient(app) as c:
        payload = {
            "title": "Get test ticket",
            "description": "For get test",
            "category": "software",
            "reporter_email": "test@campus.edu",
            "priority": "medium",
        }
        r = c.post("/tickets", json=payload)
        ticket_id = r.json()["id"]
        response = c.get(f"/tickets/{ticket_id}")
        assert response.status_code == 200
        data = response.json()
        assert data["id"] == ticket_id


def test_update_ticket():
    with TestClient(app) as c:
        payload = {
            "title": "Update test ticket",
            "description": "For update test",
            "category": "network",
            "reporter_email": "test@campus.edu",
            "priority": "low",
        }
        r = c.post("/tickets", json=payload)
        ticket_id = r.json()["id"]
        response = c.patch(
            f"/tickets/{ticket_id}",
            json={"status": "resolved", "assigned_to": "tech@campus.edu"},
        )
        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "resolved"
        assert data["assigned_to"] == "tech@campus.edu"


def test_delete_ticket():
    with TestClient(app) as c:
        payload = {
            "title": "Delete test ticket",
            "description": "For delete test",
            "category": "general",
            "reporter_email": "test@campus.edu",
            "priority": "low",
        }
        r = c.post("/tickets", json=payload)
        ticket_id = r.json()["id"]
        response = c.delete(f"/tickets/{ticket_id}")
        assert response.status_code == 204
        response = c.get(f"/tickets/{ticket_id}")
        assert response.status_code == 404


def test_export_report():
    with TestClient(app) as c:
        # Create at least one ticket so report is non-empty
        c.post("/tickets", json={
            "title": "Export test ticket",
            "description": "For export test",
            "category": "general",
            "reporter_email": "test@campus.edu",
            "priority": "medium",
        })
        payload = {"file_name": "test_report.csv", "status_filter": None}
        response = c.post("/reports/export", json=payload)
        assert response.status_code == 200
        data = response.json()
        assert data["message"] == "Report exported successfully"
        assert data["file_name"] == "test_report.csv"
        assert "sharepoint_file" in data


if __name__ == "__main__":
    test_health_check()
    test_create_ticket()
    test_list_tickets()
    test_get_ticket()
    test_update_ticket()
    test_delete_ticket()
    test_export_report()
    print("All smoke tests passed!")
