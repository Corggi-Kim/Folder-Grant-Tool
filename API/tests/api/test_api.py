from pathlib import Path

from fastapi.testclient import TestClient

from folder_grant_api.main import create_app
from folder_grant_api.settings import Settings, get_settings


def build_client(tmp_path: Path) -> TestClient:
    settings = Settings(database_url=f"sqlite:///{tmp_path / 'test.db'}")
    app = create_app()
    app.dependency_overrides[get_settings] = lambda: settings
    return TestClient(app)


def valid_payload() -> dict:
    return {
        "request_id": "TEST-001",
        "source": "swagger",
        "operation": "grant",
        "project_status": "progress",
        "employee_id": "12345",
        "project_code": "26012",
        "level2": "Study",
        "level3": "STAT",
        "role": "Trial STAT/SP",
        "requested_by": "tester",
    }


def test_health_and_ready(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        assert client.get("/health").json()["status"] == "ok"
        assert client.get("/ready").json() == {"status": "ready", "database": "ok"}


def test_preview_returns_operation_plan(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        response = client.post("/api/v1/access-jobs/preview", json=valid_payload())

    assert response.status_code == 200
    body = response.json()
    assert body["valid"] is True
    assert body["target_group"] == "LSK 26-012"
    assert body["steps"][0]["type"] == "add_ad_group_member"


def test_preview_returns_structured_domain_error(tmp_path: Path) -> None:
    payload = valid_payload()
    payload["role"] = "Not A Role"
    with build_client(tmp_path) as client:
        response = client.post("/api/v1/access-jobs/preview", json=payload)

    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "INVALID_ROLE"
