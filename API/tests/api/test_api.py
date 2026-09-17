from pathlib import Path

from fastapi.testclient import TestClient

from folder_grant_api.main import create_app
from folder_grant_api.settings import Settings, get_settings


def build_client(tmp_path: Path) -> TestClient:
    settings = Settings(database_url=f"sqlite:///{tmp_path / 'test.db'}")
    app = create_app(settings)
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


def test_register_job_runs_mock_worker_and_can_be_queried(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        created = client.post("/api/v1/access-jobs", json=valid_payload())
        assert created.status_code == 202
        job_id = created.json()["job_id"]

        fetched = client.get(f"/api/v1/access-jobs/{job_id}")

    assert fetched.status_code == 200
    body = fetched.json()
    assert body["status"] == "simulated"
    assert body["executor_mode"] == "mock"
    assert body["result"]["changed_external_system"] is False
    assert all(step["status"] == "simulated" for step in body["steps"])


def test_duplicate_request_id_returns_original_job(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        first = client.post("/api/v1/access-jobs", json=valid_payload())
        second = client.post("/api/v1/access-jobs", json=valid_payload())

    assert first.status_code == 202
    assert second.status_code == 200
    assert first.json()["job_id"] == second.json()["job_id"]
    assert second.json()["replayed"] is True


def test_job_can_be_queried_by_request_id(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        created = client.post("/api/v1/access-jobs", json=valid_payload())
        fetched = client.get("/api/v1/access-jobs/by-request/TEST-001")

    assert fetched.status_code == 200
    assert fetched.json()["job_id"] == created.json()["job_id"]


def test_completed_mock_job_cannot_be_cancelled(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        created = client.post("/api/v1/access-jobs", json=valid_payload())
        response = client.post(
            f"/api/v1/access-jobs/{created.json()['job_id']}/cancel",
            json={"requested_by": "tester"},
        )

    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "JOB_NOT_CANCELLABLE"
