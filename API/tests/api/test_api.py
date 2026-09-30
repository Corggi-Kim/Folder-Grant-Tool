from pathlib import Path
from unittest.mock import patch

import pytest
from fastapi.testclient import TestClient

from folder_grant_api.main import create_app
from folder_grant_api.infrastructure.powershell import PowerShellExecutionError, PowerShellResult
from folder_grant_api.settings import Settings, get_settings


def build_client(tmp_path: Path, *, executor_mode: str = "mock") -> TestClient:
    settings = Settings(
        database_url=f"sqlite:///{tmp_path / 'test.db'}",
        api_key="test-api-key",
        executor_mode=executor_mode,
    )
    app = create_app(settings)
    app.dependency_overrides[get_settings] = lambda: settings
    client = TestClient(app)
    client.headers["X-API-Key"] = "test-api-key"
    return client


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


def valid_project_payload() -> dict:
    return {
        "request_id": "PROJECT-TEST-001",
        "source": "swagger",
        "project_code": "98765",
        "project_name": "API Test Project",
        "requested_by": "tester",
    }


def test_health_and_ready(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        assert client.get("/health").json()["status"] == "ok"
        assert client.get("/ready").json() == {"status": "ready", "database": "ok"}


def test_protected_endpoint_rejects_invalid_api_key(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        response = client.get("/ready", headers={"X-API-Key": "wrong"})

    assert response.status_code == 401
    assert response.json()["detail"]["code"] == "INVALID_API_KEY"


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


def test_duplicate_request_id_with_different_payload_is_rejected(tmp_path: Path) -> None:
    changed = valid_payload()
    changed["employee_id"] = "54321"
    with build_client(tmp_path) as client:
        client.post("/api/v1/access-jobs", json=valid_payload())
        response = client.post("/api/v1/access-jobs", json=changed)

    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "REQUEST_ID_CONFLICT"


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


def test_powershell_mode_records_success_without_running_real_process(tmp_path: Path) -> None:
    with patch(
        "folder_grant_api.workers.powershell_worker.PowerShellRunner.run",
        return_value=PowerShellResult(0, "OK", ""),
    ) as run:
        with build_client(tmp_path, executor_mode="powershell") as client:
            created = client.post("/api/v1/access-jobs", json=valid_payload())
            fetched = client.get(f"/api/v1/access-jobs/{created.json()['job_id']}")

    assert fetched.status_code == 200
    assert fetched.json()["status"] == "succeeded"
    assert fetched.json()["executor_mode"] == "powershell"
    assert fetched.json()["result"]["changed_external_system"] is True
    assert run.call_count == len(fetched.json()["steps"])


def test_powershell_failure_after_completed_step_is_partial(tmp_path: Path) -> None:
    with patch(
        "folder_grant_api.workers.powershell_worker.PowerShellRunner.run",
        side_effect=[PowerShellResult(0, "OK", ""), PowerShellExecutionError("POWERSHELL_FAILED", "test failure")],
    ):
        with build_client(tmp_path, executor_mode="powershell") as client:
            created = client.post("/api/v1/access-jobs", json=valid_payload())
            fetched = client.get(f"/api/v1/access-jobs/{created.json()['job_id']}")

    assert fetched.json()["status"] == "partially_succeeded"
    assert fetched.json()["steps"][0]["status"] == "succeeded"
    assert fetched.json()["steps"][1]["status"] == "failed"
    assert fetched.json()["steps"][2]["status"] == "skipped"


def test_powershell_mode_rejects_default_api_key() -> None:
    settings = Settings(executor_mode="powershell", api_key="change-me")
    with pytest.raises(ValueError, match="기본 API Key"):
        settings.validate_runtime_safety()


def test_project_preview_returns_expected_creation_plan(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        response = client.post("/api/v1/project-jobs/preview", json=valid_project_payload())

    assert response.status_code == 200
    body = response.json()
    assert body["group_name"] == "LSK 98-765"
    assert body["root_path"].endswith(r"Study Folder\98765")
    assert body["steps"][-1]["type"] == "verify_project_creation"


def test_project_job_runs_all_steps_in_mock_mode(tmp_path: Path) -> None:
    with build_client(tmp_path) as client:
        created = client.post("/api/v1/project-jobs", json=valid_project_payload())
        fetched = client.get(f"/api/v1/project-jobs/{created.json()['job_id']}")

    assert created.status_code == 202
    assert fetched.json()["status"] == "simulated"
    assert len(fetched.json()["steps"]) == 6
    assert all(step["status"] == "simulated" for step in fetched.json()["steps"])


def test_project_job_powershell_mode_records_success_without_real_process(tmp_path: Path) -> None:
    with patch(
        "folder_grant_api.workers.powershell_worker.PowerShellRunner.run",
        return_value=PowerShellResult(0, "OK", ""),
    ) as run:
        with build_client(tmp_path, executor_mode="powershell") as client:
            created = client.post("/api/v1/project-jobs", json=valid_project_payload())
            fetched = client.get(f"/api/v1/project-jobs/{created.json()['job_id']}")

    assert fetched.json()["status"] == "succeeded"
    assert fetched.json()["executor_mode"] == "powershell"
    assert run.call_count == 6


def test_project_request_id_conflicts_with_different_project(tmp_path: Path) -> None:
    changed = valid_project_payload()
    changed["project_code"] = "98766"
    with build_client(tmp_path) as client:
        client.post("/api/v1/project-jobs", json=valid_project_payload())
        response = client.post("/api/v1/project-jobs", json=changed)

    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "REQUEST_ID_CONFLICT"
