"""Validation must not depend on how much request data fits in Windows argv.

The process boundary models CreateProcess' command-line limit and supplies a
read-only resource-query response. It never launches PowerShell or mutates AD.
"""
import json
import subprocess
from unittest.mock import patch

import pytest
from sqlalchemy import func, select

from folder_grant_api.api.routes.preview import create_plan
from folder_grant_api.api.routes.project_jobs import create_project_plan
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.schemas.access_job import AccessJobRequest
from folder_grant_api.schemas.project_job import ProjectJobRequest
from folder_grant_api.services.validation_service import access_probes, project_probes
from test_api import build_client, valid_payload, valid_project_payload


LONG_SHARE = r"\\파일서버\연구 자료" + "\\공유경로 확장 시험 자료" * 12


@pytest.mark.parametrize(
    "level2,level3,role,operation,project_code,extended_root",
    [
        pytest.param("Study", "STAT", "Manager", "grant", "26066", False, id="reported-stat-manager"),
        pytest.param("Study", "STAT", "Manager", "revoke", "26066", False, id="stat-manager-revoke"),
        pytest.param("Study", "STAT", "Trial STAT/SP", "grant", "26066", False, id="study-trial-sp"),
        pytest.param("Study", "STAT", "Verification SP", "grant", "26066", False, id="study-verification"),
        pytest.param("Study", "STAT", "SDTM", "grant", "26066", False, id="study-sdtm"),
        pytest.param("Study", "STAT", "Manager", "grant", "1987", False, id="legacy-study-manager"),
        pytest.param("Isolated", "STAT", "Unblind Reviewer", "grant", "26066", False, id="isolated-reviewer"),
        pytest.param("Isolated", "STAT", "Randomization Statistician", "grant", "26066", False, id="isolated-randomization"),
        pytest.param("Isolated", "STAT_IDMC", "Manager", "grant", "26066", False, id="isolated-idmc-manager"),
        pytest.param("Isolated", "STAT_IDMC", "Trial STAT/SP", "grant", "26066", False, id="isolated-idmc-trial"),
        pytest.param("Study", "DM_IWRS", None, "grant", "26066", True, id="custom-folder-long-share"),
        pytest.param("Isolated", "한글 폴더", None, "grant", "26066", True, id="unicode-folder-long-share"),
    ],
)
def test_access_validation_survives_windows_command_line_limit(
    tmp_path, level2, level3, role, operation, project_code, extended_root,
):
    payload = {
        **valid_payload(),
        "employee_id": "20260101",
        "project_code": project_code,
        "level2": level2,
        "level3": level3,
        "role": role,
        "operation": operation,
    }
    with build_client(tmp_path, executor_mode="powershell") as client:
        settings = client.app.state.settings
        settings.powershell_path = r"C:\Program Files\PowerShell\7\pwsh.exe"
        if extended_root:
            settings.share_root = LONG_SHARE + r"\{proj_seg}\{lv2}\{lv3}"
        request = AccessJobRequest(**payload)
        probes = access_probes(request, create_plan(request, settings), settings)
        _assert_validation_succeeds_without_job(client, "/api/v1/access-jobs/validate", payload, probes)


@pytest.mark.parametrize("operation", ["grant", "revoke"])
def test_closed_validation_with_long_candidate_paths_survives_windows_limit(tmp_path, operation):
    payload = {
        **valid_payload(), "operation": operation, "project_status": "closed",
        "project_code": "26066", "level2": None, "level3": None, "role": None,
    }
    with build_client(tmp_path, executor_mode="powershell") as client:
        settings = client.app.state.settings
        settings.powershell_path = r"C:\Program Files\PowerShell\7\pwsh.exe"
        settings.closed_root = LONG_SHARE + r"\Study_Closed"
        settings.closed_archive_root = LONG_SHARE + r"\Study_Archive"
        request = AccessJobRequest(**payload)
        probes = access_probes(request, create_plan(request, settings), settings)
        _assert_validation_succeeds_without_job(client, "/api/v1/access-jobs/validate", payload, probes)


def test_project_validation_with_long_template_paths_survives_windows_limit(tmp_path):
    payload = {**valid_project_payload(), "project_code": "26066"}
    with build_client(tmp_path, executor_mode="powershell") as client:
        settings = client.app.state.settings
        settings.powershell_path = r"C:\Program Files\PowerShell\7\pwsh.exe"
        settings.share_root = LONG_SHARE + r"\{proj_seg}\{lv2}\{lv3}"
        settings.template_root = LONG_SHARE + r"\템플릿 자료"
        request = ProjectJobRequest(**payload)
        probes = project_probes(request, create_project_plan(request, settings), settings)
        _assert_validation_succeeds_without_job(client, "/api/v1/project-jobs/validate", payload, probes)


def _assert_validation_succeeds_without_job(client, endpoint, payload, probes):
    lengths = []
    checks = [
        {"id": p["id"], "label": p["label"], "target": p["target"],
         "status": "passed", "message": "조회 완료"}
        for p in probes
    ]
    query_response = json.dumps({"checks": checks, "checked_by": "DOMAIN\\service"}, ensure_ascii=False).encode("utf-8")

    def windows_process_boundary(argv, **kwargs):
        # Include the terminating NUL in CreateProcess' documented 32,767 limit.
        length = len(subprocess.list2cmdline(argv).encode("utf-16-le")) // 2
        lengths.append(length)
        if length >= 32767:
            error = FileNotFoundError(2, "The filename or extension is too long")
            error.winerror = 206
            raise error
        return subprocess.CompletedProcess(argv, 0, query_response, b"")

    with patch("folder_grant_api.infrastructure.powershell.subprocess.run", side_effect=windows_process_boundary):
        response = client.post(endpoint, json=payload)

    assert response.status_code == 200
    body = response.json()
    assert body["validation_status"] == "passed", body
    assert body["checked_by"] == "DOMAIN\\service"
    assert [(c["id"], c["target"]) for c in body["checks"]] == [(c["id"], c["target"]) for c in checks]
    assert lengths and all(length < 32767 for length in lengths)
    with client.app.state.session_factory() as db:
        assert db.scalar(select(func.count()).select_from(JobRecord)) == 0
