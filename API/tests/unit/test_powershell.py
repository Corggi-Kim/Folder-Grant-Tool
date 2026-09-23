import subprocess
from types import SimpleNamespace
from unittest.mock import patch

import pytest

from folder_grant_api.infrastructure.powershell import (
    PowerShellExecutionError,
    PowerShellRunner,
    encode_command,
)
from folder_grant_api.infrastructure.powershell_commands import build_step_command, quote


def step(step_type: str, target: str, **details):
    return SimpleNamespace(step_type=step_type, target=target, details=details)


def test_quote_escapes_single_quotes() -> None:
    assert quote("A'B") == "'A''B'"


def test_ad_group_command_uses_known_step_fields() -> None:
    command = build_step_command(step("add_ad_group_member", "LSK 26-012", employee_id="12345"))
    assert command == "Add-ADGroupMember -Identity 'LSK 26-012' -Members '12345' -ErrorAction Stop;"


def test_recursive_acl_removal_command() -> None:
    command = build_step_command(
        step("revoke_folder_acl", r"\\server\share", principal="12345@lskglobal.com", recursive=True)
    )
    assert "/remove '12345@lskglobal.com' /T /C" in command


def test_legacy_study_command_requires_and_grants_mapped_subfolders() -> None:
    command = build_step_command(
        step(
            "grant_folder_acl",
            r"\\server\share\STAT",
            principal="12345@lskglobal.com",
            permission="modify",
            legacy_subfolder_numbers=[3, 5],
        )
    )
    assert "3.*" in command
    assert "5.*" in command
    assert "필수 하위 폴더" in command


def test_closed_acl_command_checks_all_candidate_paths() -> None:
    command = build_step_command(
        step(
            "grant_closed_acl",
            r"\\server\closed\26012",
            principal="12345@lskglobal.com",
            candidate_paths=[r"\\server\closed\26012", r"\\server\archive\26012"],
        )
    )
    assert "Test-Path" in command
    assert r"\\server\archive\26012" in command


def test_encoded_command_uses_powershell_utf16_encoding() -> None:
    assert encode_command("Write-Output '테스트'")


def test_runner_does_not_use_shell_and_decodes_utf8() -> None:
    completed = subprocess.CompletedProcess([], 0, "정상".encode(), b"")
    with patch("subprocess.run", return_value=completed) as run:
        result = PowerShellRunner("powershell.exe", 30).run("Write-Output '정상'")

    assert result.stdout == "정상"
    assert run.call_args.kwargs["timeout"] == 30
    assert "shell" not in run.call_args.kwargs


def test_runner_maps_timeout_to_stable_error_code() -> None:
    with patch("subprocess.run", side_effect=subprocess.TimeoutExpired("powershell.exe", 1)):
        with pytest.raises(PowerShellExecutionError) as exc_info:
            PowerShellRunner("powershell.exe", 1).run("Start-Sleep 10")
    assert exc_info.value.code == "POWERSHELL_TIMEOUT"
