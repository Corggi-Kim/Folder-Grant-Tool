from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from folder_grant_api.infrastructure.models import JobStepRecord
else:
    JobStepRecord = Any


class UnsupportedStepError(ValueError):
    pass


def quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def build_step_command(step: JobStepRecord) -> str:
    details = step.details or {}
    if step.step_type == "add_ad_group_member":
        return (
            f"Add-ADGroupMember -Identity {quote(step.target)} "
            f"-Members {quote(details['employee_id'])} -ErrorAction Stop;"
        )
    if step.step_type == "remove_ad_group_member":
        return (
            f"Remove-ADGroupMember -Identity {quote(step.target)} "
            f"-Members {quote(details['employee_id'])} -Confirm:$false -ErrorAction Stop;"
        )
    if step.step_type == "grant_closed_acl":
        return _closed_path_command(step, grant=True)
    if step.step_type == "revoke_closed_acl":
        return _closed_path_command(step, grant=False)
    if step.step_type == "grant_folder_acl":
        return _grant_acl_command(step)
    if step.step_type == "revoke_folder_acl":
        return (
            f"icacls {quote(step.target)} /remove {quote(details['principal'])} /T /C;"
            "if ($LASTEXITCODE -ne 0) { throw ('icacls remove failed: ' + $LASTEXITCODE) }"
        )
    raise UnsupportedStepError(f"지원하지 않는 작업 단계입니다: {step.step_type}")


def _icacls_grant(path_expression: str, principal: str, permission: str, recursive: bool = False) -> str:
    recursive_flags = " /T /C" if recursive else ""
    grant_value = quote(f"{principal}:{permission}")
    return (
        f"icacls {path_expression} /grant {grant_value}{recursive_flags};"
        "if ($LASTEXITCODE -ne 0) { throw ('icacls grant failed: ' + $LASTEXITCODE) };"
    )


def _grant_acl_command(step: JobStepRecord) -> str:
    details = step.details or {}
    principal = details["principal"]
    base_grant = _icacls_grant(
        quote(step.target),
        principal,
        "(CI)(OI)RXM",
        recursive=bool(details.get("recursive")),
    )
    legacy_numbers = details.get("legacy_subfolder_numbers") or []
    if legacy_numbers:
        commands = [base_grant]
        for number in legacy_numbers:
            variable = f"$folder{int(number)}"
            commands.append(
                f"{variable}=Get-ChildItem -LiteralPath {quote(step.target)} -Directory | "
                f"Where-Object {{ $_.Name -like {quote(f'{int(number)}.*')} }} | Select-Object -First 1;"
                f"if (-not {variable}) {{ throw {quote(f'필수 하위 폴더 {int(number)}.* 없음')} }};"
                + _icacls_grant(f"{variable}.FullName", principal, "(CI)(OI)RXM")
            )
        return "".join(commands)

    excluded = details.get("excluded_subfolders")
    if excluded is not None:
        excluded_array = "@(" + ",".join(quote(str(item)) for item in excluded) + ")"
        child_grant = _icacls_grant("$_.FullName", principal, "(CI)(OI)RXM")
        return (
            base_grant
            + f"$excluded={excluded_array};"
            + f"Get-ChildItem -LiteralPath {quote(step.target)} -Directory | "
            + "Where-Object { $excluded -notcontains $_.Name } | ForEach-Object {"
            + child_grant
            + "}"
        )
    return base_grant


def _closed_path_command(step: JobStepRecord, *, grant: bool) -> str:
    details = step.details or {}
    candidates = details.get("candidate_paths") or [step.target]
    candidates_array = "@(" + ",".join(quote(str(path)) for path in candidates) + ")"
    prefix = (
        f"$candidates={candidates_array};"
        "$target=$candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Container } | Select-Object -First 1;"
        "if (-not $target) { throw '종료 프로젝트 폴더를 찾을 수 없습니다.' };"
    )
    if grant:
        return prefix + _icacls_grant("$target", details["principal"], "(CI)(OI)RX")
    return (
        prefix
        + f"icacls $target /remove {quote(details['principal'])} /T /C;"
        + "if ($LASTEXITCODE -ne 0) { throw ('icacls remove failed: ' + $LASTEXITCODE) }"
    )
