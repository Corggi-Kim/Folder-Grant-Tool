from dataclasses import dataclass, field
from enum import StrEnum

from .constants import (
    DOMAIN_EMAIL_SUFFIX,
    ISOLATED_ROLE_MAP,
    ISOLATED_STAT_IDMC_BLOCKED_MAP,
    LEGACY_STUDY_MAP,
    ROLE_MAP,
)
from .paths import build_active_path, build_closed_candidates, format_group_name, normalize_level2
from .project_codes import canonical_project_code, is_new_template, is_stat_idmc_new_policy


class Operation(StrEnum):
    GRANT = "grant"
    REVOKE = "revoke"


class ProjectStatus(StrEnum):
    PROGRESS = "progress"
    CLOSED = "closed"


class PlanValidationError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class PlanStep:
    type: str
    description: str
    target: str
    permission: str | None = None
    metadata: dict = field(default_factory=dict)


@dataclass(frozen=True)
class OperationPlan:
    canonical_project_code: str
    target_group: str | None
    target_paths: list[str]
    steps: list[PlanStep]
    warnings: list[str] = field(default_factory=list)


def _validate_common(employee_id: str, project_code: str, level3: str | None) -> None:
    if not employee_id or not employee_id.strip():
        raise PlanValidationError("INVALID_EMPLOYEE_ID", "대상자 사번이 필요합니다.")
    if not project_code or not project_code.strip():
        raise PlanValidationError("INVALID_PROJECT_CODE", "프로젝트 코드가 필요합니다.")
    if level3 is not None and any(char in level3 for char in ("\\", "/")):
        raise PlanValidationError("INVALID_LEVEL3", "Level3에는 경로 구분자를 사용할 수 없습니다.")


def build_operation_plan(
    *,
    operation: Operation,
    project_status: ProjectStatus,
    employee_id: str,
    project_code: str,
    level2: str | None,
    level3: str | None,
    role: str | None,
    share_root: str,
    closed_root: str,
    archive_root: str,
) -> OperationPlan:
    _validate_common(employee_id, project_code, level3)
    employee_id = employee_id.strip()
    role = (role or "").strip()
    if project_status == ProjectStatus.CLOSED:
        candidates = build_closed_candidates(project_code, closed_root, archive_root)
        closed_code = candidates[0].rsplit("\\", 1)[-1].split("A", 1)[0]
        canonical_code = canonical_project_code(closed_code)
        step_type = "grant_closed_acl" if operation == Operation.GRANT else "revoke_closed_acl"
        permission = "read_execute" if operation == Operation.GRANT else None
        return OperationPlan(
            canonical_project_code=canonical_code,
            target_group=None,
            target_paths=candidates,
            steps=[PlanStep(
                type=step_type,
                description="종료 프로젝트 폴더 권한 부여" if operation == Operation.GRANT else "종료 프로젝트 폴더 권한 제거",
                target=candidates[0],
                permission=permission,
                metadata={"candidate_paths": candidates, "principal": f"{employee_id}@{DOMAIN_EMAIL_SUFFIX}"},
            )],
            warnings=["실행 시 존재하는 종료/Archive 경로를 선택합니다."],
        )

    canonical_code = canonical_project_code(project_code)
    if not level2 or not level3:
        raise PlanValidationError("MISSING_FOLDER_LEVEL", "진행 프로젝트에는 Level2와 Level3가 필요합니다.")

    level2_normalized = normalize_level2(level2)
    level3_normalized = level3.strip().upper()
    group_name = format_group_name(project_code, level2_normalized)
    target_path = build_active_path(project_code, level2_normalized, level3_normalized, share_root)
    membership_type = "add_ad_group_member" if operation == Operation.GRANT else "remove_ad_group_member"
    steps = [PlanStep(
        type=membership_type,
        description="프로젝트 AD 그룹에 대상자 추가" if operation == Operation.GRANT else "프로젝트 AD 그룹에서 대상자 제거",
        target=group_name,
        metadata={"employee_id": employee_id},
    )]

    if level3_normalized == "ETC":
        return OperationPlan(canonical_code, group_name, [], steps)

    if operation == Operation.REVOKE:
        steps.append(PlanStep(
            type="revoke_folder_acl",
            description="대상 폴더와 하위 항목의 사용자 ACL 제거",
            target=target_path,
            metadata={"recursive": True, "principal": f"{employee_id}@{DOMAIN_EMAIL_SUFFIX}"},
        ))
        return OperationPlan(canonical_code, group_name, [target_path], steps)

    steps.extend(_grant_acl_steps(
        project_code=project_code,
        level2=level2_normalized,
        level3=level3_normalized,
        role=role,
        employee_id=employee_id,
        target_path=target_path,
    ))
    return OperationPlan(canonical_code, group_name, [target_path], steps)


def _grant_acl_steps(
    *, project_code: str, level2: str, level3: str, role: str,
    employee_id: str, target_path: str,
) -> list[PlanStep]:
    principal = f"{employee_id}@{DOMAIN_EMAIL_SUFFIX}"

    def grant(path: str, description: str, **metadata) -> PlanStep:
        return PlanStep(
            type="grant_folder_acl",
            description=description,
            target=path,
            permission="modify",
            metadata={"principal": principal, **metadata},
        )

    if not role:
        return [grant(target_path, "대상 폴더 수정 권한 부여")]

    new_template = is_new_template(project_code)
    if level2 == "Study":
        if role not in ROLE_MAP:
            raise PlanValidationError("INVALID_ROLE", f"Study에서 허용되지 않는 Role입니다: {role}")
        if new_template:
            return [grant(target_path, "Study 폴더 수정 권한 부여")] + [
                grant(f"{target_path}\\{subfolder}", f"Role 대상 하위 폴더 권한 부여: {subfolder}")
                for subfolder in ROLE_MAP[role]
            ]
        return [grant(
            target_path,
            "레거시 Study 폴더 권한 부여",
            legacy_subfolder_numbers=LEGACY_STUDY_MAP[role],
        )]

    if level3 == "STAT_IDMC":
        if is_stat_idmc_new_policy(project_code):
            if role not in ISOLATED_STAT_IDMC_BLOCKED_MAP:
                raise PlanValidationError("INVALID_ROLE", f"STAT_IDMC에서 허용되지 않는 Role입니다: {role}")
            return [grant(
                target_path,
                "STAT_IDMC 허용 폴더 권한 부여",
                excluded_subfolders=ISOLATED_STAT_IDMC_BLOCKED_MAP[role],
            )]
        return [grant(target_path, "레거시 STAT_IDMC 재귀 권한 부여", recursive=True)]

    if new_template:
        if role not in ISOLATED_ROLE_MAP:
            raise PlanValidationError("INVALID_ROLE", f"Isolated에서 허용되지 않는 Role입니다: {role}")
        return [grant(target_path, "Isolated 폴더 수정 권한 부여")] + [
            grant(f"{target_path}\\{subfolder}", f"Role 대상 하위 폴더 권한 부여: {subfolder}")
            for subfolder in ISOLATED_ROLE_MAP[role]
        ]

    if role != "Randomization Statistician":
        raise PlanValidationError(
            "INVALID_ROLE",
            "레거시 Isolated는 Randomization Statistician Role만 허용합니다.",
        )
    return [grant(target_path, "레거시 Isolated 폴더 권한 부여")]
