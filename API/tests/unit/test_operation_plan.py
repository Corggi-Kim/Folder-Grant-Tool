import pytest

from folder_grant_api.domain.operation_plan import (
    Operation,
    PlanValidationError,
    ProjectStatus,
    build_operation_plan,
)

SHARE_ROOT = r"\\LSK_S010\Study Folder\{proj_seg}\{lv2}\{lv3}"
CLOSED_ROOT = r"\\192.168.1.95\Study_Closed"
ARCHIVE_ROOT = r"\\192.168.1.95\Study_Archive"


def make_plan(**overrides):
    values = {
        "operation": Operation.GRANT,
        "project_status": ProjectStatus.PROGRESS,
        "employee_id": "12345",
        "project_code": "26012",
        "level2": "Study",
        "level3": "STAT",
        "role": "Trial STAT/SP",
        "share_root": SHARE_ROOT,
        "closed_root": CLOSED_ROOT,
        "archive_root": ARCHIVE_ROOT,
    }
    values.update(overrides)
    return build_operation_plan(**values)


def test_new_study_grant_builds_group_and_role_paths() -> None:
    plan = make_plan()

    assert plan.canonical_project_code == "26-012"
    assert plan.target_group == "LSK 26-012"
    assert plan.target_paths == [r"\\LSK_S010\Study Folder\26012\Study\STAT"]
    assert [step.type for step in plan.steps] == ["add_ad_group_member"] + ["grant_folder_acl"] * 5
    assert plan.steps[-1].target.endswith(r"STAT\6.Validation")


def test_revoke_uses_recursive_acl_removal() -> None:
    plan = make_plan(operation=Operation.REVOKE)
    assert [step.type for step in plan.steps] == ["remove_ad_group_member", "revoke_folder_acl"]
    assert plan.steps[-1].metadata["recursive"] is True


def test_etc_only_changes_group_membership() -> None:
    plan = make_plan(level3="ETC")
    assert [step.type for step in plan.steps] == ["add_ad_group_member"]
    assert plan.target_paths == []


def test_closed_project_exposes_both_candidate_paths() -> None:
    plan = make_plan(
        project_status=ProjectStatus.CLOSED,
        level2=None,
        level3=None,
        role=None,
    )
    assert plan.target_group is None
    assert plan.target_paths == [
        r"\\192.168.1.95\Study_Closed\26012",
        r"\\192.168.1.95\Study_Archive\26012",
    ]
    assert plan.steps[0].permission == "read_execute"


def test_closed_project_preserves_existing_a_suffix() -> None:
    plan = make_plan(
        project_status=ProjectStatus.CLOSED,
        project_code="26012A2",
        level2=None,
        level3=None,
        role=None,
    )
    assert plan.canonical_project_code == "26-012"
    assert plan.target_paths[0].endswith(r"Study_Closed\26012A2")


def test_invalid_study_role_is_rejected() -> None:
    with pytest.raises(PlanValidationError) as exc_info:
        make_plan(role="Randomization Statistician")
    assert exc_info.value.code == "INVALID_ROLE"


def test_new_isolated_role_targets_expected_subfolder() -> None:
    plan = make_plan(
        level2="Isolated",
        level3="STAT",
        role="Blind Reviewer",
    )
    assert plan.target_group == "LSK 26-012 Isolated"
    assert plan.steps[-1].target.endswith(r"STAT\Reviewer")
