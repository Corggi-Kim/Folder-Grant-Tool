from folder_grant_api.domain.project_plan import build_project_creation_plan


def test_project_plan_contains_group_folder_acl_and_verification_steps() -> None:
    plan = build_project_creation_plan(
        project_code="98765",
        project_name="API Test Project",
        share_root=r"\\LSK_S010\Study Folder\{proj_seg}\{lv2}\{lv3}",
        template_root=r"\\LSK_S010\Study folder\_Template",
        group_ou_path="OU=Groups,DC=example,DC=com",
    )

    assert plan.canonical_project_code == "98-765"
    assert plan.group_name == "LSK 98-765"
    assert plan.root_path == r"\\LSK_S010\Study Folder\98765"
    assert plan.study_all_path == r"\\LSK_S010\Study Folder\98765\Study\All"
    assert [step.type for step in plan.steps] == [
        "ensure_ad_group",
        "ensure_project_folder",
        "cleanup_unknown_acl",
        "grant_project_root_acl",
        "grant_study_all_acl",
        "verify_project_creation",
    ]


def test_project_plan_preserves_numeric_suffix() -> None:
    plan = build_project_creation_plan(
        project_code="98765-2",
        project_name="Amendment",
        share_root=r"\\server\share\{proj_seg}\{lv2}\{lv3}",
        template_root=r"\\server\template",
        group_ou_path="OU=Groups,DC=example,DC=com",
    )
    assert plan.group_name == "LSK 98-765-2"
    assert plan.root_path == r"\\server\share\98765A2"
