from dataclasses import dataclass, field

from .paths import format_group_name
from .project_codes import canonical_project_code, project_segment_for_folder


@dataclass(frozen=True)
class ProjectPlanStep:
    type: str
    description: str
    target: str
    metadata: dict = field(default_factory=dict)


@dataclass(frozen=True)
class ProjectCreationPlan:
    canonical_project_code: str
    group_name: str
    root_path: str
    study_all_path: str
    template_path: str
    steps: list[ProjectPlanStep]
    warnings: list[str] = field(default_factory=list)


def build_project_creation_plan(
    *,
    project_code: str,
    project_name: str,
    share_root: str,
    template_root: str,
    group_ou_path: str,
) -> ProjectCreationPlan:
    canonical_code = canonical_project_code(project_code)
    group_name = format_group_name(project_code, "Study")
    segment = project_segment_for_folder(project_code)
    root_path = share_root.format(proj_seg=segment, lv2="", lv3="").rstrip("\\")
    study_all_path = root_path + r"\Study\All"
    steps = [
        ProjectPlanStep(
            type="ensure_ad_group",
            description="프로젝트 AD 보안 그룹 확인 또는 생성",
            target=group_name,
            metadata={
                "project_name": project_name.strip(),
                "group_ou_path": group_ou_path,
            },
        ),
        ProjectPlanStep(
            type="ensure_project_folder",
            description="템플릿을 이용한 프로젝트 폴더 확인 또는 생성",
            target=root_path,
            metadata={"template_path": template_root},
        ),
        ProjectPlanStep(
            type="cleanup_unknown_acl",
            description="프로젝트 기본 경로의 해석 불가능 ACL 정리",
            target=root_path,
            metadata={"paths": [root_path, study_all_path]},
        ),
        ProjectPlanStep(
            type="grant_project_root_acl",
            description="프로젝트 루트에 그룹 읽기/실행 권한 보장",
            target=root_path,
            metadata={"principal": group_name, "permission": "read_execute"},
        ),
        ProjectPlanStep(
            type="grant_study_all_acl",
            description="Study/All에 그룹 수정 권한 보장",
            target=study_all_path,
            metadata={"principal": group_name, "permission": "modify"},
        ),
        ProjectPlanStep(
            type="verify_project_creation",
            description="AD 그룹·프로젝트 폴더·기본 ACL 검증",
            target=root_path,
            metadata={"group_name": group_name, "study_all_path": study_all_path},
        ),
    ]
    return ProjectCreationPlan(
        canonical_project_code=canonical_code,
        group_name=group_name,
        root_path=root_path,
        study_all_path=study_all_path,
        template_path=template_root,
        steps=steps,
    )
