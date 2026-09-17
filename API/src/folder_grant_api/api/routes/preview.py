from fastapi import APIRouter, Depends, HTTPException, status

from folder_grant_api.domain.operation_plan import OperationPlan, PlanValidationError, build_operation_plan
from folder_grant_api.domain.project_codes import InvalidProjectCode
from folder_grant_api.schemas.access_job import AccessJobRequest, ErrorResponse, PreviewResponse
from folder_grant_api.settings import Settings, get_settings

router = APIRouter(prefix="/api/v1/access-jobs", tags=["access-jobs"])


def create_plan(payload: AccessJobRequest, settings: Settings | None = None) -> OperationPlan:
    settings = settings or get_settings()
    try:
        return build_operation_plan(
            operation=payload.operation,
            project_status=payload.project_status,
            employee_id=payload.employee_id,
            project_code=payload.project_code,
            level2=payload.level2,
            level3=payload.level3,
            role=payload.role,
            share_root=settings.share_root,
            closed_root=settings.closed_root,
            archive_root=settings.closed_archive_root,
        )
    except InvalidProjectCode as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={"code": "INVALID_PROJECT_CODE", "message": str(exc)},
        ) from exc
    except PlanValidationError as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={"code": exc.code, "message": str(exc)},
        ) from exc


@router.post(
    "/preview",
    response_model=PreviewResponse,
    responses={422: {"model": ErrorResponse}},
)
def preview_access_job(
    payload: AccessJobRequest,
    settings: Settings = Depends(get_settings),
) -> PreviewResponse:
    plan = create_plan(payload, settings)

    return PreviewResponse(
        request_id=payload.request_id,
        canonical_project_code=plan.canonical_project_code,
        target_group=plan.target_group,
        target_paths=plan.target_paths,
        steps=[step.__dict__ for step in plan.steps],
        warnings=plan.warnings,
    )
