from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Request, Response, status
from sqlalchemy.orm import Session

from folder_grant_api.api.dependencies import get_session
from folder_grant_api.api.security import require_api_key
from folder_grant_api.domain.project_codes import InvalidProjectCode
from folder_grant_api.domain.project_plan import ProjectCreationPlan, build_project_creation_plan
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.schemas.project_job import (
    CancelProjectJobRequest,
    ProjectJobRequest,
    ProjectJobResponse,
    ProjectJobStepResponse,
    ProjectPreviewResponse,
)
from folder_grant_api.services.job_service import JobNotCancellableError, JobNotFoundError, RequestIdConflictError
from folder_grant_api.services.project_job_service import ProjectJobService
from folder_grant_api.settings import Settings, get_settings
from folder_grant_api.workers.dispatcher import execute_job

router = APIRouter(
    prefix="/api/v1/project-jobs",
    tags=["project-jobs"],
    dependencies=[Depends(require_api_key)],
)


def create_project_plan(payload: ProjectJobRequest, settings: Settings) -> ProjectCreationPlan:
    try:
        return build_project_creation_plan(
            project_code=payload.project_code,
            project_name=payload.project_name,
            share_root=settings.share_root,
            template_root=settings.template_root,
            group_ou_path=settings.group_ou_path,
        )
    except InvalidProjectCode as exc:
        raise HTTPException(
            status_code=422,
            detail={"code": "INVALID_PROJECT_CODE", "message": str(exc)},
        ) from exc


def serialize_project_job(job: JobRecord, *, executor_mode: str, replayed: bool = False) -> ProjectJobResponse:
    recorded_mode = (job.result_payload or {}).get("executor_mode", executor_mode)
    payload = job.request_payload
    return ProjectJobResponse(
        job_id=job.id,
        request_id=job.request_id,
        replayed=replayed,
        status=job.status,
        executor_mode=recorded_mode,
        project_code=payload["project_code"],
        project_name=payload.get("project_name", ""),
        requested_by=job.requested_by,
        cancel_requested=job.cancel_requested,
        result=job.result_payload,
        error_code=job.error_code,
        error_message=job.error_message,
        created_at=job.created_at,
        started_at=job.started_at,
        finished_at=job.finished_at,
        steps=[
            ProjectJobStepResponse(
                order=step.step_order,
                type=step.step_type,
                status=step.status,
                target=step.target,
                description=step.description,
                details=step.details,
                result=step.result,
                error_code=step.error_code,
                error_message=step.error_message,
                started_at=step.started_at,
                finished_at=step.finished_at,
            )
            for step in job.steps
        ],
    )


@router.post("/preview", response_model=ProjectPreviewResponse)
def preview_project_job(
    payload: ProjectJobRequest,
    settings: Settings = Depends(get_settings),
) -> ProjectPreviewResponse:
    plan = create_project_plan(payload, settings)
    return ProjectPreviewResponse(
        request_id=payload.request_id,
        canonical_project_code=plan.canonical_project_code,
        group_name=plan.group_name,
        root_path=plan.root_path,
        study_all_path=plan.study_all_path,
        template_path=plan.template_path,
        steps=[step.__dict__ for step in plan.steps],
        warnings=plan.warnings,
    )


@router.post("", response_model=ProjectJobResponse, status_code=status.HTTP_202_ACCEPTED)
def register_project_job(
    payload: ProjectJobRequest,
    background_tasks: BackgroundTasks,
    request: Request,
    response: Response,
    session: Session = Depends(get_session),
) -> ProjectJobResponse:
    plan = create_project_plan(payload, request.app.state.settings)
    try:
        job, replayed = ProjectJobService(session).register(payload, plan)
    except RequestIdConflictError as exc:
        raise HTTPException(
            status_code=409,
            detail={"code": "REQUEST_ID_CONFLICT", "message": "같은 request_id에 다른 요청 데이터를 사용할 수 없습니다."},
        ) from exc
    if replayed:
        response.status_code = status.HTTP_200_OK
    elif job.status == "queued":
        background_tasks.add_task(
            execute_job,
            request.app.state.session_factory,
            request.app.state.settings,
            job.id,
        )
    return serialize_project_job(job, executor_mode=request.app.state.settings.executor_mode, replayed=replayed)


@router.get("/{job_id}", response_model=ProjectJobResponse)
def get_project_job(job_id: str, request: Request, session: Session = Depends(get_session)) -> ProjectJobResponse:
    try:
        job = ProjectJobService(session).get(job_id)
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "JOB_NOT_FOUND", "message": "작업을 찾을 수 없습니다."}) from exc
    return serialize_project_job(job, executor_mode=request.app.state.settings.executor_mode)


@router.get("/by-request/{request_id}", response_model=ProjectJobResponse)
def get_project_job_by_request(
    request_id: str,
    request: Request,
    session: Session = Depends(get_session),
) -> ProjectJobResponse:
    try:
        job = ProjectJobService(session).get_by_request_id(request_id)
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "JOB_NOT_FOUND", "message": "작업을 찾을 수 없습니다."}) from exc
    return serialize_project_job(job, executor_mode=request.app.state.settings.executor_mode)


@router.post("/{job_id}/cancel", response_model=ProjectJobResponse)
def cancel_project_job(
    job_id: str,
    payload: CancelProjectJobRequest,
    request: Request,
    session: Session = Depends(get_session),
) -> ProjectJobResponse:
    try:
        job = ProjectJobService(session).cancel(job_id, payload.requested_by)
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "JOB_NOT_FOUND", "message": "작업을 찾을 수 없습니다."}) from exc
    except JobNotCancellableError as exc:
        raise HTTPException(
            status_code=409,
            detail={"code": "JOB_NOT_CANCELLABLE", "message": f"현재 상태에서는 취소할 수 없습니다: {exc}"},
        ) from exc
    return serialize_project_job(job, executor_mode=request.app.state.settings.executor_mode)
