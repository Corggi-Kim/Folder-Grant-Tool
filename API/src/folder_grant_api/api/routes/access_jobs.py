from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Request, Response, status
from sqlalchemy.orm import Session

from folder_grant_api.api.dependencies import get_session
from folder_grant_api.api.routes.preview import create_plan
from folder_grant_api.api.security import require_api_key
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.schemas.access_job import AccessJobRequest, CancelJobRequest, JobResponse, JobStepResponse
from folder_grant_api.services.job_service import (
    JobNotCancellableError,
    JobNotFoundError,
    JobService,
    RequestIdConflictError,
)
from folder_grant_api.workers.dispatcher import execute_job

router = APIRouter(
    prefix="/api/v1/access-jobs",
    tags=["access-jobs"],
    dependencies=[Depends(require_api_key)],
)


def serialize_job(job: JobRecord, *, executor_mode: str, replayed: bool = False) -> JobResponse:
    recorded_mode = (job.result_payload or {}).get("executor_mode", executor_mode)
    return JobResponse(
        job_id=job.id,
        request_id=job.request_id,
        replayed=replayed,
        status=job.status,
        executor_mode=recorded_mode,
        operation=job.operation,
        project_status=job.project_status,
        requested_by=job.requested_by,
        cancel_requested=job.cancel_requested,
        result=job.result_payload,
        error_code=job.error_code,
        error_message=job.error_message,
        created_at=job.created_at,
        started_at=job.started_at,
        finished_at=job.finished_at,
        steps=[
            JobStepResponse(
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


@router.post("", response_model=JobResponse, status_code=status.HTTP_202_ACCEPTED)
def register_access_job(
    payload: AccessJobRequest,
    background_tasks: BackgroundTasks,
    request: Request,
    response: Response,
    session: Session = Depends(get_session),
) -> JobResponse:
    plan = create_plan(payload, getattr(request.app.state, "settings", None))
    try:
        job, replayed = JobService(session).register(payload, plan)
    except RequestIdConflictError as exc:
        raise HTTPException(
            status_code=409,
            detail={
                "code": "REQUEST_ID_CONFLICT",
                "message": "같은 request_id에 다른 요청 데이터를 사용할 수 없습니다.",
            },
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
    return serialize_job(job, executor_mode=request.app.state.settings.executor_mode, replayed=replayed)


@router.get("/{job_id}", response_model=JobResponse)
def get_access_job(job_id: str, request: Request, session: Session = Depends(get_session)) -> JobResponse:
    try:
        return serialize_job(JobService(session).get(job_id), executor_mode=request.app.state.settings.executor_mode)
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "JOB_NOT_FOUND", "message": "작업을 찾을 수 없습니다."}) from exc


@router.get("/by-request/{request_id}", response_model=JobResponse)
def get_access_job_by_request(
    request_id: str,
    request: Request,
    session: Session = Depends(get_session),
) -> JobResponse:
    try:
        return serialize_job(
            JobService(session).get_by_request_id(request_id),
            executor_mode=request.app.state.settings.executor_mode,
        )
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "JOB_NOT_FOUND", "message": "작업을 찾을 수 없습니다."}) from exc


@router.post("/{job_id}/cancel", response_model=JobResponse)
def cancel_access_job(
    job_id: str,
    payload: CancelJobRequest,
    request: Request,
    session: Session = Depends(get_session),
) -> JobResponse:
    try:
        return serialize_job(
            JobService(session).cancel(job_id, payload.requested_by),
            executor_mode=request.app.state.settings.executor_mode,
        )
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "JOB_NOT_FOUND", "message": "작업을 찾을 수 없습니다."}) from exc
    except JobNotCancellableError as exc:
        raise HTTPException(
            status_code=409,
            detail={"code": "JOB_NOT_CANCELLABLE", "message": f"현재 상태에서는 취소할 수 없습니다: {exc}"},
        ) from exc
