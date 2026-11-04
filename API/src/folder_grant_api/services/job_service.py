from sqlalchemy.orm import Session

from folder_grant_api.domain.operation_plan import OperationPlan
from folder_grant_api.infrastructure.job_repository import DuplicateRequestError, JobRepository
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.schemas.access_job import AccessJobRequest


class JobNotFoundError(LookupError):
    pass


class JobNotCancellableError(RuntimeError):
    pass


class JobNotRetryableError(RuntimeError):
    pass


class RequestIdConflictError(RuntimeError):
    pass


def require_retryable_attempt(
    job: JobRecord, executor_mode: str, restart_all: bool = False, repository: JobRepository | None = None,
) -> None:
    cancelled_before_start = (
        job.status == "cancelled"
        and job.started_at is None
        and job.result_payload is None
        and bool(job.steps)
        and (all(step.status == "cancelled" for step in job.steps)
             or (repository is not None and repository.has_trusted_reused_steps(job)))
    )
    stop_unconfirmed = job.error_code == "POWERSHELL_STOP_FAILED" or any(
        step.error_code == "POWERSHELL_STOP_FAILED" for step in job.steps
    )
    if (
        (job.status not in {"failed", "partially_succeeded", "cancelled"}
         and not (restart_all and job.status == "succeeded"))
        or executor_mode != "powershell"
        or stop_unconfirmed
        or not (cancelled_before_start or (job.result_payload or {}).get("executor_mode") == "powershell")
    ):
        raise JobNotRetryableError(job.status)


class JobService:
    def __init__(self, session: Session):
        self.repository = JobRepository(session)

    def register(self, request: AccessJobRequest, plan: OperationPlan) -> tuple[JobRecord, bool]:
        existing = self.repository.get_by_request_id(request.request_id)
        if existing:
            if existing.request_payload != request.model_dump(mode="json"):
                raise RequestIdConflictError(request.request_id)
            return self.repository.latest_attempt(existing), True
        try:
            return self.repository.create(request, plan), False
        except DuplicateRequestError:
            existing = self.repository.get_by_request_id(request.request_id)
            if existing is None:
                raise
            if existing.request_payload != request.model_dump(mode="json"):
                raise RequestIdConflictError(request.request_id)
            return self.repository.latest_attempt(existing), True

    def get(self, job_id: str) -> JobRecord:
        job = self.repository.get(job_id)
        if job is None:
            raise JobNotFoundError(job_id)
        return job

    def get_by_request_id(self, request_id: str) -> JobRecord:
        job = self.repository.get_by_request_id(request_id)
        if job is None:
            raise JobNotFoundError(request_id)
        return self.repository.latest_attempt(job)

    def retry(self, job_id: str, actor: str, executor_mode: str, restart_all: bool = False) -> tuple[JobRecord, bool]:
        parent = self.get(job_id)
        if parent.request_payload.get("confirm_group_creation") or parent.operation not in {"grant", "revoke"}:
            raise JobNotRetryableError(parent.status)
        require_retryable_attempt(parent, executor_mode, restart_all, self.repository)
        return self.repository.create_retry(parent, actor, restart_all)

    def cancel(self, job_id: str, actor: str) -> JobRecord:
        job = self.get(job_id)
        if job.status not in {"queued", "running"}:
            raise JobNotCancellableError(job.status)
        result = self.repository.request_cancel(job, actor)
        if result.status not in {'cancel_requested','cancelled'}:
            raise JobNotCancellableError(result.status)
        return result
