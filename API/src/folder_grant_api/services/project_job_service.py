from sqlalchemy.orm import Session

from folder_grant_api.domain.project_plan import ProjectCreationPlan
from folder_grant_api.infrastructure.job_repository import DuplicateRequestError, JobRepository
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.schemas.project_job import ProjectJobRequest

from .job_service import JobNotCancellableError, JobNotFoundError, RequestIdConflictError


class ProjectJobService:
    def __init__(self, session: Session):
        self.repository = JobRepository(session)

    def register(self, request: ProjectJobRequest, plan: ProjectCreationPlan) -> tuple[JobRecord, bool]:
        expected_payload = request.model_dump(mode="json")
        existing = self.repository.get_by_request_id(request.request_id)
        if existing:
            if existing.operation != "create_project" or existing.request_payload != expected_payload:
                raise RequestIdConflictError(request.request_id)
            return existing, True
        try:
            return self.repository.create_project(request, plan), False
        except DuplicateRequestError:
            existing = self.repository.get_by_request_id(request.request_id)
            if existing is None:
                raise
            if existing.operation != "create_project" or existing.request_payload != expected_payload:
                raise RequestIdConflictError(request.request_id)
            return existing, True

    def get(self, job_id: str) -> JobRecord:
        job = self.repository.get(job_id)
        if job is None or job.operation != "create_project":
            raise JobNotFoundError(job_id)
        return job

    def get_by_request_id(self, request_id: str) -> JobRecord:
        job = self.repository.get_by_request_id(request_id)
        if job is None or job.operation != "create_project":
            raise JobNotFoundError(request_id)
        return job

    def cancel(self, job_id: str, actor: str) -> JobRecord:
        job = self.get(job_id)
        if job.status not in {"queued", "running"}:
            raise JobNotCancellableError(job.status)
        return self.repository.request_cancel(job, actor)
