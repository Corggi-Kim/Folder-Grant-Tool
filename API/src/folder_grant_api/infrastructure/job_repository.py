from sqlalchemy import select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session, selectinload

from folder_grant_api.domain.operation_plan import OperationPlan
from folder_grant_api.schemas.access_job import AccessJobRequest

from .models import AuditLogRecord, JobRecord, JobStepRecord, utc_now


class DuplicateRequestError(RuntimeError):
    pass


class JobRepository:
    def __init__(self, session: Session):
        self.session = session

    def get(self, job_id: str) -> JobRecord | None:
        statement = select(JobRecord).options(selectinload(JobRecord.steps)).where(JobRecord.id == job_id)
        return self.session.scalar(statement)

    def get_by_request_id(self, request_id: str) -> JobRecord | None:
        statement = (
            select(JobRecord)
            .options(selectinload(JobRecord.steps))
            .where(JobRecord.request_id == request_id)
        )
        return self.session.scalar(statement)

    def create(self, request: AccessJobRequest, plan: OperationPlan) -> JobRecord:
        job = JobRecord(
            request_id=request.request_id,
            source=request.source,
            operation=request.operation.value,
            project_status=request.project_status.value,
            status="queued",
            request_payload=request.model_dump(mode="json"),
            requested_by=request.requested_by,
            steps=[
                JobStepRecord(
                    step_order=index,
                    step_type=step.type,
                    status="queued",
                    target=step.target,
                    description=step.description,
                    details={"permission": step.permission, **step.metadata},
                )
                for index, step in enumerate(plan.steps, start=1)
            ],
        )
        try:
            self.session.add(job)
            self.session.flush()
            self.session.add(AuditLogRecord(
                job_id=job.id,
                event_type="job.created",
                actor=request.requested_by,
                message="권한 작업이 접수되었습니다.",
                details={"request_id": request.request_id, "executor_mode": "mock"},
            ))
            self.session.commit()
        except IntegrityError as exc:
            self.session.rollback()
            raise DuplicateRequestError(request.request_id) from exc
        return self.get(job.id)  # type: ignore[return-value]

    def request_cancel(self, job: JobRecord, actor: str) -> JobRecord:
        job.cancel_requested = True
        if job.status == "queued":
            job.status = "cancelled"
            job.finished_at = utc_now()
            for step in job.steps:
                step.status = "cancelled"
                step.finished_at = job.finished_at
        elif job.status == "running":
            job.status = "cancel_requested"
        self.session.add(AuditLogRecord(
            job_id=job.id,
            event_type="job.cancel_requested",
            actor=actor,
            message="작업 취소가 요청되었습니다.",
            details={},
        ))
        self.session.commit()
        return self.get(job.id)  # type: ignore[return-value]
