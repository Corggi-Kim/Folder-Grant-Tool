from sqlalchemy.orm import sessionmaker

from folder_grant_api.infrastructure.job_repository import JobRepository
from folder_grant_api.infrastructure.models import AuditLogRecord, utc_now


def execute_mock_job(session_factory: sessionmaker, job_id: str) -> None:
    """Simulate every step and persist a non-production `simulated` result."""
    with session_factory() as session:
        repository = JobRepository(session)
        job = repository.get(job_id)
        if job is None or job.status != "queued":
            return
        if job.cancel_requested:
            repository.request_cancel(job, "mock-worker")
            return

        job.status = "running"
        job.started_at = utc_now()
        session.commit()

        for step in job.steps:
            session.refresh(job)
            if job.cancel_requested:
                job.status = "cancelled"
                job.finished_at = utc_now()
                step.status = "cancelled"
                step.finished_at = job.finished_at
                session.commit()
                return
            step.status = "running"
            step.started_at = utc_now()
            session.commit()
            step.status = "simulated"
            step.result = {"executor_mode": "mock", "changed_external_system": False}
            step.finished_at = utc_now()
            session.commit()

        job.status = "simulated"
        job.result_payload = {
            "executor_mode": "mock",
            "changed_external_system": False,
            "message": "Mock Worker가 작업 계획을 모의 처리했습니다.",
        }
        job.finished_at = utc_now()
        session.add(AuditLogRecord(
            job_id=job.id,
            event_type="job.simulated",
            actor="mock-worker",
            message="외부 시스템 변경 없이 작업을 모의 처리했습니다.",
            details={"changed_external_system": False},
        ))
        session.commit()
