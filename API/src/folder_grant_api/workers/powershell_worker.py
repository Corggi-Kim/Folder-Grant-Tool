from sqlalchemy.orm import sessionmaker

from folder_grant_api.infrastructure.job_repository import JobRepository
from folder_grant_api.infrastructure.models import AuditLogRecord, utc_now
from folder_grant_api.infrastructure.powershell import PowerShellExecutionError, PowerShellRunner
from folder_grant_api.infrastructure.powershell_commands import UnsupportedStepError, build_step_command
from folder_grant_api.settings import Settings


def execute_powershell_job(session_factory: sessionmaker, settings: Settings, job_id: str) -> None:
    runner = PowerShellRunner(settings.powershell_path, settings.powershell_timeout_seconds)
    with session_factory() as session:
        repository = JobRepository(session)
        job = repository.get(job_id)
        if job is None or job.status != "queued":
            return
        if job.cancel_requested:
            repository.request_cancel(job, "powershell-worker")
            return

        job.status = "running"
        job.started_at = utc_now()
        session.commit()

        for step in job.steps:
            session.refresh(job)
            if job.cancel_requested:
                _cancel_remaining_steps(job)
                session.commit()
                return

            step.status = "running"
            step.started_at = utc_now()
            session.commit()
            try:
                command = build_step_command(step)
                result = runner.run(command)
            except Exception as exc:
                if isinstance(exc, PowerShellExecutionError):
                    code = _specific_error_code(str(exc), exc.code)
                elif isinstance(exc, (UnsupportedStepError, KeyError)):
                    code = "INVALID_STEP"
                else:
                    code = "INTERNAL_EXECUTOR_ERROR"
                _fail_job(job, step, code, str(exc))
                session.add(AuditLogRecord(
                    job_id=job.id,
                    event_type="job.failed",
                    actor="powershell-worker",
                    message=str(exc),
                    details={"step_order": step.step_order, "error_code": code},
                ))
                session.commit()
                return

            step.status = "succeeded"
            step.result = {
                "executor_mode": "powershell",
                "return_code": result.return_code,
                "stdout": result.stdout[-4000:],
            }
            step.finished_at = utc_now()
            session.commit()

        job.status = "succeeded"
        job.result_payload = {
            "executor_mode": "powershell",
            "changed_external_system": True,
            "message": "모든 권한 작업 단계가 완료되었습니다.",
        }
        job.finished_at = utc_now()
        session.add(AuditLogRecord(
            job_id=job.id,
            event_type="job.succeeded",
            actor="powershell-worker",
            message="모든 권한 작업 단계가 완료되었습니다.",
            details={},
        ))
        session.commit()


def _fail_job(job, step, code: str, message: str) -> None:
    now = utc_now()
    step.status = "failed"
    step.error_code = code
    step.error_message = message[-4000:]
    step.finished_at = now
    completed_steps = any(item.status == "succeeded" for item in job.steps)
    job.status = "partially_succeeded" if completed_steps else "failed"
    job.error_code = code
    job.error_message = message[-4000:]
    job.result_payload = {"executor_mode": "powershell", "changed_external_system": True}
    job.finished_at = now
    for remaining in job.steps:
        if remaining.status == "queued":
            remaining.status = "skipped"
            remaining.finished_at = now


def _cancel_remaining_steps(job) -> None:
    now = utc_now()
    job.status = "cancelled"
    job.finished_at = now
    job.result_payload = {"executor_mode": "powershell", "changed_external_system": True}
    for step in job.steps:
        if step.status == "queued":
            step.status = "cancelled"
            step.finished_at = now


def _specific_error_code(message: str, fallback: str) -> str:
    known_codes = (
        "PROJECT_ROOT_NOT_FOUND",
        "STUDY_ALL_NOT_FOUND",
        "ROOT_ACL_NOT_FOUND",
        "STUDY_ALL_ACL_NOT_FOUND",
        "ROBOCOPY_FAILED",
        "ACL_CLEANUP_FAILED",
        "PATH_NOT_FOUND",
    )
    return next((code for code in known_codes if code in message), fallback)
