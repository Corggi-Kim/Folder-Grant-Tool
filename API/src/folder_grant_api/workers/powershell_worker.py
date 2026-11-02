from sqlalchemy.orm import sessionmaker
from sqlalchemy import select, update

from folder_grant_api.infrastructure.job_repository import JobRepository
from folder_grant_api.infrastructure.models import AuditLogRecord, JobRecord, utc_now
from folder_grant_api.infrastructure.powershell import PowerShellCancelled, PowerShellExecutionError, PowerShellRunner
from folder_grant_api.infrastructure.powershell_commands import UnsupportedStepError, build_step_command
from folder_grant_api.settings import Settings


def execute_powershell_job(session_factory: sessionmaker, settings: Settings, job_id: str) -> None:
    def cancellation_requested():
        # A fresh short transaction sees cancellation from another API request.
        with session_factory() as cancel_session:
            flag = cancel_session.scalar(select(JobRecord.cancel_requested).where(JobRecord.id == job_id))
            return flag is None or bool(flag)
    runner = PowerShellRunner(settings.powershell_path, settings.powershell_timeout_seconds,
                             cancel_check=cancellation_requested)
    with session_factory() as session:
        repository = JobRepository(session)
        job = repository.get(job_id)
        if job is None or job.status != "queued":
            return
        if job.cancel_requested:
            repository.request_cancel(job, "powershell-worker")
            return

        started = session.execute(update(JobRecord).where(
            JobRecord.id == job_id, JobRecord.status == 'queued', JobRecord.cancel_requested.is_(False)
        ).values(status='running',started_at=utc_now()),execution_options={'synchronize_session':False})
        session.commit()
        if not started.rowcount:
            return
        session.refresh(job)

        for step in job.steps:
            session.refresh(job)
            if job.cancel_requested:
                _cancel_remaining_steps(job)
                _audit_cancel(session, job)
                session.commit()
                return

            if step.status == "succeeded":
                continue
            step.status = "running"
            step.started_at = utc_now()
            session.commit()
            try:
                command = build_step_command(step)
                result = runner.run(command)
            except PowerShellCancelled as exc:
                session.refresh(job)
                step.status = 'cancelled'
                step.error_code = exc.code
                step.error_message = str(exc)
                step.finished_at = utc_now()
                _cancel_remaining_steps(job)
                _audit_cancel(session, job, step.step_order)
                session.commit()
                return
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

        payload = {
            "executor_mode": "powershell",
            "changed_external_system": True,
            "message": ("Isolated 그룹 준비가 완료되었습니다. 권한 요청을 다시 검증하세요."
                        if job.request_payload.get("confirm_group_creation") else "모든 권한 작업 단계가 완료되었습니다."),
            **({"job_kind": "group_preparation"} if job.request_payload.get("confirm_group_creation") else {}),
        }
        completed = session.execute(update(JobRecord).where(
            JobRecord.id == job_id, JobRecord.status == 'running', JobRecord.cancel_requested.is_(False)
        ).values(status='succeeded',result_payload=payload,finished_at=utc_now()),
            execution_options={'synchronize_session':False})
        if not completed.rowcount:
            session.refresh(job)
            _cancel_remaining_steps(job)
            _audit_cancel(session, job)
            session.commit()
            return
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
    job.error_code = 'POWERSHELL_CANCELLED'
    job.error_message = '작업을 중지했습니다. 이미 적용된 권한은 자동 복구되지 않습니다.'
    job.result_payload = {"executor_mode": "powershell", "changed_external_system": True,
                          "partial_changes_possible": True, "message": job.error_message,
                          **({"job_kind": "group_preparation"} if job.request_payload.get('confirm_group_creation') else {})}
    for step in job.steps:
        if step.status == "queued":
            step.status = "cancelled"
            step.finished_at = now


def _audit_cancel(session, job, step_order=None):
    session.add(AuditLogRecord(job_id=job.id, event_type='job.cancelled', actor='powershell-worker',
                              message=job.error_message, details={'step_order':step_order, 'partial_changes_possible':True}))


def _specific_error_code(message: str, fallback: str) -> str:
    known_codes = (
        "PROJECT_ROOT_NOT_FOUND",
        "STUDY_ALL_NOT_FOUND",
        "ROOT_ACL_NOT_FOUND",
        "STUDY_ALL_ACL_NOT_FOUND",
        "GROUP_SID_NOT_FOUND",
        "ROBOCOPY_FAILED",
        "ACL_CLEANUP_FAILED",
        "PATH_NOT_FOUND",
    )
    return next((code for code in known_codes if code in message), fallback)
