from sqlalchemy.orm import sessionmaker

from folder_grant_api.settings import Settings

from .mock_worker import execute_mock_job
from .powershell_worker import execute_powershell_job


def execute_job(session_factory: sessionmaker, settings: Settings, job_id: str) -> None:
    if settings.executor_mode == "powershell":
        execute_powershell_job(session_factory, settings, job_id)
    else:
        execute_mock_job(session_factory, job_id)
