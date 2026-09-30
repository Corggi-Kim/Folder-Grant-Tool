"""Windows Service wrapper for the Folder Grant API.

Install this module through ``python -m folder_grant_api.windows_service``.
The optional ``windows-service`` dependency supplies pywin32.
"""

import os
import subprocess
from pathlib import Path

import servicemanager
import win32event
import win32service
import win32serviceutil

from folder_grant_api.settings import Settings


SERVICE_NAME = "FolderGrantApi"
SERVICE_DISPLAY_NAME = "Folder Grant API"


def api_root() -> Path:
    configured = os.environ.get("FGT_SERVICE_WORKDIR")
    return Path(configured).resolve() if configured else Path(__file__).resolve().parents[2]


def uvicorn_command(settings: Settings, root: Path) -> list[str]:
    python_executable = root / ".venv" / "Scripts" / "python.exe"
    return [
        str(python_executable),
        "-m",
        "uvicorn",
        "folder_grant_api.main:app",
        "--host",
        settings.host,
        "--port",
        str(settings.port),
    ]


class FolderGrantApiService(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY_NAME
    _svc_description_ = "Folder Grant Tool FastAPI server"
    _svc_reg_class_ = "folder_grant_api.windows_service.FolderGrantApiService"

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.process: subprocess.Popen | None = None

    def SvcStop(self) -> None:
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        if self.process and self.process.poll() is None:
            self.process.terminate()

    def SvcDoRun(self) -> None:
        root = api_root()
        os.chdir(root)
        settings = Settings()
        settings.validate_runtime_safety()
        settings.ensure_local_directories()

        log_dir = root / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "service.log"
        servicemanager.LogInfoMsg(f"{SERVICE_DISPLAY_NAME} starting in {root}")

        with log_path.open("a", encoding="utf-8", buffering=1) as log:
            self.process = subprocess.Popen(
                uvicorn_command(settings, root),
                cwd=root,
                stdin=subprocess.DEVNULL,
                stdout=log,
                stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            while self.process.poll() is None:
                if win32event.WaitForSingleObject(self.stop_event, 1000) == win32event.WAIT_OBJECT_0:
                    self.process.terminate()
                    break
            try:
                return_code = self.process.wait(timeout=15)
            except subprocess.TimeoutExpired:
                self.process.kill()
                return_code = self.process.wait(timeout=5)

        if return_code and win32event.WaitForSingleObject(self.stop_event, 0) != win32event.WAIT_OBJECT_0:
            raise RuntimeError(f"Uvicorn exited unexpectedly with code {return_code}; see {log_path}")
        servicemanager.LogInfoMsg(f"{SERVICE_DISPLAY_NAME} stopped")


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(FolderGrantApiService)
