import base64
import os
import signal
import subprocess
import time
from dataclasses import dataclass
from typing import Callable

from .process_tree import WindowsJob

WINDOWS = os.name == 'nt'


@dataclass(frozen=True)
class PowerShellResult:
    return_code: int
    stdout: str
    stderr: str

    @property
    def succeeded(self) -> bool:
        return self.return_code == 0


class PowerShellExecutionError(RuntimeError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


class PowerShellCancelled(PowerShellExecutionError):
    def __init__(self):
        super().__init__('POWERSHELL_CANCELLED', '작업을 중지했습니다. 이미 적용된 권한은 자동 복구되지 않습니다.')


def encode_command(command: str) -> str:
    return base64.b64encode(command.encode("utf-16-le")).decode("ascii")


class PowerShellRunner:
    def __init__(self, executable: str, timeout_seconds: int, cancel_check: Callable[[], bool] | None = None):
        self.executable = executable
        self.timeout_seconds = timeout_seconds
        self.cancel_check = cancel_check

    def run(self, command: str) -> PowerShellResult:
        wrapped = (
            "$ErrorActionPreference='Stop';"
            "$OutputEncoding=[System.Text.Encoding]::UTF8;"
            "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
            + command
        )
        if self.cancel_check is not None:
            return self._run_cancellable(wrapped)
        try:
            completed = subprocess.run(
                [
                    self.executable,
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encode_command(wrapped),
                ],
                capture_output=True,
                check=False,
                timeout=self.timeout_seconds,
            )
        except FileNotFoundError as exc:
            raise PowerShellExecutionError(
                "POWERSHELL_NOT_FOUND",
                f"PowerShell 실행 파일을 찾을 수 없습니다: {self.executable}",
            ) from exc
        except subprocess.TimeoutExpired as exc:
            raise PowerShellExecutionError(
                "POWERSHELL_TIMEOUT",
                f"PowerShell 실행 시간이 {self.timeout_seconds}초를 초과했습니다.",
            ) from exc

        stdout = completed.stdout.decode("utf-8", errors="replace").strip()
        stderr = completed.stderr.decode("utf-8", errors="replace").strip()
        result = PowerShellResult(completed.returncode, stdout, stderr)
        if not result.succeeded:
            message = stderr or stdout or f"PowerShell 종료 코드: {completed.returncode}"
            raise PowerShellExecutionError("POWERSHELL_FAILED", message)
        return result

    def _run_cancellable(self, wrapped: str) -> PowerShellResult:
        if self.cancel_check():
            raise PowerShellCancelled()
        job = None
        process = None
        contained = False
        try:
            if WINDOWS:
                job = WindowsJob()
            # No external changes are possible before the process is contained.
            gated = "if ([Console]::In.ReadLine() -cne 'FGT_START') { exit 125 };" + wrapped
            process = subprocess.Popen(
                [self.executable, '-NoLogo', '-NoProfile', '-NonInteractive',
                 '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encode_command(gated)],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                start_new_session=not WINDOWS,
            )
            if job:
                try:
                    job.attach(process)
                except OSError as exc:
                    raise PowerShellExecutionError('POWERSHELL_PROCESS_CONTROL_FAILED',
                        '작업 프로세스를 안전하게 관리할 수 없어 실행하지 않았습니다.') from exc
            contained = True
            deadline = time.monotonic() + self.timeout_seconds
            first_input = b'FGT_START\n'
            while True:
                cancelled = self.cancel_check()
                timed_out = time.monotonic() >= deadline
                if cancelled or timed_out:
                    self._terminate(process, job)
                    if cancelled:
                        raise PowerShellCancelled()
                    raise PowerShellExecutionError('POWERSHELL_TIMEOUT',
                        f'PowerShell 실행 시간이 {self.timeout_seconds}초를 초과했습니다.')
                try:
                    stdout, stderr = process.communicate(input=first_input, timeout=min(.5, max(.01, deadline-time.monotonic())))
                    break
                except subprocess.TimeoutExpired:
                    first_input = None
            # Do not claim success if cancellation arrived as the process exited.
            if self.cancel_check():
                self._terminate(process, job)
                raise PowerShellCancelled()
            result = PowerShellResult(process.returncode, stdout.decode('utf-8', errors='replace').strip(),
                                      stderr.decode('utf-8', errors='replace').strip())
            if not result.succeeded:
                raise PowerShellExecutionError('POWERSHELL_FAILED', result.stderr or result.stdout or f'PowerShell 종료 코드: {result.return_code}')
            return result
        except FileNotFoundError as exc:
            raise PowerShellExecutionError('POWERSHELL_NOT_FOUND', f'PowerShell 실행 파일을 찾을 수 없습니다: {self.executable}') from exc
        finally:
            try:
                if process is not None:
                    if contained:
                        # Also collect orphaned descendants after the parent exits.
                        self._terminate(process, job)
                    else:
                        # Assignment failed: the stdin gate is still closed.
                        process.kill()
                        process.communicate(timeout=5)
            finally:
                if job:
                    job.close()

    @staticmethod
    def _terminate(process, job):
        try:
            if job:
                job.terminate()
            else:
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
            process.communicate(timeout=5)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise PowerShellExecutionError('POWERSHELL_STOP_FAILED',
                '작업 프로세스 종료를 확인하지 못했습니다. 서버에서 해당 작업 상태를 확인하세요.') from exc
