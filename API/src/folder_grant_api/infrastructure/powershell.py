import base64
import os
import signal
import subprocess
import threading
import time
from dataclasses import dataclass
from typing import Callable

from .process_tree import WindowsJob

WINDOWS = os.name == 'nt'

# Keep the Windows command line independent of script/path/Role size. Only the
# owning worker opens the gate; a complete ASCII payload is required to run.
_STDIN_BOOTSTRAP = (
    "$ErrorActionPreference='Stop';"
    "if ([Console]::In.ReadLine() -cne 'FGT_START') { exit 125 };"
    "$fgtExpectedLength=0;"
    "if (-not [int]::TryParse([Console]::In.ReadLine(),[ref]$fgtExpectedLength) "
    "-or $fgtExpectedLength -le 0) { exit 125 };"
    "$fgtPayload=[Console]::In.ReadToEnd();"
    "if ($fgtPayload.Length -ne $fgtExpectedLength) { exit 125 };"
    "$fgtScriptSource=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($fgtPayload));"
    "& ([ScriptBlock]::Create($fgtScriptSource));"
)


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
    def __init__(self, executable: str, timeout_seconds: int, cancel_check: Callable[[], bool] | None = None,
                 on_event: Callable[[str, dict], None] | None = None):
        self.executable = executable
        self.timeout_seconds = timeout_seconds
        self.cancel_check = cancel_check
        self.on_event = on_event

    def _arguments(self) -> list[str]:
        return [self.executable, '-NoLogo', '-NoProfile', '-NonInteractive',
                '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encode_command(_STDIN_BOOTSTRAP)]

    def _launch_error(self, error: FileNotFoundError) -> PowerShellExecutionError:
        winerror = getattr(error, 'winerror', None)
        if winerror == 206:
            return PowerShellExecutionError('POWERSHELL_COMMAND_TOO_LONG',
                'PowerShell 실행 명령이 Windows 길이 제한을 초과했습니다. (Windows 오류 206)')
        suffix = f' (Windows 오류 {winerror})' if winerror is not None else ''
        if winerror not in {None, 2, 3}:
            return PowerShellExecutionError('POWERSHELL_START_FAILED',
                f'PowerShell 실행을 시작하지 못했습니다: {self.executable}{suffix}')
        return PowerShellExecutionError('POWERSHELL_NOT_FOUND',
            f'PowerShell 실행 파일을 찾을 수 없습니다: {self.executable}{suffix}')

    def run(self, command: str) -> PowerShellResult:
        wrapped = (
            "$ErrorActionPreference='Stop';"
            "$OutputEncoding=[System.Text.Encoding]::UTF8;"
            "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
            + command
        )
        payload = base64.b64encode(wrapped.encode('utf-8'))
        input_data = b'FGT_START\n' + str(len(payload)).encode('ascii') + b'\n' + payload
        if self.cancel_check is not None:
            return self._run_cancellable(input_data)
        try:
            completed = subprocess.run(
                self._arguments(),
                input=input_data,
                capture_output=True,
                check=False,
                timeout=self.timeout_seconds,
            )
        except FileNotFoundError as exc:
            raise self._launch_error(exc) from exc
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

    def _run_cancellable(self, input_data: bytes) -> PowerShellResult:
        if self.cancel_check():
            raise PowerShellCancelled()
        job = None
        process = None
        contained = False
        terminated = False
        io_started = False
        io_done = threading.Event()
        output = {}
        started_at = time.monotonic()

        def report(phase):
            if self.on_event:
                try:
                    self.on_event(phase, {'pid':getattr(process,'pid',None),
                                         'elapsed_seconds':round(time.monotonic()-started_at,3)})
                except Exception:
                    pass  # Diagnostics must never prevent stopping a Job.

        def collect_output():
            try:
                output['streams'] = process.communicate(input=input_data)
            except BaseException as exc:
                output['error'] = exc
            finally:
                io_done.set()

        try:
            if WINDOWS:
                job = WindowsJob()
            # No external changes are possible before the process is contained.
            process = subprocess.Popen(
                self._arguments(),
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
            report('process.contained')
            deadline = time.monotonic() + self.timeout_seconds
            while True:
                cancelled = self.cancel_check()
                timed_out = time.monotonic() >= deadline
                if cancelled or (timed_out and not io_done.is_set()):
                    report('cancel.observed' if cancelled else 'timeout.observed')
                    report('stop.begin')
                    self._terminate(process, job)
                    terminated = True
                    report('stop.confirmed')
                    if io_started and not io_done.wait(.1):
                        report('output.cleanup_pending')
                    if cancelled:
                        raise PowerShellCancelled()
                    raise PowerShellExecutionError('POWERSHELL_TIMEOUT',
                        f'PowerShell 실행 시간이 {self.timeout_seconds}초를 초과했습니다.')
                if io_done.is_set():
                    break
                if not io_started:
                    # Only this collector owns the pipes. The monitor never waits
                    # inside communicate(), stdin flushing, or output collection.
                    threading.Thread(target=collect_output, name='fgt-process-output', daemon=True).start()
                    io_started = True
                    report('process.started')
                io_done.wait(min(.25,max(.01,deadline-time.monotonic())))
            # Do not claim success if cancellation arrived as the process exited.
            if self.cancel_check():
                report('cancel.observed')
                report('stop.begin')
                self._terminate(process, job)
                terminated = True
                report('stop.confirmed')
                raise PowerShellCancelled()
            if 'error' in output:
                raise PowerShellExecutionError('POWERSHELL_IO_FAILED',
                    'PowerShell 출력 수집에 실패했습니다: '+str(output['error'])) from output['error']
            stdout, stderr = output['streams']
            result = PowerShellResult(process.returncode, stdout.decode('utf-8', errors='replace').strip(),
                                      stderr.decode('utf-8', errors='replace').strip())
            if not result.succeeded:
                raise PowerShellExecutionError('POWERSHELL_FAILED', result.stderr or result.stdout or f'PowerShell 종료 코드: {result.return_code}')
            return result
        except FileNotFoundError as exc:
            raise self._launch_error(exc) from exc
        finally:
            try:
                if process is not None:
                    if contained:
                        # Also collect orphaned descendants after the parent exits.
                        if not terminated:
                            report('cleanup.begin')
                            self._terminate(process, job)
                            report('cleanup.confirmed')
                    else:
                        # Assignment failed: the stdin gate is still closed.
                        process.kill()
                        process.communicate(timeout=5)
                    if not io_started:
                        # No collector is using these streams in this path.
                        for stream in (process.stdin,process.stdout,process.stderr):
                            if stream is not None: stream.close()
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
            # Pipe collection belongs to its own thread and may itself be stuck.
            # Confirm process exit independently; never communicate concurrently.
            process.wait(timeout=5)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise PowerShellExecutionError('POWERSHELL_STOP_FAILED',
                '작업 프로세스 종료를 확인하지 못했습니다. 서버에서 해당 작업 상태를 확인하세요.') from exc
