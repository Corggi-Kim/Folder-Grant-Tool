import base64
import subprocess
from dataclasses import dataclass


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


def encode_command(command: str) -> str:
    return base64.b64encode(command.encode("utf-16-le")).decode("ascii")


class PowerShellRunner:
    def __init__(self, executable: str, timeout_seconds: int):
        self.executable = executable
        self.timeout_seconds = timeout_seconds

    def run(self, command: str) -> PowerShellResult:
        wrapped = (
            "$ErrorActionPreference='Stop';"
            "$OutputEncoding=[System.Text.Encoding]::UTF8;"
            "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
            + command
        )
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
