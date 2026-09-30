[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$apiRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $apiRoot ".venv\Scripts\python.exe"

if (Get-Service FolderGrantApi -ErrorAction SilentlyContinue) {
    Stop-Service FolderGrantApi -ErrorAction SilentlyContinue
    & $python -m folder_grant_api.windows_service remove
    if ($LASTEXITCODE -ne 0) { throw "Windows 서비스 제거에 실패했습니다." }
}
