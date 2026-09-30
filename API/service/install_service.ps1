[CmdletBinding()]
param(
    [PSCredential]$Credential,
    [switch]$OpenFirewall
)

$ErrorActionPreference = "Stop"
$apiRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $apiRoot ".venv\Scripts\python.exe"
$envFile = Join-Path $apiRoot ".env"

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "관리자 권한 PowerShell에서 실행해야 합니다."
}
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw ".venv가 없습니다: $python"
}
if (-not (Test-Path -LiteralPath $envFile -PathType Leaf)) {
    throw ".env 파일이 없습니다: $envFile"
}

$executorLine = Get-Content -LiteralPath $envFile | Where-Object { $_ -match '^FGT_EXECUTOR_MODE=' } | Select-Object -Last 1
$executorMode = if ($executorLine) { ($executorLine -split '=', 2)[1].Trim() } else { "mock" }
if ($executorMode -eq "powershell" -and -not $Credential) {
    throw "PowerShell 모드는 AD와 공유 폴더 권한이 있는 서비스 계정이 필요합니다. -Credential (Get-Credential)을 지정하세요."
}

Set-Location $apiRoot
& $python -m pip install -e ".[windows-service]"
if ($LASTEXITCODE -ne 0) { throw "서비스 의존성 설치에 실패했습니다." }

$serviceArgs = @("-m", "folder_grant_api.windows_service", "--startup", "auto")
if ($Credential) {
    $network = $Credential.GetNetworkCredential()
    $serviceArgs += @("--username", $Credential.UserName, "--password", $network.Password)
}
$serviceArgs += "install"
& $python @serviceArgs
if ($LASTEXITCODE -ne 0) { throw "Windows 서비스 등록에 실패했습니다." }

& sc.exe description FolderGrantApi "Folder Grant Tool FastAPI server" | Out-Null
& sc.exe failure FolderGrantApi reset= 86400 actions= restart/5000/restart/15000/restart/30000 | Out-Null
if ($OpenFirewall) {
    $portLine = Get-Content -LiteralPath $envFile | Where-Object { $_ -match '^FGT_PORT=' } | Select-Object -Last 1
    $port = if ($portLine) { [int](($portLine -split '=', 2)[1].Trim()) } else { 8000 }
    New-NetFirewallRule -DisplayName "Folder Grant API TCP $port" -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -ErrorAction SilentlyContinue | Out-Null
}

Start-Service FolderGrantApi
Start-Sleep -Seconds 3
Get-Service FolderGrantApi
