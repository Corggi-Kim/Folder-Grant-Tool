# Folder Grant API

기존 `Folder Grant Tool_v3.3.5.pyw`를 변경하지 않고, BUS와 과도기 테스트 클라이언트가 공통으로 사용할 서버 API를 개발하는 독립 프로젝트입니다.

현재 개발 단계는 Preview/Job API와 선택 가능한 Mock 또는 PowerShell 실행기를 제공합니다. 입력값을 검증한 뒤 대상 AD 그룹, 폴더 경로와 실행 계획을 만들고 SQLite에 작업 이력을 저장합니다.

> **주의:** 기본값은 `mock`이며 외부 시스템을 변경하지 않습니다. `powershell` 모드는 실제 AD 및 폴더 ACL을 변경하므로 테스트 계정·그룹·폴더에서만 사용하십시오.

## 개발 환경 실행

```powershell
cd API
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
uvicorn folder_grant_api.main:app --reload
```

Swagger UI: <http://127.0.0.1:8000/docs>

Swagger의 **Authorize** 버튼에 `.env`의 `FGT_API_KEY` 값을 입력해야 보호된 API를 호출할 수 있습니다. `/health`만 인증 없이 접근할 수 있습니다.

## 테스트

```powershell
pytest
```

## 제공 엔드포인트

| Method | Path | 설명 |
|---|---|---|
| `GET` | `/health` | 프로세스 상태 확인 |
| `GET` | `/ready` | SQLite 연결을 포함한 준비 상태 확인 |
| `POST` | `/api/v1/access-jobs/preview` | 실제 변경 없이 권한 작업 계획 생성 |
| `POST` | `/api/v1/access-jobs` | Mock Job 등록, `request_id` 중복 방지 |
| `GET` | `/api/v1/access-jobs/{job_id}` | 작업과 단계별 상태 조회 |
| `GET` | `/api/v1/access-jobs/by-request/{request_id}` | 요청번호로 작업 조회 |
| `POST` | `/api/v1/access-jobs/{job_id}/cancel` | 대기/실행 중 작업 취소 요청 |

## 설정

`.env.example`을 `.env`로 복사하여 환경별 값을 설정합니다. 기본 DB는 `API/data/folder_grant.db` SQLite 파일이며 Git에 포함되지 않습니다.

API Key는 다음과 같이 생성할 수 있습니다.

```powershell
py -c "import secrets; print(secrets.token_urlsafe(32))"
```

처음에는 반드시 다음 설정으로 실행합니다.

```dotenv
FGT_API_KEY=생성한-긴-임의-문자열
FGT_EXECUTOR_MODE=mock
```

Mock 검증 후 테스트 자원에 실제 작업을 적용할 때만 서버를 중지하고 다음으로 변경합니다.

```dotenv
FGT_EXECUTOR_MODE=powershell
FGT_POWERSHELL_PATH=powershell.exe
FGT_POWERSHELL_TIMEOUT_SECONDS=300
```

PowerShell 모드는 기본 API Key인 `change-me`로 시작할 수 없습니다. 사용자 사번·프로젝트 코드·Level2·Level3를 제한된 형식으로 검증하고 서버가 허용된 작업 단계만 PowerShell 명령으로 변환합니다. API에서 임의 PowerShell 문자열은 받지 않습니다.

## 현재 실행 방식의 제한

현재 Worker는 FastAPI 프로세스의 Background Task로 실행됩니다. 서버 프로세스가 종료되면 실행 중인 작업을 자동 복구하지 않으므로 테스트용으로만 사용해야 합니다. 또한 취소 요청은 작업 단계 사이에서 확인하며, 이미 실행 중인 단일 PowerShell 프로세스를 즉시 종료하지는 않습니다. 운영 전에는 별도 Worker 프로세스와 재시작 복구 기능을 추가해야 합니다.

## 다음 단계

1. 테스트 AD 그룹과 공유 폴더에서 PowerShell 모드 검증
2. 별도 Worker 프로세스와 서버 재시작 복구
3. PyQt/Selenium 테스트 클라이언트
4. BUS 연동 계약 확정
