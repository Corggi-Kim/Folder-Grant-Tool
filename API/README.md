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
| `POST` | `/api/v1/project-jobs/preview` | 신규 프로젝트 그룹·폴더 생성 계획 확인 |
| `POST` | `/api/v1/project-jobs` | 신규 프로젝트 생성 Job 등록 |
| `GET` | `/api/v1/project-jobs/{job_id}` | 신규 프로젝트 Job 상태 조회 |
| `GET` | `/api/v1/project-jobs/by-request/{request_id}` | 요청번호로 신규 프로젝트 Job 조회 |
| `POST` | `/api/v1/project-jobs/{job_id}/cancel` | 신규 프로젝트 Job 취소 요청 |

## 설정

`.env`는 비밀 API Key를 포함하므로 저장소에 들어 있지 않습니다. 서버의 `API` 폴더에서 예제 파일을 복사해 직접 생성합니다.

```powershell
Copy-Item .env.example .env
```

기본 DB는 `API/data/folder_grant.db` SQLite 파일이며 Git에 포함되지 않습니다.

API Key는 다음과 같이 생성할 수 있습니다.

```powershell
py -c "import secrets; print(secrets.token_urlsafe(32))"
```

출력된 값을 `.env`의 `FGT_API_KEY`에 넣습니다. 처음 생성된 `.env`에서 반드시 확인할 핵심 부분은 다음과 같습니다.

처음에는 반드시 다음 설정으로 실행합니다.

```dotenv
FGT_API_KEY=생성한-긴-임의-문자열
FGT_EXECUTOR_MODE=mock
```

`mock`은 API 요청, SQLite 저장, 작업 상태 변경만 모의 실행하고 AD 및 폴더 ACL은 변경하지 않는 안전 확인 모드입니다. `.env`를 저장한 뒤 API 서버를 재시작해야 설정이 적용됩니다.

Mock 검증 후 테스트 자원에 실제 작업을 적용할 때만 서버를 중지하고 다음으로 변경합니다.

```dotenv
FGT_EXECUTOR_MODE=powershell
FGT_POWERSHELL_PATH=powershell.exe
FGT_POWERSHELL_TIMEOUT_SECONDS=300
FGT_TEMPLATE_ROOT=\\LSK_S010\Study folder\_Template
FGT_GROUP_OU_PATH=OU=Group Project Folder,OU=0.Management Object Group,OU=lskglobal,DC=lskglobal,DC=com
```

PowerShell 모드는 기본 API Key인 `change-me`로 시작할 수 없습니다. 사용자 사번·프로젝트 코드·Level2·Level3를 제한된 형식으로 검증하고 서버가 허용된 작업 단계만 PowerShell 명령으로 변환합니다. API에서 임의 PowerShell 문자열은 받지 않습니다.

## 현재 실행 방식의 제한

현재 Worker는 FastAPI 프로세스의 Background Task로 실행됩니다. 서버 프로세스가 종료되면 실행 중인 작업을 자동 복구하지 않으므로 테스트용으로만 사용해야 합니다. 또한 취소 요청은 작업 단계 사이에서 확인하며, 이미 실행 중인 단일 PowerShell 프로세스를 즉시 종료하지는 않습니다. 운영 전에는 별도 Worker 프로세스와 재시작 복구 기능을 추가해야 합니다.

## 다음 단계

1. 신규 프로젝트 생성 API를 테스트 그룹·경로에서 검증
2. 별도 Worker 프로세스와 서버 재시작 복구
3. 기존 BUS Selenium 기능을 유지한 단일 파일 클라이언트
4. BUS 연동 계약 확정

## 신규 프로젝트 생성

`project-jobs`는 프로젝트 코드와 이름을 받아 다음 단계를 수행합니다.

1. Study AD 보안 그룹 확인 또는 생성
2. `_Template`에서 프로젝트 루트 폴더 복사
3. 복사된 기본 경로의 해석 불가능 SID 정리
4. 프로젝트 루트에 그룹 읽기/실행 권한 적용
5. `Study/All`에 그룹 수정 권한 적용
6. 그룹·폴더·기본 ACL 재검증

그룹과 폴더가 이미 존재하면 재사용하므로 같은 프로젝트에 대해 다시 실행할 수 있습니다. `robocopy` 종료 코드는 `0~7`을 정상 범위로 처리하고 `8 이상`을 실패로 처리합니다. 실제 실행 전에는 `/api/v1/project-jobs/preview` 결과를 먼저 확인하십시오.

## 수동 테스트 클라이언트

여러 Level2, Level3, Role 조합을 반복 시험할 수 있는 PyQt 클라이언트가 `test_client/`에 있습니다.

```powershell
python -m pip install -e ".[client]"
python -m test_client.main
```

현재 버전은 수동 요청 추가, Preview, 다중 Job 등록, 상태 자동 조회, 취소 및 결과 로그를 제공합니다. Selenium BUS 조회 연결은 후속 단계입니다.
