# Folder Grant API

권한 부여·해제와 신규 프로젝트 생성을 처리하는 독립 API 서버입니다. 현재 패키지 버전은 **0.4.0**입니다. BUS 조회·완료는 [정식 Client](../Client/README.md)가 담당하고, API는 AD·폴더 권한 실행과 작업 이력을 관리합니다. 기존 `Folder Grant Tool_v3.3.5.pyw`는 수정하지 않습니다.

현재 구현은 실제 자원 검증·Preview·Job API와 Mock/PowerShell 실행기를 제공합니다. 안내 기준은 2026-10-27 `main b09100d2e6d5`입니다. 입력값을 검증한 뒤 대상 AD 그룹, 폴더 경로와 실행 계획을 만들고 SQLite에 작업 이력을 저장합니다.

> **주의:** 기본값은 `mock`이며 외부 시스템을 변경하지 않습니다. `powershell` 모드는 실제 AD 및 폴더 ACL을 변경하므로 오픈 전 테스트 계정·그룹·폴더에서 먼저 검증하십시오.

[API 설계·운영·유지보수 Word](../docs/API_설계_운영_유지보수.docx) · [운영 문서 목록](../docs/README.md)

## 개발 환경 실행

```powershell
cd API
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
uvicorn folder_grant_api.main:app --reload
```

Swagger UI: <http://127.0.0.1:8000/docs>

Swagger의 **Authorize** 버튼에 `.env`의 `FGT_API_KEY` 값을 입력해야 보호된 API를 호출할 수 있습니다. 상태·업무 엔드포인트 중 `/health`만 인증 없이 접근할 수 있습니다. `/docs`와 OpenAPI 문서는 별도로 접근 제한되어 있지 않습니다. `/ready`는 인증과 DB 연결을 검사하지만 AD·공유 폴더 접근까지 확인하지는 않습니다. Preview도 실행 계획만 만들며 실제 자원의 존재·권한은 검사하지 않습니다.

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
| `POST` | `/api/v1/access-jobs/validate` | 실제 사용자·그룹·폴더·현재 ACL 읽기 전용 조회 |
| `POST` | `/api/v1/access-jobs` | 설정된 Mock/PowerShell 모드로 Job 등록, `request_id` 중복 방지 |
| `GET` | `/api/v1/access-jobs/{job_id}` | 작업과 단계별 상태 조회 |
| `GET` | `/api/v1/access-jobs/by-request/{request_id}` | 요청번호로 작업 조회 |
| `POST` | `/api/v1/access-jobs/{job_id}/cancel` | 대기/실행 중 작업 취소 요청 |
| `POST` | `/api/v1/access-jobs/prepare-group` | 명시적 동의 후 누락 Isolated 그룹·기본 RX 준비, 별도 Job |
| `POST` | `/api/v1/access-jobs/{job_id}/retry` | 실제 실패·부분 성공 권한 Job의 미완료 단계 재실행 |
| `POST` | `/api/v1/project-jobs/preview` | 신규 프로젝트 그룹·폴더 생성 계획 확인 |
| `POST` | `/api/v1/project-jobs/validate` | 그룹·생성 위치·템플릿 구조·기존 폴더 읽기 전용 검증 |
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

외부 PC에서 접속할 서버의 기본 수신 주소와 포트는 다음과 같습니다.

```dotenv
FGT_HOST=0.0.0.0
FGT_PORT=8000
```

위 수신 설정은 Windows 서비스 래퍼가 사용합니다. 외부 PC에서 수동 Uvicorn으로 접속을 받으려면 `python -m uvicorn folder_grant_api.main:app --host 0.0.0.0 --port 8000`처럼 CLI 인자도 지정해야 합니다.

Mock 검증 후 테스트 자원에 실제 작업을 적용할 때만 서버를 중지하고 다음으로 변경합니다.

```dotenv
FGT_EXECUTOR_MODE=powershell
FGT_POWERSHELL_PATH=powershell.exe
FGT_POWERSHELL_TIMEOUT_SECONDS=10800
FGT_TEMPLATE_ROOT=\\LSK_S010\Study folder\_Template
FGT_GROUP_OU_PATH=OU=Group Project Folder,OU=0.Management Object Group,OU=lskglobal,DC=lskglobal,DC=com
```

PowerShell 모드는 기본 API Key인 `change-me`로 시작할 수 없습니다. 사용자 사번·프로젝트 코드·Level2·Level3를 제한된 형식으로 검증하고 서버가 허용된 작업 단계만 PowerShell 명령으로 변환합니다. API에서 임의 PowerShell 문자열은 받지 않습니다.

### 오래된 과제의 장시간 ACL 작업

PowerShell의 기본 실행 제한은 **단계당 10,800초(3시간)**입니다. 오래된 과제의 재귀 ACL 변경이 300초를 넘더라도 이 범위에서는 서버가 기다립니다. 클라이언트의 HTTP 요청 제한(30초)이나 전체 Job 조회 대기(4시간)와는 별도 설정입니다.

기존 서버의 `.env`에 `FGT_POWERSHELL_TIMEOUT_SECONDS=300`이 있으면 코드 업데이트 후에도 그 값이 우선합니다. 다음과 같이 변경하고, 실행 중인 작업이 없는 상태에서 API 서비스를 재시작해야 적용됩니다.

```dotenv
FGT_POWERSHELL_TIMEOUT_SECONDS=10800
```

```powershell
Restart-Service FolderGrantApi
```

수동 Uvicorn 실행이면 서버를 종료하고 다시 실행합니다. `POWERSHELL_TIMEOUT`은 단순한 클라이언트 조회 실패가 아니라 서버가 실행 제한에 도달하여 PowerShell 프로세스를 종료한 결과입니다. 일부 권한이 적용됐을 수 있으며 이후 단계는 실행되지 않습니다. 실행 제한을 늘려도 기존의 실패·부분 성공 Job은 자동으로 재개되거나 성공으로 바뀌지 않습니다.

같은 `request_id`를 다시 등록하면 기존 또는 최신 재실행 Job을 반환하므로 권한을 다시 실행하지 않습니다. 원인과 실제 권한을 확인한 뒤 **실패한 권한 작업 재실행** 절차의 `/retry`를 사용하세요. 정식 Client에서는 실패·부분 성공 항목을 선택하고 **선택 실행**의 재실행 안내를 확인합니다. 임의 새 요청번호로 중복 접수하지 마세요. **BUS 완료 재확인**은 서버 결과를 조회하고 BUS 완료만 처리하며 권한 작업을 재등록하지 않습니다.

## 실제 요청 검증

정식 Client의 **권한 요청 선택 실행**은 `/api/v1/access-jobs/validate`를 먼저 호출하고 실제 검증이 통과한 경우에만 등록·재실행합니다. **신규 프로젝트 선택 실행은 사전 검증 없이 `/api/v1/project-jobs`를 바로 등록**합니다. `/api/v1/project-jobs/validate`는 개발·진단용 API로 유지하며 사용자 실행을 차단하는 선행 조건이 아닙니다. 본문은 해당 Job 등록과 같은 입력 모델을 사용하며 API Key가 필요합니다. 기존 `/preview`는 실행 계획 계산용으로 유지하고 실제 자원을 조회하지 않습니다.

- 권한 요청: AD 사용자, 보안 그룹, 그룹 등록 여부, 프로젝트 Level1·Level2·Level3, Role 대상 하위 폴더와 현재 ACL을 조회합니다. ETC는 그룹만, 종료 과제는 Closed·Archive 후보를 확인합니다.
- 신규 프로젝트 개발용 검증 API: 그룹 존재·생성 OU, 공유 경로, 프로젝트·Study, 새로 복사할 때 필요한 템플릿·Study을 확인합니다. 기본 All 폴더는 별도 검증에서 제외합니다. 새 그룹·프로젝트 폴더가 없으면 생성 예정으로 표시하며 기존 자원 재사용 시 불필요한 템플릿·OU 검사를 생략합니다.
- 검증은 **API 서비스 계정**으로 읽기 전용 수행합니다. AD·ACL 변경, 파일 생성·삭제·복사, Job 등록과 BUS 완료는 하지 않습니다.
- Mock 모드는 실제 조회를 하지 않고 `unknown`과 안내를 반환합니다. PowerShell 모드에서도 검증 경로는 읽기 전용입니다.
- 응답은 `validation_status`(passed / failed / warning / unknown), `checked_at`, `checked_by`, `checks`와 `warnings`를 반환합니다. 각 checks에는 id·label·target·status·message가 있으며 해당 없는 검사는 not_applicable입니다. 대상 없음과 접근·조회 실패를 구분합니다.
- 현재 ACL은 사용자·프로젝트 그룹과 일치하는 ACE를 보여 주는 참고 정보입니다. 중첩 그룹을 포함한 유효 권한 계산이나 실제 변경 권한 보장은 아닙니다. 검증 후 환경이 바뀔 수도 있습니다.

`FGT_VALIDATION_TIMEOUT_SECONDS=20`은 검증 전체 PowerShell 조회의 제한입니다(1~25초). 실제 작업의 단계당 10800초와 별도이며 Client HTTP 30초보다 짧게 제한합니다. 제한 초과·응답 이상은 확인 불가로 처리합니다. 전체 파일 트리를 재귀 검사하지 않습니다.

API와 Client를 함께 업데이트하고 서비스를 재시작하세요. 구버전 서버에는 검증 API가 없어 Client에서 업데이트 안내를 표시합니다. 이 기능은 DB 스키마 변경이 없습니다. 현재 동작은 docs/CURRENT_BEHAVIOR.md를 참고하고, Word·PPT는 작성 당시 기준 자료로 구분합니다. 사용자가 수정한 Office 파일은 자동 재생성하지 않습니다.

## Windows 서비스 등록

API 서버를 로그인 세션의 콘솔 명령으로 계속 실행하지 않고 Windows 부팅 시 자동 시작되는 `FolderGrantApi` 서비스로 등록할 수 있습니다. 서비스 래퍼는 Uvicorn을 별도 프로세스로 실행하고 중지 요청 시 해당 프로세스를 종료하며, 비정상 종료 시 Windows 서비스 복구 설정으로 다시 시작합니다.

먼저 실행 중인 수동 Uvicorn 프로세스를 종료합니다. 그다음 **관리자 권한 PowerShell**에서 `API` 폴더로 이동합니다.

Mock 모드에서는 다음 명령으로 등록할 수 있습니다.

```powershell
cd C:\Folder-Grant-Tool\API
.\service\install_service.ps1 -OpenFirewall
```

PowerShell 실행 모드에서는 LocalSystem 대신 AD 조회, 공유 폴더 접근 및 ACL 변경 권한이 있는 전용 도메인 서비스 계정을 지정해야 합니다.

```powershell
cd C:\Folder-Grant-Tool\API
$credential = Get-Credential 'LSKGLOBAL\서비스계정'
.\service\install_service.ps1 -Credential $credential -OpenFirewall
```

`-OpenFirewall`은 `.env`의 `FGT_PORT`에 대한 인바운드 TCP 규칙을 생성합니다. 방화벽을 별도로 관리한다면 이 옵션을 생략합니다. 설치 스크립트는 다음 작업을 수행합니다.

1. `.[windows-service]` 의존성과 pywin32 설치
2. 서비스 시작 유형을 `Automatic`으로 등록
3. 비정상 종료 시 5초, 15초, 30초 간격 재시작 설정
4. 선택적으로 Windows 방화벽 규칙 추가
5. 서비스 시작

등록 후 상태와 API 응답을 확인합니다.

```powershell
Get-Service FolderGrantApi
Invoke-RestMethod http://127.0.0.1:8000/health
Invoke-RestMethod http://127.0.0.1:8000/ready -Headers @{ 'X-API-Key' = '실제-API-Key' }
```

서비스 제어 및 로그 확인 명령은 다음과 같습니다.

```powershell
Restart-Service FolderGrantApi
Stop-Service FolderGrantApi
Start-Service FolderGrantApi
Get-Content .\logs\service.log -Tail 100
```

`.env`를 변경하면 `Restart-Service FolderGrantApi`를 실행해야 반영됩니다. 서비스 제거 시 `.env`, SQLite DB 및 로그는 삭제되지 않습니다.

```powershell
.\service\uninstall_service.ps1
```

서비스가 시작되지 않으면 먼저 `logs\service.log`와 Windows 이벤트 뷰어의 Application 로그를 확인합니다. PowerShell 모드에서는 서비스 계정에 **서비스로 로그온**, AD 그룹 관리, `\\LSK_S010\Study Folder` 공유 및 NTFS 권한이 모두 있어야 합니다.

## 현재 실행 방식의 제한

현재 Worker는 FastAPI 프로세스의 Background Task로 실행됩니다. 서버 프로세스가 종료되면 실행 중인 작업을 자동 복구하지 않으므로 현장 인수 후 승인한 범위에서 사용하고 실행 중 서비스 재시작을 피해야 합니다. 또한 취소 요청은 작업 단계 사이에서 확인하며, 이미 실행 중인 단일 PowerShell 프로세스를 즉시 종료하지는 않습니다. 중단 없이 작업 복구가 필요한 운영 범위로 확대하려면 별도 Worker와 재시작 복구가 선행되어야 합니다. 현재 범위에서는 접수·중단·장애 조사를 운영 절차로 통제합니다.

## 내부 운영 범위와 후속 개선

정식 BUS Client와 신규 프로젝트 생성 API는 구현되어 있습니다. 제한된 시험 자원에서 먼저 검증하며, 서버 재시작 전 진행 Job과 잔존 PowerShell을 확인해야 합니다. 서비스 자동 재시작은 Job 자동 복구를 의미하지 않습니다.

별도 Worker·재시작 복구, 사용자별 인증·권한, 프로젝트 단위 동시성 제어, 로그 보관·모니터링과 DB 마이그레이션은 후속 개선 항목입니다. 현재 공유 API Key 인증에는 사용자별 권한 분리가 없고 `requested_by`는 요청자가 입력한 값입니다. 기본 HTTP의 전송 보호와 방화벽 접근 범위는 배포 환경에서 별도로 구성해야 합니다.

## 신규 프로젝트 생성

`project-jobs`는 프로젝트 코드와 이름을 받아 다음 단계를 수행합니다.

1. Study AD 보안 그룹 확인 또는 생성
2. `_Template`에서 프로젝트 루트 폴더 복사
3. 복사된 기본 경로의 해석 불가능 SID 정리
4. 프로젝트 루트에 그룹 읽기/실행 권한 적용
5. `Study/All`에 그룹 수정 권한 적용
6. 그룹·폴더·기본 ACL 재검증

그룹과 폴더가 이미 존재하면 재사용하므로 같은 프로젝트에 대해 다시 실행할 수 있습니다. `robocopy` 종료 코드는 `0~7`을 정상 범위로 처리하고 `8 이상`을 실패로 처리합니다. 정식 Client는 이 생성 Job을 바로 등록합니다. 생성 중 각 단계와 마지막 `verify_project_creation`으로 결과를 확인하며, 개발자가 필요할 때만 `/validate`·`/preview`를 별도로 사용합니다.

## 수동 테스트 클라이언트

여러 Level2, Level3, Role 조합을 수동 시험하는 개발 도구가 `test_client/`에 있습니다. BUS 사용자는 [정식 Client](../Client/README.md)를 사용하세요.

```powershell
python -m pip install -e ".[client]"
python -m test_client.main
```

이 개발 도구는 수동 요청 추가, Preview, 다중 Job 등록, 상태 자동 조회, 취소 및 결과 로그를 제공합니다. BUS 연동은 이 도구에 없으며 별도 `Client/`에서 제공합니다.


### 실패한 권한 작업 재실행

`POST /api/v1/access-jobs/{job_id}/retry`에 `{"requested_by":"operator"}`를 보내면 PowerShell 모드의 `failed`·`partially_succeeded` 권한 Job을 새 Job으로 재실행합니다. API Key 인증이 필요하며, 성공·진행 중·취소·Mock·신규 프로젝트 Job은 이 엔드포인트의 대상이 아닙니다.

성공한 단계의 결과는 `reused_from_job_id`와 함께 보존하고 실패한 단계 및 미실행 단계만 다시 실행합니다. 기존 Job은 변경하지 않습니다. 같은 이전 Job에 대한 중복 접수는 하나의 재실행 Job을 반환합니다. 재실행도 실패하면 그 재실행 Job ID를 대상으로 다시 접수할 수 있습니다.

응답의 `original_request_id`는 최초 요청번호입니다. `GET /api/v1/access-jobs/{job_id}`는 지정 이력 그대로, `GET /api/v1/access-jobs/by-request/{request_id}`와 원래 요청의 중복 등록 응답은 최신 재실행 Job을 반환합니다. 이 연결은 DB에 보존되며 기존 DB 테이블·컬럼 변경은 필요하지 않습니다. 서버·클라이언트 모두 업데이트 후 API 서비스를 재시작하세요. 이미 성공한 단계는 재검증하지 않으므로 해당 성공 상태를 외부에서 변경했다면 별도 확인이 필요합니다.

재실행도 기존 in-process BackgroundTasks 실행기를 사용합니다. **서버 프로세스가 실행 도중 종료되어 queued/running 상태로 남은 작업은 이 기능이 자동 복구하지 않습니다.** 남아 있는 PowerShell 프로세스와 실제 권한 상태를 확인하지 않고 중복 실행하지 않도록, 진행 중·불확실한 작업은 재실행 대상으로 취급하지 않습니다. 완료된 재실행 이력의 원래 요청번호 연결은 서버 재시작 후에도 유지됩니다.

## 오픈·운영 인수 기록

2026-10-27 `main b09100d2e6d5` 기준 자동 시험은 **98개 통과·선택 PowerShell 시험 6개 건너뜀**입니다. 선택 시험은 `FGT_TEST_POWERSHELL`에 인터프리터를 지정해야 실행되며 실제 AD·SMB 변경을 대신하지 않습니다.

[오픈·운영 인수인계 Word](../docs/오픈_운영_인수인계.docx)에 서비스 계정·환경·백업·실제 시험·승인·문의·중단 기준을 기록합니다. 비밀값은 문서에 넣지 않습니다. API Key를 가진 호출자의 `requested_by`가 신뢰 가능한 개인 인증 정보는 아니므로 사용자별 권한이 필요한 범위로 확대할 때는 인증 구조를 보완해야 합니다.

업데이트는 접수 중단 → 활성 Job·잔존 PowerShell 확인 → 서비스 중지·일관된 DB/설정/코드 백업 → 의존성·설정 비교 → 서비스 시작 → health·인증 ready·검증·제한된 실제 시험 순서로 진행합니다. 코드·DB 복원만으로 AD·ACL 변경이 되돌아가지는 않습니다.

### Missing Isolated group preparation

`POST /api/v1/access-jobs/validate` adds `group_preparation` (`null` or
`{group_name, root_path, isolated_path}`). It is offered only for PowerShell
progress/grant/Isolated requests with a confirmed missing group and otherwise
successful resource checks, including the configured OU and existing project and
Isolated roots (also for ETC).

`POST /api/v1/access-jobs/prepare-group` uses the access request fields plus required
`confirm_group_creation: true` and a separate preparation `request_id`. It returns
the usual asynchronous JobResponse. The server revalidates readiness and creates a
Security DomainLocal group in the configured OU, copying the Study group's
Description (empty if unavailable), waits 15 seconds, and grants that group's exact
SID RX on the two existing roots without inheritance or recursion. It never creates
folders, adds users, or grants user ACLs. Matching requests replay idempotently;
ordinary access-job request IDs conflict. Preparation jobs cannot use `/retry`.
GET, by-request and cancel work normally. A successful preparation is not a
successful permission request: revalidate and submit the original access request
separately. No database migration is required.
