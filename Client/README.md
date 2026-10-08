# Folder Grant API Client

BUS 조회·완료 처리와 기존 PyQt 화면을 유지하면서, 권한 부여·해제 및 신규 프로젝트 생성은 API 서버에 맡기는 독립 Windows 클라이언트입니다. 기존 프로그램과 `API/`를 수정하거나 import하지 않습니다. 이전 작업의 요약을 기준으로 재구성했으며 이전 커밋의 코드를 복구한 것은 아닙니다.

## 설치와 실행

Python 3.11 이상과 Google Chrome이 필요합니다. 클라이언트 PC에서 AD 관리 도구 또는 PowerShell 권한 작업 환경은 필요하지 않습니다. BUS 접속 및 API 서버 연결이 가능한 PC에서 실행하세요.

```powershell
cd C:\Folder-Grant-Tool\Client
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python ".\Folder Grant API Client.pyw"
```

실행 코드는 `Folder Grant API Client.pyw` 한 파일입니다. `assets/`는 아이콘·로고이며 같은 폴더에 유지하세요. Selenium이 Chrome 드라이버를 자동으로 찾거나 설치하므로, 제한된 네트워크에서는 Chrome에 맞는 드라이버를 미리 준비해야 합니다. 초기 실행 오류를 확인할 때는 위처럼 PowerShell에서 `python`으로 실행하세요.

## 처음 설정

첫 실행의 API 설정 화면에서 아래 값을 입력하고 **연결 확인** 후 저장합니다. 주소나 키는 메인 화면의 **API 설정**에서 다시 변경할 수 있습니다.

| 설정 | 값 |
|---|---|
| API 주소 | 기본값 `http://192.168.1.10:8000`, 실제 서버 주소로 변경 |
| X-API-Key | 서버 `API\.env`의 `FGT_API_KEY` 값 |
| BUS 계정 | 기존 **설정** 화면에서 입력 |

연결 확인은 `/health`와 API Key가 필요한 `/ready`를 함께 호출합니다. API Key는 `C:\FGT\conf\api_client.json`에 평문 저장됩니다. 이 파일을 저장소나 공유 폴더에 넣지 마세요. BUS 계정 저장 기능 및 `C:\FGT\conf\login.json`은 기존 프로그램과 공유하며, 저장을 선택하면 BUS 비밀번호도 해당 파일에 저장됩니다. API Key는 로그에서 마스킹합니다.

로그는 `C:\FGT\Log`, BUS 다운로드는 `C:\FGT\ef`, BUS 디버그 자료는 `C:\FGT\debug`를 사용합니다. 해당 경로에 쓰기 권한이 필요합니다.

## 기능

- **요청 확인**: Selenium으로 BUS 진행·종료 권한 요청을 조회합니다.
- **신규 확인**: BUS 신규 프로젝트 생성 요청을 불러옵니다. 화면의 수동 추가는 API 생성할 요청을 목록에 추가합니다.
- **파일 선택 / 드래그 / 수동 입력**: 기존 Excel(`.xlsx`, HTML 형식 `.xls`)·수동 요청 흐름을 유지합니다.
- **실행**: 선택된 권한 요청을 API에 등록하고 Job ID로 최종 상태를 조회합니다. 진행·종료, 권한 부여·해제, 사번, 프로젝트 코드, Level2·Level3·Role을 서버 계약으로 변환합니다. `26-012` 같은 BUS 표기는 `26012`로 변환합니다.
- **Dry Run**: 실제 Job을 등록하지 않고 해당 access/project Preview만 호출합니다.
- **중지**: 다음 작업 등록을 멈추고 실행 Job 취소를 요청한 뒤 최종 상태를 조회합니다. 서버의 실행 중인 단일 PowerShell 단계는 즉시 종료되지 않을 수 있습니다.
- **완료 처리**: API 실행이 실제 성공한 BUS 요청의 완료만 재시도합니다. API가 성공하고 BUS만 실패했다면 권한 작업을 다시 실행할 필요가 없습니다.
- **테마 / 테이블 / 로그**: 기존 화면 기능을 유지합니다. 실행 중에는 행 편집·삭제 및 파일 드롭을 제한합니다.

## 완료 처리 조건

자동·수동 완료 처리 모두 **BUS에서 조회한 요청**이며 `executor_mode=powershell`, `status=succeeded`일 때만 허용합니다. 신규 프로젝트에도 같은 조건을 적용합니다. 중지 후 자동 완료는 실행하지 않습니다.

| API 결과 | 화면 / BUS 처리 |
|---|---|
| Preview | `DryRun`, BUS 미처리 |
| Mock | `모의완료`, BUS 미처리 |
| 실제 성공 | `API 성공`, BUS 요청만 완료 가능 |
| 부분 성공 / 실패 / 취소 | 결과 및 단계 로그 표시, BUS 미처리 |
| API 조회 오류 | 등록 결과 확인 필요 표시, BUS 미처리 |

직접 연 Excel 파일과 수동 입력은 `source=test-client`이고 BUS 자동·수동 완료 대상에서 제외됩니다. BUS 조회 행을 편집하면 별도의 수동 요청으로 전환되고 이전 성공 결과를 무효화합니다. BUS 완료가 끝난 행은 같은 세션에서 다시 완료 처리하지 않습니다.

신규 프로젝트 생성 화면의 **실제 성공 후 BUS 완료**를 끄면 서버 작업만 수행합니다. 이후 켠 상태에서 동일 항목을 다시 실행하면 보관된 실제 성공 결과를 사용해 BUS 완료만 재시도합니다. Dry Run은 이 경우에도 항상 Preview를 호출합니다.

## 요청번호와 재시도

BUS 원본 요청번호가 다운로드에 포함되면 이를 사용합니다. 없으면 원본 행의 요청일 등 필드와 대상자·프로젝트·작업·경로·Role의 해시를 사용합니다. 행 순번은 식별에서 제외합니다. 동일 요청을 다시 불러와도 같은 API 요청번호를 사용하고, 등록 응답이 유실되면 `by-request` 조회로 원래 Job을 복구합니다. 임의로 새 요청번호를 만들어 재등록하지 않습니다.

원본 식별자가 없고 모든 필드가 같은 별도 요청은 구분할 수 없습니다. 반대로 원본 행 필드가 변경되면 식별자가 달라질 수 있습니다. BUS 완료 대상 선택은 기존 Selenium의 사번·프로젝트·경로 및 요청 종류 매칭을 유지하므로, 동일 조건의 복수 요청은 실제 BUS에서 확인해야 합니다. 다른 운영자가 동일 요청을 등록해 서버에서 payload 충돌(HTTP 409)이 발생하면 자동으로 무시하지 않습니다.

수동 요청번호는 항목별로 새로 생성되며 같은 화면 항목의 재시도에서는 유지됩니다. 앱을 재시작해서 다시 수동 입력하면 별도 요청이 됩니다. Mock 작업도 서버 DB에서 요청번호를 차지하므로 동일 BUS 요청을 실제 실행하려면 Mock 검증용 서버·DB와 실제 실행 서버·DB를 분리하세요.

## 검증 순서

1. 서버를 Mock 모드로 실행하고 API 연결·인증을 확인합니다.
2. 수동 권한 요청 1건을 추가하고 Dry Run으로 서버 그룹·경로·단계 결과를 확인합니다.
3. Dry Run을 해제해 Mock Job 결과와 SQLite 이력을 확인합니다. BUS 완료는 실행되지 않아야 합니다.
4. 신규 프로젝트 화면의 수동 추가·Dry Run·Mock 결과도 확인합니다.
5. BUS 조회로 진행·종료·신규 요청을 불러오고 대상 행과 원본 BUS 정보를 비교합니다.
6. 테스트용 AD·폴더 자원이 준비된 별도 서버·DB에서 PowerShell 모드를 사용합니다. 자동 완료를 끈 상태로 1건 실행합니다.
7. 서버 결과 및 실제 그룹·폴더 권한을 확인한 후 BUS 완료만 실행합니다.
8. 검증된 테스트 요청으로 자동 완료·중지·취소·BUS 실패 후 완료 재시도를 확인합니다.

서버가 `FGT_EXECUTOR_MODE=powershell`이면 Dry Run을 해제한 실행은 실제 AD·폴더 ACL을 변경합니다.

## 자동 테스트

저장소 루트에서 클라이언트 가상환경으로 실행합니다. 실제 BUS·AD 작업은 호출하지 않습니다.

```powershell
.\Client\.venv\Scripts\python.exe -m unittest discover -s Client\tests -v
.\Client\.venv\Scripts\python.exe -m py_compile ".\Client\Folder Grant API Client.pyw"
```

계약 변환, Preview·Mock 완료 차단, 실제 성공의 완료 허용, 요청 식별, 모의 HTTP 서버를 통한 오류·복구·취소 및 offscreen Qt 작업 스레드 수명을 검사합니다. Linux 환경에서는 `QT_QPA_PLATFORM=offscreen`을 사용할 수 있습니다.

Windows GUI·Chrome·실제 BUS 사이트는 해당 PC에서 확인해야 합니다. 서버의 Background Task 및 재시작 시 작업 복구 제한은 그대로 유지됩니다. 작업 조회 오류가 발생하면 로그의 request_id/job_id로 서버 이력을 확인하세요. 창 닫기는 API 및 BUS 스레드 종료를 기다리며, 실행 중인 서버 작업이나 브라우저 응답이 늦으면 종료도 늦어질 수 있습니다.
