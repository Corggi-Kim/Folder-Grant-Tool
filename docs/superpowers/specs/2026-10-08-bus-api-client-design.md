# BUS API 클라이언트 재구성 설계

## 목적과 산출물
이전 작업 요약에 따라 저장소 루트의 `Client/`에 독립 클라이언트를 재구성한다. 기존 `Folder Grant Tool_v3.3.5.pyw`와 `API/`는 수정하지 않는다. 이전 커밋 `5c08be0`은 로컬 및 GitHub에서 확인되지 않았으므로 원본 복구가 아닌 기존 코드 기반 재구성이다.

산출물은 `Client/Folder Grant API Client.pyw`, `Client/requirements.txt`, `Client/README.md`, `Client/assets/`이다. 실행 코드는 단일 파일이며 API 서버 패키지를 import하지 않는다. PyQt5, Selenium, openpyxl을 사용하고 HTTP 호출은 표준 라이브러리로 구현한다.

## 유지할 사용자 흐름
기존 프로그램의 진행·종료 BUS 권한 요청 조회, 신규 프로젝트 요청 조회, Excel 불러오기, 수동 입력, 테이블, 로그, BUS 계정 설정, 테마를 유지한다. 원본의 UI 및 Selenium 동작을 복사해 재사용하고, 권한 실행 경로를 API 호출로 교체한다. `test_client`의 HTTP 오류 처리와 Qt 작업 스레드 패턴을 재사용한다.

## 서버 호출과 상태
권한 작업은 `/api/v1/access-jobs`, 신규 프로젝트는 `/api/v1/project-jobs`로 등록한다. 진행·종료, grant·revoke, 사번, 프로젝트 코드, Level2·Level3·Role 및 요청자를 서버 계약으로 변환한다. 종료 과제의 Level2·Level3·Role은 null이다. BUS 요청은 source=bus, 수동 및 파일 입력은 source=test-client로 구분한다.

Dry Run은 해당 종류의 preview만 호출한다. 실제 Job은 job_id로 최종 상태까지 조회하며 simulated, succeeded, partially_succeeded, failed, cancelled를 구분한다. 등록 응답 및 단계별 오류를 로그에 남긴다. HTTP 오류, 인증 실패, 연결 실패, 시간 초과는 작업 실패 또는 조회 실패로 표시하며 성공으로 취급하지 않는다.

API 통신은 Qt 작업 스레드에서 실행한다. 중지 시 새 작업 등록을 중단하고 등록된 실행 작업에 취소를 요청한다. 취소 요청은 서버 작업의 즉시 종료를 의미하지 않으므로 최종 상태를 조회한다. 창 종료는 실행 중인 스레드를 강제 파괴하지 않도록 처리한다.

## 요청 식별과 재시도
같은 요청의 등록 재시도에는 같은 request_id를 사용한다. BUS 원본 요청 식별자가 있는 경우 그 식별자와 항목별 구분값을 사용한다. 식별자가 없는 경우 요청 내용과 BUS 매칭 정보를 사용한 안정적인 식별 규칙을 구현하고, 충돌 가능성 및 한계를 README에 명시한다. 수동 요청은 독립 ID를 생성한다. 등록 응답을 잃은 경우 by-request 조회로 복구하여 중복 등록을 피한다.

## BUS 완료 처리 조건
자동 및 수동 완료 처리 모두 executor_mode=powershell이고 status=succeeded인 서버 작업에만 허용한다. Preview, simulated, 부분 성공, 실패, 취소는 BUS 완료 대상으로 넘기지 않는다. 신규 프로젝트도 같은 조건을 적용한다. API 성공 후 BUS 완료만 실패했다면 권한 작업을 다시 등록하지 않고 BUS 완료만 재시도할 수 있어야 한다. BUS에 대응되는 요청이 없는 수동 입력은 자동 완료 대상에서 제외한다.

## 설정
기본 API 주소는 `http://192.168.1.10:8000`이다. 처음 실행 시 주소 및 X-API-Key 입력 화면을 제공한다. API 설정은 기존 요약과 동일하게 `C:\FGT\conf\api_client.json`을 사용한다. API Key는 설정 파일에 저장되므로 README에 저장 위치를 명시하고 로그에는 출력하지 않는다. 설정 화면을 다시 열어 주소 및 키를 변경할 수 있어야 한다. API 연결 확인은 health와 인증된 ready를 사용한다.

## 검증과 제한
자동 검증은 단일 파일 구문 검사, API payload 변환, Mock·Preview 완료 차단, 실제 성공의 완료 허용, 오류·취소·중복 요청 복구를 대상으로 한다. 모의 HTTP 서버와 가짜 BUS 동작으로 검증하며 실제 AD·ACL·BUS 변경은 실행하지 않는다. 가능하면 offscreen Qt로 화면 초기화 및 작업 스레드 실행을 확인한다.

Windows GUI, Chrome 및 실제 BUS 페이지 연동은 이 Linux 환경에서 완전 검증할 수 없으므로 README에 수동 확인 절차를 제공한다. 기존 서버의 Background Task 및 재시작 복구 제한도 유지된다. 코드 검증 완료와 실제 업무 환경 검증 완료를 구분해서 보고한다.
