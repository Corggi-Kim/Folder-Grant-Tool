# BUS API Client Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans or superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox syntax for tracking.

**Goal:** 기존 BUS 프로그램을 기반으로 서버 API에 권한 작업을 맡기는 독립 Client 클라이언트를 만든다.

**Architecture:** 기존 단일 파일의 UI와 Selenium 흐름을 복사하고 로컬 PowerShell 실행 경로를 제거한다. 파일 안의 API 통신 계층, 작업 스레드, UI 연결부를 구분하고 API 결과가 BUS 완료를 허용하는 유일한 근거가 되도록 한다.

**Tech Stack:** Python 3.11+, PyQt5, Selenium, openpyxl, urllib, unittest.

**Spec:** `docs/superpowers/specs/2026-10-08-bus-api-client-design.md`

## Global Constraints

- `Client/Folder Grant API Client.pyw` 단일 실행 파일이며 API 서버 패키지를 import하지 않는다.
- 기존 `Folder Grant Tool_v3.3.5.pyw`와 `API/`는 수정하지 않는다.
- 기본 주소는 `http://192.168.1.10:8000`, 설정 경로는 `C:\FGT\conf\api_client.json`이다.
- HTTP 통신은 표준 라이브러리, UI는 PyQt5, BUS는 Selenium, Excel은 openpyxl을 사용한다.
- 실제 AD·ACL·BUS 변경 없이 검증한다.
- 자동 및 수동 BUS 완료는 executor_mode=powershell, status=succeeded인 연결된 BUS 요청만 허용한다.

## Review Focus

- 동일 BUS 요청 재조회 및 등록 응답 유실: 동일 ID로 작업 복구, 중복 실행 방지.
- 엑셀·수동 입력을 BUS 요청으로 오인: 서버 성공이어도 BUS 완료 금지.
- 중지와 서버 성공이 교차: 취소 결과 조회, 중지 후 자동 BUS 완료 금지.
- 신규 프로젝트 Mock 성공: 폴더 생성 완료 및 BUS 완료로 표시하지 않기.
- 창 종료 중 네트워크 지연: 실행 스레드 파괴 및 GUI 멈춤 방지.

---

### Task 1: 독립 파일과 API 계약

**Files:** Create `Client/Folder Grant API Client.pyw`, `Client/requirements.txt`, `Client/assets/`, `Client/tests/test_api_contract.py`.

**Interfaces:**
- `FolderGrantApiClient(base_url: str, api_key: str, timeout_seconds: int = 30)`
- `request(method: str, path: str, payload: dict | None = None) -> dict`
- `make_access_payload(values: dict, request_id: str, source: str, requested_by: str) -> dict`
- `make_project_payload(project_code: str, project_name: str, request_id: str, source: str, requested_by: str) -> dict`
- `can_complete_bus(result: dict, source: str, stopped: bool = False) -> bool`
- `make_bus_request_id(values: dict, bus_identity: str = '') -> str`

- [x] 테스트 작성: progress grant/revoke, closed null 필드, project payload를 서버 스키마와 대조한다. BUS ID는 재조회 시 동일하고 수동 ID는 독립이어야 한다. `can_complete_bus`는 simulated/preview/부분 성공/실패/취소/수동/중지에서 False, BUS 실제 성공에서만 True다.
- [x] `python -m unittest discover -s Client/tests -v` 실행 후 누락 구현에 따른 실패 확인.
- [x] 원본을 Client 파일에 복사하고 assets 및 의존성을 배치한다. HTTP 클라이언트는 test_client 패턴에서 가져오되 access/project 종류를 모두 지원하고 오류에 API Key를 포함하지 않는다. 단일 파일 내부에 위 함수를 구현한다. 테스트는 GUI 시작 없이 import한다.
- [x] 위 명령을 재실행하고 Task 1 테스트 통과 확인.
- [x] 변경 파일을 커밋한다.

### Task 2: Job 스레드와 복구·취소

**Files:** Modify `Client/Folder Grant API Client.pyw`; Create `Client/tests/test_job_worker.py`.

**Interfaces:**
- `ApiJobWorker(client: FolderGrantApiClient, kind: str, payload: dict, dry_run: bool = False)`는 QObject이며 progress(dict), succeeded(dict), failed(str), finished() signal을 제공한다.
- `ApiJobWorker.stop() -> None`은 스레드 간 중지 플래그를 설정한다.
- `ApiJobWorker.run() -> None`은 preview 또는 등록→조회→최종 결과를 실행한다.

- [x] 모의 HTTP 서버 테스트 작성: Dry Run은 preview만 호출; 등록 응답 유실은 by-request로 복구; 중지 시 cancel 호출 후 최종 상태 조회; 조회 오류는 성공 signal을 발생시키지 않음; 401 및 JSON 오류 전달. 네트워크 지연 중 stop 호출이 UI 스레드를 기다리게 하지 않는지도 확인한다.
- [x] unittest 실행 후 Task 2 실패 확인.
- [x] access-jobs/project-jobs 공통 경로를 구현한다. 대기에는 중지 가능한 이벤트를 사용하고 HTTP timeout을 유지한다. 최종 상태와 executor_mode를 보존해 UI에 전달한다.
- [x] 모의 HTTP 서버 테스트 통과 확인.
- [x] 변경 파일을 커밋한다.

### Task 3: 기존 UI와 BUS 흐름 연결

**Files:** Modify `Client/Folder Grant API Client.pyw`; Create `Client/tests/test_ui_flow.py`.

**Interfaces:**
- `ApiSettingsDialog`는 주소·키 입력, health/ready 확인, 설정 저장을 제공한다.
- `AccessManager.run_execute()`는 선택 행별 payload 및 ApiJobWorker를 만든다.
- `AccessManager.run_complete()`는 Task 1의 완료 조건을 다시 확인한다.
- 기존 `CreateWorker`는 Task 2 worker를 사용하고 실제 성공일 때만 `_bus_click_process`를 호출한다.

- [x] 가짜 BUS 및 offscreen Qt 테스트 작성: BUS 실제 성공만 자동·수동 완료 허용; Mock 프로젝트 성공 차단; 수동·Excel 완료 차단; API 성공 후 BUS 실패는 BUS만 재시도; 중지 후 성공 응답 차단; 실행 중 closeEvent는 스레드 종료 전 파괴하지 않음.
- [x] unittest 실행 후 UI 흐름 실패 확인.
- [x] 기존 테이블 행에 source, stable request_id, API result 및 BUS matching 정보를 보관한다. BUS 다운로드 콜백과 일반 파일 입력을 구분하고 기존 Selenium 세부 동작을 유지한다. 설정 시작 화면 및 다시 열기 기능을 연결한다. 기존 권한·신규 프로젝트 로컬 PowerShell 실행 경로를 제거하고 DryRun 수동 완료 허용도 제거한다.
- [x] `QT_QPA_PLATFORM=offscreen python -m unittest discover -s Client/tests -v`로 계약·worker·UI 검증 통과 확인.
- [x] 변경 파일을 커밋한다.

### Task 4: 문서와 최종 확인

**Files:** Create `Client/README.md`, `Client/.gitignore`; Modify 테스트가 발견한 결함의 해당 Client 파일.

- [x] README에 Windows 설치·실행, API 설정 파일과 키 저장, BUS 계정 설정, Preview→Mock→테스트 자원 실제 작업→BUS 완료 순서를 적는다. 안정적 BUS 식별자 부재 시 동일 내용의 별도 요청 구분 한계와 서버 재시작 복구 제한을 명시한다.
- [x] `.venv`, `__pycache__`, 로컬 설정·다운로드·로그를 제외한다.
- [x] 전체 Client unittest 및 `py_compile` 실행. 원본 파일과 API에 변경이 없는지 git diff로 확인하고 로컬 PowerShell 실행 연결이 남지 않았는지 검사한다.
- [x] 코드 검토에서 API 성공의 잘못된 판정, QObject 스레드 수명, BUS 매칭 및 민감정보 로그를 집중 확인한다. Windows·실제 BUS 미검증 사항을 명시한다.
- [x] 수정이 있으면 관련 검증만 다시 실행하고 최종 변경을 커밋한다.

## 실행 방식 제안

단일 파일의 기존 UI·worker 간 결합이 높으므로 현재 세션에서 직접 순서대로 구현하는 방식을 권장한다. 실행 전 사용자가 계획과 실행 방식을 확인한다.
