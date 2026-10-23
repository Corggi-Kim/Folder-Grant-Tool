"""Reviewed Korean content for the beta handover documents. No credentials."""
DATE = '2026-10-23'
BASE = '26f3f761d4f1d243f2a2aa14ec8470ec03c369ec'
API_PAGES = [
('문서 활용과 서비스 개요', [
 ('p','이 문서는 API 서버 설치 담당자와 유지보수 개발자를 위한 베타 기준서입니다. 현재 구현과 운영 권장 사항을 구분하며, 실제 운영 환경의 검증 완료를 의미하지 않습니다.'),
 ('table',(['항목','기준'],[
 ['프로그램','Folder Grant API / Python 패키지 0.4.0'],['구현 기준','main 26f3f76 / 2026-10-23'],['담당 범위','권한 부여·해제, 신규 프로젝트 생성, 작업 이력·단계 결과 관리'],['주요 기술','Python 3.11+, FastAPI, Pydantic, SQLAlchemy, SQLite, Uvicorn'],['실행 환경','Mock: 개발 환경 / 실제 변경: Windows + AD·공유 폴더 접근 환경'],['읽는 순서','설계·구조 → 계약·설정 → 설치·운영 → 장애·변경 관리']])),
 ('h','베타에서 제공하는 가치'),
 ('p','BUS 조회와 완료 처리는 Client에 두고, AD 그룹과 폴더 권한 변경은 API 서버에 모읍니다. 입력 검증, 실행 계획, Job 상태와 단계별 결과를 공통 기준으로 관리합니다.'),
 ('callout',('범위 구분','API 서버는 BUS 로그인이나 완료 버튼 클릭을 수행하지 않습니다. BUS 완료 여부는 Client가 실제 서버 성공을 확인한 후 별도로 처리합니다.')),
 ('h','주요 장과 찾아볼 내용'),
 ('p','설계·코드 지도: 3–4쪽 / 권한·프로젝트 정책: 5–6쪽 / API 계약: 7–8쪽 / 설정·배포: 9–10쪽 / 운영·장애: 11–12쪽 / 유지보수·베타 점검: 13–14쪽'),
]),
('전체 설계와 처리 흐름', [
 ('image',('architecture.png',5.9)),
 ('h','요청 한 건의 처리 과정'),
 ('steps',['Client가 X-API-Key와 업무 입력값을 전송합니다.','Pydantic과 도메인 로직이 형식·정책을 검증하고 실행 계획을 만듭니다.','Job과 단계들을 DB에 기록하고, 새 요청만 BackgroundTasks에 등록합니다.','Worker가 Mock 또는 PowerShell 모드로 단계를 순서대로 처리합니다.','Client는 Job을 조회하고, 실제 성공한 BUS 원본 요청만 완료 처리합니다.']),
 ('callout',('설계상 경계','HTTP 접수 성공은 권한 작업 성공을 뜻하지 않습니다. 최종 Job 결과와 steps를 확인해야 하며, Preview는 대상 폴더의 실재 여부나 서비스 계정 권한까지 검증하지 않습니다.')),
]),
('코드 구조와 변경 지점', [
 ('table',(['경로 · API/src/folder_grant_api/ 기준','책임'],[
 ['main.py / settings.py','앱 수명 주기, 라우터 등록, 환경 설정과 실행 안전 검사'],['api/routes/ · api/security.py','HTTP 계약, API Key 검증, 응답 변환'],['schemas/','권한·프로젝트 요청 및 결과 모델'],['domain/operation_plan.py','진행·종료, 부여·해제의 정책과 단계 구성'],['domain/project_plan.py','신규 프로젝트 생성 6단계 구성'],['domain/constants.py · project_codes.py · paths.py','Role, 레거시 예외, 코드·그룹·경로 변환'],['services/','등록·중복·취소·명시적 재실행 정책'],['infrastructure/','DB, Repository, PowerShell 명령 작성과 실행'],['workers/','실행 모드 선택, 단계 실행·상태 갱신'],['windows_service.py / API/service/','서비스 실행 래퍼와 설치·제거 스크립트']])),
 ('h','이력을 저장하는 네 테이블'),
 ('p','jobs는 요청·최종 결과·취소 여부, job_steps는 순서·대상·단계 결과, audit_logs는 접수·취소·실패 등 이벤트, service_metadata는 메타데이터용 테이블입니다. service_metadata의 별도 운영 절차는 현재 제공하지 않습니다.'),
 ('callout',('DB 변경 주의','시작 시 create_all()로 테이블을 생성하지만 기존 컬럼을 바꾸는 마이그레이션 도구는 없습니다. 필드 추가·변경 시 백업·변환·롤백 절차를 별도로 설계하세요.')),
]),
('권한 처리 정책과 프로젝트 코드', [
 ('table',(['요청 조건','생성되는 주요 작업'],[
 ['진행 · 부여','프로젝트 AD 그룹 가입 후 대상 폴더 ACL 적용. Role·정책에 따라 하위 폴더 처리'],['진행 · 해제','프로젝트 AD 그룹 탈퇴 후 대상 폴더와 하위 항목의 사용자 ACL 제거'],['Level3 = ETC','그룹 가입·탈퇴만 수행하며 별도 폴더 ACL 단계 없음'],['종료 · 부여','Closed / Archive 후보 중 존재하는 경로에 읽기·실행 권한 부여'],['종료 · 해제','선택된 종료 경로의 사용자 ACL 제거; AD 그룹 변경 단계 없음']])),
 ('h','코드 표기와 경로는 다릅니다'),
 ('p','API에는 4~5자리 숫자와 허용된 접미사 형식을 보냅니다. Client의 BUS 표기 26-012는 26012로 변환됩니다. 1987은 폴더 세그먼트 19087, 그룹 LSK 19-087로 해석합니다. API에 26-012를 그대로 전달하지 마세요.'),
 ('h','정책 경계값을 유지하세요'),
 ('p','신규 템플릿 기준은 25069, STAT_IDMC 신규 정책 기준은 26012이며 강제 적용 예외 목록이 있습니다. 신형 Study Role은 지정 하위 폴더, 레거시는 번호 패턴, Isolated는 별도 Role 정책을 사용합니다. 제외 목록은 새 권한 부여 대상을 고르는 조건이며 이미 존재하는 권한을 자동 회수하는 기능이 아닙니다.'),
 ('callout',('정책 변경 시 확인','constants.py와 operation_plan.py, PowerShell 명령, 단위 테스트를 함께 확인하세요. Client에도 선택지·변환 코드가 남아 있으므로 서버와 화면 계약을 함께 검증해야 합니다.')),
]),
('신규 프로젝트 생성 설계', [
 ('steps',['ensure_ad_group: Study AD 보안 그룹을 확인하거나 생성합니다.','ensure_project_folder: 서버 템플릿에서 프로젝트 루트 폴더를 확인·생성합니다.','cleanup_unknown_acl: 프로젝트 루트와 Study/All의 해석 불가능 SID를 정리합니다.','grant_project_root_acl: 프로젝트 루트에 그룹 읽기·실행 권한을 보장합니다.','grant_study_all_acl: Study/All에 그룹 수정 권한을 보장합니다.','verify_project_creation: 그룹·폴더·기본 ACL을 재확인합니다.']),
 ('h','기존 자원을 사용하는 방식'),
 ('p','이미 존재하는 그룹과 폴더는 재사용합니다. 템플릿 복사에 쓰는 robocopy의 0~7 종료 코드는 정상 범위, 8 이상은 실패로 처리합니다. 같은 요청번호의 재등록은 재실행하지 않고 기존 Job을 반환합니다.'),
 ('table',(['확인 항목','운영 시 판단'],[
 ['템플릿 변경','신규 생성 결과가 달라지므로 배포 전 별도 시험 프로젝트에서 검증'],['폴더와 그룹 이름','코드 변환·OU·공유 경로의 실제 결과를 Preview로 먼저 검토'],['프로젝트 Job 실패','권한 Job용 /retry를 사용할 수 없음. 실제 생성 상태 확인 후 담당자가 복구 방식을 결정'],['검증 실패','자동 롤백되지 않으므로 부분 생성된 그룹·폴더를 확인']])),
 ('callout',('권장','프로젝트 폴더를 수동 정리하거나 새 요청으로 다시 접수하기 전, 기존 Job과 실제 자원을 비교해 중복 생성·기존 권한 변경의 영향을 확인하세요.')),
]),
('API 엔드포인트와 상태', [
 ('table',(['Method · 경로','용도'],[
 ['GET /health','프로세스 응답 확인; API Key 불필요'],['GET /ready','인증과 DB SELECT 1 확인; AD·SMB 준비 상태는 검사하지 않음'],['POST /api/v1/access-jobs/preview','권한 실행 계획만 생성'],['POST /api/v1/access-jobs','권한 Job 등록: 신규 202 / 동일 요청 재응답 200'],['GET /api/v1/access-jobs/{job_id}','지정된 Job 이력 조회'],['GET /api/v1/access-jobs/by-request/{request_id}','요청번호로 최신 재실행 결과 조회'],['POST /api/v1/access-jobs/{job_id}/cancel','진행·대기 작업의 취소 요청'],['POST /api/v1/access-jobs/{job_id}/retry','실패·부분 성공한 실제 권한 Job 재실행'],['/api/v1/project-jobs','POST 등록, POST /preview, GET /{job_id}, GET /by-request/{request_id}, POST /{job_id}/cancel; /retry 없음']])),
 ('h','상태 흐름'),
 ('p','queued → running → succeeded / partially_succeeded / failed. 취소 요청 시 cancel_requested를 거쳐 cancelled가 될 수 있습니다. Mock은 최종 simulated입니다. Preview는 Job을 만들지 않으며 Client에서 preview로 표시합니다.'),
 ('callout',('오류 응답','401: 인증 / 404: 작업 없음 / 409: 요청 충돌·취소·재시도 불가 / 422: 입력·정책 오류 / 503: ready DB 오류. 처리 전 status뿐 아니라 executor_mode와 steps를 함께 확인하세요.')),
]),
('요청 계약·중복 방지·재실행', [
 ('code','{\n  "request_id": "BETA-ACCESS-001", "source": "swagger",\n  "operation": "grant", "project_status": "progress",\n  "employee_id": "TEST001", "project_code": "26012",\n  "level2": "Study", "level3": "STAT",\n  "role": "Manager", "requested_by": "beta-operator"\n}'),
 ('p','위 값은 형식 예시입니다. TEST001은 실재 계정이 아니므로 실제 실행용 값으로 사용하지 마세요. 프로젝트 요청은 request_id, source, project_code, project_name(선택), requested_by를 사용합니다. 취소·재실행 본문은 {"requested_by":"beta-operator"}입니다.'),
 ('table',(['구분','계약'],[
 ['입력 제한','source: test-client / bus / swagger. extra 필드는 거부. 사번은 영문·숫자·._-'],['Level2 / Level3','Study·Isolated / ARS, CO, DM, ER, MW, PM, PV, RA, SSU, STAT, STAT_IDMC, ETC'],['동일 request_id + 동일 입력','기존 또는 최신 재실행 Job 반환; 재실행 안 함'],['동일 request_id + 다른 입력','HTTP 409 REQUEST_ID_CONFLICT; requested_by 변경도 충돌 가능'],['명시적 /retry','PowerShell failed·partially_succeeded 권한 Job만 허용. 성공 단계 복사, 나머지 재실행'],['재시도 이력','새 request_id는 RETRY-<부모 Job UUID>. original_request_id로 최초 요청 연결']])),
 ('callout',('재실행의 한계','부모 Job은 보존하며, 복사된 성공 단계는 다시 검증하지 않습니다. 외부에서 그룹·권한을 바꿨다면 실제 상태를 먼저 확인하세요. 진행 중·취소·Mock·프로젝트 Job에는 권한 재실행 API를 사용하지 않습니다.')),
]),
('서버 설정과 보안 경계', [
 ('table',(['환경 변수','기본값 · 의미'],[
 ['FGT_ENVIRONMENT','development / 환경 표시'],['FGT_DATABASE_URL','sqlite:///./data/folder_grant.db / API 작업 폴더 기준'],['FGT_API_KEY','Settings 기본 change-me / 반드시 별도 긴 임의 값 설정'],['FGT_HOST / FGT_PORT','0.0.0.0 / 8000 / Windows 서비스 수신 설정'],['FGT_EXECUTOR_MODE','mock / powershell은 실제 AD·ACL 변경'],['FGT_POWERSHELL_PATH','powershell.exe / 서버 실행 파일'],['FGT_POWERSHELL_TIMEOUT_SECONDS','10800 / PowerShell 단계당 3시간'],['FGT_TEMPLATE_ROOT / FGT_GROUP_OU_PATH','템플릿 UNC / 그룹 생성 OU'],['FGT_SHARE_ROOT','진행 폴더 형식: {proj_seg}, {lv2}, {lv3}'],['FGT_CLOSED_ROOT / FGT_CLOSED_ARCHIVE_ROOT','종료·Archive 검색 루트']])),
 ('p','.env.example은 설정 항목의 예시입니다. 실제 .env의 키·계정 값을 문서나 저장소에 복사하지 마세요. 변경 후 프로세스 재시작이 필요합니다. 기존 .env의 timeout=300은 코드 기본값을 업데이트해도 유지됩니다.'),
 ('h','현재 보안 모델과 배포 권장 사항'),
 ('p','보호 API는 공유 X-API-Key로 인증합니다. 사용자별 권한 분리·SSO·requested_by의 본인 검증은 구현되지 않았습니다. 기본 HTTP에는 암호화가 없으므로 제한된 내부망에서 사용하고, 필요 시 별도 HTTPS 프록시와 접근 제어를 구성하세요. .env와 DB 접근은 운영 담당자로 제한합니다.'),
]),
('설치와 Windows 서비스 배포', [
 ('h','1. 환경 준비 · 관리자 또는 서버 담당자'),
 ('p','Windows, Python 3.11+, API 포트 접근, AD PowerShell 모듈, icacls·robocopy, 도메인 연결을 준비합니다. 서비스 계정에는 서비스로 로그온, 필요한 AD 그룹 관리와 공유·NTFS 권한이 있어야 합니다.'),
 ('code','cd C:\\Folder-Grant-Tool\\API\npy -m venv .venv\n.\\.venv\\Scripts\\Activate.ps1\npython -m pip install -e ".[dev]"\nCopy-Item .env.example .env\n# .env의 API Key를 교체하고 mock으로 시작\npython -m uvicorn folder_grant_api.main:app --host 0.0.0.0 --port 8000'),
 ('p','위 수동 실행의 host·port는 CLI 인자입니다. FGT_HOST·FGT_PORT를 변경해도 수동 Uvicorn 명령의 인자를 자동 변경하지 않습니다. 서비스 래퍼는 해당 환경 설정을 사용합니다.'),
 ('h','2. Mock 검증 후 서비스 등록'),
 ('code','cd C:\\Folder-Grant-Tool\\API\n# 관리자 PowerShell에서 실행; 수동 Uvicorn은 먼저 종료\n$credential = Get-Credential "DOMAIN\\서비스계정"\n.\\service\\install_service.ps1 -Credential $credential -OpenFirewall\nGet-Service FolderGrantApi'),
 ('p','설치 스크립트는 가상환경과 .env가 필요하며 pywin32 의존성, 자동 시작과 장애 복구를 설정합니다. -OpenFirewall은 인바운드 규칙을 추가하므로 허용 대상 범위는 방화벽 정책으로 별도 제한하세요.'),
 ('callout',('실제 모드 전환','테스트용 별도 DB·자원에서 검증한 뒤 .env의 executor를 powershell로 변경하고 서비스 재시작. Mock DB에 등록된 BUS 요청번호를 실제 실행에 그대로 재사용하지 마세요.')),
]),
('일상 운영·백업·업데이트', [
 ('table',(['주기 · 상황','운영 절차'],[
 ['시작 점검','Get-Service → /health → 인증한 /ready → 서비스 계정의 AD·공유 폴더 접근 점검'],['작업 중','job_id와 request_id, 단계별 상태·error_message 확인. 같은 프로젝트 작업은 운영자가 직렬 처리 권장'],['매일 또는 운영 기준 주기','SQLite·서비스 로그·클라이언트 로그 용량 점검, 실패·장시간 대기 이력 확인'],['업데이트 전','새 접수 중지, 활성 Job·PowerShell 종료 확인, 서비스 중지, 코드·설정·DB·로그 백업'],['업데이트 후','의존성 설치, 설정 비교, 서비스 시작, ready와 Preview 및 제한된 시험 검증'],['복구·롤백','코드만 되돌려도 AD·ACL은 복구되지 않음. DB와 실제 자원 상태를 맞춰 판단']])),
 ('code','Get-Service FolderGrantApi\nInvoke-RestMethod http://127.0.0.1:8000/health\n$headers = @{ "X-API-Key" = "<실제 키를 로컬에서 입력>" }\nInvoke-RestMethod http://127.0.0.1:8000/ready -Headers $headers\nGet-Content .\\logs\\service.log -Tail 100'),
 ('h','백업 대상과 보관 원칙'),
 ('p','API/.env, API/data/folder_grant.db, API/logs/service.log 및 배포 코드·의존성 버전을 보관합니다. DB는 작업 종료 후 서비스 중지 상태에서 복사하거나 SQLite 일관성 백업 기능을 사용하세요. 실제 AD·ACL 복원 수단은 조직 운영 절차로 별도 준비합니다.'),
 ('callout',('자동화되지 않은 부분','작업 자동 복구, 자동 롤백, 로그 순환, 보관 기간 정리와 프로젝트 단위 실행 잠금은 제공되지 않습니다. 대량 동시 요청을 전제로 운영하지 마세요.')),
]),
('장애 대응과 결과 해석', [
 ('table',(['증상 · 코드','확인과 조치'],[
 ['연결 실패 / 401','서비스·포트·주소 확인 후 서버와 Client 키 일치 여부 확인'],['422 / INVALID_ROLE','코드 형식·Level2·Level3·Role을 Preview로 검증'],['POWERSHELL_FAILED / 파일 없음','실패 step.target의 서버 UNC를 서비스 계정으로 확인. 폴더 복구 후 권한 Job 재실행'],['POWERSHELL_TIMEOUT','단계 제한 초과. 기존 .env의 300 여부 확인. 외부 상태 확인 후 필요 시 제한 변경·재시도'],['부분 성공','성공 단계와 실패 단계 모두 확인. 자동 롤백·BUS 완료 없음'],['수정했는데 같은 결과','동일 request_id가 기존 결과를 반환하는 정상 동작인지 replayed 확인. 명시적 /retry 사용'],['queued / running이 계속됨','서버 종료·재시작 이력, 잔존 PowerShell 확인. 자동 복구 대상이 아니므로 수동 조사'],['신규 프로젝트 실패','템플릿·OU·SID·기본 경로·robocopy 결과 확인. 권한 전용 /retry 사용 금지']])),
 ('h','추적에 필요한 최소 자료'),
 ('p','발생 시각·서버/Client 버전·request_id·job_id·executor_mode·status·실패한 steps의 target/error_message를 모읍니다. Client 상세 로그와 API Job·audit_logs를 연결하세요. 서버 시각은 UTC 기반, Client 파일 로그는 PC 로컬 시각이므로 시간대를 맞춰 비교합니다.'),
 ('callout',('외부 변경 판단','changed_external_system=true는 정확한 변경 목록이나 복구 완료 증명이 아닙니다. 실패·취소 중에도 외부 변경 가능성을 보수적으로 나타내므로 실제 그룹과 ACL을 점검해야 합니다.')),
]),
('유지보수와 변경 검증', [
 ('table',(['변경 종류','함께 검토할 부분'],[
 ['입력 필드·응답 변경','schemas, routes 직렬화, Client payload·응답 처리, test_client'],['권한·Role 정책 변경','constants, operation_plan, powershell_commands, Client 선택지·변환'],['프로젝트 생성 변경','project_plan, 템플릿·OU·검증 명령, 기존 자원 재사용'],['중복·재실행 변경','JobService·Repository, request_payload 연결, 원본 요청·BUS 완료 일치'],['실행·취소 변경','workers, PowerShellRunner, 부분 성공·취소·시간 초과 처리'],['DB 모델 변경','create_all 한계, 데이터 마이그레이션·백업·구버전 호환성'],['서비스 변경','작업 폴더, .env 로드, 실행 계정·포트, 시작·중지 로그']])),
 ('h','개발자 확인 명령'),
 ('code','cd C:\\Folder-Grant-Tool\\API\n.\\.venv\\Scripts\\python.exe -m pytest tests -q'),
 ('p','문서 작성 시 API 자동 테스트 60개가 통과했습니다. 테스트는 계획·검증·API 계약·실패·재실행·PowerShell 명령 구성을 검사하며 실제 Windows AD·SMB 변경 성공을 증명하지는 않습니다.'),
 ('h','릴리스 시 버전과 호환성'),
 ('p','API/pyproject.toml과 패키지 __version__을 함께 관리하세요. /api/v1 경로는 현재 계약입니다. 실패 권한 재실행 기능은 API와 Client가 모두 지원해야 하므로 함께 배포하고 서비스 재시작 후 확인하세요.'),
]),
('베타 오픈 점검과 현재 제한', [
 ('table',(['담당','오픈 전 완료 조건'],[
 ['서버 담당','전용 실행 계정·접근 범위·키·DB·로그 백업 위치 확정'],['서버 담당','Mock과 실제 검증 환경/DB 분리, 서비스 재시작 절차 점검'],['업무 담당','시험 대상의 프로젝트·사번·폴더·Role·완료 처리 범위 확인'],['공동 검증','진행/종료 × 부여/해제, 레거시·신규·STAT_IDMC·ETC 사례 점검'],['공동 검증','신규 프로젝트 생성·부분 실패·취소·복구 후 권한 재실행 확인'],['업무 담당','자동 완료 OFF로 실제 상태 검증 후 BUS 수동 완료·자동 완료 확인'],['운영 담당','장애 접수 담당·중단 기준·피드백 채널·로그 보관 기준 지정']])),
 ('callout',('현재 베타 제한','Worker는 API 프로세스 안의 BackgroundTasks입니다. 재시작 후 작업 자동 재개·전용 작업 큐·사용자별 권한·자동 복구는 없습니다. Windows 서비스의 자동 재시작은 Job 복구를 의미하지 않습니다.')),
 ('h','오픈 후 권장 개선 순서'),
 ('p','① 별도 Worker와 재시작 복구 ② 사용자 인증·권한 분리와 HTTPS 운용 ③ 작업 단위 동시성 제어 ④ 로그 보관·모니터링 ⑤ DB 마이그레이션·배포 자동화. 개선 항목은 제안이며 현재 구현된 기능이 아닙니다.'),
 ('h','소스 근거'),
 ('p','API/src/folder_grant_api/{api,schemas,domain,services,infrastructure,workers}, settings.py, windows_service.py, API/service/, API/pyproject.toml, API/tests/. 문서 기준 이후 코드를 변경하면 계약·운영 절차도 함께 갱신하세요.'),
]),
]

CLIENT_PAGES = [
('문서 활용과 프로그램 개요', [
 ('p','이 문서는 베타 사용자, 배포 담당자와 Client 유지보수 개발자가 함께 보는 안내서입니다. 사용자 절차를 먼저 설명하고 뒤에서 내부 구조와 변경 시 유의 사항을 다룹니다.'),
 ('table',(['항목','기준'],[
 ['프로그램','Folder Grant API Client'],['버전 표시','V1.1.0 Build:20261022 / 문서 작성 2026-10-23'],['구현 기준','main 26f3f76'],['담당 범위','BUS 조회·완료 처리, 수동·Excel 입력, API 요청, 결과와 로그 표시'],['권한 실행 위치','API 서버. Client PC는 AD·폴더 권한 PowerShell을 실행하지 않음'],['필수 환경','Windows 사용 PC, Python 3.11+, Chrome, BUS·API 연결'],['읽는 순서','설치·사용 → 결과·장애 → 구조·데이터 → 유지보수·베타 점검']])),
 ('h','두 프로그램의 역할'),
 ('p','Client는 작업을 고르고 결과를 확인하는 화면입니다. API는 실제 그룹·폴더 작업을 처리하는 서버입니다. API가 성공했더라도 BUS 완료 처리까지 성공했는지는 별도로 확인해야 합니다.'),
 ('h','주요 장과 찾아볼 내용'),
 ('p','설치·시작: 3쪽 / 화면: 4쪽 / 권한 요청: 5쪽 / 신규 프로젝트: 6쪽 / 결과·재시도: 7쪽 / 내부 구조: 8쪽 / 식별·저장: 9쪽 / 장애: 10쪽 / 유지보수: 11쪽 / 베타 점검: 12쪽'),
]),
('설치·연결·시작 준비', [
 ('code','cd C:\\Folder-Grant-Tool\\Client\npy -m venv .venv\n.\\.venv\\Scripts\\Activate.ps1\npython -m pip install -r requirements.txt\npython ".\\Folder Grant API Client.pyw"'),
 ('p','실행 파일과 assets/를 함께 유지하세요. 첫 실행은 콘솔에서 위 명령으로 실행하면 초기 오류를 확인하기 쉽습니다. 제한된 네트워크에서는 Chrome 버전과 호환되는 드라이버를 미리 준비해야 합니다.'),
 ('table',(['입력 항목','사용 방법'],[
 ['API 주소','관리자가 안내한 주소 입력. 기본 예시 192.168.1.10:8000을 실제 주소로 확인'],['API Key','관리자에게 받은 키 입력 후 연결 확인. /health와 인증한 /ready 호출'],['BUS 계정·암호','저장된 정보가 없으면 로딩 중 전용 로그인창에 입력'],['계정 정보 저장','기본 해제. 선택 시 저장, 해제 시 이번 실행에서만 사용하고 저장 계정 제거']])),
 ('h','시작 순서와 실패 시 선택'),
 ('p','작은 로딩창에서 API 연결 → BUS 로그인 → 권한 요청 조회 → 신규 프로젝트 조회를 진행합니다. 회전 표시와 완료 체크로 진행을 보여 줍니다. 단계당 120초 지연 시 다시 시도 또는 메인 화면 열기를 선택할 수 있습니다. 기존 작업이 끝나기 전 중복 재시도는 제한됩니다.'),
 ('callout',('로그인 동작','로그인 입력창은 일반 설정과 같은 계정 저장 함수를 사용합니다. 입력 확정 시 저장 선택을 적용하며, 로그인 성공 후에만 저장하는 방식은 아닙니다. 취소하면 저장 파일은 바뀌지 않습니다.')),
]),
('화면 구성과 기본 조작', [
 ('image',('access.png',5.65)),
 ('table',(['영역','할 수 있는 일'],[
 ['상단 연결 상태·설정','API·BUS 상태 확인, 톱니바퀴에서 연결·계정 설정과 사용 안내'],['권한 요청 / 신규 프로젝트','두 작업 목록 전환. 조회한 목록은 탭 이동만으로 다시 조회하지 않음'],['목록 / 오른쪽 상세','체크박스로 실행 대상 선택 / 행 선택으로 대상·Role·진행·오류 확인'],['하단 작업 로그·버튼','로그 접기·펼치기, 선택 실행·중지. 화면 로그 삭제는 파일 기록 유지']])),
 ('p','밝은·어두운 테마는 오른쪽 위 달/해 아이콘으로 바꿉니다. 버전은 오른쪽 하단에 표시됩니다. 요청 알림·숫자 배지는 현재 비활성화되어 있으며, 새로운 요청은 각 탭의 BUS 조회로 갱신합니다. 화면은 샘플 데이터입니다.'),
]),
('권한 요청 처리 절차', [
 ('steps',['BUS 요청 조회로 진행·종료 권한 요청을 불러오거나 수동 입력·엑셀 불러오기로 준비합니다.','사번, 프로젝트, 진행/종료, 부여/해제, Level2·Level3·Role을 확인하고 처리 항목을 체크합니다.','미리보기를 눌러 서버가 계산한 경로·그룹·실행 계획을 확인합니다. 실제 권한은 바뀌지 않습니다.','베타 첫 시험은 성공 후 BUS 자동 완료를 끈 뒤 선택 실행합니다.','서버 최종 결과와 실제 그룹·폴더 권한을 확인합니다. 검증 후 BUS 완료 재확인을 눌러 완료 처리합니다.','자동 완료는 제한된 시험이 끝난 후 사용합니다. 실제 성공한 BUS 원본 요청만 완료됩니다.']),
 ('table',(['입력 경로','특징'],[
 ['BUS 요청 조회','원본 BUS 연결 정보 보존. 실제 성공 확인 후 완료 처리 가능'],['수동 입력 / Excel','API 요청을 만들지만 BUS 완료 대상은 아님'],['BUS 행 직접 편집','수동 요청으로 전환하고 이전 결과 무효화. BUS 완료 연결 해제'],['상세 열 표시','톱니바퀴에서 이름·부서·Role 등 편집 열 표시']])),
 ('callout',('중지와 재조회','중지는 다음 항목 접수를 멈추고 진행 작업에 취소를 요청합니다. 이미 적용된 권한을 되돌리지 않습니다. 서버가 처리 중일 수 있으므로 중복 등록 전에 결과를 확인하세요.')),
]),
('신규 프로젝트 처리 절차', [
 ('image',('projects.png',5.3)),
 ('steps',['수동 추가 또는 BUS 신규조회로 요청을 준비하고 목록에서 대상을 체크합니다.','오른쪽 상세에서 프로젝트 코드·이름과 BUS 원본 정보를 확인합니다.','미리보기 (Dry Run)를 켜고 선택 실행해 생성 계획을 확인합니다.','실제 시험은 Dry Run을 해제하고 실제 성공 후 BUS 완료를 끈 상태에서 실행합니다.','실제 자원을 검증한 뒤 BUS 완료 옵션을 켜고 같은 항목을 실행하면 보관된 성공 결과로 완료만 재시도할 수 있습니다.']),
 ('callout',('구분','신규 프로젝트는 서버에서 그룹·템플릿 폴더·기본 ACL을 생성·검증합니다. 실패한 권한 요청의 부분 재실행 API는 프로젝트 Job에 제공되지 않습니다. 화면은 샘플 데이터입니다.')),
]),
('결과 확인·BUS 완료·권한 재시도', [
 ('table',(['결과','의미 · 다음 동작'],[
 ['미리보기 / DryRun','계획만 확인. 실제 실행과 BUS 완료 없음'],['모의 실행 / simulated','Mock 서버 결과. 실제 자원 변경·BUS 완료 없음'],['API 성공','실제 서버 작업 성공. BUS 원본이면 완료 처리 가능'],['BUS 완료','서버 성공 확인과 BUS 완료 처리가 끝남'],['부분 성공','일부 단계만 성공. 실패 원인 조치 후 권한 작업 명시적 재시도'],['실패 / 시간 초과','상세 오류·파일 로그 확인. 기존 Job은 원인 조치만으로 성공으로 변하지 않음'],['결과 확인 필요','접수나 조회 결과가 불확실. 같은 요청번호로 서버 이력 확인']])),
 ('h','권한 요청을 다시 실행할 때'),
 ('p','실패·부분 성공 항목을 체크하고 선택 실행을 누른 뒤 재실행 안내를 확인합니다. 서버가 새 Job을 만들고 이전 성공 단계는 유지하며 실패·미실행 단계를 처리합니다. 폴더를 복구한 뒤 일반 등록만 반복하면 기존 결과가 반환될 수 있습니다.'),
 ('h','완료 재확인은 권한 작업 재실행이 아닙니다'),
 ('p','BUS 완료 재확인은 서버 최신 상태를 확인하고 BUS 완료만 시도합니다. source=bus, executor_mode=powershell, status=succeeded이며 중지되지 않은 요청만 허용합니다. 조회 실패나 다른 요청의 응답을 근거로 완료하지 않습니다.'),
 ('callout',('재시도 전제','API와 Client가 모두 최신 재실행 계약을 지원해야 합니다. 성공했던 단계를 외부에서 바꿨다면 자동 재검증되지 않으므로 실제 권한을 먼저 확인하세요.')),
]),
('Client 내부 설계와 스레드 구조', [
 ('table',(['클래스 · 단일 실행 파일 내','책임'],[
 ['AccessManager','메인 UI, 권한 목록, 연결 상태, 실행·완료 흐름 조정'],['StartupLoadingDialog / BusySpinner','시작 단계와 진행 애니메이션, 지연·재시도 안내'],['BusLoginDialog / SettingsDialog','간단 로그인 / 일반 설정. 계정 처리 함수 공유'],['BusSessionManager','BUS 로그인 세션, 요청 다운로드와 완료 처리'],['FolderGrantApiClient','HTTP JSON, 인증 헤더, 오류 변환·키 마스킹'],['ApiJobWorker / ApiJobRefreshWorker','API 접수·조회·취소·재실행 / 완료 전 최신 결과 조회'],['NewItemsViewer / CreateWorker','프로젝트 요약·상세 화면 / API 생성과 BUS 완료'],['RequestTable / RequestDelegate / HelpDialog','표 편집·상태 표현 / 사용자 안내']])),
 ('h','UI와 외부 작업의 분리'),
 ('p','HTTP와 Selenium 작업은 Qt 작업 스레드에서 수행하고 신호로 화면을 갱신합니다. 권한 작업은 선택 항목을 순서대로 접수합니다. 화면 위젯은 UI 스레드에서만 조작하고, 창 종료 시 실행 스레드 정리를 기다립니다.'),
 ('table',(['대기 설정','현재 값'],[
 ['HTTP 요청 제한','요청당 30초'],['Job 조회','1.5초 간격 / 최대 4시간'],['시작 단계 지연 안내','120초 / 기존 작업을 강제로 종료하지 않음'],['서버 PowerShell 제한','별도 서버 설정. 기본 단계당 10800초']])),
]),
('요청 식별·저장 파일·보안', [
 ('h','중복 처리를 줄이는 식별 규칙'),
 ('p','BUS 원본 번호가 있으면 사용하고, 없으면 원본 필드·대상·작업·경로·Role 등을 해시해 request_id를 만듭니다. 행 순서는 식별에서 제외합니다. 응답 유실 시 by-request로 복구하며 임의 새 번호로 우회하지 않습니다. 원본 필드가 모두 같은 별도 요청은 구분하지 못할 수 있습니다.'),
 ('table',(['PC 경로','역할 · 취급'],[
 ['C:\\FGT\\conf\\api_client.json','서버 주소·API Key 평문. 파일 접근 제한·공유 금지'],['C:\\FGT\\conf\\login.json','저장을 선택한 BUS 계정·암호와 테마 등 설정. 기존 프로그램과 공유'],['C:\\FGT\\Log\\access.log','날짜·시간을 포함한 상세 기록을 단일 파일에 누적'],['C:\\FGT\\ef','BUS 다운로드 자료'],['C:\\FGT\\debug','BUS 오류 조사 자료. 화면·업무 정보 포함 가능']])),
 ('h','화면 로그와 파일 로그는 다릅니다'),
 ('p','화면은 결과를 짧게, 파일은 API 응답·오류·단계 결과를 상세히 기록합니다. 화면의 로그 지우기는 파일을 삭제하지 않습니다. 자동 파일 순환은 없으므로 보관·정리 기준을 정하세요.'),
 ('callout',('보안 운영 권장','API Key와 저장 암호는 암호화 저장이 아닙니다. 공용 PC에서는 계정 저장을 해제하고, 설정·다운로드·로그를 배포본에 넣지 마세요. 키 마스킹이 있어도 상세 로그의 사번·경로 등 업무 정보는 별도로 보호해야 합니다.')),
]),
('문제 해결과 문의 시 준비 자료', [
 ('table',(['증상','확인 순서'],[
 ['시작 연결 실패','API 주소·키·서버 서비스·네트워크 확인. 필요 시 메인 화면으로 진입해 설정'],['BUS 로그인 실패','계정·암호, Chrome·드라이버·BUS 접근 확인. 로딩 재시도 시 계정 재입력'],['BUS 조회 오류 / 빈 목록','연결 상태 확인 후 해당 BUS 조회. 실제 요청 없음과 조회 실패 상태를 구분'],['부분 성공 / 파일 없음','오른쪽 실패 단계·경로 확인. 서버 담당자가 UNC·실제 폴더를 확인한 뒤 권한 재시도'],['300초 초과','서버 .env의 timeout을 담당자가 확인. Client 대기와 서버 제한은 다른 값'],['고쳤는데 이전 실패 반환','이전 작업 결과·재실행 안 함 표시 확인. 원인 조치 후 명시적 권한 재실행'],['서버 성공 / BUS 완료 실패','권한 탭은 BUS 완료 재확인. 프로젝트는 완료 옵션을 켜고 같은 성공 항목 처리'],['버튼 비활성 / 종료 지연','API·BUS 작업이 진행 중일 수 있음. 중지 후 상태 확인, 강제 종료 전 담당자 확인']])),
 ('h','문의할 때 함께 보내 주세요'),
 ('p','발생 시각, Client·API 버전, 어떤 버튼을 눌렀는지, 프로젝트·요청 종류, 화면 상태, access.log의 request_id/job_id와 실패 단계·오류 문구를 전달하세요. API Key·암호·불필요한 개인정보는 제거합니다.'),
 ('callout',('결과가 불확실할 때','서버가 작업 중일 수 있습니다. 같은 내용을 새 수동 요청으로 다시 입력하거나 BUS를 직접 완료하기 전에 기존 Job과 실제 권한 상태를 확인하세요.')),
]),
('유지보수·배포·회귀 검증', [
 ('table',(['변경 대상','유의 사항'],[
 ['단일 .pyw / assets','기존 프로그램·API 코드를 직접 import하지 않음. 배포 시 아이콘·로고 동봉'],['BUS 화면·Excel 양식','URL·iframe·셀렉터·헤더 별칭·해제 판정·완료 매칭을 실제 BUS에서 확인'],['프로젝트·Role 규칙','Client 변환과 서버 정책을 함께 확인. 표 표시용 합치기가 실제 필드값을 훼손하면 안 됨'],['계정 입력·저장','전용 로그인과 일반 설정은 _apply_login_credentials 공유 유지'],['스레드·중지·종료','UI 스레드 접근, busy 상태, 종료 대기, 중복 작업 방지 확인'],['BUS 완료·재실행','원본 request_id·source와 최신 실제 성공을 보존. cached 성공만으로 완료하지 않음'],['알림·배지','주석으로 비활성화 상태 유지. 복원 시 감시 세션·타이머·설정·UI 모두 검토']])),
 ('code','cd C:\\Folder-Grant-Tool\n.\\Client\\.venv\\Scripts\\python.exe -m unittest discover -s Client\\tests -v\n.\\Client\\.venv\\Scripts\\python.exe -m py_compile ".\\Client\\Folder Grant API Client.pyw"'),
 ('p','문서 작성 시 Client 자동 테스트 93개가 통과했습니다. API 계약·Mock/Preview 완료 차단·재실행·UI 흐름·시작 로그인 등을 검사합니다. Windows 실제 GUI, Chrome, BUS 및 AD·공유 폴더 연동은 별도 확인 대상입니다.'),
 ('callout',('배포 권장','requirements.txt는 버전 범위를 지정합니다. 베타 배포마다 실제 설치 버전·Python·Chrome 버전과 코드를 함께 기록하세요. 현재 저장소에는 정식 EXE 빌드 절차가 확정되어 있지 않으므로 소스 실행 기준으로 배포합니다.')),
]),
('베타 참여 안내와 확인 목록', [
 ('table',(['단계','확인할 사항'],[
 ['환경 준비','관리자에게 서버 주소·API Key·시험 대상·문의 경로 수령'],['연결·로그인','저장 해제/선택, 취소, 실패 후 재입력, 양쪽 탭 초기 목록 확인'],['조회·화면','권한·프로젝트 수동 조회, 요약·상세, 행 번호, 스크롤, 밝은·어두운 테마'],['계획 확인','수동·Excel·BUS 요청의 사번·코드·경로·Role을 Preview로 비교'],['실제 시험','자동 완료 OFF, 제한된 테스트 대상 1건 처리 후 실제 권한 확인'],['실패 복구','누락 폴더·부분 성공·시간 초과·조회 오류·명시적 재실행 확인'],['완료 확인','BUS 완료 재확인·자동 완료, 수동·Preview·Mock 완료 제외 확인'],['기록 확인','화면 요약과 access.log 상세 분리, 민감 정보 제거 후 피드백 제출']])),
 ('h','피드백을 남기는 간단한 형식'),
 ('p','환경·버전 → 작업·버튼 → 기대한 결과와 실제 결과 → 발생 시각·request_id·job_id·필요한 로그 순서로 정리해 주세요.'),
 ('callout',('베타의 의미','현재 기능을 제한된 대상에서 검증하는 단계입니다. 자동 테스트 통과를 실제 업무 환경 검증 완료로 해석하지 마세요. 서버 재시작 시 Job 자동 복구와 사용자별 권한 분리는 제공되지 않습니다.')),
 ('h','소스 근거'),
 ('p','Client/Folder Grant API Client.pyw, Client/requirements.txt, Client/tests/, API/src/folder_grant_api/api/routes/. 기존 README의 화면 위치·초기 설명과 차이가 있으면 이 문서의 기준 커밋과 현재 코드를 확인하세요.'),
]),
]
