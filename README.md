# Folder Grant Tool · Beta

BUS의 권한 요청과 신규 프로젝트 요청을 처리하는 **Windows Client + API 서버**입니다. Client는 BUS 조회·완료와 사용자 화면을 담당하고, API 서버는 AD 그룹·폴더 권한 변경과 작업 이력을 관리합니다.

| 구성 | 현재 버전 | 역할 | 안내 |
|---|---|---|---|
| `API/` | 0.4.0 | 요청 검증, Preview, Job 실행·조회·취소·권한 재시도 | [서버 설치·운영](API/README.md) |
| `Client/` | V1.1.0 Build:20261022 | BUS 조회·완료, 수동·Excel 입력, 결과 확인 | [Client 설치·사용](Client/README.md) |
| `API/test_client/` | 개발 도구 | 수동 API 시험, BUS 연동 없음 | [개발 도구 안내](API/test_client/README.md) |

기존 `Folder Grant Tool_v3.3.5.pyw`는 별도로 보존합니다. 정식 BUS Client의 실행 파일은 `Client/Folder Grant API Client.pyw`입니다.

## 베타 문서

[문서 목록과 읽는 순서](docs/README.md)

- [API 설계·운영·유지보수 · Word](docs/API_설계_운영_유지보수_Beta.docx)
- [Client 설계·사용·유지보수 · Word](docs/Client_설계_사용_유지보수_Beta.docx)
- [Folder Grant 프로그램 소개서 · PowerPoint](docs/Folder_Grant_프로그램_소개서_Beta.pptx)

## 시작 순서

1. 서버를 `mock`으로 실행해 인증·DB·Preview·상태 흐름을 확인합니다.
2. [Client 안내](Client/README.md)에 따라 설치하고 서버 주소·API Key·BUS 계정을 설정합니다.
3. Mock과 분리한 DB·테스트 자원에서 PowerShell 모드를 검증합니다. 첫 실제 시험은 자동 BUS 완료를 끄고 수행합니다.
4. 실제 그룹·폴더 권한과 BUS 완료를 확인한 뒤 제한된 베타 대상에 배포합니다.

`powershell` 모드는 실제 AD·폴더 ACL을 변경합니다. Preview와 Mock은 BUS 완료를 처리하지 않습니다. API 성공과 BUS 완료 성공은 별도 결과입니다.

현재 Worker는 API 프로세스 안에서 실행되며 서버 재시작 후 진행 Job을 자동 복구하지 않습니다. 서비스 자동 재시작·중지는 권한 롤백이나 Job 복구를 뜻하지 않습니다. 베타 오픈 점검과 장애 대응은 Word 문서를 참고하세요.
