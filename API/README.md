# Folder Grant API

기존 `Folder Grant Tool_v3.3.5.pyw`를 변경하지 않고, BUS와 과도기 테스트 클라이언트가 공통으로 사용할 서버 API를 개발하는 독립 프로젝트입니다.

현재 개발 단계는 실제 AD/ACL을 변경하지 않는 **Preview API와 Mock Job API**를 제공합니다. 입력값을 검증한 뒤 대상 AD 그룹, 폴더 경로와 실행 계획을 만들고 SQLite에 작업 이력을 저장합니다.

> **주의:** 현재 Job 실행 결과는 `simulated`입니다. Mock Worker는 AD, 폴더 또는 기타 외부 시스템을 변경하지 않습니다.

## 개발 환경 실행

```powershell
cd API
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
uvicorn folder_grant_api.main:app --reload
```

Swagger UI: <http://127.0.0.1:8000/docs>

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

## 다음 단계

1. Windows PowerShell 실행기와 테스트 자원 검증
2. 실제 Worker의 AD 그룹 및 폴더 ACL 처리
3. API 인증과 운영 오류 처리
4. PyQt/Selenium 테스트 클라이언트
