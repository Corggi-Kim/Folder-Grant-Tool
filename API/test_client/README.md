# API Test Client

BUS 정식 연동 전 수동 요청으로 Folder Grant API의 다양한 프로젝트, Level2, Level3, Role 조합을 시험하는 PyQt 클라이언트입니다. 기존 `Folder Grant Tool_v3.3.5.pyw`는 수정하거나 import하지 않습니다.

## 설치

API 가상환경에서 클라이언트 선택 의존성을 설치합니다.

```powershell
cd C:\Folder-Grant-Tool\API
.\.venv\Scripts\python.exe -m pip install -e ".[client]"
```

## 실행

```powershell
.\.venv\Scripts\python.exe -m test_client.main
```

## 사용 순서

1. 서버 주소와 API Key를 입력하고 **연결 확인**을 누릅니다.
2. 사번, 프로젝트 코드, 작업, 과제 상태, Level2, Level3, Role, 요청자를 입력합니다.
3. **목록에 추가**를 누릅니다. `request_id`는 매번 자동 생성됩니다.
4. 한 행을 선택하여 **선택 Preview**로 대상 그룹과 경로를 확인합니다.
5. 시험할 행들을 선택하고 **선택 실행**을 누릅니다.
6. 클라이언트가 Job 상태를 주기적으로 조회하여 최종 결과와 단계별 결과를 로그에 표시합니다.
7. 서버가 `powershell` 모드이면 실제 권한이 변경됩니다. 반드시 테스트 자원만 사용하십시오.

API Key는 디스크에 저장하지 않으며 프로그램을 다시 실행할 때 재입력해야 합니다. 서버 URL만 Windows 사용자 설정에 저장합니다.
