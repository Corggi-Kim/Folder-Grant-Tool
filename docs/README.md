# 베타 배포 문서

2026-10-23 기준 코드(`main`의 `26f3f761d4f1`)를 확인해 작성했습니다. 프로그램 버전은 API **0.4.0**, Client **V1.1.0 Build:20261022**입니다. 문서 작성일과 프로그램 빌드 날짜는 다릅니다.

| 문서 | 대상 독자 | 주요 내용 |
|---|---|---|
| [API 설계·운영·유지보수 · Word](API_설계_운영_유지보수_Beta.docx) | 서버 설치·운영 담당자, 개발자 | 전체 설계, 모듈, 정책, API 계약, 설정, 서비스 배포, 백업, 장애 대응, 유지보수, 베타 확인 목록 |
| [Client 설계·사용·유지보수 · Word](Client_설계_사용_유지보수_Beta.docx) | 베타 사용자, 배포 담당자, 개발자 | 설치·로그인, 화면, 권한·프로젝트 처리, 결과·재시도, 스레드 구조, 설정·로그, 장애 대응, 변경 검증 |
| [Folder Grant 프로그램 소개서 · PowerPoint](Folder_Grant_프로그램_소개서_Beta.pptx) | 베타 참여자, 업무 담당자, 내부 설명회 | 도입 목적, API·Client 역할, 실제 화면 예시, 처리 흐름, 결과 판단, 베타 도입 절차와 현재 제한 |

Word는 각각 **14쪽 / 12쪽**, PPT는 **14장**으로 구성했습니다. 표지·페이지 번호·테마를 적용했고, PPT의 제목·본문·카드는 편집할 수 있습니다. 화면과 전체 구성도는 이미지입니다. 페이지 수는 작성 환경의 PDF 변환 기준이며 설치 글꼴에 따라 달라질 수 있습니다.

## 읽는 순서

1. **업무 사용자**: PPT → Client Word의 설치·화면·작업·결과·문제 해결.
2. **서버 담당자**: API Word의 설정·서비스 배포·운영·장애·베타 점검.
3. **개발자**: 두 Word의 구조·계약·유지보수 장과 [API README](../API/README.md), [Client README](../Client/README.md).

## 문서의 기준과 검증 범위

현재 구현과 운영 권장 사항을 구분했습니다. 미구현 개선 항목은 현재 기능으로 소개하지 않습니다. 실제 API Key나 BUS 계정·암호는 포함하지 않았습니다. 화면은 현재 Qt 프로그램을 오프라인 샘플 데이터로 실행해 캡처했습니다.

- 자동 테스트: API **60개**, Client **93개** 통과.
- Office 파일: 생성 후 LibreOffice로 PDF 변환해 페이지와 슬라이드의 배치 확인.
- 실제 Windows·Chrome·BUS·AD·SMB 연동은 이번 문서 작업에서 실행하지 않았습니다.
- 서버 재시작 후 Job 자동 복구, 사용자별 권한 분리, 자동 롤백은 현재 제공하지 않습니다.

## 화면 예시

[권한 요청](beta-assets/access.png) · [다크 테마](beta-assets/access-dark.png) · [신규 프로젝트](beta-assets/projects.png) · [시작 로딩](beta-assets/startup.png) · [전체 구성도](beta-assets/architecture.png)

이전 UI 변경 과정의 캡처와 시안은 [screenshots/](screenshots/README.md)에 보존되어 있습니다. 현재 베타 문서의 화면 기준은 `beta-assets/`입니다.

## 문서 수정과 재생성

내용 원본은 [tools/beta_content.py](tools/beta_content.py), 문서 생성 코드는 [tools/build_beta_documents.py](tools/build_beta_documents.py)에 있습니다. Word·PPT만 직접 고치면 재생성 시 변경이 사라지므로, 원본도 함께 갱신하세요.

```bash
python -m pip install -r docs/tools/requirements.txt
python docs/tools/build_beta_documents.py
```

생성 코드는 제품 실행 환경과 분리된 문서 도구입니다. Linux의 `/usr/share/fonts/opentype/noto/`에 Noto Sans CJK 글꼴이 필요합니다. 다른 환경에서는 생성 코드의 구성도 글꼴 경로를 변경하세요. Word·PPT에도 같은 글꼴을 사용하므로 수신 PC에 설치하면 배치를 더 일관되게 볼 수 있습니다.

API 계약·정책·설정·버튼·저장 경로가 바뀌면 해당 문서와 README를 함께 갱신하고, Office 파일을 다시 열어 배치를 확인하세요. `beta_content.py`의 기준 커밋·날짜와 이 문서의 버전·검증 기록도 함께 수정합니다.
