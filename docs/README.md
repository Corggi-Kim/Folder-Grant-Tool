# 오픈·운영 문서

2026-10-26, `main ebbdfb9d4b8ffe09456bc2c9cc5c4dd4a16328d7` 기준입니다. 프로그램 버전은 API **0.4.0**, Client **V1.1.0 Build:20261022**로 유지합니다. 문서 작성일·문서판 2.0과 프로그램 빌드는 구분합니다.

## 최신 운영본

| 문서 | 대상 | 주요 내용 | 분량 |
|---|---|---|---|
| [API 설계·운영·유지보수 · Word](API_설계_운영_유지보수.docx) | 서버 운영자·개발자 | 설계·코드 구조·정책·계약·실제 검증·서비스·백업·장애·유지보수 | 17쪽 |
| [Client 설계·사용·유지보수 · Word](Client_설계_사용_유지보수.docx) | 사용자·배포 담당자·개발자 | 설치·로그인·두 탭·요청 검증·진행·완료·재실행·로그·변경 검증 | 16쪽 |
| [오픈·운영 인수인계 · Word](오픈_운영_인수인계.docx) | 업무 책임자·운영·배포 담당자 | 담당자 기록·오픈/현장 시험 체크리스트·일상 운영·장애·백업·복구·인수 | 8쪽 |
| [프로그램 소개서 · PowerPoint](Folder_Grant_프로그램_소개서.pptx) | 사용자·내부 설명회 | 목적·역할·화면·사용 흐름·결과 판단·설치·현재 제한 | 14장 |
| [개발 완료 보고서 · PowerPoint](Folder_Grant_개발완료_보고서.pptx) | 개발 완료 보고·인수 담당자 | 완료 범위·구조·화면·검증·실패 복구·운영 전환·제약·후속 과제 | 18장 |

분량은 Noto Sans CJK KR을 사용하는 LibreOffice PDF 변환 기준이며 다른 PC의 글꼴·Office 버전에 따라 Word 페이지 수는 달라질 수 있습니다. PPT의 제목·본문·카드는 편집 가능한 도형이며 실제 화면과 구성도는 이미지입니다.

## 읽는 순서

1. **업무 사용자**: 프로그램 소개서 → Client Word → 프로그램의 도움말.
2. **오픈 책임자**: 개발 완료 PPT → 오픈·운영 인수인계 → [오픈 확인 목록](OPEN_CHECKLIST.md).
3. **서버 운영자**: API Word → [API README](../API/README.md) → 인수인계의 백업·장애 절차.
4. **유지보수 개발자**: 두 설계 Word → [변경·검증 기록](RELEASE_NOTES.md) → 코드·테스트와 [Client README](../Client/README.md).

## 반영 내용과 검증 범위

- 기존 베타본의 사람이 수정한 문장·구성을 바탕으로 운영본을 만들고, 바뀐 기능·안내를 수정했습니다.
- 실제 요청 검증과 개발용 Preview를 구분합니다. 요청 검증은 `All`을 별도 검사하지 않으며 실제 생성 시 기본 권한 처리는 유지합니다.
- 양쪽 탭의 BUS 완료 재확인, 회전 표시와 처리 행 자동 선택, 공통 상세 열, 행 삭제·편집 제한과 타이머를 반영합니다.
- 코드 자동 시험을 다시 실행했습니다: **API 76개 통과·선택 PowerShell 시험 5개 건너뜀 / Client 125개 통과**.
- Word·PPT ZIP/XML 구조, 기준 커밋·최신 기능 문구와 링크를 확인하고 PDF 변환 후 배치를 검토했습니다.
- 실제 Windows·Chrome·BUS·AD·SMB 인수 시험과 부하 시험은 이번 문서 작업에서 수행하지 않았습니다. 오픈 승인·담당자·시험 결과 칸은 담당자가 작성합니다.
- 현재 제한: API 프로세스 내 Worker, 재시작 후 진행 Job 복구 없음, 자동 롤백·개인별 권한·프로젝트 잠금·로그 순환 없음. 새 요청 자동 감지는 보류이며 알림·배지는 비활성화입니다.

현재 화면 예시: [권한 요청](release-assets/access.png) · [다크 테마](release-assets/access-dark.png) · [신규 프로젝트](release-assets/projects.png).
이 이미지는 최신 코드에 오프라인 샘플 데이터를 넣어 캡처했으며 실제 권한 처리 증거가 아닙니다.

## 기존 사용자 수정본 보존

다음 파일은 **2026-10-23 당시 베타 기록**으로 보존하며 최신 동작의 안내는 위 운영본을 사용합니다. 사용자가 수정한 파일·기존 원본·자산은 덮어쓰지 않았습니다.

- [API 베타본](API_설계_운영_유지보수_Beta.docx)
- [Client 베타본](Client_설계_사용_유지보수_Beta.docx)
- [프로그램 소개서 베타본](Folder_Grant_프로그램_소개서_Beta.pptx)
- 기존 [beta-assets/](beta-assets/)와 [screenshots/](screenshots/README.md), [베타 내용 원본](tools/beta_content.py), [베타 생성 도구](tools/build_beta_documents.py).

이전 베타 생성 도구는 과거 산출물을 다시 만드는 코드이므로 **현재 사용자가 수정한 Office 파일 위에서 실행하지 마세요.** 직접 수정한 Office 내용이 항상 자동 생성 원본에 들어 있는 것은 아닙니다.

## 운영 문서 수정·생성

[build_release_documents.py](tools/build_release_documents.py)는 기존 베타 Office 편집본을 입력으로 읽고, 변경된 안내와 [release_content.py](tools/release_content.py)의 최신 기능·인수 내용을 적용해 별도 운영본을 만듭니다. 기존 베타 생성기를 실행하지 않습니다. 수정 내용은 기존 Office 서식·문장 기준으로 반영하며 범용 Word 병합 도구는 아닙니다.

```bash
python -m pip install -r docs/tools/requirements.txt
# 현재 운영본과 사용자 수정본을 덮어쓰지 않고 검토용 복사본 생성
python docs/tools/build_release_documents.py --output /tmp/fgt-doc-review
# 저장소의 운영본 구조·로컬 링크·베타본 보존 검사
python docs/tools/verify_release_documents.py
```

기본 출력은 `docs/`지만 동일 파일이 있으면 중단합니다. `--refresh`는 [release-manifest.json](release-manifest.json)의 생성 시 해시와 현재 파일이 **모두 일치할 때만** 갱신합니다. Office에서 직접 고친 파일은 해시가 달라지므로 덮어쓰지 않습니다. 변경된 파일은 직접 편집하거나 별도 출력 폴더에서 새 내용을 비교·병합하세요. 원본 파일도 Git에 함께 보존합니다.

문서 도구는 제품 런타임과 분리되어 있습니다. 한국어 글꼴은 Noto Sans CJK KR, 이미지 입력은 `release-assets/`이며 기존 `beta-assets/architecture.png`를 구성도에 사용합니다. 실제 계정·암호·API Key·실환경 로그는 문서·배포 자료에 넣지 않습니다.

계약·정책·설정·버튼·저장 경로·제한이 바뀌면 해당 README, Word, PPT와 기준 커밋·검증 기록을 같이 갱신하고 변환한 페이지·슬라이드의 배치를 다시 확인하세요.
