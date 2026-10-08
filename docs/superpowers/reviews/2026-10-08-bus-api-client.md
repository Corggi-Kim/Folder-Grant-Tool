# BUS API Client 검증 기록

## 결과

- 독립 구현: `Client/Folder Grant API Client.pyw`, requirements, assets, README, tests.
- 기존 프로그램과 API: origin/main 대비 변경 없음.
- Client unittest: 29개 통과 (모의 HTTP 서버, offscreen Qt).
- 기존 API pytest: 54개 통과, 기존 라이브러리 deprecation 경고 2개.
- 단일 파일 py_compile 및 origin/main 대비 git diff --check 통과.
- 실제 AD·ACL·BUS 변경은 실행하지 않았고 Windows GUI·실제 Chrome/BUS는 현장 검증이 필요하다.

## 독립 코드 리뷰와 수정

별도 검토자의 전체 변경 검토에서 Critical은 없고 Important 4건이 확인됐다. 각 문제를 테스트로 재현한 뒤 수정했고 전체 Client 테스트를 다시 통과했다.

1. 취소 POST가 서버 완료와 경합하여 409를 반환하면 최종 상태를 조회한다. 중지 플래그를 유지하므로 자동 BUS 완료는 실행하지 않는다.
2. 신규 프로젝트 화면을 하나만 유지하여 여러 창의 worker가 종료 추적에서 빠지지 않도록 한다.
3. 프로젝트 시작 상태와 worker를 먼저 등록하고 불필요한 Qt 이벤트 처리를 제거하여 시작 직전 중지 요청을 보존한다.
4. 다른 운영자의 동일 BUS 요청 충돌은 by-request와 Preview의 실제 실행 계획을 비교한 후에만 복구한다. 기존 서버 요청자의 이력은 변경하지 않는다.

검증 중 추가로 발견한 취소 중간 상태 cancel_requested, 원본 요청번호를 유지하는 Excel 식별, BUS 세션 준비 중 API 큐 재진입, 수동 완료 처리 준비 중 조작 잠금, 프로젝트 수동 입력의 대체 이름 헤더도 재현 테스트로 확인하고 수정했다.

## 실행 판단

- 기존 관리형 work 브랜치에서 작업했다. 별도 중첩 worktree를 만들지 않아 호스트의 변경 추적을 유지하며 main은 수정하지 않았다.
- 클라우드 스킬 스크립트를 로컬 경로로 실행할 수 없어 동일한 진행 기록을 work/에 작성했다. 제품 동작에는 영향이 없다.
- 입력 계약, UI 및 BUS 원본 흐름을 유지하는 범위로 재구성했으며 이전 커밋 5c08be0의 원본 복구나 해당 PR 생성 성공을 주장하지 않는다.

보류한 Minor 수정은 없다. 원본 BUS 매칭과 원본 식별자 누락 시의 한계는 README에 기록했다.
