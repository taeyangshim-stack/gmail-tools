# Gmail 중복 메일 정리 프로젝트

## 배경

- 계정: taeyang.shim@gmail.com
- 총 메일: 4,817건 / 스레드: 3,347건
- 용량 이슈로 인해 중복 메일 정리 필요

## 분석 결과 (2026-04-05)

Gmail MCP를 통해 메일함을 분석한 결과, 3가지 유형의 중복/불필요 메일이 확인됨.

### 1. 엑소스피어 보안알림 (전체 삭제 대상)

- 발신: `no-reply@exosp.com`
- 내용: PC 로그인 상태 변경 알림 (SHIMTAEYANG)
- 패턴: 하루 5~10통씩 동일 내용 반복 수신
- 규모: 약 200통 이상 (label:중복메일 태그된 것만)

### 2. IMAP 동기화 중복

- 원인: 회사 메일(taeyang@spsystems.kr)을 Gmail로 IMAP 동기화 시 동일 메일이 2건씩 생성
- 특징: messageId는 다르지만 제목/발신자/시간이 동일
- 확인된 사례:
  - 한화오션 제안서 (김유미) — 각 ~25MB x 2건
  - 간섭검사 이슈 (안창우, 삼성중공업) — 각 ~13MB x 2건
  - 인체감지 센서 (이호일, 더원시스템) — 2건
  - 주간업무보고 (정시득) — 2건
  - Anti-Spatter Fluid 관련 (이동혁, 삼성중공업) — 2건
- 예상 용량 절약: 200MB+

### 3. Microsoft Outlook 테스트 메시지

- 발신: Microsoft Outlook (taeyang@spsystems.kr → 자기 자신)
- 내용: "계정 설정을 테스트하는 동안 Microsoft Outlook에서 자동으로 보낸 전자 메일 메시지입니다."
- 규모: 3~4통

## 대용량 메일 현황 (참고)

10MB 이상 첨부파일 메일이 50건 이상 존재. 주요 항목:
- 한화오션 소부재 용접시스템 제안서 (3/12~3/19) — 11~26MB, 4건
- 소부재 과제 lay-out / 이슈 공유 — 15~28MB, 5건+
- 라이트커튼 인증/품의 관련 — 11~20MB, 3건
- 갠트리 토크/패킷 데이터 — 12~20MB, 3건

## 스크립트: gmail-cleanup.gs

### 기능

| 함수 | 동작 |
|------|------|
| `step1_exosphere()` | `from:no-reply@exosp.com` 전체 휴지통 이동 |
| `step2_outlook()` | Outlook 테스트 메시지 휴지통 이동 |
| `step3_duplicates()` | 같은 제목+발신자+시간(분 단위) 중복 탐지, 첫 번째만 보존 |
| `runAll()` | 위 3개를 순서대로 실행 |

### 설정

- `DRY_RUN = true` : 미리보기 모드 (삭제하지 않음, 로그만 출력)
- `DRY_RUN = false` : 실제 삭제 모드

### 사용법

1. https://script.google.com 접속 (taeyang.shim@gmail.com)
2. 새 프로젝트 생성 후 gmail-cleanup.gs 코드 붙여넣기
3. 단계별 실행 권장: `step1_exosphere()` → `step2_outlook()` → `step3_duplicates()`
4. Gmail API 일일 한도 초과 시 24시간 후 재실행

### 주의사항

- 휴지통 이동이므로 30일 이내 복구 가능
- Gmail API 일일 호출 한도 존재 — `moveThreadsToTrash()` 배치 처리로 최소화
- `step3_duplicates()`는 1회당 최대 2000건 스캔, 메일이 많으면 반복 실행 필요

## 실행 이력

| 일시 | 작업 | 결과 |
|------|------|------|
| 2026-04-05 23:27 | `runAll()` (DRY_RUN=false) | step1 진행 중 Gmail API 일일 한도 초과로 중단 |
| - | 스크립트 v2로 개선 | 배치 처리(`moveThreadsToTrash`) 적용 |
| - | 재실행 대기 중 | 24시간 후 재실행 필요 |
