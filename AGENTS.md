# Project Development Rules

## Project Overview
- 본사 당직명령 자동 배정 시스템 (docs/[연속근무 스킵형 제도] 안내자료.pdf 참조)
- 핵심 스크립트: `dangjik_basic1.py` — SQLite 기반 순번 배정, 28일 간격 체크, 스킵 큐(익익월 복귀), 수동 배정(벌당직/대근) 로직
  - 스킵정책: /admin의 '해당 월에만 스킵 정책 반영'이 기본ON. system_settings.skip_current_month_only=1이면당월중복스킵,0이면기존28일기준. 사유에벌당직/대근이명시된수동근무는양쪽정책모두판정에서제외하며이력보존. `_has_recent_regular_duty`를정규후보/큐복귀가공통사용. 월삭제시설정유지, 전체DB초기화는기본ON.
  - CLI 실행: `.venv/bin/python dangjik_basic1.py` (메뉴 루프: 초기데이터/배정생성/수동배정/조회)
- 웹 앱: `app.py` (Flask) — 모든 DB 로직은 `dangjik_basic1.py`를 import해 재사용하므로 배정 규칙 변경은 이 모듈에서만 (app.py 라우트는 얇게 유지)
  - 실행: `.venv/bin/python app.py` → http://localhost:5001
  - 기능: 엑셀 입력양식 업로드 / 배정 생성(월 단위) / 수동 예외 배정(벌당직) / 직원 변동(퇴사·승진) 처리 / 시스템 관리(/admin — 전체 DB 초기화, 월 데이터 삭제: 확인 문구 + 백업 사본)
  - 월 삭제 안전성: 생성 시 포인터 스냅샷(system_state_snapshot) 자동 기록 → 지우고 재생성하면 처음과 동일. 스냅샷 없는 구월은 이력 기반 폴백(결과 달라질 수 있음)
  - 웹 테스트: Flask test_client + 임시 DB(`dy.DB_NAME` 교체) 스모크 테스트 — 실데이터 무영향
  - 직원 변동 이력: `/employees/history`, `employee_change_history`에 전입·전출·제외/복구·정보/엑셀변경의 전후 정보·사유·변동일·KST기록시각 저장. 월 삭제 시 보존. 변동일은 기록용이며 현재 상태에는 즉시 적용(날짜별 상태를 사용한 과거 재생성은 미구현).
- DB: `duty_system_v3.db`(현재) / `duty_system_v2.db`(원본 보존). 주의: v3가 없으면 코드가 v2를 복사해 재이관하므로 v3 삭제 시 수동 배정 이력·4월 이후 데이터가 사라짐
- 참고 자료: `docs/`(월별 당직명령 엑셀 — 2·3월은 `.xls`, 4월부터 `.xlsx`, 당직공문.txt), 루트의 `당직명령시스템_2월통합_입력양식.xlsx`(업로드 양식 기준)
- 검증 기준(2026-10-26 기준): docs 3월 실제 공문 대조 현재 49/51 일치. 잔여 2건(3/28·3/29 일직 부)은 조보영(사번22261037) 명부 누락 때문이며, 조보영을 순번54에 임시 등록하면 51/51. 옛 수치 "v2 회귀 100%", "0/51" 을 현재 상태 설명에 쓰지 말 것. `scripts/audit_march.py`, `audit_results/2026-02·2026-03/` 참조. 배정 규칙 변경 후 v2 회귀뿐 아니라 실제 공문 대조도 반복할 것.
- 원본 대조 검증/재생성 실험은 반드시 DB 사본+백업으로 수행(기존 작업들이 DB SHA‑256 보존 확인을 기록함). 조보영/류승현 사번 중복의 입력양식 정정과 동일 월 직접 재생성 시 포인터/상태 복원은 미완료 TASKS 참조.

## Python Environment
- 반드시 .venv 가상환경 사용: `.venv/bin/python`
- 시스템 Python에 패키지 설치 금지
- requirements.txt 유지 — 의존성 변경 시 갱신
- 파이선 버전은 3.12.15로 해줘.

## Coding Standards
- 함수에는 타입 힌트 작성
- 변경 후 테스트 실행:
  - 전체 유닛테스트: `.venv/bin/python -m unittest discover -s tests` (tests/test_*.py 4종 — 배정 간격/직원변동/변동이력/상태정정)
  - 문법 확인: `.venv/bin/python -m py_compile <파일>`
  - 테스트는 임시 DB(`dy.DB_NAME`를 tempfile 경로로 교체) 패턴 사용 — 실데이터 무영향
  - 웹 화면 검증은 Flask test_client + 임시 DB(스모크테스트). 라우트는 `if __name__ == '__main__'` 블록 **앞**에 정의(뒤에 두면 라이브 서버 404)

## Git Rules
- 기존 기능을 임의로 삭제하지 말 것
- 다른 에이전트의 변경사항을 덮어쓰지 말 것
- 대규모 변경 전에 영향 범위 확인
- DB 파일과 엑셀 원본은 데이터 자산이므로 임의로 수정/삭제 금지
- 과거 기간을 재생성하면 duty_history는 덮어써지지만 `system_state`의 `ptr_*` 포인터는 롤백되지 않음 — 과거 데이터 정정 시 이후 배정 중복/누락 가능, 주의

## 작업 관리 규칙
- 작업 시작 전 TASKS.md와 PROGRESS.md를 읽는다.
- TASKS.md에서 진행할 작업을 확인한다.
- 작업 완료 후 TASKS.md의 상태를 갱신한다.
- 주요 변경 사항과 테스트 결과를 PROGRESS.md에 기록한다.
- 다른 에이전트가 이어서 작업할 수 있도록
  미완료 사항과 알려진 문제를 명시한다.
