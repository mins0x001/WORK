# 업무 앱 구조 (V1.15 · 2026-10-29)

화면 3개 + 독립 모듈 8개. 화면은 UI만, 계산·저장·동기화는 모듈이 담당합니다.

| 파일 | 역할 | 의존 |
|---|---|---|
| index.html | 업무허브 (카드 올리기·이동·요약) | base, card, sync, ledger |
| worklog.html | 업무일지 (대리점 조회·회전 저장·달력·집계 화면) | base, sync, holiday, dealer, stats, export |
| expenses.html | 가계부 (분석·분류·고정지출) | base, card, sync, ledger |
| mod-base.js | 공통 도구, 저장소(ls), 알림 버스(bus), 엑셀 쓰기(XL) | 없음 |
| mod-card.js | 카드 엑셀·CSV·붙여넣기 읽기, 취소 상계, 업무 비용 분류 | base |
| mod-sync.js | GitHub 비공개 Gist 동기화 엔진 (문서를 register) | base |
| mod-ledger.js | 가계부 기록·분류·카드 반영(importCore) | base, card, (sync) |
| mod-holiday.js | 공휴일·음력 계산 | base |
| mod-dealer.js | 대리점·거리 데이터, 경로 계산, 주소 자동 정리 | 없음 |
| mod-stats.js | 업무일지 기록 모델·근무 기간·집계 (data를 매개변수로 받는 순수 계산) | base, holiday, dealer |
| mod-export.js | 업무일지 엑셀·CSV 내용 만들기 | base, holiday, dealer, stats |

규칙: 모듈은 `M.<이름>`으로 공개 API만 내보내고 필요한 것만 가져옵니다. 모듈은 화면 변수를 직접 쓰지 않고 `bus`(sync:state, data:changed)로 알립니다.
코드를 고칠 때는 mod-*.js 와 화면 파일을 함께 올리세요.
