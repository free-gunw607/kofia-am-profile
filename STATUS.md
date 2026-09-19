# STATUS

## Repository
- repo: `kofia-am-profile`
- workspace path: `~/agent-coding/agent-projects/A4-worker-repos/kofia-am-profile`

## Current objective
KOFIA 자산운용사 종합 프로파일 V4: KOFIA 회원사 자동 수집 + 매칭 복구

## Current phase
V4 완료: KOFIA 회원사 수집 + 매칭/대표자/VLOOKUP 전부 정상 작동

## V4 changes (from V3)
- **KOFIA 회원사 자동 수집**: `collector.py`에 `collect_kofia_members()` 추가
  - 소스: `kofia.or.kr/members/m_61/sub020201.do?srchCate=5` (GET, 34페이지)
  - beautifulsoup4 + requests로 HTML 파싱
  - 334개사 전체 수집 (회사명/대표자/대표전화/주소/웹사이트/로고)
- **CSV 저장**: `data/raw/KOFIA_자산운용사_리스트.csv` (xlsx 대신, 재삭제 방지)
- **joiner.py CSV 지원**: `_build_raw_members_sheet()`가 CSV 우선, 없으면 xlsx에서 읽기
- **매칭 복구**: COUNTIF(RAW_회원사!) 정상 작동 → "매칭"/"비매칭" 분류 복원
- **대표자 VLOOKUP 복원**: RAW_회원사에서 대표자/대표전화 조회 정상

## V4 architecture
```
collector.py → collect_kofia_members() → requests + BS4 → CSV 저장
             → collect_freesis_all_modes() → FreeSIS 3모드

joiner.py   → _build_raw_members_sheet() → CSV 우선, xlsx 대체
             → AUM/NAV/펀드수 통합 시트 → COUNTIF + VLOOKUP RAW_회원사 참조

cli.py      → run --date 20260916
```

## 검증 결과 (2026-09-18)
### 1단계: KOFIA 회원사 수집
- 1페이지 샘플: 10개사 정확 파싱, &nbsp; 없음
- 페이지네이션: 페이지 1/10/34 각각 다른 데이터, 마지막=흥국자산운용
- 빈/이상 케이스: 페이지 0/35/100 모두 정상 처리
- 전체 수집: 34페이지 → 334개사, 빈 필드 0건
- FreeSIS 교차검증: 332개사 양쪽 일치, 핵심 5개사 전부 매칭

### 2단계: FreeSIS 점검
- FreeSIS 3모드 수집: 520개사 정상
- 멀티사 검증: AUM 6/6, NAV 6/6 통과
- AUM vs NAV 일관성: 6/6 양호 (100~137%)
- 삼성 NAV 주식: 1,224,864억원 (FreeSIS 정확 일치)
- RAW_AUM 컬럼 순서: 21/21 FreeSIS UI 100% 일치
- formula 참조: AUM/NAV/펀드수 통합 시트 전부 정상

## Key files
- `src/kofia_am_profile/collector.py` — FreeSIS 수집 + KOFIA 회원사 수집
- `src/kofia_am_profile/joiner.py` — 엑셀 생성 V4 (CSV 지원)
- `src/kofia_am_profile/cli.py` — CLI (--date 옵션)
- `data/raw/KOFIA_자산운용사_리스트.csv` — KOFIA 회원사 원본 (334개사)
- `data/processed/KOFIA_자산운용사_종합프로파일.xlsx` — 최종 출력

## Known limits
- FreeSIS 기관전용사모펀드: 삼성 NAV=0 (FreeSIS 화면=39,282) — API vs 화면 표시 차이
- 기준일은 매일 변동 → `--date` 옵션으로 확인 필요
- KOFIA 회원사 페이지는 GET 요청 (WAF 차단 없음), 주기적 업데이트 필요
