# CODEMAP

## src/kofia_am_profile/
- `__init__.py` — 패키지
- `collector.py` — FreeSIS API 수집 (3모드) + KOFIA 회원사 수집 (`collect_kofia_members()`)
  - FreeSIS: `urllib` + JSON API (AUM/NAV/펀드수 3모드)
  - KOFIA 회원사: `requests` + `BeautifulSoup` HTML 파싱 (`kofia.or.kr/members/m_61/`)
- `joiner.py` — 엑셀 조인 + 함수 연결 + 대시보드 + 검색 시트
  - `_build_raw_members_sheet()`: CSV 우선, 없으면 xlsx에서 읽기
- `cli.py` — CLI 진입점 (`run`, `join`)

## data/
- `raw/` — 원본 수집 데이터
  - `KOFIA_자산운용사_리스트.csv` — KOFIA 자산운용 회원사 (334개사, 회사명/대표자/대표전화/주소/웹사이트/로고)
- `processed/` — 최종 조인 엑셀
