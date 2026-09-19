# KOFIA 자산운용사 종합 프로파일

FreeSIS 설정규모 데이터와 KOFIA 회원사 정보를 조인하여 만든 종합 자산운용사 프로파일.

## 구조
- `data/raw/` — 원본 수집 데이터 (FreeSIS 3모드 + KOFIA 회원사)
- `data/processed/` — 조인 + 대시보드 + 검색이 포함된 최종 엑셀
- `src/kofia_am_profile/` — Python 패키지 (수집, 조인, 대시보드, 검색, CLI)

## 사용법
```bash
# 전체 파이프라인 실행
PYTHONPATH=src python3 -m kofia_am_profile.cli run

# 조인만 실행 (raw 데이터가 이미 있을 때)
PYTHONPATH=src python3 -m kofia_am_profile.cli join
```

## 엑셀 시트 구성
1. **통합_전체** — FreeSIS LEFT JOIN KOFIA (523행, Excel 함수로 RAW 참조)
2. **RAW_회원사** — KOFIA 회원사 원본 (334행)
3. **RAW_AUM** — FreeSIS 설정원본 원본 (524행)
4. **RAW_NAV** — FreeSIS 순자산 원본 (524행)
5. **RAW_펀드수** — FreeSIS 펀드수 원본 (524행)
6. **대시보드** — 요약 통계, TOP10, 투자유형별 비중
7. **검색** — 조건부 검색 (회사명, 대표자, AUM 등)
8. **비매칭분석** — FreeSIS에만 있는 회사 목록
