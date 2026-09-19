# ENTRY

## 빠른 시작
```bash
cd ~/agent-coding/agent-projects/A4-worker-repos/kofia-am-profile
PYTHONPATH=src python3 -m kofia_am_profile.cli run
```

## 산출물
- `data/processed/KOFIA_자산운용사_종합프로파일.xlsx`

## 엑셀 시트 안내
1. **통합_전체** — 메인 조인 시트 (523행, Excel 함수로 RAW 참조)
2. **RAW_AUM/NAV/펀드수** — FreeSIS 원본
3. **RAW_회원사** — KOFIA 회원사 원본
4. **대시보드** — 요약 통계
5. **검색** — 조건부 검색
6. **비매칭분석** — FreeSIS 전용 회사 목록
