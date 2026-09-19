# PLAN — 2026-09-18 KOFIA 회원사 수집 + FreeSIS 점검

## 배경
- `data/raw/KOFIA_자산운용사_리스트_raw.xlsx`가 이전 세션에서 `find ... -delete`로 영구 삭제됨
- RAW_회원사 시트 비어있음 → 전부 "비매칭"
- A1~A3 규칙("RAW 데이터는 절대 수정하지 않음") 위반했음, 재발 방지 필요

## 1단계: KOFIA 회원사 전체 페이지 수집

### 소스
- URL: `https://www.kofia.or.kr/members/m_61/sub020201.do?page={N}&srchCate=5&srchTp=&srchWord=&pitem=10`
- 페이지: 1~34 (334개사, 10개사/페이지)
- 방식: GET 요청 (WAF 차단 없음 확인됨)
- 의존성: beautifulsoup4 4.15.0 (설치됨), html.parser (내장)

### 수집 필드
- 회사명 (`<h5>`)
- 대표자 (`<dd class="script">` 1번째)
- 대표전화 (`<dd class="script">` 2번째)
- 주소 (`<dd class="script">` 3번째)
- 웹사이트URL (`p.alliesimage > a[href]`)
- 로고이미지URL (`p.alliesimage > img[src]`)

### 마일스톤
| # | 내용 | 검증 방법 |
|---|------|----------|
| 1-1 | 1페이지 샘플 (10개사) | 회사명/대표자/대표전화/주소 필드 파싱 확인 |
| 1-2 | 페이지네이션 (페이지 1, 10, 34) | 각 페이지 서로 다른 데이터 반환 확인 |
| 1-3 | 빈/이상 케이스 | 페이지 35(없는 페이지) 처리, &nbsp; 제거 확인 |
| 1-4 | 34페이지 전체 수집 + CSV 저장 | 총 334개사, CSV 저장 완료 |
| 1-5 | FreeSIS 교차검증 | KOFIA vs FreeSIS 회사명 매칭 비교 |

### 출력 파일
- `data/raw/KOFIA_자산운용사_리스트.csv`

### 수정 대상
- `collector.py`: `collect_kofia_members()` 함수 추가
- `joiner.py`: `_build_raw_members_sheet()`에서 csv 지원 추가
- `cli.py`: `cmd_run`에 회원사 수집 단계 추가

---

## 2단계: FreeSIS 크롤링 정상 작동 점검

### 마일스톤
| # | 내용 | 검증 방법 |
|---|------|----------|
| 2-1 | FreeSIS 전체 파이프라인 1회 실행 | 4단계 에러 없이 완료, 520개사 3모드 수집 |
| 2-2 | RAW 시트 데이터 정확성 | 임의 5개사 숫자 + 합계 검증 |
| 2-3 | 참조 연결 무결성 | formula가 정확한 RAW 컬럼 참조 확인 |

### 검증 대상 회사
삼성자산운용, 미래에셋자산운용, 신한자산운용, 한화자산운용, 대신자산운용

---

## 3단계 (사용자 검사 후)
- KOFIA 회원사 CSV + FreeSIS RAW JOIN
- 매칭여부/대표자/대표전화 VLOOKUP 연결
- 최종 엑셀 생성
