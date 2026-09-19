"""FreeSIS + KOFIA 회원사 데이터 수집 모듈."""

import json
import time
from http.cookiejar import CookieJar
from urllib import parse, request

BASE_URL = "https://freesis.kofia.or.kr"
SITEMAP_PATH = "/stats/siteMap.do"
FREESIS_ENTRYPOINT = "/stat/FreeSIS.do"
SERVICE_METADATA_ENDPOINT = "/meta/getSrvData.do"
GRID_DATA_ENDPOINT = "/meta/getMetaDataList.do"
SERVICE_ID = "STATFND0200100130"
PARENT_DIV_ID = "MSIS40200000000000"
DIVISION_ID = ""
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

# FreeSIS API TMPV 키 → 숨김 여부 (합계/위탁/전일/전년은 RAW에 포함하되 integrated에서 별도 처리)
_HIDDEN_TMPV = {"TMPV98"}


class FreesisClient:
    def __init__(self):
        self._cookie_jar = CookieJar()
        self._opener = request.build_opener(
            request.HTTPCookieProcessor(self._cookie_jar)
        )

    def reset_session(self):
        self._cookie_jar.clear()
        self._opener = request.build_opener(
            request.HTTPCookieProcessor(self._cookie_jar)
        )

    def _headers(self, referer=None):
        h = {"User-Agent": USER_AGENT, "Accept": "text/html,*/*"}
        if referer:
            h["Referer"] = referer
        return h

    def get_text(self, url):
        req = request.Request(url, headers=self._headers())
        with self._opener.open(req, timeout=20) as resp:
            body = resp.read().decode(
                resp.headers.get_content_charset() or "utf-8", errors="replace"
            )
            return resp.status, body

    def post_form(self, url, payload, referer=None):
        headers = self._headers(referer)
        headers["Content-Type"] = "application/x-www-form-urlencoded"
        encoded = parse.urlencode(payload).encode("utf-8")
        req = request.Request(url, data=encoded, headers=headers)
        with self._opener.open(req, timeout=20) as resp:
            body = resp.read().decode(
                resp.headers.get_content_charset() or "utf-8", errors="replace"
            )
            return resp.status, body

    def post_json(self, url, payload, referer=None):
        headers = self._headers(referer)
        headers["Content-Type"] = "application/json;charset=UTF-8"
        body_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = request.Request(url, data=body_bytes, headers=headers)
        with self._opener.open(req, timeout=60) as resp:
            raw = resp.read()
            charset = resp.headers.get_content_charset() or "utf-8"
            return resp.status, raw.decode(charset, errors="replace")


def bootstrap_session(client):
    client.get_text(BASE_URL + SITEMAP_PATH)
    time.sleep(0.3)
    client.post_form(
        BASE_URL + FREESIS_ENTRYPOINT,
        {"parentDivId": PARENT_DIV_ID, "serviceId": SERVICE_ID, "divisionId": DIVISION_ID},
        referer=BASE_URL + SITEMAP_PATH,
    )
    time.sleep(0.3)


def fetch_metadata(client):
    payload = {
        "dmSearchData": {
            "strSvrId": SERVICE_ID,
            "strDivId": DIVISION_ID,
            "strGetCode": "N",
            "language_gb": "KOR",
            "app_peron_yn": "Y",
        }
    }
    url = BASE_URL + SERVICE_METADATA_ENDPOINT
    referer = BASE_URL + FREESIS_ENTRYPOINT
    status, body = client.post_json(url, payload, referer=referer)
    if not body.strip().startswith("{"):
        return None
    return json.loads(body)


def build_column_map(metadata):
    """ORDER_SEQ 기준으로 정렬된 컬럼 리스트와 TMPV→라벨 매핑을 반환.

    Returns:
        ordered_keys: [(tmpv_key, label, order_seq), ...] FreeSIS UI 순서
        key_to_label: {tmpv_key: label}
    """
    headers = [
        row
        for row in metadata.get("dsGrid", [])
        if str(row.get("HEADER_ID", "")).startswith("H")
    ]
    headers.sort(key=lambda r: r.get("ORDER_SEQ", 9999))

    ordered_keys = []
    key_to_label = {}
    for row in headers:
        hid = str(row["HEADER_ID"])
        idx = int(hid.replace("H", ""))
        tmpv_key = f"TMPV{idx}"
        label = str(row.get("HEADER_NM", hid)).replace("\n", "")
        ordered_keys.append((tmpv_key, label, row.get("ORDER_SEQ", 0)))
        key_to_label[tmpv_key] = label
    return ordered_keys, key_to_label


def build_dm_search(metadata, date_override=None):
    latest_date = None
    for row in metadata.get("dsLatestDate", []):
        if row.get("TMPV1") == "RD":
            latest_date = row.get("TMPV2")
            break
    if not latest_date:
        for row in metadata.get("dsLatestDate", []):
            latest_date = row.get("TMPV2")
            break

    if date_override:
        latest_date = date_override

    grid_sql = metadata.get("dsGridSQL", [])
    obj_nm = (
        grid_sql[0].get("OBJ_NM")
        if grid_sql and grid_sql[0].get("OBJ_NM")
        else f"{SERVICE_ID}BO"
    )

    code_groups = {}
    for row in metadata.get("dsSearchCdList", []):
        group_cd = str(row.get("GROUP_CD", ""))
        if group_cd:
            code_groups.setdefault(group_cd, []).append(row)

    dm_search = {"OBJ_NM": obj_nm}

    for row in metadata.get("dsSearch", []):
        param = (row.get("SRCH_METAVAR") or "").strip()
        if not param:
            continue
        default_raw = (row.get("D_DISPLAY_DSPVLU") or "").strip()
        code_desc = (row.get("CODE_DESC") or "").strip()
        output_typ = str(row.get("OUTPUT_TYP") or "")
        output_func = (row.get("DISPLAY_OUTPUTFUNC") or "").strip()

        if code_desc == "최근일자" or output_typ == "D":
            if latest_date:
                dm_search[param] = latest_date
            continue

        if output_func and output_func in code_groups:
            options = code_groups[output_func]
            if options:
                if (
                    default_raw
                    and len(default_raw) >= 2
                    and default_raw.startswith("'")
                    and default_raw.endswith("'")
                ):
                    val = default_raw[1:-1]
                elif default_raw and default_raw != "*":
                    val = default_raw
                else:
                    continue
                if val and val != "*":
                    dm_search[param] = val
                continue

        if default_raw:
            if (
                len(default_raw) >= 2
                and default_raw.startswith("'")
                and default_raw.endswith("'")
            ):
                dm_search[param] = default_raw[1:-1]
            elif default_raw != "*":
                dm_search[param] = default_raw

    if "tmpV40" not in dm_search:
        dm_search["tmpV40"] = "1"
    if "tmpV41" not in dm_search:
        dm_search["tmpV41"] = "1"

    return dm_search


def fetch_grid_data(client, dm_search, max_retries=5):
    url = BASE_URL + GRID_DATA_ENDPOINT
    referer = BASE_URL + FREESIS_ENTRYPOINT

    for attempt in range(max_retries):
        if attempt > 0:
            client.reset_session()
            bootstrap_session(client)
            time.sleep(2 + attempt)

        status, body = client.post_json(url, {"dmSearch": dm_search}, referer=referer)
        if status != 200 or not body or not body.strip().startswith("{"):
            continue
        try:
            data = json.loads(body)
        except json.JSONDecodeError:
            continue
        raw_rows = data.get("ds1", []) or []
        rows = [r for r in raw_rows if str(r.get("TMPV98", "")) in ("00", "20")]
        if rows:
            return rows
    return []


def write_raw_sheet(ws, rows, ordered_keys, key_to_label):
    """ORDER_SEQ 순서대로 RAW 시트에 데이터 기록.

    컬럼: A=회사명, B=TMPV98(숨김), C~ = ordered_keys 순서대로
    TMPV98은 숨김 처리 (필터용)
    """
    col_headers = [("TMPV1", "회사명"), ("TMPV98", "상태코드")]
    for tmpv_key, label, _ in ordered_keys:
        if tmpv_key == "TMPV1":
            continue  # 회사명은 이미 첫 번째
        if tmpv_key == "TMPV98":
            continue  # 상태코드는 이미 두 번째
        col_headers.append((tmpv_key, label))

    for ci, (_, label) in enumerate(col_headers, 1):
        _apply_header(ws, 1, ci, label)

    for ri, row in enumerate(rows, 2):
        for ci, (key, _) in enumerate(col_headers, 1):
            val = row.get(key, "")
            if key == "TMPV98":
                val = ""
            _apply_data_cell(
                ws, ri, ci, val, is_number=isinstance(val, (int, float))
            )

    for ci, (_, label) in enumerate(col_headers, 1):
        ws.column_dimensions[_col_letter(ci)].width = max(len(label) * 2.5, 14)

    ws.column_dimensions["B"].hidden = True
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "C2"


def _apply_header(ws, row, col, value):
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    DARK_BLUE = "1F3864"
    BORDER_GRAY = "D9D9D9"
    HEADER_FONT = Font(bold=True, color="FFFFFF", size=10, name="맑은 고딕")
    HEADER_FILL = PatternFill(
        start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid"
    )
    THIN_BORDER = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.border = THIN_BORDER
    cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    return cell


def _apply_data_cell(ws, row, col, value, is_number=False):
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    BORDER_GRAY = "D9D9D9"
    STRIPE_GRAY = "F2F2F2"
    DATA_FONT = Font(size=10, name="맑은 고딕")
    STRIPE_FILL = PatternFill(
        start_color=STRIPE_GRAY, end_color=STRIPE_GRAY, fill_type="solid"
    )
    THIN_BORDER = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )
    NUM_FMT = "#,##0"
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = DATA_FONT
    cell.border = THIN_BORDER
    if row % 2 == 0:
        cell.fill = STRIPE_FILL
    if is_number:
        cell.number_format = NUM_FMT
        cell.alignment = Alignment(horizontal="right")
    return cell


def _col_letter(n):
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _validate_multi_company(results, ordered_keys):
    """설정원본/NAV/펀드수 교차검증. AUM 합계 = 유형별 합계 확인."""
    TARGETS = [
        "삼성자산운용", "미래에셋자산운용", "한국투자자산운용",
        "KB자산운용", "신한자산운용", "NH아グリ-investment자산운용",
        "DGB자산운용", "한화자산운용", "대신자산운용", "유진자산운용",
    ]
    aum_map = {r["TMPV1"]: r for r in results["설정원본(AUM)"]["rows"]}
    nav_map = {r["TMPV1"]: r for r in results["순자산(NAV)"]["rows"]}
    fund_map = {r["TMPV1"]: r for r in results["펀드수"]["rows"]}

    skip_keys = {"TMPV1", "TMPV98", "TMPV15", "TMPV16", "TMPV17", "TMPV18"}
    invest_tmpv = [k for k, _, _ in ordered_keys if k not in skip_keys]

    print("\n┌─ 멀티사 설정원본 검증 ─────────────────────────────────────┐")
    checked = 0
    passed = 0
    for name in TARGETS:
        aum = aum_map.get(name)
        if not aum:
            continue
        checked += 1
        types_sum = sum(aum.get(k, 0) for k in invest_tmpv)
        aum_total = aum.get("TMPV15", 0)
        if types_sum == aum_total:
            passed += 1
            print(f"│ ✓ {name}: 유형합계={types_sum:>18,} = 합계컬럼={aum_total:>18,}")
        else:
            diff = aum_total - types_sum
            print(f"│ ✗ {name}: 유형합계={types_sum:>18,} ≠ 합계컬럼={aum_total:>18,} (차={diff:,})")
    print(f"│ ── 결과: {passed}/{checked} 통과 ──")
    print("└─────────────────────────────────────────────────────────────┘")

    print("\n┌─ 멀티사 NAV 검증 ──────────────────────────────────────────┐")
    nav_checked = 0
    nav_passed = 0
    for name in TARGETS:
        nav = nav_map.get(name)
        if not nav:
            continue
        nav_checked += 1
        types_sum = sum(nav.get(k, 0) for k in invest_tmpv)
        nav_total = nav.get("TMPV15", 0)
        if types_sum == nav_total:
            nav_passed += 1
            print(f"│ ✓ {name}: 유형합계={types_sum:>18,} = 합계컬럼={nav_total:>18,}")
        else:
            diff = nav_total - types_sum
            print(f"│ ✗ {name}: 유형합계={types_sum:>18,} ≠ 합계컬럼={nav_total:>18,} (차={diff:,})")
    print(f"│ ── 결과: {nav_passed}/{nav_checked} 통과 ──")
    print("└─────────────────────────────────────────────────────────────┘")

    print("\n┌─ 멀티사 AUM vs NAV 일관성 ──────────────────────────────────┐")
    for name in TARGETS:
        aum = aum_map.get(name)
        nav = nav_map.get(name)
        fund = fund_map.get(name)
        if not aum or not nav:
            continue
        aum_t = aum.get("TMPV15", 0)
        nav_t = nav.get("TMPV15", 0)
        ratio = (nav_t / aum_t * 100) if aum_t else 0
        fund_t = fund.get("TMPV15", 0) if fund else 0
        flag = "✓" if 50 < ratio < 200 else "⚠"
        print(f"│ {flag} {name}: AUM={aum_t:>18,}  NAV={nav_t:>18,}  비율={ratio:.1f}%  펀드수={fund_t:,}")
    print("└─────────────────────────────────────────────────────────────┘")


def collect_kofia_members():
    """KOFIA 자산운용 회원사 리스트 수집 (kofia.or.kr).

    Returns:
        list[dict]: [{"company": str, "ceo": str, "phone": str, "address": str,
                       "website": str|None, "logo": str|None}, ...]
    """
    import requests as _requests
    from bs4 import BeautifulSoup as _BS

    base_url = "https://www.kofia.or.kr/members/m_61/sub020201.do"
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    }

    def _fetch_page(page_num):
        params = {
            "page": str(page_num),
            "srchCate": "5",
            "srchTp": "",
            "srchWord": "",
            "pitem": "10",
        }
        resp = _requests.get(base_url, params=params, headers=headers, timeout=20)
        resp.raise_for_status()
        return _BS(resp.text, "html.parser")

    def _get_total_pages(soup):
        last_img = soup.select_one("img.next2")
        if last_img:
            parent_a = last_img.find_parent("a")
            if parent_a and parent_a.get("href"):
                from urllib.parse import parse_qs, urlparse
                qs = parse_qs(urlparse(parent_a["href"]).query)
                return int(qs.get("page", [0])[0])
        return 1

    def _parse_companies(soup):
        companies = []
        for li in soup.select("ul.member11 > li"):
            h5 = li.select_one("h5")
            if not h5:
                continue
            name = h5.get_text(strip=True)
            dds = li.select("dd.script")
            ceo = dds[0].get_text(strip=True).replace("\xa0", "") if len(dds) > 0 else ""
            phone = dds[1].get_text(strip=True).replace("\xa0", "") if len(dds) > 1 else ""
            address = dds[2].get_text(strip=True).replace("\xa0", "") if len(dds) > 2 else ""

            website_tag = li.select_one("p.alliesimage a")
            website = website_tag["href"].strip() if website_tag and website_tag.get("href") else None
            logo_tag = li.select_one("p.alliesimage img")
            logo = logo_tag["src"].strip() if logo_tag and logo_tag.get("src") else None

            companies.append({
                "company": name, "ceo": ceo, "phone": phone,
                "address": address, "website": website, "logo": logo,
            })
        return companies

    print("  [KOFIA 회원사] 페이지 1 수집 중...")
    soup = _fetch_page(1)
    total_pages = _get_total_pages(soup)
    all_companies = _parse_companies(soup)
    print(f"  [KOFIA 회원사] 전체 {total_pages}페이지, 1/{total_pages} 완료 ({len(all_companies)}개사)")

    for page in range(2, total_pages + 1):
        time.sleep(0.5)
        soup = _fetch_page(page)
        page_companies = _parse_companies(soup)
        all_companies.extend(page_companies)
        if page % 5 == 0 or page == total_pages:
            print(f"  [KOFIA 회원사] {page}/{total_pages} 완료 (누적 {len(all_companies)}개사)")

    print(f"  [KOFIA 회원사] 수집 완료: {len(all_companies)}개사")
    return all_companies


def collect_freesis_all_modes(client, metadata, ordered_keys, key_to_label, date_override=None):
    modes = {
        "설정원본(AUM)": "1",
        "순자산(NAV)": "2",
        "펀드수": "3",
    }
    results = {}
    for label, code in modes.items():
        dm_search = build_dm_search(metadata, date_override=date_override)
        dm_search["tmpV100"] = code
        if results:
            time.sleep(2)
            client.reset_session()
            bootstrap_session(client)
            time.sleep(1)
        else:
            time.sleep(0.5)
        rows = fetch_grid_data(client, dm_search)
        results[label] = {"rows": rows, "dm_search": dm_search}
        print(f"  {label}: {len(rows)}개사 수집 (기준일={dm_search.get('tmpV34')})")

    _validate_multi_company(results, ordered_keys)
    return results


def save_freesis_csv(results, raw_dir, ordered_keys):
    """FreeSIS 수집 결과를 data/raw/에 CSV로 저장.

    각 모드(AUM/NAV/펀드수)별로 별도 CSV 파일 생성.
    RAW 데이터 보존을 위해 절대 기존 CSV를 덮어쓰지 않음 (检查 후 저장).
    """
    import csv as _csv
    import os

    os.makedirs(raw_dir, exist_ok=True)

    mode_to_filename = {
        "설정원본(AUM)": "KOFIA_FreeSIS_설정원본.csv",
        "순자산(NAV)": "KOFIA_FreeSIS_순자산.csv",
        "펀드수": "KOFIA_FreeSIS_펀드수.csv",
    }

    # ordered_keys에서 회사명(TMPV1)과 상태코드(TMPV98) 제외, 투자유형+합계+위탁+전일+전년만
    tmpv_keys = [(k, label) for k, label, _ in ordered_keys if k not in ("TMPV1", "TMPV98")]

    # CSV 헤더: 회사명 + 투자유형 컬럼들
    header = ["회사명"] + [label for _, label in tmpv_keys]

    for mode_label, data in results.items():
        filename = mode_to_filename.get(mode_label)
        if not filename:
            continue
        filepath = os.path.join(raw_dir, filename)

        rows = data["rows"]
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = _csv.writer(f)
            writer.writerow(header)
            for row in rows:
                company = row.get("TMPV1", "")
                values = [row.get(k, "") for k, _ in tmpv_keys]
                writer.writerow([company] + values)

        print(f"  CSV 저장: {filename} ({len(rows)}행)")
