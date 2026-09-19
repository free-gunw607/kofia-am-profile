"""엑셀 조인 + 함수 연결 모듈 (V4 — 올드스쿨 INDEX/MATCH, Excel 전 버전 호환)."""

import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

# ── 스타일 상수 ──────────────────────────────────────────────
DARK_BLUE = "1F3864"
MID_BLUE = "4472C4"
LIGHT_BLUE = "D6E4F0"
ACCENT_RED = "CC0000"
STRIPE_GRAY = "F2F2F2"
BORDER_GRAY = "D9D9D9"

HEADER_FONT = Font(bold=True, color="FFFFFF", size=10, name="맑은 고딕")
HEADER_FILL = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
SUB_HEADER_FILL = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
DATA_FONT = Font(size=10, name="맑은 고딕")
STRIPE_FILL = PatternFill(start_color=STRIPE_GRAY, end_color=STRIPE_GRAY, fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin", color=BORDER_GRAY),
    right=Side(style="thin", color=BORDER_GRAY),
    top=Side(style="thin", color=BORDER_GRAY),
    bottom=Side(style="thin", color=BORDER_GRAY),
)
NUM_FMT = "#,##0"
ACCENT_FONT = Font(bold=True, color=ACCENT_RED, size=10, name="맑은 고딕")
BOLD_FONT = Font(bold=True, size=10, name="맑은 고딕")
TITLE_FONT = Font(bold=True, size=14, name="맑은 고딕", color=DARK_BLUE)
SECTION_FONT = Font(bold=True, size=12, name="맑은 고딕", color=DARK_BLUE)
NOTE_FONT = Font(italic=True, color="666666", size=9, name="맑은 고딕")

# ── FreeSIS UI 컬럼 순서 ────────────────────────────────────
INVEST_TYPES = [
    "주식", "혼합주식", "혼합채권", "채권",
    "투자계약", "재간접", "단기금융", "파생형", "부동산", "실물",
    "특별자산", "혼합자산",
    "기업성장", "기관전용사모펀드", "투자일임기타",
]

RAW_INVEST_START = 3
RAW_TOTAL_COL = RAW_INVEST_START + len(INVEST_TYPES)  # 18 = R
RAW_CUSTODY_COL = RAW_TOTAL_COL + 1  # 19 = S
RAW_PREVYEAR_COL = RAW_TOTAL_COL + 3  # 21 = U


def _col_letter(n):
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _apply_header(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell


def _apply_data_cell(ws, row, col, value, is_number=False, is_accent=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = ACCENT_FONT if is_accent else DATA_FONT
    cell.border = THIN_BORDER
    if row % 2 == 0:
        cell.fill = STRIPE_FILL
    if is_number:
        cell.number_format = NUM_FMT
        cell.alignment = Alignment(horizontal="right")
    return cell


def _apply_formula_cell(ws, row, col, formula, is_number=False, is_accent=False):
    cell = ws.cell(row=row, column=col, value=formula)
    cell.font = ACCENT_FONT if is_accent else DATA_FONT
    cell.border = THIN_BORDER
    if row % 2 == 0:
        cell.fill = STRIPE_FILL
    if is_number:
        cell.number_format = NUM_FMT
        cell.alignment = Alignment(horizontal="right")
    return cell


def _set_col_widths(ws, widths):
    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[_col_letter(col_idx)].width = w


def _set_tab_color(ws, color):
    ws.sheet_properties.tabColor = color


# ── 메인 진입점 ──────────────────────────────────────────────
def build_joined_excel(freesis_results, ordered_keys, key_to_label, metadata, output_path):
    wb = Workbook()

    aum_data = freesis_results["설정원본(AUM)"]["rows"]
    nav_data = freesis_results["순자산(NAV)"]["rows"]
    fund_data = freesis_results["펀드수"]["rows"]
    total_rows = len(aum_data)
    last_row = total_rows + 1

    _build_raw_aum_sheet(wb, aum_data, ordered_keys, key_to_label)
    _build_raw_nav_sheet(wb, nav_data, ordered_keys, key_to_label)
    _build_raw_fund_sheet(wb, fund_data, ordered_keys, key_to_label)
    _build_raw_members_sheet(wb, output_path)

    _build_aum_sheet(wb, total_rows)
    _build_nav_sheet(wb, total_rows)
    _build_fund_sheet(wb, total_rows)
    _build_summary_sheet(wb, total_rows)
    _build_search_sheet(wb, total_rows)
    _build_unmatched_sheet(wb, total_rows)

    wb.save(output_path)
    print(f"Saved: {output_path}")
    return wb


# ── RAW 시트 ─────────────────────────────────────────────────
def _build_raw_aum_sheet(wb, rows, ordered_keys, key_to_label):
    ws = wb.active
    ws.title = "RAW_AUM"
    from kofia_am_profile.collector import write_raw_sheet
    write_raw_sheet(ws, rows, ordered_keys, key_to_label)
    _set_tab_color(ws, "A5A5A5")
    print(f"  RAW_AUM: {len(rows)} rows x {ws.max_column} cols")


def _build_raw_nav_sheet(wb, rows, ordered_keys, key_to_label):
    ws = wb.create_sheet("RAW_NAV")
    from kofia_am_profile.collector import write_raw_sheet
    write_raw_sheet(ws, rows, ordered_keys, key_to_label)
    _set_tab_color(ws, "A5A5A5")
    print(f"  RAW_NAV: {len(rows)} rows x {ws.max_column} cols")


def _build_raw_fund_sheet(wb, rows, ordered_keys, key_to_label):
    ws = wb.create_sheet("RAW_펀드수")
    from kofia_am_profile.collector import write_raw_sheet
    write_raw_sheet(ws, rows, ordered_keys, key_to_label)
    _set_tab_color(ws, "A5A5A5")
    print(f"  RAW_펀드수: {len(rows)} rows x {ws.max_column} cols")


def _build_raw_members_sheet(wb, output_path):
    ws = wb.create_sheet("RAW_회원사")
    headers = ["회사명", "대표자", "대표전화", "주소", "웹사이트URL", "로고이미지URL"]
    for col_idx, h in enumerate(headers, 1):
        _apply_header(ws, 1, col_idx, h)

    raw_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(output_path))),
        "data", "raw",
    )

    csv_path = os.path.join(raw_dir, "KOFIA_자산운용사_리스트.csv")
    xlsx_path = os.path.join(raw_dir, "KOFIA_자산운용사_리스트_raw.xlsx")

    if os.path.exists(csv_path):
        import csv as _csv
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = list(_csv.DictReader(f))
        for ri, row in enumerate(reader, 2):
            _apply_data_cell(ws, ri, 1, row.get("company", ""))
            _apply_data_cell(ws, ri, 2, row.get("ceo", ""))
            _apply_data_cell(ws, ri, 3, row.get("phone", ""))
            _apply_data_cell(ws, ri, 4, row.get("address", ""))
            _apply_data_cell(ws, ri, 5, row.get("website", ""))
            _apply_data_cell(ws, ri, 6, row.get("logo", ""))
    elif os.path.exists(xlsx_path):
        src_wb = load_workbook(xlsx_path)
        src_ws = src_wb.active
        for row in range(2, src_ws.max_row + 1):
            for col in range(1, 7):
                _apply_data_cell(ws, row, col, src_ws.cell(row, col).value)
        src_wb.close()
    _set_tab_color(ws, "A5A5A5")
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    print(f"  RAW_회원사: {ws.max_row - 1} rows")


# ── AUM 투자유형별 시트 ──────────────────────────────────────
def _build_aum_sheet(wb, total_rows):
    ws = wb.create_sheet("AUM_투자유형별", 0)
    _set_tab_color(ws, MID_BLUE)
    last_row = total_rows + 1

    headers = ["회사명", "매칭여부", "대표자", "대표전화"]
    for t in INVEST_TYPES:
        headers.append(f"AUM_{t}")
    headers.extend(["AUM_합계", "AUM_위탁운용", "AUM_전년대비"])

    for col_idx, h in enumerate(headers, 1):
        _apply_header(ws, 1, col_idx, h)

    for row_idx in range(2, last_row + 1):
        r = row_idx
        _apply_formula_cell(ws, r, 1, f"=RAW_AUM!A{r}")
        _apply_formula_cell(ws, r, 2, f'=IF(COUNTIF(RAW_회원사!A:A,A{r})>0,"매칭","비매칭")')
        _apply_formula_cell(ws, r, 3, f'=IFERROR(VLOOKUP(A{r},RAW_회원사!A:B,2,FALSE),"")')
        _apply_formula_cell(ws, r, 4, f'=IFERROR(VLOOKUP(A{r},RAW_회원사!A:C,3,FALSE),"")')

        for i in range(len(INVEST_TYPES)):
            raw_col = _col_letter(RAW_INVEST_START + i)
            _apply_formula_cell(ws, r, 5 + i, f"=RAW_AUM!{raw_col}{r}", is_number=True)

        total_col = _col_letter(RAW_TOTAL_COL)
        custody_col = _col_letter(RAW_CUSTODY_COL)
        prevyear_col = _col_letter(RAW_PREVYEAR_COL)
        _apply_formula_cell(ws, r, 5 + len(INVEST_TYPES), f"=RAW_AUM!{total_col}{r}", is_number=True, is_accent=True)
        _apply_formula_cell(ws, r, 6 + len(INVEST_TYPES), f"=RAW_AUM!{custody_col}{r}", is_number=True)
        _apply_formula_cell(ws, r, 7 + len(INVEST_TYPES), f"=RAW_AUM!{prevyear_col}{r}", is_number=True)

    widths = [28, 10, 15, 18] + [16] * len(INVEST_TYPES) + [20, 16, 16]
    _set_col_widths(ws, widths)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "E2"
    print(f"  AUM_투자유형별: {total_rows} rows x {len(headers)} cols")


# ── NAV 투자유형별 시트 ──────────────────────────────────────
def _build_nav_sheet(wb, total_rows):
    ws = wb.create_sheet("NAV_투자유형별", 1)
    _set_tab_color(ws, "70AD47")
    last_row = total_rows + 1

    headers = ["회사명", "매칭여부"]
    for t in INVEST_TYPES:
        headers.append(f"NAV_{t}")
    headers.append("NAV_합계")

    for col_idx, h in enumerate(headers, 1):
        _apply_header(ws, 1, col_idx, h)

    for row_idx in range(2, last_row + 1):
        r = row_idx
        _apply_formula_cell(ws, r, 1, f"=RAW_NAV!A{r}")
        _apply_formula_cell(ws, r, 2, f'=IF(COUNTIF(RAW_회원사!A:A,A{r})>0,"매칭","비매칭")')

        for i in range(len(INVEST_TYPES)):
            raw_col = _col_letter(RAW_INVEST_START + i)
            _apply_formula_cell(ws, r, 3 + i, f"=RAW_NAV!{raw_col}{r}", is_number=True)

        total_col = _col_letter(RAW_TOTAL_COL)
        _apply_formula_cell(ws, r, 3 + len(INVEST_TYPES), f"=RAW_NAV!{total_col}{r}", is_number=True, is_accent=True)

    widths = [28, 10] + [16] * len(INVEST_TYPES) + [20]
    _set_col_widths(ws, widths)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "C2"
    print(f"  NAV_투자유형별: {total_rows} rows x {len(headers)} cols")


# ── 펀드수 투자유형별 시트 ──────────────────────────────────
def _build_fund_sheet(wb, total_rows):
    ws = wb.create_sheet("펀드수_투자유형별", 2)
    _set_tab_color(ws, "ED7D31")
    last_row = total_rows + 1

    headers = ["회사명", "매칭여부"]
    for t in INVEST_TYPES:
        headers.append(f"펀드수_{t}")
    headers.append("펀드수_합계")

    for col_idx, h in enumerate(headers, 1):
        _apply_header(ws, 1, col_idx, h)

    for row_idx in range(2, last_row + 1):
        r = row_idx
        _apply_formula_cell(ws, r, 1, f"=RAW_펀드수!A{r}")
        _apply_formula_cell(ws, r, 2, f'=IF(COUNTIF(RAW_회원사!A:A,A{r})>0,"매칭","비매칭")')

        for i in range(len(INVEST_TYPES)):
            raw_col = _col_letter(RAW_INVEST_START + i)
            _apply_formula_cell(ws, r, 3 + i, f"=RAW_펀드수!{raw_col}{r}", is_number=True)

        total_col = _col_letter(RAW_TOTAL_COL)
        _apply_formula_cell(ws, r, 3 + len(INVEST_TYPES), f"=RAW_펀드수!{total_col}{r}", is_number=True, is_accent=True)

    widths = [28, 10] + [14] * len(INVEST_TYPES) + [16]
    _set_col_widths(ws, widths)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "C2"
    print(f"  펀드수_투자유형별: {total_rows} rows x {len(headers)} cols")


# ── 요약 시트 ────────────────────────────────────────────────
def _build_summary_sheet(wb, total_rows):
    ws = wb.create_sheet("요약", 3)
    _set_tab_color(ws, "7030A0")
    last_row = total_rows + 1

    ws.cell(1, 1, "KOFIA 자산운용사 종합 프로파일").font = TITLE_FONT
    ws.cell(2, 1, "기준일:").font = BOLD_FONT
    ws.cell(2, 2, '=RAW_AUM!A2&" 설정원본 기준"').font = DATA_FONT
    ws.cell(3, 1, "단위: 원(설정/NAV), 개(펀드수)").font = NOTE_FONT

    ws.cell(5, 1, "■ 전체 현황").font = SECTION_FONT
    ws.cell(5, 1).fill = SUB_HEADER_FILL
    for c in range(2, 4):
        ws.cell(5, c).fill = SUB_HEADER_FILL

    total_col = _col_letter(RAW_TOTAL_COL)

    summary_items = [
        ("총 운용사 수", f'=COUNTA(AUM_투자유형별!A2:A{last_row})'),
        ("매칭 회사 수", f'=COUNTIF(AUM_투자유형별!B2:B{last_row},"매칭")'),
        ("비매칭 회사 수", f'=COUNTIF(AUM_투자유형별!B2:B{last_row},"비매칭")'),
        ("전체 AUM 합계", f"=RAW_AUM!{total_col}{last_row}"),
        ("전체 순자산(NAV)", f"=RAW_NAV!{total_col}{last_row}"),
        ("전체 펀드 수", f"=RAW_펀드수!{total_col}{last_row}"),
    ]
    for i, (label, formula) in enumerate(summary_items):
        r = 6 + i
        ws.cell(r, 1, label).font = BOLD_FONT
        ws.cell(r, 2, formula).font = DATA_FONT
        ws.cell(r, 2).number_format = NUM_FMT
        ws.cell(r, 1).border = THIN_BORDER
        ws.cell(r, 2).border = THIN_BORDER

    ws.cell(13, 1, "■ 투자 유형별 AUM 비중").font = SECTION_FONT
    ws.cell(13, 1).fill = SUB_HEADER_FILL
    for c in range(2, 4):
        ws.cell(13, c).fill = SUB_HEADER_FILL

    ws.cell(14, 1, "투자 유형").font = HEADER_FONT
    ws.cell(14, 1).fill = HEADER_FILL
    ws.cell(14, 2, "AUM 합계").font = HEADER_FONT
    ws.cell(14, 2).fill = HEADER_FILL
    ws.cell(14, 3, "비중").font = HEADER_FONT
    ws.cell(14, 3).fill = HEADER_FILL
    for c in range(1, 4):
        ws.cell(14, c).border = THIN_BORDER

    for i, name in enumerate(INVEST_TYPES):
        r = 15 + i
        raw_col = _col_letter(RAW_INVEST_START + i)
        ws.cell(r, 1, name).font = BOLD_FONT
        ws.cell(r, 1).border = THIN_BORDER
        ws.cell(r, 2, f"=RAW_AUM!{raw_col}{last_row}").font = DATA_FONT
        ws.cell(r, 2).number_format = NUM_FMT
        ws.cell(r, 2).border = THIN_BORDER
        ws.cell(r, 3, f"=IF(B$6=0,0,B{r}/B$6)").font = DATA_FONT
        ws.cell(r, 3).number_format = "0.0%"
        ws.cell(r, 3).border = THIN_BORDER
        if r % 2 == 0:
            for c in range(1, 4):
                ws.cell(r, c).fill = STRIPE_FILL

    ws.cell(31, 1, "■ TOP 10 운용사 (AUM 합계 기준)").font = SECTION_FONT
    ws.cell(31, 1).fill = SUB_HEADER_FILL
    for c in range(2, 6):
        ws.cell(31, c).fill = SUB_HEADER_FILL

    top_headers = ["순위", "회사명", "AUM 합계", "NAV 합계", "펀드 수"]
    for col_idx, h in enumerate(top_headers, 1):
        _apply_header(ws, 32, col_idx, h)

    for rank in range(1, 11):
        r = 32 + rank
        ws.cell(r, 1, rank).font = DATA_FONT
        ws.cell(r, 1).border = THIN_BORDER
        ws.cell(r, 2, f'=IFERROR(INDEX(AUM_투자유형별!A$2:A${last_row},MATCH(LARGE(AUM_투자유형별!{total_col}$2:{total_col}${last_row},{rank}),AUM_투자유형별!{total_col}$2:{total_col}${last_row},0)),"")').font = DATA_FONT
        ws.cell(r, 2).border = THIN_BORDER
        ws.cell(r, 3, f"=IFERROR(LARGE(AUM_투자유형별!{total_col}$2:{total_col}${last_row},{rank}),0)").font = DATA_FONT
        ws.cell(r, 3).number_format = NUM_FMT
        ws.cell(r, 3).border = THIN_BORDER
        ws.cell(r, 4, f"=IFERROR(LARGE(NAV_투자유형별!{total_col}$2:{total_col}${last_row},{rank}),0)").font = DATA_FONT
        ws.cell(r, 4).number_format = NUM_FMT
        ws.cell(r, 4).border = THIN_BORDER
        ws.cell(r, 5, f"=IFERROR(LARGE(펀드수_투자유형별!{total_col}$2:{total_col}${last_row},{rank}),0)").font = DATA_FONT
        ws.cell(r, 5).number_format = NUM_FMT
        ws.cell(r, 5).border = THIN_BORDER
        if r % 2 == 0:
            for c in range(1, 6):
                ws.cell(r, c).fill = STRIPE_FILL

    _set_col_widths(ws, [25, 28, 20, 20, 15])
    print("  요약: created")


# ── 검색 시트 (OLD SCHOOL: INDEX/MATCH + AGGREGATE) ─────────
def _build_search_sheet(wb, total_rows):
    ws = wb.create_sheet("검색", 4)
    _set_tab_color(ws, "FFC000")
    last_row = total_rows + 1

    ws.cell(1, 1, "검색 조건").font = SECTION_FONT
    ws.cell(1, 1).fill = SUB_HEADER_FILL
    ws.cell(1, 2).fill = SUB_HEADER_FILL

    ws.cell(2, 1, "회사명:").font = BOLD_FONT
    ws.cell(2, 1).border = THIN_BORDER
    ws.cell(2, 2).border = THIN_BORDER
    ws.cell(2, 2).font = DATA_FONT

    ws.cell(3, 1, "매칭여부:").font = BOLD_FONT
    ws.cell(3, 1).border = THIN_BORDER
    ws.cell(3, 2).border = THIN_BORDER
    ws.cell(3, 2, "전체").font = DATA_FONT
    ws.cell(3, 2).border = THIN_BORDER

    dv = DataValidation(type="list", formula1='"전체,매칭,비매칭"', allow_blank=False)
    dv.error = "선택해주세요"
    ws.add_data_validation(dv)
    dv.add(ws["B3"])

    ws.cell(5, 1, "▼ 검색 결과 (회사명 입력 후 Enter)").font = SECTION_FONT
    ws.cell(5, 1).fill = SUB_HEADER_FILL
    for c in range(2, 20):
        ws.cell(5, c).fill = SUB_HEADER_FILL

    result_headers = ["회사명", "매칭여부", "대표자", "대표전화"]
    for t in INVEST_TYPES:
        result_headers.append(f"AUM_{t}")
    result_headers.extend(["AUM_합계"])

    for col_idx, h in enumerate(result_headers, 1):
        _apply_header(ws, 6, col_idx, h)

    # B2 = 회사명 검색어, B3 = 매칭여부
    # 검색 결과는행 7부터 50행까지 (최대 50개)

    # 매칭 회사의 행 번호를 찾는 공식 (AGGREGATE + ROW)
    # AGGREGATE(15,6,ROW(A2:A521) / (조건), k) → k번째 매칭 행 번호 반환
    # 단, AGGREGATE는 Excel 2010+ 필요 (구버전은 SMALL/IF 배열수식)

    aum_range = f"AUM_투자유형별!A$2:A${last_row}"
    b_range = f"AUM_투자유형별!B$2:B${last_row}"
    repm_range = f"AUM_투자유형별!C$2:C${last_row}"
    phone_range = f"AUM_투자유형별!D$2:D${last_row}"

    # MATCH("*검색어*", range, 0) → 와일드카드 부분일치, 첫 번째 매칭 위치 반환
    match_pos = f'MATCH("*"&B$2&"*",{aum_range},0)'

    # 행 7에만 첫 번째 매칭 회사 표시 (INDEX + MATCH)
    _apply_formula_cell(ws, 7, 1, f'=IFERROR(INDEX({aum_range},{match_pos}),"")')
    _apply_formula_cell(ws, 7, 2, f'=IFERROR(INDEX({b_range},{match_pos}),"")')
    _apply_formula_cell(ws, 7, 3, f'=IFERROR(INDEX({repm_range},{match_pos}),"")')
    _apply_formula_cell(ws, 7, 4, f'=IFERROR(INDEX({phone_range},{match_pos}),"")')

    for i in range(len(INVEST_TYPES)):
        src_col = _col_letter(5 + i)
        src_range = f"AUM_투자유형별!{src_col}$2:{src_col}${last_row}"
        _apply_formula_cell(ws, 7, 5 + i,
            f'=IFERROR(INDEX({src_range},{match_pos}),0)', is_number=True)

    total_src_col = _col_letter(5 + len(INVEST_TYPES))
    _apply_formula_cell(ws, 7, 5 + len(INVEST_TYPES),
        f'=IFERROR(INDEX(AUM_투자유형별!{total_src_col}$2:{total_src_col}${last_row},{match_pos}),0)', is_number=True, is_accent=True)

    # 행 8~56: 안내문
    for r in range(8, 57):
        for c in range(1, 6 + len(INVEST_TYPES)):
            ws.cell(r, c).value = ""
    ws.cell(8, 1, "※ 정확한 회사명을 입력하세요 (예: 삼성자산운용)").font = NOTE_FONT
    ws.cell(9, 1, "  부분 검색 시 첫 번째 매칭 회사만 표시됩니다").font = NOTE_FONT
    ws.cell(10, 1, "  전체 데이터는 RAW_AUM 시트의 필터를 이용하세요").font = NOTE_FONT

    widths = [28, 10, 15] + [16] * len(INVEST_TYPES) + [20]
    _set_col_widths(ws, widths)
    ws.freeze_panes = "A7"
    print(f"  검색: created (1 row MATCH formula)")


# ── 비매칭분석 시트 (OLD SCHOOL: INDEX/MATCH) ───────────────
def _build_unmatched_sheet(wb, total_rows):
    ws = wb.create_sheet("비매칭분석", 5)
    _set_tab_color(ws, "FF0000")
    last_row = total_rows + 1

    headers = ["회사명", "매칭여부", "대표자", "AUM_합계", "NAV_합계", "펀드수_합계"]
    for col_idx, h in enumerate(headers, 1):
        _apply_header(ws, 1, col_idx, h)

    total_col_ref = _col_letter(RAW_TOTAL_COL)

    # 비매칭 회사의 행 번호를 찾는 공식 (AUM_투자유형별!B열 매칭여부 기준)
    cond_unmatch = '(AUM_투자유형별!B$2:B${lr}="비매칭")'.format(lr=last_row)

    for rank in range(1, 201):
        r = rank + 1
        row_formula = f'AGGREGATE(15,6,ROW(AUM_투자유형별!A$2:A${last_row})/({cond_unmatch}),{rank})'

        _apply_formula_cell(ws, r, 1,
            f'=IFERROR(INDEX(AUM_투자유형별!A$2:A${last_row},{row_formula}-1),"")')
        _apply_formula_cell(ws, r, 2,
            f'=IFERROR(INDEX(AUM_투자유형별!B$2:B${last_row},{row_formula}-1),"")')
        _apply_formula_cell(ws, r, 3,
            f'=IFERROR(INDEX(AUM_투자유형별!C$2:C${last_row},{row_formula}-1),"")')
        _apply_formula_cell(ws, r, 4,
            f'=IFERROR(INDEX(AUM_투자유형별!{total_col_ref}$2:{total_col_ref}${last_row},{row_formula}-1),0)', is_number=True)
        _apply_formula_cell(ws, r, 5,
            f'=IFERROR(INDEX(NAV_투자유형별!{total_col_ref}$2:{total_col_ref}${last_row},{row_formula}-1),0)', is_number=True)
        _apply_formula_cell(ws, r, 6,
            f'=IFERROR(INDEX(펀드수_투자유형별!{total_col_ref}$2:{total_col_ref}${last_row},{row_formula}-1),0)', is_number=True)

    _set_col_widths(ws, [28, 10, 15, 20, 20, 16])
    ws.freeze_panes = "A2"
    print("  비매칭분석: created (200 rows, INDEX/MATCH+AGGREGATE)")
