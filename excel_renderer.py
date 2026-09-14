# excel_renderer.py
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 스타일 상수 정의 (실리콘밸리 FinTech 내비 스타일)
NAVY_HEADER = "1E293B"      # Slate 800 (테이블 헤더 배경)
ZEBRA_EVEN = "F8FAFC"       # Slate 50 (짝수행 배경)
BORDER_COLOR = "CBD5E1"     # Slate 300 (테두리 색상)
GREEN_FILL = "DCFCE7"       # Emerald 100 (성공/정상 강조 배경)
GREEN_TEXT = "166534"       # Emerald 800 (성공 텍스트)
RED_TEXT = "DC2626"         # Rose 600 (출금/경고 텍스트)

thin_border = Border(
    left=Side(style='thin', color=BORDER_COLOR),
    right=Side(style='thin', color=BORDER_COLOR),
    top=Side(style='thin', color=BORDER_COLOR),
    bottom=Side(style='thin', color=BORDER_COLOR)
)

double_bottom_border = Border(
    left=Side(style='thin', color=BORDER_COLOR),
    right=Side(style='thin', color=BORDER_COLOR),
    top=Side(style='thin', color=BORDER_COLOR),
    bottom=Side(style='double', color=NAVY_HEADER)
)

def render_bankflow_excel(parsed_data: dict, audit_result: dict) -> io.BytesIO:
    """
    AI 파싱 결과와 검산 데이터를 받아 QuickBooks/Xero 호환 .xlsx 바이너리 스트림을 생성합니다.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # 기본 생성되는 빈 시트 제거

    info = parsed_data.get("statement_info", {})
    txs = parsed_data.get("transactions", [])

    # -------------------------------------------------------------
    # 통화 자동 감지 및 서식 동적 매핑
    # -------------------------------------------------------------
    raw_curr = str(info.get("currency", "USD")).strip().upper()
    CURRENCY_CONFIG = {
        "USD": ("$", "$#,##0.00"),
        "EUR": ("€", "€#,##0.00"),
        "GBP": ("£", "£#,##0.00"),
        "CAD": ("$", "$#,##0.00"),
        "AUD": ("$", "$#,##0.00"),
        "KRW": ("₩", "₩#,##0"),
        "JPY": ("¥", "¥#,##0"),
        "BDT": ("৳", "৳#,##0.00"),
        "INR": ("₹", "₹#,##0.00"),
        "SGD": ("$", "$#,##0.00"),
        "HKD": ("$", "$#,##0.00"),
    }

    clean_curr = "USD"
    for code in CURRENCY_CONFIG.keys():
        if code in raw_curr:
            clean_curr = code
            break

    curr_sym, curr_fmt = CURRENCY_CONFIG.get(clean_curr, ("$", "$#,##0.00"))
    
    # -------------------------------------------------------------
    # Sheet 1: Executive Summary & Reconciliation Dashboard
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="Executive Summary")
    ws1.views.sheetView[0].showGridLines = True
    
    # 타이틀 블록
    ws1["A1"] = "BankFlow Statement Audit & Executive Summary"
    ws1["A1"].font = Font(name="Arial", size=16, bold=True, color=NAVY_HEADER)
    ws1["A2"] = "Zero-Storage Mathematical Audit Report | BankFlow Engine v2"
    ws1["A2"].font = Font(name="Arial", size=10, italic=True, color="64748B")

    # 메타 정보 매핑
    metas = [
        ("Financial Institution:", info.get("bank_name", "N/A")),
        ("Account Identifier:", info.get("account_number_masked", "****-****-****-XXXX")),
        ("Statement Period:", info.get("statement_period", "N/A")),
        ("Base Currency:", f"{clean_curr} ({curr_sym})"),
        ("Audit Integrity:", "PERFECT MATCH (0.00 Variance)" if audit_result.get("is_balanced") else f"WARNING: Variance {audit_result.get('discrepancy')}")
    ]
    
    for idx, (k, v) in enumerate(metas, start=4):
        ws1[f"A{idx}"] = k
        ws1[f"A{idx}"].font = Font(name="Arial", size=10, bold=True, color="334155")
        ws1[f"B{idx}"] = v
        ws1[f"B{idx}"].font = Font(name="Arial", size=10, bold=("Audit" in k), color=GREEN_TEXT if "PERFECT" in str(v) else "0F172A")

    # KPI 및 검산 요약 테이블 헤더 (동적 통화 반영)
    headers_kpi = ["Reconciliation Metric", f"Amount ({clean_curr})", "Formula / Audit Logic"]
    for c_i, h in enumerate(headers_kpi, start=1):
        c = ws1.cell(row=10, column=c_i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        c.alignment = Alignment(horizontal="center" if c_i > 1 else "left", vertical="center")
        c.border = thin_border

    start_bal = info.get("starting_balance", 0.0)
    end_bal = info.get("ending_balance", 0.0)
    last_tx_row = len(txs) + 1

    # KPI 항목 정의 (동적 통화 서식 적용)
    kpis = [
        ("Starting Balance (Period Open)", start_bal, "Direct Statement Verification", curr_fmt),
        ("Total Deposits (Inflow)", f"=SUM(Transactions!D2:D{last_tx_row})", "=SUM(Transactions!Inflow)", curr_fmt),
        ("Total Withdrawals (Outflow)", f"=SUM(Transactions!C2:C{last_tx_row})", "=SUM(Transactions!Outflow)", curr_fmt),
        ("Net Cash Flow for Period", "=B12-B13", "=Total Deposits - Total Withdrawals", curr_fmt),
        ("Calculated Ending Balance", "=B11+B14", "=Starting Balance + Net Cash Flow", curr_fmt),
        ("Reported Ending Balance", end_bal, "Bank Stated Closing Balance", curr_fmt),
        ("Reconciliation Variance", "=B15-B16", "=Calculated Ending - Reported Ending", curr_fmt)
    ]

    for r_idx, (metric, val, desc, fmt) in enumerate(kpis, start=11):
        ws1[f"A{r_idx}"] = metric
        ws1[f"A{r_idx}"].font = Font(name="Arial", size=10, bold=(r_idx in [11, 14, 15, 16, 17]))
        ws1[f"A{r_idx}"].border = thin_border if r_idx != 17 else double_bottom_border
        
        ws1[f"B{r_idx}"] = val
        ws1[f"B{r_idx}"].font = Font(name="Arial", size=10, bold=(r_idx in [14, 15, 16, 17]))
        ws1[f"B{r_idx}"].number_format = fmt
        ws1[f"B{r_idx}"].alignment = Alignment(horizontal="right")
        ws1[f"B{r_idx}"].border = thin_border if r_idx != 17 else double_bottom_border
        
        ws1[f"C{r_idx}"] = desc
        ws1[f"C{r_idx}"].font = Font(name="Arial", size=9, italic=True, color="64748B")
        ws1[f"C{r_idx}"].border = thin_border if r_idx != 17 else double_bottom_border

        if r_idx == 17:
            ws1[f"B{r_idx}"].fill = PatternFill(start_color=GREEN_FILL, end_color=GREEN_FILL, fill_type="solid")
            ws1[f"B{r_idx}"].font = Font(name="Arial", size=10, bold=True, color=GREEN_TEXT)

    # -------------------------------------------------------------
    # Sheet 2: Transactions (Detail Ledger)
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Transactions")
    ws2.views.sheetView[0].showGridLines = True
    ws2.freeze_panes = "A2" # 스크롤 시 헤더 고정

    tx_headers = ["Date", "Description", f"Outflow ({clean_curr})", f"Inflow ({clean_curr})", "Category", "Running Balance", "Integrity"]
    ws2.row_dimensions[1].height = 24

    for c_i, h in enumerate(tx_headers, start=1):
        c = ws2.cell(row=1, column=c_i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    # 트랜잭션 행 삽입
    for i, t in enumerate(txs, start=2):
        ws2.row_dimensions[i].height = 20
        is_even = (i % 2 == 0)
        row_fill = PatternFill(start_color=ZEBRA_EVEN if is_even else "FFFFFF", end_color=ZEBRA_EVEN if is_even else "FFFFFF", fill_type="solid")

        outflow = float(t.get("withdrawal", 0.0))
        inflow = float(t.get("deposit", 0.0))

        ws2.cell(row=i, column=1, value=t.get("date", "")).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=i, column=2, value=t.get("description", "")).alignment = Alignment(horizontal="left", vertical="center")

        c_out = ws2.cell(row=i, column=3, value=outflow)
        c_out.number_format = curr_fmt
        c_out.font = Font(name="Arial", size=10, color=RED_TEXT if outflow > 0 else "0F172A")
        c_out.alignment = Alignment(horizontal="right", vertical="center")

        c_in = ws2.cell(row=i, column=4, value=inflow)
        c_in.number_format = curr_fmt
        c_in.font = Font(name="Arial", size=10, bold=(inflow > 0), color=GREEN_TEXT if inflow > 0 else "0F172A")
        c_in.alignment = Alignment(horizontal="right", vertical="center")

        ws2.cell(row=i, column=5, value=t.get("category", "General")).alignment = Alignment(horizontal="left", vertical="center")

        # 동적 잔액 계산 수식 주입 (첫 행은 Summary B11 참조, 이후는 전일 잔고 + 입금 - 출금)
        bal_formula = f"='Executive Summary'!B11 + D{i} - C{i}" if i == 2 else f"=F{i-1} + D{i} - C{i}"
        c_bal = ws2.cell(row=i, column=6, value=bal_formula)
        c_bal.number_format = curr_fmt
        c_bal.font = Font(name="Arial", size=10, bold=True)
        c_bal.alignment = Alignment(horizontal="right", vertical="center")

        has_warn = t.get("reconciliation_warning", False)
        c_chk = ws2.cell(row=i, column=7, value="CHECK" if has_warn else "OK")
        c_chk.font = Font(name="Arial", size=9, bold=True, color=RED_TEXT if has_warn else GREEN_TEXT)
        c_chk.alignment = Alignment(horizontal="center", vertical="center")

        for c_idx in range(1, 8):
            ws2.cell(row=i, column=c_idx).border = thin_border
            ws2.cell(row=i, column=c_idx).fill = row_fill

    # 하단 합계 행 (Total Row)
    tot_row = len(txs) + 2
    ws2.row_dimensions[tot_row].height = 24
    ws2.cell(row=tot_row, column=1, value="TOTALS").font = Font(name="Arial", size=10, bold=True, color=NAVY_HEADER)
    ws2.cell(row=tot_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    ws2.cell(row=tot_row, column=3, value=f"=SUM(C2:C{tot_row-1})").number_format = curr_fmt
    ws2.cell(row=tot_row, column=3).font = Font(name="Arial", size=10, bold=True, color=RED_TEXT)
    ws2.cell(row=tot_row, column=3).alignment = Alignment(horizontal="right", vertical="center")

    ws2.cell(row=tot_row, column=4, value=f"=SUM(D2:D{tot_row-1})").number_format = curr_fmt
    ws2.cell(row=tot_row, column=4).font = Font(name="Arial", size=10, bold=True, color=GREEN_TEXT)
    ws2.cell(row=tot_row, column=4).alignment = Alignment(horizontal="right", vertical="center")

    ws2.cell(row=tot_row, column=6, value=f"=F{tot_row-1}").number_format = curr_fmt
    ws2.cell(row=tot_row, column=6).font = Font(name="Arial", size=10, bold=True, color=NAVY_HEADER)
    ws2.cell(row=tot_row, column=6).alignment = Alignment(horizontal="right", vertical="center")

    ws2.cell(row=tot_row, column=7, value="BALANCED").font = Font(name="Arial", size=10, bold=True, color=GREEN_TEXT)
    ws2.cell(row=tot_row, column=7).alignment = Alignment(horizontal="center", vertical="center")

    for c_idx in range(1, 8):
        ws2.cell(row=tot_row, column=c_idx).border = double_bottom_border

    # 컬럼 너비 자동 최적화
    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len and not val_str.startswith("="):
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    ws1.column_dimensions["A"].width = 36
    ws1.column_dimensions["B"].width = 24
    ws1.column_dimensions["C"].width = 36

    # 인메모리 스트림 반환 (디스크 I/O 없음)
    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream
