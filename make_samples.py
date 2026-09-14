# make_samples.py
# 표준 라이브러리 기반 샘플 명세서 PDF 생성기 (외부 설치 불필요)

def generate_pdf(filename, title, meta, rows, summary):
    content_lines = [
        "BT",
        "/F1 15 Tf",
        "40 750 Td",
        f"({title}) Tj",
        "/F1 9 Tf",
        "0 -22 Td",
        f"(Financial Institution : {meta['bank']}) Tj",
        "0 -13 Td",
        f"(Account Identifier   : {meta['account']}) Tj",
        "0 -13 Td",
        f"(Statement Period     : {meta['period']} | Currency: {meta['currency']}) Tj",
        "0 -13 Td",
        f"(Starting Balance     : ${summary['starting']:,.2f}) Tj",
        "0 -20 Td",
        "(========================================================================================================) Tj",
        "0 -13 Td",
        "(Date          Description                                       Withdrawal        Deposit        Balance) Tj",
        "0 -11 Td",
        "(--------------------------------------------------------------------------------------------------------) Tj",
    ]

    for r in rows:
        d_str = r[0].ljust(13)
        desc = r[1][:45].ljust(48)
        w_val = (f"-${r[2]:,.2f}" if r[2] > 0 else "-").rjust(14)
        dep_val = (f"+${r[3]:,.2f}" if r[3] > 0 else "-").rjust(14)
        bal_val = f"${r[4]:,.2f}".rjust(15)
        row_line = f"{d_str}{desc}{w_val}{dep_val}{bal_val}".replace("(", "\\(").replace(")", "\\)")
        content_lines.append("0 -15 Td")
        content_lines.append(f"({row_line}) Tj")

    content_lines.extend([
        "0 -15 Td",
        "(========================================================================================================) Tj",
        "0 -15 Td",
        "(AUDIT PROOF & RECONCILIATION SUMMARY:) Tj",
        "0 -14 Td",
        f"(Starting Balance      : ${summary['starting']:,.2f}) Tj",
        "0 -13 Td",
        f"(Total Deposits (+)    : +${summary['total_deposits']:,.2f}) Tj",
        "0 -13 Td",
        f"(Total Withdrawals (-) : -${summary['total_withdrawals']:,.2f}) Tj",
        "0 -13 Td",
        f"(Calculated Ending     : ${summary['ending']:,.2f}) Tj",
        "0 -13 Td",
        f"(Reported Bank Ending  : ${summary['ending']:,.2f}  |  Variance: $0.00 [BALANCED - 100% MATCH]) Tj",
        "ET"
    ])

    stream_bytes = "\n".join(content_lines).encode("latin1")
    stream_len = len(stream_bytes)

    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {stream_len} >>\nstream\n".encode("latin1") + stream_bytes + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>"
    ]

    pdf = bytearray(b"%PDF-1.4\n")
    xref_offsets = [0]
    for i, obj in enumerate(objects, 1):
        xref_offsets.append(len(pdf))
        pdf.extend(f"{i} 0 obj\n".encode('latin1'))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")

    xref_pos = len(pdf)
    pdf.extend(f"xref\n0 {len(objects)+1}\n0000000000 65535 f \n".encode('latin1'))
    for offset in xref_offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode('latin1'))
    pdf.extend(f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode('latin1'))

    with open(filename, "wb") as f:
        f.write(pdf)
    print(f"✅ 생성 완료: {filename}")

# 1. Chase 비즈니스 체킹 명세서 (기초: $25,400.00 -> 기말: $35,509.50)
generate_pdf(
    "sample1_chase_business.pdf",
    "JPMorgan Chase Bank, N.A. - Commercial Checking Statement",
    {"bank": "Chase Bank", "account": "****-****-****-4821", "period": "2026-09-01 to 2026-09-30", "currency": "USD ($)"},
    [
        ("2026-09-02", "STRIPE PAYOUT BATCH #8821", 0.0, 6500.00, 31900.00),
        ("2026-09-06", "AMAZON WEB SERVICES CLOUD INFRA", 480.50, 0.0, 31419.50),
        ("2026-09-12", "GOOGLE WORKSPACE & GEMINI API", 210.00, 0.0, 31209.50),
        ("2026-09-18", "INBOUND WIRE - ACME VENTURES CORP", 0.0, 14200.00, 45409.50),
        ("2026-09-24", "GUSTO PAYROLL AUTOMATED TAX DEBIT", 8650.00, 0.0, 36759.50),
        ("2026-09-29", "WEWORK DEDICATED OFFICE LEASE", 1250.00, 0.0, 35509.50)
    ],
    {"starting": 25400.00, "total_deposits": 20700.00, "total_withdrawals": 10590.50, "ending": 35509.50}
)

# 2. Mercury 스타트업 트레저리 계좌 명세서 (기초: $82,000.00 -> 기말: $134,620.00)
generate_pdf(
    "sample2_mercury_treasury.pdf",
    "Mercury Technologies / Choice Financial - Treasury Statement",
    {"bank": "Mercury Technologies", "account": "****-****-****-9014", "period": "2026-08-01 to 2026-08-31", "currency": "USD ($)"},
    [
        ("2026-08-03", "ANGEL INVESTMENT TRANCHE A WIRE", 0.0, 50000.00, 132000.00),
        ("2026-08-10", "VERCEL PRO EDGE HOSTING & WORKERS", 360.00, 0.0, 131640.00),
        ("2026-08-17", "OPENAI PLATFORM API INVOICE #8912", 920.00, 0.0, 130720.00),
        ("2026-08-22", "LEMON SQUEEZY MERCHANT SETTLEMENT", 0.0, 8400.00, 139120.00),
        ("2026-08-28", "CONTRACTOR DEVELOPMENT PAYMENT", 4500.00, 0.0, 134620.00)
    ],
    {"starting": 82000.00, "total_deposits": 58400.00, "total_withdrawals": 5780.00, "ending": 134620.00}
)

# 3. 실리콘밸리 뱅크 기업 원장 명세서 (기초: $15,350.00 -> 기말: $24,680.00)
generate_pdf(
    "sample3_svb_commercial.pdf",
    "Silicon Valley Bank - Commercial Account Monthly Statement",
    {"bank": "Silicon Valley Bank", "account": "****-****-****-3318", "period": "2026-07-01 to 2026-07-31", "currency": "USD ($)"},
    [
        ("2026-07-04", "CLIENT INVOICE #1092 WIRE PAYMENT", 0.0, 7200.00, 22550.00),
        ("2026-07-11", "GITHUB ENTERPRISE & COPILOT SEATS", 180.00, 0.0, 22370.00),
        ("2026-07-16", "FIGMA DESIGN TEAM ANNUAL LICENSE", 340.00, 0.0, 22030.00),
        ("2026-07-23", "STRIPE RECURRING SAAS BILLING", 0.0, 3850.00, 25880.00),
        ("2026-07-30", "ACCOUNTING CPA QUARTERLY AUDIT FEE", 1200.00, 0.0, 24680.00)
    ],
    {"starting": 15350.00, "total_deposits": 11050.00, "total_withdrawals": 1720.00, "ending": 24680.00}
)