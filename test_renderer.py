# test_renderer.py
from excel_renderer import render_bankflow_excel

mock_parsed_data = {
    "statement_info": {
        "bank_name": "Silicon Valley Bank",
        "account_number_masked": "****-8842",
        "statement_period": "2026-08-01 ~ 2026-08-31",
        "starting_balance": 10000.00,
        "ending_balance": 12500.00,
        "currency": "USD ($)"
    },
    "transactions": [
        {"date": "2026-08-01", "description": "Stripe Payout", "withdrawal": 0.0, "deposit": 3000.0, "balance": 13000.0, "category": "Revenue"},
        {"date": "2026-08-02", "description": "AWS Hosting", "withdrawal": 500.0, "deposit": 0.0, "balance": 12500.0, "category": "SaaS"}
    ]
}

mock_audit = {
    "is_balanced": True,
    "discrepancy": 0.0,
    "total_deposits": 3000.0,
    "total_withdrawals": 500.0,
    "transaction_count": 2
}

stream = render_bankflow_excel(mock_parsed_data, mock_audit)

with open("output_test.xlsx", "wb") as f:
    f.write(stream.getbuffer())

print("테스트 엑셀 파일 생성 완료: output_test.xlsx")