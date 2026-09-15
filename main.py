import io
import json
import os
import urllib.parse
from dotenv import load_dotenv

from fastapi import FastAPI, File, UploadFile, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse

from google import genai
from google.genai import types

# Phase 2에서 작성한 엑셀 렌더러 모듈 불러오기
from excel_renderer import render_bankflow_excel

load_dotenv()

# -------------------------------------------------------------
# 1. 환경 변수 및 Gemini AI 클라이언트 설정
# -------------------------------------------------------------
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    print("WARNING: GEMINI_API_KEY is not set. Please check your .env file.")

ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

# 모델 우선순위 순차 폴백 목록 (503 과부하 또는 일시적 장애 대비)
TARGET_MODELS = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"]

app = FastAPI(
    title="BankFlow Core Engine",
    description="Zero-Storage Bank Statement to Reconciled Excel SaaS",
    version="2.0.0"
)

# CORS 설정 (글로벌 프론트엔드 및 로컬 테스트 호환)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------------------------------------------------
# 2. AI 파싱 시스템 프롬프트 및 스키마
# -------------------------------------------------------------
BANKFLOW_SYSTEM_PROMPT = """
You are a precision-grade Financial Data Extraction Engine specialized in Bank and Credit Card Statements.
Extract every single transaction row from the document with 100% mathematical integrity.

RULES:
1. Preserve every row chronologically. Do NOT skip or aggregate repeating transactions.
2. In multi-line descriptions, clean and concatenate them into a single string.
3. If deposit/withdrawal is omitted in a row, represent it as 0.0.
4. Starting Balance + Total Deposits - Total Withdrawals MUST equal Ending Balance.
5. Strict Privacy: Never log, store, or output unmasked account numbers.
6. Currency Detection: Detect the 3-letter ISO currency code (USD, EUR, GBP, CAD, AUD, KRW, JPY, etc.) based on currency symbols ($, €, £, ₩) or the issuing financial institution. If unspecified for US or international banks (e.g., Chase, Mercury, SVB), default to 'USD'.
"""

BANK_STATEMENT_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "statement_info": {
            "type": "OBJECT",
            "properties": {
                "bank_name": {"type": "STRING"},
                "account_number_masked": {"type": "STRING"},
                "statement_period": {"type": "STRING"},
                "starting_balance": {"type": "NUMBER"},
                "ending_balance": {"type": "NUMBER"},
                "currency": {"type": "STRING"}
            },
            "required": ["starting_balance", "ending_balance", "currency"]
        },
        "transactions": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "date": {"type": "STRING"},
                    "description": {"type": "STRING"},
                    "withdrawal": {"type": "NUMBER"},
                    "deposit": {"type": "NUMBER"},
                    "balance": {"type": "NUMBER"},
                    "category": {"type": "STRING"}
                },
                "required": ["date", "description", "withdrawal", "deposit", "balance"]
            }
        }
    },
    "required": ["statement_info", "transactions"]
}

# -------------------------------------------------------------
# 3. 2중 수학적 무결성 검산기 (Reconciliation Engine)
# -------------------------------------------------------------
def verify_statement_integrity(parsed_data: dict) -> dict:
    info = parsed_data.get("statement_info", {})
    txs = parsed_data.get("transactions", [])

    start_bal = round(float(info.get("starting_balance", 0.0)), 2)
    end_bal = round(float(info.get("ending_balance", 0.0)), 2)

    total_deposits = round(sum(float(t.get("deposit", 0.0)) for t in txs), 2)
    total_withdrawals = round(sum(float(t.get("withdrawal", 0.0)) for t in txs), 2)
    calculated_ending = round(start_bal + total_deposits - total_withdrawals, 2)

    discrepancy = round(calculated_ending - end_bal, 2)
    is_balanced = abs(discrepancy) < 0.01

    curr = start_bal
    for t in txs:
        expected = round(curr + float(t.get("deposit", 0.0)) - float(t.get("withdrawal", 0.0)), 2)
        row_bal = round(float(t.get("balance", expected)), 2)
        t["reconciliation_warning"] = (abs(expected - row_bal) >= 0.01)
        curr = row_bal

    return {
        "is_balanced": is_balanced,
        "discrepancy": discrepancy,
        "total_deposits": total_deposits,
        "total_withdrawals": total_withdrawals,
        "transaction_count": len(txs)
    }

# -------------------------------------------------------------
# 4. 정적 프론트엔드 파일(static) 서빙 마운트
# -------------------------------------------------------------
if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def serve_homepage():
    if os.path.exists("static/index.html"):
        return FileResponse("static/index.html")
    return {"status": "ok", "message": "BankFlow API is running. static/index.html not found."}

@app.get("/health")
def health_check():
    return {"status": "ok", "service": "BankFlow API v2.0"}

# -------------------------------------------------------------
# 5. PDF/이미지 -> 엑셀 즉시 변환 API 엔드포인트
# -------------------------------------------------------------
@app.post("/api/convert")
async def convert_statement(file: UploadFile = File(...)):
    if not ai_client:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Server configuration error: GEMINI_API_KEY is missing."
        )

    # 허용 포맷 확인 (PDF, PNG, JPG)
    allowed = ["application/pdf", "image/png", "image/jpeg"]
    if file.content_type not in allowed:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Unsupported file format. Please upload a PDF, PNG, or JPG file."
        )

    # 100% 인메모리(RAM) 로드 (디스크 저장 원천 차단)
    file_bytes = await file.read()
    if len(file_bytes) > 20 * 1024 * 1024:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail="File size exceeds the 20MB limit."
        )

    # Gemini Vision AI 순차 폴백(Fallback) 구조화 파싱 호출
    parsed_json = None
    last_error = None

    for model_name in TARGET_MODELS:
        try:
            response = ai_client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=file_bytes, mime_type=file.content_type),
                    "Extract all transactional rows and reconciliation summary according to the schema."
                ],
                config=types.GenerateContentConfig(
                    system_instruction=BANKFLOW_SYSTEM_PROMPT,
                    response_mime_type="application/json",
                    response_schema=BANK_STATEMENT_SCHEMA,
                    temperature=0.1
                )
            )
            parsed_json = json.loads(response.text)
            break  # 파싱 성공 시 루프 탈출
        except Exception as e:
            last_error = e
            print(f"[경고] {model_name} 호출 실패 (다음 순위 모델로 전환): {str(e)}")
            continue

    if not parsed_json:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"모든 AI 모델 호출 실패 (최종 오류: {str(last_error)})"
        )

    # 2중 수학적 무결성 검산 실행
    audit_report = verify_statement_integrity(parsed_json)

    # excel_renderer.py를 통해 인메모리 .xlsx 바이너리 스트림 생성
    excel_stream = render_bankflow_excel(parsed_json, audit_report)

    base_name = os.path.splitext(file.filename or "statement")[0]
    out_name = f"BankFlow_{base_name}.xlsx"
    encoded_name = urllib.parse.quote(out_name)

    return StreamingResponse(
        excel_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}",
            "X-Audit-Status": "PASSED" if audit_report["is_balanced"] else "WARNING",
            "X-Audit-Variance": str(audit_report["discrepancy"]),
            "X-Transaction-Count": str(audit_report["transaction_count"])
        }
    )
