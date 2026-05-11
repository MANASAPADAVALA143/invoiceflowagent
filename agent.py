"""
InvoiceFlow — Agentic AI Backend
Serves both use cases from one FastAPI server.

USE CASE 1: Excel VBA calls /api/agent/process-invoice
USE CASE 2: n8n calls /api/agent/process-email-invoice
USE CASE 3: InvoiceFlow web app Scan invoice → POST /api/agent/extract-image (multipart file)

Install: pip install -r requirements.txt
Run:     uvicorn agent:app --reload --port 8000
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware

from extract_image_routes import router as extract_image_router
from pydantic import BaseModel
from typing import Optional, List, Any
import anthropic
import httpx
import json
import os
import base64
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="InvoiceFlow Agentic AI", version="3.0")
# Specific origins + Vercel preview regex (wildcard subdomains are not valid in allow_origins).
# Do not combine allow_origins=["*"] with allow_credentials=True — browsers reject it.
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5175",
        "http://localhost:5173",
        "http://localhost:5176",
        "https://invoiceflow.vercel.app",
        "https://apinvoiceflow.vercel.app",
        "https://invoiceflow.ai",
    ],
    allow_origin_regex=r"https://.*\.vercel\.app",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(extract_image_router, prefix="/api/agent", tags=["agent"])

claude = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

SUPABASE_URL = os.getenv("SUPABASE_URL", "https://xuaaqonmaarldzklocax.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_ANON_KEY", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
N8N_WEBHOOK = os.getenv("N8N_WEBHOOK_URL", "http://localhost:5678/webhook/ap-invoice")


# ══════════════════════════════════════════════════════════
# REQUEST MODELS
# ══════════════════════════════════════════════════════════

class InvoiceFromExcel(BaseModel):
    """USE CASE 1 — sent by VBA from Excel row"""
    row_number: int
    invoice_number: str
    vendor_name: str
    invoice_date: str
    due_date: Optional[str] = ""
    amount: float
    currency: str = "INR"
    description: Optional[str] = ""
    po_number: Optional[str] = ""
    department: Optional[str] = ""
    accounting_standard: str = "IFRS"

class EmailInvoiceRequest(BaseModel):
    """USE CASE 2 — sent by n8n from email attachment"""
    from_email: str
    sender_name: str
    subject: str
    filename: str
    file_base64: str          # PDF/image as base64
    media_type: str = "application/pdf"
    accounting_standard: str = "IFRS"
    write_to_sheets: bool = True
    notify_email: str = ""


# ══════════════════════════════════════════════════════════
# CLAUDE AGENT TOOLS
# These are the functions Claude decides to call
# ══════════════════════════════════════════════════════════

AGENT_TOOLS = [
    {
        "name": "classify_invoice",
        "description": "Classify invoice GL code, IFRS category, and generate journal entry based on vendor, amount, and description.",
        "input_schema": {
            "type": "object",
            "properties": {
                "vendor_name": {"type": "string"},
                "amount": {"type": "number"},
                "description": {"type": "string"},
                "accounting_standard": {"type": "string", "enum": ["IFRS", "Ind AS", "US GAAP", "IGAAP"]},
                "currency": {"type": "string"}
            },
            "required": ["vendor_name", "amount", "accounting_standard"]
        }
    },
    {
        "name": "check_duplicate",
        "description": "Check if this invoice is a duplicate by looking up invoice number and vendor in historical records.",
        "input_schema": {
            "type": "object",
            "properties": {
                "invoice_number": {"type": "string"},
                "vendor_name": {"type": "string"},
                "amount": {"type": "number"}
            },
            "required": ["invoice_number", "vendor_name", "amount"]
        }
    },
    {
        "name": "score_risk",
        "description": "Score the fraud and compliance risk of an invoice based on multiple signals.",
        "input_schema": {
            "type": "object",
            "properties": {
                "invoice_number": {"type": "string"},
                "vendor_name": {"type": "string"},
                "amount": {"type": "number"},
                "invoice_date": {"type": "string"},
                "po_number": {"type": "string"},
                "description": {"type": "string"}
            },
            "required": ["vendor_name", "amount", "invoice_date"]
        }
    },
    {
        "name": "verify_vendor",
        "description": "Verify if vendor exists in master data, check GSTIN/TIN validity, and flag new vendors.",
        "input_schema": {
            "type": "object",
            "properties": {
                "vendor_name": {"type": "string"},
                "amount": {"type": "number"}
            },
            "required": ["vendor_name"]
        }
    },
    {
        "name": "check_po_match",
        "description": "Check if a purchase order exists matching this invoice and verify amounts match.",
        "input_schema": {
            "type": "object",
            "properties": {
                "po_number": {"type": "string"},
                "vendor_name": {"type": "string"},
                "amount": {"type": "number"}
            },
            "required": ["po_number", "vendor_name", "amount"]
        }
    },
    {
        "name": "calculate_tax",
        "description": "Calculate applicable tax (GST/VAT/Sales Tax) based on vendor location, amount and invoice type.",
        "input_schema": {
            "type": "object",
            "properties": {
                "amount": {"type": "number"},
                "currency": {"type": "string"},
                "vendor_name": {"type": "string"},
                "description": {"type": "string"}
            },
            "required": ["amount", "currency"]
        }
    },
    {
        "name": "determine_approval_route",
        "description": "Determine who should approve this invoice based on amount, vendor type, risk level and department.",
        "input_schema": {
            "type": "object",
                "properties": {
                "amount": {"type": "number"},
                "currency": {"type": "string"},
                "risk_level": {"type": "string"},
                "vendor_name": {"type": "string"},
                "department": {"type": "string"},
                "is_new_vendor": {"type": "boolean"}
            },
            "required": ["amount", "risk_level"]
        }
    },
    {
        "name": "save_to_supabase",
        "description": "Save the fully processed invoice to InvoiceFlow database.",
        "input_schema": {
            "type": "object",
            "properties": {
                "invoice_data": {"type": "object", "description": "Complete processed invoice object"}
            },
            "required": ["invoice_data"]
        }
    },
    {
        "name": "write_to_google_sheets",
        "description": "Append the processed invoice row to the Google Sheet for the CA firm to review.",
        "input_schema": {
            "type": "object",
            "properties": {
                "invoice_data": {"type": "object", "description": "Complete processed invoice object"},
                "sheet_id": {"type": "string"}
            },
            "required": ["invoice_data"]
        }
    },
    {
        "name": "send_alert",
        "description": "Send email or Slack alert for high risk invoices, new vendors, or invoices needing CFO approval.",
        "input_schema": {
            "type": "object",
            "properties": {
                "alert_type": {"type": "string", "enum": ["high_risk", "duplicate", "new_vendor", "cfo_approval", "overdue"]},
                "invoice_number": {"type": "string"},
                "vendor_name": {"type": "string"},
                "amount": {"type": "number"},
                "reason": {"type": "string"},
                "notify_email": {"type": "string"}
            },
            "required": ["alert_type", "vendor_name", "amount", "reason"]
        }
    }
]


# ══════════════════════════════════════════════════════════
# TOOL EXECUTION ENGINE
# When Claude decides to call a tool, this executes it
# ══════════════════════════════════════════════════════════

async def execute_tool(tool_name: str, tool_input: dict, context: dict) -> Any:
    """Execute the tool Claude requested and return the result."""

    if tool_name == "classify_invoice":
        return {
            "gl_code": classify_gl(tool_input.get("description", ""), tool_input.get("vendor_name", "")),
            "gl_account_name": get_gl_name(tool_input.get("description", ""), tool_input.get("vendor_name", "")),
            "ifrs_category": classify_ifrs(tool_input.get("description", ""), tool_input.get("vendor_name", "")),
            "journal_debit": get_gl_name(tool_input.get("description", ""), tool_input.get("vendor_name", "")),
            "journal_credit": f"Accounts Payable — {tool_input.get('vendor_name', 'Vendor')}",
            "payment_terms_days": 30,
            "confidence": 92
        }

    elif tool_name == "check_duplicate":
        # Check Supabase for existing invoice
        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"{SUPABASE_URL}/rest/v1/invoices",
                    headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"},
                    params={
                        "invoice_number": f"eq.{tool_input.get('invoice_number')}",
                        "select": "id,invoice_number,vendor_name,total_amount"
                    }
                )
                existing = response.json()
                is_duplicate = len(existing) > 0
                return {
                    "is_duplicate": is_duplicate,
                    "existing_count": len(existing),
                    "message": f"Found {len(existing)} matching invoice(s)" if is_duplicate else "No duplicate found"
                }
        except:
            return {"is_duplicate": False, "message": "Could not check — proceed with caution"}

    elif tool_name == "score_risk":
        flags = []
        score = 0
        amount = tool_input.get("amount", 0)
        inv_date = tool_input.get("invoice_date", "")
        po = tool_input.get("po_number", "")
        desc = tool_input.get("description", "")

        if amount % 1000 == 0 and amount >= 10000:
            flags.append("round_number"); score += 15
        if amount >= 500000:
            flags.append("high_value"); score += 20
        if not po:
            flags.append("no_po_number"); score += 15
        if not desc:
            flags.append("missing_description"); score += 10
        try:
            d = datetime.strptime(inv_date, "%Y-%m-%d")
            if d.weekday() in [5, 6]:
                flags.append("weekend_date"); score += 20
        except:
            pass

        score = min(score, 100)
        level = "High" if score >= 60 else "Medium" if score >= 30 else "Low"
        return {"risk_score": score, "risk_level": level, "risk_flags": flags}

    elif tool_name == "verify_vendor":
        # Check if vendor exists in Supabase vendors table
        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"{SUPABASE_URL}/rest/v1/vendors",
                    headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"},
                    params={"name": f"ilike.*{tool_input.get('vendor_name', '')}*", "select": "id,name,gst_number"}
                )
                vendors = response.json()
                is_new = len(vendors) == 0
                return {
                    "is_new_vendor": is_new,
                    "has_gstin": any(v.get("gst_number") for v in vendors),
                    "vendor_count": len(vendors),
                    "message": "New vendor — KYC required" if is_new else "Known vendor"
                }
        except:
            return {"is_new_vendor": True, "has_gstin": False, "message": "Could not verify vendor"}

    elif tool_name == "check_po_match":
        po_num = tool_input.get("po_number", "")
        if not po_num:
            return {"po_found": False, "match_status": "no_po", "message": "No PO number provided"}
        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"{SUPABASE_URL}/rest/v1/purchase_orders",
                    headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"},
                    params={"po_number": f"eq.{po_num}", "select": "id,po_number,total_amount,vendor_name"}
                )
                pos = response.json()
                if not pos:
                    return {"po_found": False, "match_status": "po_not_found", "message": f"PO {po_num} not found"}
                po = pos[0]
                amount_match = abs(po.get("total_amount", 0) - tool_input.get("amount", 0)) < 100
                return {
                    "po_found": True,
                    "match_status": "matched" if amount_match else "amount_mismatch",
                    "po_amount": po.get("total_amount"),
                    "invoice_amount": tool_input.get("amount"),
                    "message": "PO matched" if amount_match else f"Amount mismatch: PO={po.get('total_amount')}, Invoice={tool_input.get('amount')}"
                }
        except:
            return {"po_found": False, "match_status": "error", "message": "Could not check PO"}

    elif tool_name == "calculate_tax":
        amount = tool_input.get("amount", 0)
        currency = tool_input.get("currency", "INR")
        desc = (tool_input.get("description") or "").lower()

        # GST rates for India
        if currency == "INR":
            if any(k in desc for k in ["software", "service", "consulting", "professional"]):
                rate = 18
            elif any(k in desc for k in ["food", "restaurant", "hotel"]):
                rate = 5
            elif any(k in desc for k in ["rent", "lease"]):
                rate = 18
            else:
                rate = 18  # default GST
        elif currency in ["GBP", "EUR"]:
            rate = 20  # UK/EU VAT
        elif currency == "USD":
            rate = 0   # US no federal VAT
        else:
            rate = 0

        tax_amount = round(amount * rate / 100, 2)
        net_amount = round(amount - tax_amount, 2)
        return {"tax_rate": rate, "tax_amount": tax_amount, "net_amount": net_amount, "tax_type": "GST" if currency == "INR" else "VAT"}

    elif tool_name == "determine_approval_route":
        amount = tool_input.get("amount", 0)
        risk = tool_input.get("risk_level", "Low")
        is_new = tool_input.get("is_new_vendor", False)

        if is_new or risk == "High":
            return {"approver": "CFO", "route": "cfo_approval", "reason": "High risk or new vendor"}
        elif amount >= 500000:
            return {"approver": "CFO", "route": "cfo_approval", "reason": "Amount above ₹5L threshold"}
        elif amount >= 100000:
            return {"approver": "Finance Controller", "route": "controller_approval", "reason": "Amount ₹1L-₹5L"}
        elif amount >= 10000:
            return {"approver": "AP Manager", "route": "manager_approval", "reason": "Standard approval"}
        else:
            return {"approver": "Auto-approved", "route": "auto_approved", "reason": "Below ₹10K threshold"}

    elif tool_name == "save_to_supabase":
        inv = tool_input.get("invoice_data", {})
        try:
            async with httpx.AsyncClient() as client:
                response = await client.post(
                    f"{SUPABASE_URL}/rest/v1/invoices",
                    headers={
                        "apikey": SUPABASE_KEY,
                        "Authorization": f"Bearer {SUPABASE_KEY}",
                        "Content-Type": "application/json",
                        "Prefer": "return=minimal"
                    },
                    json={
                        "company_id": "11fab3d0-7374-4205-8c10-4a61f49cd60d",
                        "invoice_number": inv.get("invoice_number", ""),
                        "vendor_name": inv.get("vendor_name", ""),
                        "total_amount": inv.get("amount", 0),
                        "currency": inv.get("currency", "INR"),
                        "tax_amount": inv.get("tax_amount", 0),
                        "ifrs_category": inv.get("ifrs_category", "Operating Expense"),
                        "accounting_category": inv.get("ifrs_category", "Operating Expense"),
                        "gl_code": inv.get("gl_code", "6000"),
                        "gl_account_name": inv.get("gl_account_name", "Operating Expenses"),
                        "risk_level": inv.get("risk_level", "Low"),
                        "ocr_confidence": inv.get("confidence", 85),
                        "approval_status": "pending",
                        "payment_status": "unpaid",
                        "source": inv.get("source", "agent"),
                        "notes": inv.get("agent_summary", "")
                    }
                )
            return {"saved": True, "message": "Saved to InvoiceFlow dashboard"}
        except Exception as e:
            return {"saved": False, "message": str(e)}

    elif tool_name == "write_to_google_sheets":
        inv = tool_input.get("invoice_data", {})
        sheet_id = tool_input.get("sheet_id") or GOOGLE_SHEET_ID
        if not sheet_id:
            return {"written": False, "message": "No Google Sheet ID configured"}
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_file(
                "google_service_account.json",
                scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            )
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(sheet_id)
            ws = sh.sheet1
            row = [
                inv.get("invoice_number", ""),
                inv.get("vendor_name", ""),
                inv.get("invoice_date", ""),
                inv.get("due_date", ""),
                inv.get("amount", 0),
                inv.get("currency", "INR"),
                inv.get("gl_code", ""),
                inv.get("gl_account_name", ""),
                inv.get("ifrs_category", ""),
                inv.get("tax_rate", 0),
                inv.get("tax_amount", 0),
                inv.get("net_amount", 0),
                inv.get("journal_debit", ""),
                inv.get("journal_credit", ""),
                inv.get("risk_level", ""),
                inv.get("risk_score", 0),
                ", ".join(inv.get("risk_flags", [])),
                inv.get("is_duplicate", False),
                inv.get("approval_route", ""),
                inv.get("approver", ""),
                inv.get("confidence", 0),
                inv.get("agent_summary", ""),
                datetime.now().strftime("%Y-%m-%d %H:%M")
            ]
            ws.append_row(row)
            return {"written": True, "message": "Row added to Google Sheet"}
        except Exception as e:
            return {"written": False, "message": f"Sheets error: {str(e)}"}

    elif tool_name == "send_alert":
        # In production — send via Gmail node in n8n or direct SMTP
        # Here we just log and return
        print(f"ALERT [{tool_input.get('alert_type')}]: {tool_input.get('vendor_name')} — {tool_input.get('reason')}")
        return {
            "alert_sent": True,
            "alert_type": tool_input.get("alert_type"),
            "message": f"Alert queued for {tool_input.get('notify_email', 'default recipient')}"
        }

    return {"error": f"Unknown tool: {tool_name}"}


# ══════════════════════════════════════════════════════════
# CORE AGENT RUNNER
# Runs Claude in an agentic loop until it finishes
# ══════════════════════════════════════════════════════════

async def run_agent(system_prompt: str, user_message: str, context: dict) -> dict:
    """
    Run Claude agent with tool use.
    Claude decides which tools to call and in what order.
    Loops until Claude returns a final text response.
    """
    messages = [{"role": "user", "content": user_message}]
    tool_calls_log = []
    max_iterations = 10
    iteration = 0

    while iteration < max_iterations:
        iteration += 1

        response = claude.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system=system_prompt,
            tools=AGENT_TOOLS,
            messages=messages
        )

        # If Claude is done (no more tool calls)
        if response.stop_reason == "end_turn":
            final_text = ""
            for block in response.content:
                if hasattr(block, "text"):
                    final_text = block.text
            return {
                "agent_summary": final_text,
                "tool_calls": tool_calls_log,
                "iterations": iteration
            }

        # Claude wants to use tools
        if response.stop_reason == "tool_use":
            # Add Claude's response to messages
            messages.append({"role": "assistant", "content": response.content})

            # Execute each tool Claude requested
            tool_results = []
            for block in response.content:
                if block.type == "tool_use":
                    tool_name = block.name
                    tool_input = block.input
                    tool_id = block.id

                    print(f"Agent calling tool: {tool_name} with {tool_input}")

                    # Execute the tool
                    result = await execute_tool(tool_name, tool_input, context)
                    tool_calls_log.append({
                        "tool": tool_name,
                        "input": tool_input,
                        "result": result
                    })

                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": tool_id,
                        "content": json.dumps(result)
                    })

            # Add tool results back to conversation
            messages.append({"role": "user", "content": tool_results})

    return {
        "agent_summary": "Agent reached max iterations",
        "tool_calls": tool_calls_log,
        "iterations": iteration
    }


# ══════════════════════════════════════════════════════════
# USE CASE 1 — VBA calls this endpoint
# Excel → VBA → FastAPI → Agent → results back to Excel
# ══════════════════════════════════════════════════════════

@app.post("/api/agent/process-invoice")
async def process_invoice_from_excel(req: InvoiceFromExcel):
    """
    USE CASE 1 — Called by Excel VBA button.
    Agent processes one invoice and returns structured result.
    VBA writes result back to Excel row.
    """
    system = """You are an expert AP automation agent for InvoiceFlow.
Your job: process invoice data from Excel and determine everything needed for AP approval.

For EVERY invoice you must:
1. Call score_risk — always first
2. Call check_duplicate — always
3. Call verify_vendor — always
4. Call check_po_match — only if po_number is provided
5. Call classify_invoice — always
6. Call calculate_tax — always
7. Call determine_approval_route — always, using results from above
8. Call save_to_supabase — always at the end
9. Call send_alert — ONLY if risk is High, duplicate found, or new vendor

After all tool calls, return a JSON summary with ALL fields needed for Excel:
{
  "gl_code": "...",
  "gl_account_name": "...",
  "ifrs_category": "...",
  "tax_rate": 0,
  "tax_amount": 0,
  "net_amount": 0,
  "journal_debit": "...",
  "journal_credit": "...",
  "risk_level": "Low/Medium/High",
  "risk_score": 0,
  "risk_flags": [],
  "is_duplicate": false,
  "is_new_vendor": false,
  "approval_route": "...",
  "approver": "...",
  "confidence": 0,
  "excel_color": "CCFFCC for Low / FFF2CC for Medium / FFCCCC for High",
  "action": "approve/review/hold/reject",
  "summary": "one sentence for CFO"
}"""

    user_msg = f"""Process this invoice from Excel:
Invoice Number: {req.invoice_number}
Vendor: {req.vendor_name}
Amount: {req.currency} {req.amount}
Date: {req.invoice_date}
Due Date: {req.due_date or 'not specified'}
Description: {req.description or 'not provided'}
PO Number: {req.po_number or 'none'}
Department: {req.department or 'not specified'}
Accounting Standard: {req.accounting_standard}"""

    context = {"source": "excel_vba", "row_number": req.row_number}

    # Run the agent
    agent_result = await run_agent(system, user_msg, context)

    # Parse the JSON from agent summary
    try:
        summary_text = agent_result.get("agent_summary", "{}")
        # Extract JSON from agent response
        start = summary_text.find("{")
        end = summary_text.rfind("}") + 1
        if start >= 0 and end > start:
            result_data = json.loads(summary_text[start:end])
        else:
            result_data = {}
    except:
        result_data = {}

    # Ensure all Excel fields are present
    return {
        "row_number": req.row_number,
        "invoice_number": req.invoice_number,
        "vendor_name": req.vendor_name,
        "amount": req.amount,
        "currency": req.currency,
        # AI Results
        "gl_code": result_data.get("gl_code", "6000"),
        "gl_account_name": result_data.get("gl_account_name", "Operating Expenses"),
        "ifrs_category": result_data.get("ifrs_category", "Operating Expense"),
        "tax_rate": result_data.get("tax_rate", 0),
        "tax_amount": result_data.get("tax_amount", 0),
        "net_amount": result_data.get("net_amount", req.amount),
        "journal_debit": result_data.get("journal_debit", "Operating Expenses"),
        "journal_credit": result_data.get("journal_credit", f"Accounts Payable — {req.vendor_name}"),
        "risk_level": result_data.get("risk_level", "Low"),
        "risk_score": result_data.get("risk_score", 0),
        "risk_flags": result_data.get("risk_flags", []),
        "is_duplicate": result_data.get("is_duplicate", False),
        "is_new_vendor": result_data.get("is_new_vendor", False),
        "approval_route": result_data.get("approval_route", "manager_approval"),
        "approver": result_data.get("approver", "AP Manager"),
        "confidence": result_data.get("confidence", 85),
        "excel_color": result_data.get("excel_color", "CCFFCC"),
        "action": result_data.get("action", "review"),
        "agent_summary": result_data.get("summary", ""),
        "tool_calls_count": len(agent_result.get("tool_calls", [])),
        "agent_iterations": agent_result.get("iterations", 0)
    }


# ══════════════════════════════════════════════════════════
# USE CASE 2 — n8n calls this endpoint
# Email → n8n → FastAPI → Agent → Google Sheets auto-update
# ══════════════════════════════════════════════════════════

@app.post("/api/agent/process-email-invoice")
async def process_email_invoice(req: EmailInvoiceRequest):
    """
    USE CASE 2 — Called by n8n when email arrives with invoice attachment.
    Agent reads PDF, extracts data, processes completely,
    writes to Google Sheets automatically. No human needed.
    """
    system = """You are an expert AP automation agent for InvoiceFlow.
You receive invoice files from email and process them fully automatically.

Your job for EVERY email invoice:
1. First extract all invoice data from the file (it's provided as base64)
2. Call score_risk on the extracted data
3. Call check_duplicate
4. Call verify_vendor
5. Call check_po_match if PO number found
6. Call classify_invoice
7. Call calculate_tax
8. Call determine_approval_route
9. Call write_to_google_sheets with ALL extracted and processed data
10. Call save_to_supabase
11. Call send_alert if: High risk OR duplicate OR new vendor OR CFO approval needed

Be thorough. Extract every field visible in the invoice.
After all tools, write a brief summary for the email reply."""

    # Build user message with the invoice file
    if req.media_type == "application/pdf":
        content = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": req.file_base64
                }
            },
            {
                "type": "text",
                "text": f"""Process this invoice received from email.

Sender: {req.sender_name} ({req.from_email})
Email Subject: {req.subject}
Filename: {req.filename}
Accounting Standard: {req.accounting_standard}
Notify Email: {req.notify_email}

Extract ALL invoice data visible in the document.
Then run the complete processing workflow.
Write results to Google Sheets automatically."""
            }
        ]
    else:
        # Image invoice
        content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": req.media_type,
                    "data": req.file_base64
                }
            },
            {
                "type": "text",
                "text": f"""Process this invoice image received from email.
Sender: {req.sender_name} ({req.from_email})
Accounting Standard: {req.accounting_standard}
Extract all data and run complete processing workflow."""
            }
        ]

    context = {
        "source": "email_n8n",
        "from_email": req.from_email,
        "write_to_sheets": req.write_to_sheets,
        "notify_email": req.notify_email
    }

    # Build messages with file content
    messages_with_file = [{"role": "user", "content": content}]

    # Run agent (modified to accept initial content)
    tool_calls_log = []
    messages = messages_with_file
    max_iterations = 12
    iteration = 0
    final_summary = ""

    while iteration < max_iterations:
        iteration += 1
        response = claude.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2500,
            system=system,
            tools=AGENT_TOOLS,
            messages=messages
        )

        if response.stop_reason == "end_turn":
            for block in response.content:
                if hasattr(block, "text"):
                    final_summary = block.text
            break

        if response.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = []

            for block in response.content:
                if block.type == "tool_use":
                    print(f"Email Agent → {block.name}: {block.input}")
                    result = await execute_tool(block.name, block.input, context)
                    tool_calls_log.append({"tool": block.name, "result": result})
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": json.dumps(result)
                    })

            messages.append({"role": "user", "content": tool_results})

    # Check if Google Sheets was written
    sheets_written = any(
        tc["tool"] == "write_to_google_sheets" and tc["result"].get("written")
        for tc in tool_calls_log
    )

    # Check if saved to Supabase
    supabase_saved = any(
        tc["tool"] == "save_to_supabase" and tc["result"].get("saved")
        for tc in tool_calls_log
    )

    return {
        "success": True,
        "source": "email",
        "from_email": req.from_email,
        "filename": req.filename,
        "agent_summary": final_summary,
        "google_sheets_updated": sheets_written,
        "supabase_saved": supabase_saved,
        "tools_used": [tc["tool"] for tc in tool_calls_log],
        "total_tool_calls": len(tool_calls_log),
        "agent_iterations": iteration,
        "processed_at": datetime.now().isoformat()
    }


# ══════════════════════════════════════════════════════════
# USE CASE 2 SETUP — Initialize Google Sheet headers
# Call once to set up the sheet
# ══════════════════════════════════════════════════════════

@app.post("/api/setup/google-sheet")
async def setup_google_sheet():
    """Call once to add headers to your Google Sheet."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        creds = Credentials.from_service_account_file(
            "google_service_account.json",
            scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
        ws = sh.sheet1
        headers = [
            "Invoice #", "Vendor", "Invoice Date", "Due Date",
            "Amount", "Currency", "GL Code", "GL Account",
            "IFRS Category", "Tax Rate %", "Tax Amount", "Net Amount",
            "Journal DR", "Journal CR", "Risk Level", "Risk Score",
            "Risk Flags", "Duplicate?", "Approval Route", "Approver",
            "AI Confidence", "Agent Summary", "Processed At"
        ]
        ws.insert_row(headers, 1)
        return {"success": True, "message": "Headers added to Google Sheet"}
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.get("/")
def root():
    return {
        "service": "InvoiceFlow Agentic AI",
        "status": "online",
        "powered_by": "GnanovaPro · gnanova.pro",
        "endpoints": [
            "/api/health",
            "/api/agent/process-invoice",
            "/api/agent/process-email-invoice",
            "/api/agent/extract-image",
            "/api/setup/google-sheet",
        ],
    }


@app.get("/api/health")
def health():
    return {
        "status": "online",
        "service": "InvoiceFlow Agentic AI",
        "version": "3.0",
        "use_cases": [
            "Excel VBA (Use Case 1)",
            "Email Auto (Use Case 2)",
        ],
        "powered_by": "GnanovaPro · gnanova.pro",
        "timestamp": datetime.now().isoformat(),
    }


# ── GL Classification Helpers ──
def classify_gl(description: str, vendor: str) -> str:
    desc = (description + " " + vendor).lower()
    if any(k in desc for k in ["laptop", "server", "computer", "hardware", "equipment"]): return "1500"
    if any(k in desc for k in ["rent", "lease", "office space"]): return "6200"
    if any(k in desc for k in ["consult", "professional", "legal", "audit"]): return "6100"
    if any(k in desc for k in ["travel", "flight", "hotel", "cab", "makemytrip"]): return "6500"
    if any(k in desc for k in ["marketing", "advertising", "digital", "ads"]): return "6400"
    if any(k in desc for k in ["electricity", "internet", "utilities", "tsspdcl", "airtel"]): return "6300"
    if any(k in desc for k in ["research", "r&d", "development"]): return "7000"
    if any(k in desc for k in ["stationery", "office supplies", "printing"]): return "6600"
    return "6000"

def get_gl_name(description: str, vendor: str) -> str:
    code = classify_gl(description, vendor)
    names = {
        "1500": "Fixed Assets — IT Equipment",
        "6000": "Operating Expenses",
        "6100": "Professional Services",
        "6200": "Rent & Lease Expense",
        "6300": "Utilities",
        "6400": "Marketing & Advertising",
        "6500": "Travel & Entertainment",
        "6600": "Office Supplies",
        "7000": "Research & Development"
    }
    return names.get(code, "Operating Expenses")

def classify_ifrs(description: str, vendor: str) -> str:
    desc = (description + " " + vendor).lower()
    if any(k in desc for k in ["laptop", "server", "equipment", "machinery"]): return "Property Plant & Equipment"
    if any(k in desc for k in ["rent", "lease"]): return "Right-of-Use Asset"
    if any(k in desc for k in ["research", "r&d"]): return "Research & Development"
    if any(k in desc for k in ["interest", "finance charge", "bank"]): return "Finance Cost"
    return "Operating Expense"


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("agent:app", host="0.0.0.0", port=8000, reload=True)
