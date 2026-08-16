import os
import json
import re
import datetime
import requests
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

# Load .env from the app directory so OLLAMA_CLOUD_* and OLLAMA_API are available
BASE_DIR = os.path.dirname(__file__)
load_dotenv(os.path.join(BASE_DIR, ".env"))


RECORD_FILE = "records.xlsx"

# Comprehensive transaction categories
TRANSACTION_CATEGORIES = {
    "Credit Transactions": [
        "Salary Credit",
        "Interest Credit",
        "Refund Received",
        "Cash Deposit",
        "Cheque Deposit",
        "NEFT Credit",
        "RTGS Credit",
        "IMPS Credit",
        "UPI Credit",
        "Bank Transfer Received",
        "Dividend Credit",
        "Cashback Credit",
        "Rewards Redemption",
        "Tax Refund",
        "Loan Disbursement",
        "Insurance Claim Credit",
        "Merchant Settlement Credit",
        "Wallet Transfer Received",
    ],
    "Debit Transactions": [
        "ATM Cash Withdrawal",
        "Cash Withdrawal at Branch",
        "Cash Deposit Reversal",
        "Cheque Clearing Debit",
        "Bank Transfer Sent",
        "NEFT Debit",
        "RTGS Debit",
        "IMPS Debit",
        "UPI Payment",
        "Standing Instruction Debit",
        "Auto Debit / ECS",
        "NACH Debit",
        "Loan EMI",
        "Credit Card Bill Payment",
        "Utility Bill Payment",
    ],
    "E-commerce": [
        "Amazon Purchase",
        "Flipkart Purchase",
        "Myntra Purchase",
        "Ajio Purchase",
        "Electronics Purchase",
        "Fashion Purchase",
        "Grocery Order",
        "Pharmacy Order",
        "Online Subscription Purchase",
        "Digital Product Purchase",
    ],
    "Food & Restaurant": [
        "Restaurant Payment",
        "Cafe Payment",
        "Food Delivery (Swiggy/Zomato)",
        "Fast Food Outlet",
        "Fine Dining",
        "Bakery Purchase",
    ],
    "Fuel & Transportation": [
        "Petrol Pump",
        "Diesel Station",
        "EV Charging",
        "FASTag Toll Payment",
        "Parking Payment",
        "Metro Recharge",
        "Bus Ticket Purchase",
        "Taxi Payment",
        "Uber Payment",
        "Ola Payment",
        "Train Ticket",
        "Flight Ticket",
    ],
    "Shopping & Retail": [
        "Supermarket",
        "Hypermarket",
        "Department Store",
        "Electronics Store",
        "Mobile Store",
        "Furniture Store",
        "Jewelry Store",
        "Medical Store",
        "Book Store",
    ],
    "Entertainment": [
        "Movie Ticket",
        "OTT Subscription",
        "Gaming Purchase",
        "Event Ticket",
        "Music Subscription",
        "Sports Ticket",
    ],
    "Utilities": [
        "Electricity Bill",
        "Water Bill",
        "Gas Bill",
        "Broadband Bill",
        "Mobile Recharge",
        "DTH Recharge",
        "Municipal Tax",
    ],
    "Financial Services": [
        "Mutual Fund Investment",
        "SIP Debit",
        "Stock Purchase",
        "Bond Purchase",
        "Insurance Premium",
        "Fixed Deposit Creation",
        "Recurring Deposit Installment",
        "Demat Charges",
        "Brokerage Charges",
    ],
    "Fees & Charges": [
        "ATM Charge",
        "SMS Charge",
        "Annual Maintenance Charge",
        "Debit Card Fee",
        "Credit Card Fee",
        "Penalty Charge",
        "Late Payment Fee",
        "GST Debit",
        "Convenience Fee",
    ],
    "Wallets & Fintech": [
        "Paytm Wallet Load",
        "PhonePe Transfer",
        "Google Pay Payment",
        "Amazon Pay",
        "Mobikwik Payment",
        "Wallet Withdrawal",
    ],
    "Travel": [
        "Hotel Booking",
        "Flight Booking",
        "Train Booking",
        "Bus Booking",
        "Holiday Package",
        "Car Rental",
    ],
    "Healthcare": [
        "Hospital Payment",
        "Doctor Consultation",
        "Laboratory Test",
        "Pharmacy Purchase",
        "Health Insurance Premium",
    ],
    "Education": [
        "School Fee",
        "College Fee",
        "Coaching Fee",
        "Online Course Purchase",
        "Examination Fee",
    ],
    "Government & Taxes": [
        "Income Tax Payment",
        "GST Payment",
        "Property Tax",
        "Traffic Fine",
        "Passport Fee",
        "Government Service Fee",
    ],
    "Investment & Wealth": [
        "Gold Purchase",
        "Digital Gold Purchase",
        "PPF Deposit",
        "NPS Contribution",
        "Mutual Fund Redemption Credit",
    ],
    "Corporate / Business": [
        "Vendor Payment",
        "Employee Salary Credit",
        "Expense Reimbursement",
        "Merchant Settlement",
        "Client Payment Received",
        "GST Refund Credit",
    ],
    "Card-Specific Transactions": [
        "Debit Card POS Purchase",
        "Credit Card POS Purchase",
        "Cardless ATM Withdrawal",
        "Contactless Payment",
        "International Card Transaction",
        "EMI Conversion",
    ],
}

# Flatten for easier lookup
ALL_CATEGORIES = []
for category_group, items in TRANSACTION_CATEGORIES.items():
    ALL_CATEGORIES.extend(items)
ALL_CATEGORIES = sorted(list(set(ALL_CATEGORIES)))


def get_category_suggestions() -> str:
    """Return formatted category list for LLM prompt"""
    result = ""
    for group, cats in TRANSACTION_CATEGORIES.items():
        result += f"{group}:\n"
        for cat in cats:
            result += f"  - {cat}\n"
    return result


def match_category(extracted_category: str) -> str:
    """Match extracted category to predefined categories using fuzzy matching"""
    if not extracted_category or extracted_category.lower() in ["unknown", "uncategorized", "none", "null"]:
        return "Uncategorized"
    
    extracted_lower = extracted_category.lower().strip()
    
    # Exact match
    for cat in ALL_CATEGORIES:
        if cat.lower() == extracted_lower:
            return cat
    
    # Substring match (if extracted is substring of category)
    for cat in ALL_CATEGORIES:
        if extracted_lower in cat.lower():
            return cat
    
    # Reverse match (if category is substring of extracted)
    for cat in ALL_CATEGORIES:
        if cat.lower() in extracted_lower:
            return cat
    
    # Keyword-based matching for common patterns
    keywords = {
        "Amazon": "Amazon Purchase",
        "Flipkart": "Flipkart Purchase",
        "Zepto": "Grocery Order",
        "Swiggy": "Food Delivery (Swiggy/Zomato)",
        "Zomato": "Food Delivery (Swiggy/Zomato)",
        "Petrol": "Petrol Pump",
        "Fuel": "Petrol Pump",
        "FASTag": "FASTag Toll Payment",
        "Uber": "Uber Payment",
        "Ola": "Ola Payment",
        "Netflix": "OTT Subscription",
        "Prime": "OTT Subscription",
        "Salary": "Salary Credit",
        "EMI": "Loan EMI",
        "Restaurant": "Restaurant Payment",
        "Hotel": "Hotel Booking",
        "Flight": "Flight Booking",
        "Hospital": "Hospital Payment",
        "Insurance": "Insurance Premium",
        "Bill": "Utility Bill Payment",
        "Electricity": "Electricity Bill",
    }
    
    for keyword, category in keywords.items():
        if keyword.lower() in extracted_lower:
            return category
    
    # Default fallback
    return "Uncategorized"



def call_ollama(prompt: str, model: str = "pi", prefer_local: bool = True) -> str:
    # Use user preference or env config to choose endpoint
    api_base = os.getenv("OLLAMA_API") if prefer_local else None
    cloud_base = os.getenv("OLLAMA_CLOUD_BASE_URL")
    cloud_key = os.getenv("OLLAMA_CLOUD_API_KEY")
    try:
        if api_base and prefer_local:
            # Use local Ollama endpoint (preferred)
            base = api_base.rstrip("/")
            url = f"{base}/api/generate"
            local_model = os.getenv("OLLAMA_LOCAL_MODEL", model)
            payload = {"model": local_model, "prompt": prompt}
            resp = requests.post(url, json=payload, timeout=60)
        elif cloud_base and cloud_key:
            # Use Ollama Cloud
            base = cloud_base.rstrip("/")
            url = f"{base}/api/generate"
            headers = {"Authorization": f"Bearer {cloud_key}", "Content-Type": "application/json"}
            payload = {"model": os.getenv("OLLAMA_CLOUD_MODEL", model), "prompt": prompt}
            resp = requests.post(url, json=payload, headers=headers, timeout=60)
        elif api_base:
            # Fall back to local if cloud not available
            base = api_base.rstrip("/")
            url = f"{base}/api/generate"
            local_model = os.getenv("OLLAMA_LOCAL_MODEL", model)
            payload = {"model": local_model, "prompt": prompt}
            resp = requests.post(url, json=payload, timeout=60)
        else:
            raise ValueError("No Ollama endpoint configured (OLLAMA_API or OLLAMA_CLOUD_BASE_URL)")

        resp.raise_for_status()
        # Handle both single JSON (local) and NDJSON streaming (cloud)
        try:
            data = resp.json()
            # Single JSON response
            if isinstance(data, dict):
                text = data.get("text") or data.get("output") or data.get("result") or ""
                if not text:
                    choices = data.get("choices")
                    if choices and isinstance(choices, list):
                        parts = []
                        for c in choices:
                            if isinstance(c, dict):
                                msg = c.get("message") or c.get("content") or c.get("text")
                                if isinstance(msg, dict):
                                    parts.append(msg.get("content", ""))
                                else:
                                    parts.append(msg or "")
                        text = "\n".join([p for p in parts if p])
                return text or json.dumps(data)
            return str(data)
        except json.JSONDecodeError:
            # Try NDJSON (streaming response)
            text_parts = []
            for line in resp.text.splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                    if isinstance(obj, dict):
                        resp_text = obj.get("response", "")
                        if resp_text:
                            text_parts.append(resp_text)
                except json.JSONDecodeError:
                    continue
            return "".join(text_parts) if text_parts else resp.text
    except Exception as e:
        return f"ERROR: {e}"


def normalize_parsed_result(parsed):
    if isinstance(parsed, list):
        for item in parsed:
            if isinstance(item, dict):
                return item
        return {}
    if isinstance(parsed, dict):
        return parsed
    return {}


def extract_json(text: str) -> dict:
    """
    Robustly extract the first valid JSON object from arbitrary LLM text.
    Handles outputs with surrounding commentary, code fences, or multiple JSON snippets.
    Returns an empty dict if no valid JSON is found.
    """
    if not text:
        return {}
    # strip common code fences
    try:
        text_clean = re.sub(r"```(?:json)?", "", text)
    except Exception:
        text_clean = text

    # Scan for balanced JSON object or array using a state machine that respects strings/escapes
    in_str = False
    esc = False
    stack = []  # will hold expected closing chars '}' or ']'
    start_idx = None
    for i, ch in enumerate(text_clean):
        if ch == '"' and not esc:
            in_str = not in_str
        if ch == '\\' and not esc:
            esc = True
            continue
        if esc:
            esc = False
            continue
        if in_str:
            continue
        if ch in '{[':
            if not stack:
                start_idx = i
            stack.append('}' if ch == '{' else ']')
        elif ch in '}]' and stack:
            expected = stack[-1]
            if ch == expected:
                stack.pop()
                if not stack and start_idx is not None:
                    snippet = text_clean[start_idx:i+1]
                    try:
                        parsed = json.loads(snippet)
                        return parsed
                    except Exception:
                        # if parsing fails, continue scanning for next candidate
                        start_idx = None
                        continue

    # fallback: try to parse the whole cleaned text
    try:
        return json.loads(text_clean)
    except Exception:
        pass

    # fallback: try line-by-line (ndjson style)
    for line in text_clean.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            return json.loads(line)
        except Exception:
            continue

    return {}


def split_bulk_sms(text: str) -> list:
    """
    Split bulk SMS input into individual transactions.
    Handles numbered records like:
      1. HDFC Bank: ...
      2. SBI Alert: ...
    and also blank-line-separated records.
    """
    if not text or not text.strip():
        return []

    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # Normalize weird spacing from pasted docs like '2\n \n 3'
    text = re.sub(r'\n\s*\n\s*(?=\d+\s*[.)])', '\n\n', text)
    lines = text.split('\n')
    transactions = []
    current = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            if current and '\n'.join(current).strip():
                transactions.append('\n'.join(current).strip())
                current = []
            continue

        if re.match(r'^\d+\s*[.)]\s*', stripped):
            if current and '\n'.join(current).strip():
                transactions.append('\n'.join(current).strip())
            current = [re.sub(r'^\d+\s*[.)]\s*', '', stripped, count=1)]
        else:
            current.append(stripped)

    if current and '\n'.join(current).strip():
        transactions.append('\n'.join(current).strip())

    if transactions:
        return [t.strip() for t in transactions if t and t.strip()]

    # fallback: blank-line split
    chunks = re.split(r'\n\s*\n+', text)
    return [c.strip() for c in chunks if c.strip()]



def coerce_amount(value):
    if value is None or value == "" or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    # handle currency strings, like INR 10,875.00 or USD 125.00 or Rs. 450.00
    match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
    if match:
        return float(match.group(0))
    return None


def normalize_record(raw: str, parsed: dict) -> dict:
    if not isinstance(parsed, dict):
        parsed = {}
    now = datetime.datetime.utcnow().isoformat()
    amount = coerce_amount(parsed.get("amount"))
    date_value = parsed.get("date") or parsed.get("transaction_date") or parsed.get("posted_date")
    merchant = parsed.get("merchant") or "Unknown"
    category = parsed.get("category") or "Uncategorized"
    description = parsed.get("description") or raw
    currency = parsed.get("currency") or "INR"
    status = "Tracked" if amount is not None or date_value or merchant != "Unknown" or category != "Uncategorized" else "Untracked"

    # For raw text like "SBI Alert: International Card transaction USD 125.00 (INR 10,875.00)"
    if amount is None:
        # find amount in either INR or USD patterns in the raw text
        amount_match = re.search(r"(?:INR|Rs\.|Rs|USD|EUR)\s*[: ]?\s*([-+]?\d[\d,]*\.?\d*)", raw, flags=re.I)
        if amount_match:
            amount = coerce_amount(amount_match.group(1))
            if amount is not None:
                status = "Tracked"

    if not date_value:
        match_date = re.search(r"(\d{1,2}-\d{1,2}-\d{4}|\d{1,2}-[A-Za-z]{3}-\d{4})", raw)
        if match_date:
            date_value = match_date.group(1)
            status = "Tracked"

    if merchant == "Unknown":
        # Try to infer merchant from the text
        merchant_patterns = [
            r"(?:to|at|from|via)\s+([A-Z][A-Za-z0-9 &.-]+)",
            r"(?:A/c|A/c\s+XX\d+)\s+(?:to|at|from|via)\s+([A-Z][A-Za-z0-9 &.-]+)",
            r"(?:to|at|from|via)\s+([A-Z][A-Za-z0-9&.-]+(?:\s+[A-Z][A-Za-z0-9&.-]+)*)"
        ]
        for p in merchant_patterns:
            m = re.search(p, raw, flags=re.I)
            if m:
                merchant = m.group(1).strip()
                status = "Tracked"
                break

    if category == "Uncategorized":
        category = match_category(raw + " " + (description or ""))

    return {
        "timestamp": now,
        "raw_text": raw,
        "parsed": json.dumps(parsed, ensure_ascii=False),
        "date": date_value,
        "amount": amount,
        "currency": currency,
        "merchant": merchant,
        "category": category,
        "description": description,
        "status": status,
    }


def save_records_to_excel(tracked_df: pd.DataFrame, untracked_df: pd.DataFrame):
    if tracked_df.empty and untracked_df.empty:
        return

    # Create workbook if it does not exist, otherwise open existing file.
    if not os.path.exists(RECORD_FILE):
        from openpyxl import Workbook
        wb = Workbook()
    else:
        from openpyxl import load_workbook
        wb = load_workbook(RECORD_FILE)

    for sheet_name, df in [("Transactions", tracked_df), ("Untracked", untracked_df)]:
        if df.empty:
            continue
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)

        if ws.max_row == 1 and ws["A1"].value is None:
            ws.append(list(df.columns))
        # append rows
        for _, row in df.iterrows():
            ws.append([row.get(col) for col in df.columns])

    wb.save(RECORD_FILE)


def append_record(raw: str, parsed: dict):
    record = normalize_record(raw, parsed)
    df = pd.DataFrame([record])
    tracked_df = df[df["status"] == "Tracked"].copy()
    untracked_df = df[df["status"] == "Untracked"].copy()
    save_records_to_excel(tracked_df, untracked_df)


def load_records() -> pd.DataFrame:
    default_columns = ["timestamp", "raw_text", "parsed", "date", "amount", "currency", "merchant", "category", "description", "status"]
    if not os.path.exists(RECORD_FILE):
        return pd.DataFrame(columns=default_columns)

    frames = []
    try:
        tracked = pd.read_excel(RECORD_FILE, sheet_name="Transactions", engine="openpyxl")
        # Handle Excel files where the header row was written as the first data row
        # (common when appending without proper header handling). Detect if the
        # first row contains the expected column names and promote it to header.
        try:
            if any(str(c).startswith("Unnamed") for c in tracked.columns):
                first_row_vals = tracked.iloc[0].astype(str).tolist()
                first_row_lc = [v.strip().lower() for v in first_row_vals]
                expected_lc = [c.lower() for c in default_columns]
                if set(expected_lc).issubset(set(first_row_lc)):
                    # promote first row to header
                    tracked.columns = [v.strip() for v in first_row_vals]
                    tracked = tracked.drop(tracked.index[0]).reset_index(drop=True)
        except Exception:
            pass
        for col in default_columns:
            if col not in tracked.columns:
                tracked[col] = None
        tracked["status"] = tracked.get("status", "Tracked")
        frames.append(tracked)
    except Exception:
        pass
    try:
        untracked = pd.read_excel(RECORD_FILE, sheet_name="Untracked", engine="openpyxl")
        try:
            if any(str(c).startswith("Unnamed") for c in untracked.columns):
                first_row_vals = untracked.iloc[0].astype(str).tolist()
                first_row_lc = [v.strip().lower() for v in first_row_vals]
                expected_lc = [c.lower() for c in default_columns]
                if set(expected_lc).issubset(set(first_row_lc)):
                    untracked.columns = [v.strip() for v in first_row_vals]
                    untracked = untracked.drop(untracked.index[0]).reset_index(drop=True)
        except Exception:
            pass
        for col in default_columns:
            if col not in untracked.columns:
                untracked[col] = None
        untracked["status"] = untracked.get("status", "Untracked")
        frames.append(untracked)
    except Exception:
        pass
    if not frames:
        return pd.DataFrame(columns=default_columns)
    combined = pd.concat(frames, ignore_index=True)
    for col in default_columns:
        if col not in combined.columns:
            combined[col] = None
    if "status" not in combined.columns:
        combined["status"] = "Tracked"
    combined["parsed"] = combined.get("parsed", "{}")
    combined["parsed"] = combined["parsed"].fillna("{}")
    return combined


def inject_custom_css():
    st.markdown("""
    <style>
    /* ===== Global ===== */
    .block-container { padding-top: 1rem !important; max-width: 1200px; }

    /* ===== Title Banner ===== */
    .app-header {
        background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
        color: #fff; padding: 1.4rem 2rem; border-radius: 12px;
        margin-bottom: 1.2rem; box-shadow: 0 4px 15px rgba(26,35,126,0.3);
        display: flex; align-items: center; justify-content: space-between;
    }
    .app-header h1 { margin: 0; font-size: 1.6rem; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    .app-header .subtitle { font-size: 0.85rem; opacity: 0.85; }

    /* ===== Sidebar ===== */
    section[data-testid="stSidebar"] { background-color: #f8f9fa; }
    section[data-testid="stSidebar"] .stHeader { background-color: transparent; }
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 { color: #1a237e !important; }

    /* ===== Buttons ===== */
    .stButton > button {
        border-radius: 8px; font-weight: 600; border: none;
        padding: 0.55rem 1.5rem; transition: all 0.2s ease;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    .stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    .stButton > button:active { transform: translateY(0); }

    /* Primary green buttons (Parse) */
    div[data-testid="stVerticalBlock"] button[kind="primary"],
    .btn-green {
        background: linear-gradient(135deg, #2e7d32, #43a047) !important; color: #fff !important;
    }
    /* Blue buttons (Send) */
    .btn-blue { background: linear-gradient(135deg, #1565c0, #1e88e5) !important; color: #fff !important; }
    /* Purple buttons (Generate) */
    .btn-purple { background: linear-gradient(135deg, #6a1b9a, #8e24aa) !important; color: #fff !important; }

    /* ===== Inputs ===== */
    .stTextArea textarea, .stTextInput input {
        border: 1.5px solid #e0e0e0 !important; border-radius: 8px !important;
        transition: border-color 0.2s, box-shadow 0.2s;
    }
    .stTextArea textarea:focus, .stTextInput input:focus {
        border-color: #3949ab !important; box-shadow: 0 0 0 3px rgba(57,73,171,0.15) !important;
    }

    /* ===== Tabs ===== */
    .stTabs [data-baseweb="tab"] { font-weight: 600; font-size: 0.95rem; }
    .stTabs [aria-selected="true"] { color: #1a237e !important; border-bottom-color: #1a237e !important; }

    /* ===== Metric Cards ===== */
    .metric-card {
        padding: 1rem 1.2rem; border-radius: 10px; margin-bottom: 0.8rem;
        border-left: 5px solid; background: #fff;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .metric-card .label { font-size: 0.82rem; color: #757575; margin-bottom: 0.3rem; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-card .value { font-size: 1.6rem; font-weight: 700; }
    .mc-red    { border-color: #c62828; } .mc-red .value    { color: #c62828; }
    .mc-green  { border-color: #2e7d32; } .mc-green .value  { color: #2e7d32; }
    .mc-blue   { border-color: #1565c0; } .mc-blue .value   { color: #1565c0; }
    .mc-purple { border-color: #6a1b9a; } .mc-purple .value  { color: #6a1b9a; }

    /* ===== Section Headers ===== */
    .section-hdr {
        border-bottom: 2.5px solid #1a237e; padding-bottom: 0.45rem;
        margin: 1.6rem 0 1rem 0;
    }
    .section-hdr h3 { margin: 0; color: #1a237e; font-size: 1.15rem; }

    /* ===== Chat Bubbles ===== */
    .chat-user {
        background: #e3f2fd; padding: 0.85rem 1rem; border-radius: 14px 14px 2px 14px;
        margin: 0.5rem 0; max-width: 82%; margin-left: auto;
        border: 1px solid #bbdefb;
    }
    .chat-asst {
        background: #f5f5f5; padding: 0.85rem 1rem; border-radius: 14px 14px 14px 2px;
        margin: 0.5rem 0; max-width: 82%; margin-right: auto;
        border: 1px solid #e0e0e0;
    }
    .chat-label { font-size: 0.78rem; font-weight: 700; margin-bottom: 0.25rem; }
    .chat-user .chat-label { color: #1565c0; }
    .chat-asst .chat-label { color: #2e7d32; }

    /* ===== Expanders ===== */
    .streamlit-expanderHeader { font-weight: 600 !important; }

    /* ===== Radio / Selectboxes ===== */
    .stRadio > div { gap: 0.5rem; }

    /* ===== Misc ===== */
    .stAlert { border-radius: 8px !important; }
    hr { border: none; border-top: 1px solid #e0e0e0; margin: 1rem 0; }
    </style>
    """, unsafe_allow_html=True)


def metric_card(title: str, value: str, color_cls: str):
    st.markdown(
        f'<div class="metric-card {color_cls}">'
        f'<div class="label">{title}</div>'
        f'<div class="value">{value}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def section_header(text: str):
    st.markdown(f'<div class="section-hdr"><h3>{text}</h3></div>', unsafe_allow_html=True)


def styled_title(model: str):
    st.markdown(
        f'<div class="app-header">'
        f'<div><h1>Daily Expense Analyser</h1>'
        f'<div class="subtitle">Powered by {model}</div></div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def chat_bubble(role: str, text: str):
    cls = "chat-user" if role == "You" else "chat-asst"
    label_color = "#1565c0" if role == "You" else "#2e7d32"
    st.markdown(
        f'<div class="{cls}">'
        f'<div class="chat-label" style="color:{label_color};">{role}</div>'
        f'<div>{text}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def main():
    st.set_page_config(page_title="Daily Expense Analyser", layout="wide")
    inject_custom_css()

    # --- Sidebar: Endpoint preference ---
    with st.sidebar:
        st.header("⚙️ Settings")
        api_base = os.getenv("OLLAMA_API")
        cloud_base = os.getenv("OLLAMA_CLOUD_BASE_URL")
        cloud_key = os.getenv("OLLAMA_CLOUD_API_KEY")
        
        # Initialize session state for preference
        if "use_local" not in st.session_state:
            # Default to Cloud if cloud credentials are available, otherwise Local if configured
            st.session_state.use_local = False if (cloud_base and cloud_key) else (True if api_base else False)
        
        available_endpoints = []
        if api_base:
            available_endpoints.append("Local Ollama")
        if cloud_base and cloud_key:
            available_endpoints.append("Ollama Cloud")
        
        if len(available_endpoints) > 1:
            st.subheader("Endpoint Preference")
            preference = st.selectbox(
                "Choose endpoint:",
                available_endpoints,
                index=0 if st.session_state.use_local else 1
            )
            st.session_state.use_local = (preference == "Local Ollama")
        elif len(available_endpoints) == 1:
            st.subheader("Endpoint")
            st.info(f"✓ Using: {available_endpoints[0]}")
        else:
            st.error("No Ollama endpoint configured in .env")
        
        # Show current endpoint details
        st.subheader("Current Config")
        if st.session_state.use_local and api_base:
            selected_model = os.getenv("OLLAMA_LOCAL_MODEL", "llama3.2:latest")
            st.write(f"**Local:**\n- URL: {api_base}\n- Model: {selected_model}")
        elif cloud_base and cloud_key:
            selected_model = os.getenv("OLLAMA_CLOUD_MODEL", "gpt-oss:120b-cloud")
            st.write(f"**Cloud:**\n- URL: {cloud_base}\n- Model: {selected_model}")
        else:
            selected_model = "Unconfigured"

    selected_model = os.getenv("OLLAMA_LOCAL_MODEL", "llama3.2:latest") if st.session_state.use_local and os.getenv("OLLAMA_API") else os.getenv("OLLAMA_CLOUD_MODEL", "gpt-oss:120b-cloud") if os.getenv("OLLAMA_CLOUD_BASE_URL") and os.getenv("OLLAMA_CLOUD_API_KEY") else "Unconfigured"
    styled_title(selected_model)

    tabs = st.tabs(["Add Transactions", "Chatbot", "Analysis"])

    # --- Input & Parse ---
    with tabs[0]:
        parse_mode = st.radio("Select mode:", ["Single Transaction", "Bulk SMS (multiple)"], horizontal=True)
        raw = st.text_area("Paste SMS / bank transaction text here", height=150)
        
        btn_cols = st.columns([6, 1])
        with btn_cols[1]:
            parse_clicked = st.button("Parse and Save", type="primary", use_container_width=True)
        
        if parse_clicked:
            if not raw.strip():
                st.error("Please provide some text to parse.")
            else:
                if parse_mode == "Bulk SMS (multiple)":
                    # Split bulk input into individual transactions
                    transactions = split_bulk_sms(raw)
                    if not transactions:
                        st.error("Could not detect any transactions in the bulk input.")
                    else:
                        total_txns = len(transactions)
                        st.info(f"Detected {total_txns} transaction(s). Parsing each separately...")
                        progress_text = st.empty()
                        progress_bar = st.progress(0)
                        results = []
                        for i, txn in enumerate(transactions, 1):
                            progress_text.write(f"Processing transaction {i}/{total_txns}...")
                            progress_bar.progress(i / total_txns)
                            category_guide = get_category_suggestions()
                            prompt = (
                                "You are a helpful parser that converts a single-line or multi-line financial transaction or SMS into a JSON object.\n"
                                "Extract the following fields when available: date, amount, currency, merchant, category, description.\n"
                                "If a field is not present, set it to null. Return ONLY valid JSON (no extra commentary).\n\n"
                                "For CATEGORY, use one of these options (important!):\n" + category_guide + "\n"
                                f"Input:\n{txn}\n\nOutput JSON:")
                            resp = call_ollama(prompt, prefer_local=st.session_state.use_local)
                            parsed = normalize_parsed_result(extract_json(resp))
                            if parsed.get("category"):
                                parsed["category"] = match_category(parsed["category"])
                            append_record(txn, parsed)
                            results.append({"txn_num": i, "raw": txn, "parsed": parsed})
                            progress_bar.progress(i / total_txns)
                        progress_text.write(f"Completed {total_txns}/{total_txns} transactions.")
                        
                        st.success(f"✓ Saved {len(results)} transactions to Excel")
                        section_header("Parsed Results")
                        for r in results:
                            with st.expander(f"Transaction {r['txn_num']}"):
                                st.write(f"**Raw:** {r['raw'][:100]}...")
                                st.json(r['parsed'] or {"warning": "Could not parse"})
                else:
                    # Single transaction mode
                    category_guide = get_category_suggestions()
                    prompt = (
                        "You are a helpful parser that converts a single-line or multi-line financial transaction or SMS into a JSON object.\n"
                        "Extract the following fields when available: date, amount, currency, merchant, category, description.\n"
                        "If a field is not present, set it to null. Return ONLY valid JSON (no extra commentary).\n\n"
                        "For CATEGORY, use one of these options (important!):\n" + category_guide + "\n"
                        f"Input:\n{raw}\n\nOutput JSON:")
                    progress_text = st.empty()
                    progress_bar = st.progress(0)
                    progress_text.write("Processing...")
                    progress_bar.progress(20)
                    resp = call_ollama(prompt, prefer_local=st.session_state.use_local)
                    progress_bar.progress(60)
                    parsed = normalize_parsed_result(extract_json(resp))
                    if parsed.get("category"):
                        parsed["category"] = match_category(parsed["category"])
                    append_record(raw, parsed)
                    progress_bar.progress(100)
                    progress_text.write("Completed")
                    section_header("Parsed JSON")
                    st.json(parsed or {"warning": "Could not parse LLM response, raw output shown below", "raw": resp})

    # --- Chatbot ---
    with tabs[1]:
        st.header("Expense Assistant Chatbot")
        if "history" not in st.session_state:
            st.session_state.history = []
        if "chat_processing" not in st.session_state:
            st.session_state.chat_processing = False

        # Display existing chat history
        for msg in reversed(st.session_state.history):
            chat_bubble("You", msg['user'])
            chat_bubble("Assistant", msg['assistant'])

        # Form enables Enter-to-submit and prevents double-clicks during processing
        with st.form("chat_form", clear_on_submit=False):
            user_input = st.text_input(
                "Ask about your transactions (e.g. 'get all ATM transactions', 'total spent on food')",
                placeholder="Type your question and press Enter...",
            )
            submitted = st.form_submit_button(
                "Search",
                type="primary",
                use_container_width=True,
                disabled=st.session_state.chat_processing,
            )

        if submitted and user_input.strip():
            st.session_state.chat_processing = True
            st.rerun()

        # Process after form submit (rerun enters here with flag set)
        if st.session_state.chat_processing and user_input.strip():
            # show live status updates while processing
            with st.status("Processing your query...", expanded=True) as status:
                st.write("Loading transaction records from Excel...")
                df = load_records()
                total_records = len(df)

                if df.empty:
                    st.write("No records found in the database.")
                    context = "No records found in the database."
                else:
                    st.write(f"Loaded **{total_records}** records. Building search context...")
                    for col in ["date", "amount", "merchant", "category", "status", "currency", "raw_text"]:
                        if col not in df.columns:
                            df[col] = ""

                    lines = []
                    for idx, r in df.iterrows():
                        date_str = str(r.get("date", "")) if pd.notna(r.get("date")) else "N/A"
                        amt = r.get("amount", "")
                        amt_str = f"₹{float(amt):,.2f}" if pd.notna(amt) and amt != "" else "N/A"
                        merchant = str(r.get("merchant", "Unknown")) if pd.notna(r.get("merchant")) else "Unknown"
                        category = str(r.get("category", "Uncategorized")) if pd.notna(r.get("category")) else "Uncategorized"
                        status_val = str(r.get("status", "")) if pd.notna(r.get("status")) else ""
                        currency = str(r.get("currency", "INR")) if pd.notna(r.get("currency")) else "INR"
                        raw_text = str(r.get("raw_text", "")) if pd.notna(r.get("raw_text")) else ""

                        line = (
                            f"[{idx+1}] Date: {date_str} | Amount: {amt_str} | "
                            f"Merchant: {merchant} | Category: {category} | "
                            f"Status: {status_val} | Currency: {currency} | "
                            f"Raw: {raw_text}"
                        )
                        lines.append(line)

                    context = f"Total records: {total_records}\n\nAll transactions:\n" + "\n".join(lines)

                st.write(f"Sending **{total_records}** records to LLM. Waiting for response...")
                system_prompt = (
                    "You are a personal finance assistant. You have access to the user's COMPLETE transaction history.\n"
                    "IMPORTANT RULES:\n"
                    "1. Search through ALL records provided to answer the user's question.\n"
                    "2. Filter records by the user's criteria (date range, amount, merchant, category, status, etc.).\n"
                    "3. When asked for specific transactions (e.g. ATM, Swiggy, salary), list ALL matching records with full details.\n"
                    "4. If the user asks for a summary, calculate totals from the matching records.\n"
                    "5. Always reference record numbers [N] when listing transactions.\n"
                    "6. If no matching records are found, say so clearly.\n\n"
                )
                prompt = system_prompt + "Context:\n" + context + "\n\nUser question: " + user_input + "\nAssistant:"
                answer = call_ollama(prompt, prefer_local=st.session_state.use_local)
                status.update(label="Response ready!", state="complete", expanded=False)

            st.session_state.history.append({"user": user_input, "assistant": answer})
            st.session_state.chat_processing = False
            st.rerun()

    # --- Analysis ---
    with tabs[2]:
        st.header("📊 Comprehensive Analytics Dashboard")
        df = load_records()
        if df.empty:
            st.info("No records found. Parse and save some transactions first.")
        else:
            if "parsed" not in df.columns:
                st.warning("The saved sheet does not contain a parsed JSON column. Re-save or re-parse records to continue.")
                st.stop()
            # expand parsed JSON into dataframe
            def safe_parse_json(x):
                # Accept native dicts, strings containing JSON, or return empty dict
                if isinstance(x, dict):
                    return x
                if isinstance(x, str):
                    if not x.strip():
                        return {}
                    try:
                        obj = json.loads(x)
                        if isinstance(obj, dict):
                            return obj
                        else:
                            return {"raw_value": str(obj)}
                    except Exception:
                        return {}
                return {}
            
            parsed_expanded = df["parsed"].apply(safe_parse_json)

            # Diagnostics: show counts to help debug why analytics may be empty
            try:
                total_rows = len(df)
                parsed_nonempty = sum(1 for p in parsed_expanded if p)
            except Exception:
                total_rows = len(df)
                parsed_nonempty = 0

            with st.expander("Diagnostics: parsed data overview", expanded=False):
                st.write(f"Total rows: {total_rows}")
                st.write(f"Parsed JSON objects (non-empty): {parsed_nonempty}")

            # If parsed JSON objects are empty (e.g. "{}"), try to derive basic fields
            # from existing dataframe columns (amount, date, merchant) so analytics can still run.
            parsed_filled = []
            for idx, p in parsed_expanded.items():
                if p:
                    parsed_filled.append(p)
                    continue
                # attempt to build a minimal parsed dict from columns in df
                row = df.iloc[idx] if idx < len(df) else None
                built = {}
                if row is not None:
                    if "amount" in row and not pd.isna(row.get("amount")):
                        built["amount"] = row.get("amount")
                    if "date" in row and not pd.isna(row.get("date")):
                        built["date"] = str(row.get("date"))
                    if "merchant" in row and row.get("merchant") and str(row.get("merchant")).strip():
                        built["merchant"] = row.get("merchant")
                    if "category" in row and row.get("category") and str(row.get("category")).strip():
                        built["category"] = row.get("category")
                parsed_filled.append(built)

            # If still no useful parsed data, warn the user
            non_empty = [p for p in parsed_filled if p and any(v not in (None, "") for v in p.values())]
            if not non_empty:
                st.warning("No valid parsed data to display. Try parsing a transaction first.")
                st.stop()

            parsed_df = pd.json_normalize(parsed_filled, errors='ignore')
            combined = pd.concat([df[["timestamp", "raw_text", "status"]].reset_index(drop=True), parsed_df.reset_index(drop=True)], axis=1)
            
            # Prepare data: convert timestamp and amount
            combined["timestamp"] = pd.to_datetime(combined["timestamp"], errors='coerce')
            combined["amount"] = pd.to_numeric(combined["amount"], errors="coerce")
            if "date" in combined.columns:
                # Parse common date formats (dd-mm-yyyy, dd-Mon-yyyy, ISO) with dayfirst
                combined["date"] = pd.to_datetime(combined["date"], errors='coerce', dayfirst=True)
            else:
                combined["date"] = pd.NaT
            if "status" not in combined.columns:
                combined["status"] = "Tracked"
            combined["merchant"] = combined.get("merchant", "Unknown").fillna("Unknown")
            combined["category"] = combined.get("category", "Uncategorized").fillna("Uncategorized")
            combined["currency"] = combined.get("currency", "INR").fillna("INR")
            combined["description"] = combined.get("description", combined["raw_text"]).fillna(combined["raw_text"])
            
            # Classify as debit or credit based on description/merchant keywords
            def classify_transaction(row):
                raw = str(row.get("raw_text", "")).lower() + " " + str(row.get("description", "")).lower()
                debit_keywords = ["debited", "spent", "payment", "withdrawal", "debit", "paid"]
                credit_keywords = ["credited", "deposited", "credit", "salary", "income"]
                
                for kw in debit_keywords:
                    if kw in raw:
                        return "Debit"
                for kw in credit_keywords:
                    if kw in raw:
                        return "Credit"
                return "Unknown"
            
            combined["txn_type"] = combined.apply(classify_transaction, axis=1)
            
            # Extract merchant from description/merchant column
            if "merchant" not in combined.columns:
                combined["merchant"] = combined.get("description", "Unknown").fillna("Unknown")
            else:
                combined["merchant"] = combined["merchant"].fillna("Unknown")
            
            combined["category"] = combined.get("category", "Uncategorized").fillna("Uncategorized")
            
            # Time-based filtering
            section_header("⏰ Filter by Time Period")
            col1, col2, col3 = st.columns(3)
            with col1:
                period = st.selectbox("Select period:", ["Daily", "Monthly", "Yearly", "All"], key="period_select")
            with col2:
                status_filter = st.selectbox("Status:", ["All", "Tracked", "Untracked"], key="status_filter")
            
            # Filter data based on period
            if period == "Daily":
                if "date" in combined.columns:
                    selected_date = st.date_input("Select date:", pd.Timestamp("today").date())
                    filtered = combined[combined["date"].dt.date == selected_date]
                else:
                    filtered = combined
            elif period == "Monthly":
                if "date" in combined.columns:
                    months = combined["date"].dt.to_period("M").dropna().astype(str).unique().tolist()
                    if months:
                        selected_month = st.selectbox("Select month:", months, key="month_select")
                        filtered = combined[combined["date"].dt.to_period("M").astype(str) == selected_month]
                    else:
                        filtered = combined
                else:
                    filtered = combined
            elif period == "Yearly":
                if "date" in combined.columns:
                    selected_year = st.selectbox("Select year:", 
                        sorted(combined["date"].dt.year.dropna().unique()), key="year_select")
                    filtered = combined[combined["date"].dt.year == selected_year]
                else:
                    filtered = combined
            else:
                filtered = combined

            if status_filter != "All":
                filtered = filtered[filtered["status"] == status_filter]
            
            if filtered.empty:
                st.info(f"No transactions found for {period.lower()}.")
                st.stop()
            
            # --- Summary Metrics Row 1 ---
            section_header("📈 Summary Metrics")
            metric_cols = st.columns(4)
            
            total_debit = filtered[filtered["txn_type"] == "Debit"]["amount"].sum()
            total_credit = filtered[filtered["txn_type"] == "Credit"]["amount"].sum()
            net = total_credit - total_debit
            txn_count = len(filtered)
            untracked_count = int((filtered["status"] == "Untracked").sum())
            tracked_count = int((filtered["status"] == "Tracked").sum())
            
            with metric_cols[0]:
                metric_card("Total Debits", f"₹{total_debit:.2f}" if total_debit > 0 else "₹0.00", "mc-red")
            with metric_cols[1]:
                metric_card("Total Credits", f"₹{total_credit:.2f}" if total_credit > 0 else "₹0.00", "mc-green")
            with metric_cols[2]:
                metric_card("Net (Credits - Debits)", f"₹{net:.2f}", "mc-blue")
            with metric_cols[3]:
                metric_card("Tracked / Untracked", f"{tracked_count} / {untracked_count}", "mc-purple")
            
            # --- Debit vs Credit Comparison ---
            section_header("💳 Debit vs Credit Breakdown")
            debit_credit = filtered.groupby("txn_type")["amount"].sum()
            col1, col2 = st.columns(2)
            with col1:
                st.bar_chart(debit_credit)
            with col2:
                st.write(debit_credit.to_frame().rename(columns={"amount": "Total Amount"}))
            
            # --- Category-wise Breakdown ---
            section_header("🏷️ Spending by Category")
            if not filtered["category"].empty:
                cat_summary = filtered.groupby("category")["amount"].agg(["sum", "count"]).sort_values("sum", ascending=False)
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Top Categories:**")
                    st.dataframe(cat_summary)
                with col2:
                    if len(cat_summary) > 0:
                        st.write("**Category Distribution:**")
                        st.bar_chart(cat_summary["sum"])
                
                # Pie chart for categories
                st.write("**Category Pie Chart:**")
                st.write(filtered.groupby("category")["amount"].sum())
                import matplotlib.pyplot as plt
                fig, ax = plt.subplots(figsize=(8, 6))
                filtered.groupby("category")["amount"].sum().plot(kind="pie", autopct="%1.1f%%", ax=ax)
                ax.set_ylabel("")
                st.pyplot(fig)
            
            # --- Merchant-wise Breakdown ---
            section_header("🏪 Spending by Merchant/Vendor")
            merchant_summary = filtered.groupby("merchant")["amount"].agg(["sum", "count"]).sort_values("sum", ascending=False)
            st.write("**Top Merchants:**")
            st.dataframe(merchant_summary.head(15))
            
            if len(merchant_summary) > 0:
                st.write("**Merchant Spending Chart:**")
                st.bar_chart(merchant_summary["sum"].head(10))
            
            # --- Transaction Trend Over Time ---
            section_header("📉 Transaction Trends")
            if "timestamp" in combined.columns and not combined["timestamp"].isna().all():
                daily_trend = filtered.groupby(filtered["timestamp"].dt.date)["amount"].sum()
                st.line_chart(daily_trend)
            
            # --- Detailed Transaction Table ---
            section_header("📋 All Transactions")
            display_cols = ["timestamp", "raw_text", "amount", "category", "merchant", "txn_type", "status"]
            available_display_cols = [c for c in display_cols if c in filtered.columns]
            st.dataframe(filtered[available_display_cols].sort_values("timestamp", ascending=False))
            
            # --- LLM-based Insights ---
            section_header("🤖 AI-Powered Insights")
            gcols = st.columns([3, 2, 3])
            with gcols[1]:
                gen_clicked = st.button("Generate LLM Analysis", type="primary", use_container_width=True)
            if gen_clicked:
                # Prepare serializable sample rows
                sample_df = filtered.tail(30).copy()
                # Convert all Timestamp/datetime objects to strings
                for col in sample_df.columns:
                    if pd.api.types.is_datetime64_any_dtype(sample_df[col]):
                        sample_df[col] = sample_df[col].astype(str)
                    elif col == "amount":
                        sample_df[col] = pd.to_numeric(sample_df[col], errors='coerce').fillna(0).astype(float)
                
                sample_rows = sample_df.to_dict(orient="records")
                
                # Prepare summary stats with all values JSON-serializable
                summary_stats = {
                    "period": str(period),
                    "total_debits": float(total_debit) if not pd.isna(total_debit) else 0.0,
                    "total_credits": float(total_credit) if not pd.isna(total_credit) else 0.0,
                    "net": float(net) if not pd.isna(net) else 0.0,
                    "transaction_count": int(txn_count),
                    "top_categories": {str(k): float(v) for k, v in filtered.groupby("category")["amount"].sum().nlargest(5).to_dict().items()},
                    "top_merchants": {str(k): float(v) for k, v in filtered.groupby("merchant")["amount"].sum().nlargest(5).to_dict().items()}
                }
                
                prompt = (
                    "You are a financial analyst. Analyze the following transaction data and provide actionable insights:\n\n"
                    f"Summary Stats: {json.dumps(summary_stats, ensure_ascii=False)}\n\n"
                    f"Sample Transactions: {json.dumps(sample_rows, ensure_ascii=False, default=str)}\n\n"
                    "Provide insights on:\n"
                    "1. Overall spending patterns\n"
                    "2. Top spending categories and merchants\n"
                    "3. Debit vs Credit patterns\n"
                    "4. Recommendations for expense management\n"
                    "5. Any unusual transactions or trends\n\n"
                    "Analysis:")
                insight = call_ollama(prompt, prefer_local=st.session_state.use_local)
                st.write(insight)


if __name__ == "__main__":
    main()
