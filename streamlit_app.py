import os
import json
import re
import random
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

    /* ===== Top Ribbon ===== */
    .top-ribbon {
        background: linear-gradient(135deg, #1565c0, #1976d2, #1e88e5);
        color: #fff; padding: 0.55rem 1.5rem; border-radius: 8px;
        margin-bottom: 1rem; display: flex; align-items: center;
        justify-content: space-between; box-shadow: 0 2px 8px rgba(21,101,192,0.3);
        font-size: 0.85rem;
    }
    .top-ribbon .ribbon-label { font-weight: 600; opacity: 0.9; }
    .top-ribbon .ribbon-model { font-weight: 700; background: rgba(255,255,255,0.2); padding: 0.15rem 0.6rem; border-radius: 4px; }

    /* ===== Story Card ===== */
    .story-card {
        background: linear-gradient(135deg, #fff8e1 0%, #fff3e0 40%, #fce4ec 70%, #f3e5f5 100%);
        border: 2px solid #ff9800; border-radius: 16px;
        padding: 1.5rem 1.8rem; margin-top: 1.5rem;
        box-shadow: 0 6px 25px rgba(255,152,0,0.18);
        position: relative; overflow: hidden;
    }
    .story-card::before {
        content: ''; position: absolute; top: -30px; right: -30px;
        width: 120px; height: 120px; border-radius: 50%;
        background: rgba(255,152,0,0.08);
    }
    .story-card::after {
        content: ''; position: absolute; bottom: -40px; left: -20px;
        width: 100px; height: 100px; border-radius: 50%;
        background: rgba(233,30,99,0.06);
    }
    .story-card h4 { margin: 0 0 1rem 0; color: #e65100; font-size: 1.15rem; }
    .story-item { display: flex; align-items: flex-start; gap: 0.7rem; margin-bottom: 0.7rem; font-size: 0.9rem; color: #333; }
    .story-icon { font-size: 1.4rem; flex-shrink: 0; margin-top: 0.05rem; }
    .story-amount { font-weight: 700; }
    .story-highlight { color: #c62828; }
    .story-good { color: #2e7d32; }
    .story-badge {
        display: inline-block; padding: 0.15rem 0.5rem; border-radius: 12px;
        font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.3px;
    }
    .badge-danger { background: #ffcdd2; color: #c62828; }
    .badge-warn { background: #fff3e0; color: #e65100; }
    .badge-success { background: #c8e6c9; color: #2e7d32; }
    .badge-info { background: #e3f2fd; color: #1565c0; }

    /* ===== Value-Add Purchase Box ===== */
    .valuebox {
        background: linear-gradient(135deg, #e8f5e9 0%, #e0f2f1 30%, #e3f2fd 60%, #ede7f6 100%);
        border: 2px solid #43a047; border-radius: 16px;
        padding: 1.5rem 1.8rem; margin-top: 1.5rem;
        box-shadow: 0 6px 25px rgba(67,160,71,0.15);
        position: relative; overflow: hidden;
    }
    .valuebox::before {
        content: ''; position: absolute; top: -25px; right: -25px;
        width: 100px; height: 100px; border-radius: 50%;
        background: rgba(67,160,71,0.08);
    }
    .valuebox h4 { margin: 0 0 0.4rem 0; color: #2e7d32; font-size: 1.15rem; }
    .valuebox .v-subtitle { font-size: 0.82rem; color: #555; margin-bottom: 1rem; }

    .vcard {
        background: #fff; border-radius: 12px; padding: 1rem 1.2rem;
        margin-bottom: 0.8rem; border-left: 5px solid #43a047;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        transition: transform 0.15s;
    }
    .vcard:hover { transform: translateX(4px); }
    .vcard-header { display: flex; align-items: center; gap: 0.6rem; margin-bottom: 0.5rem; }
    .vcard-emoji { font-size: 1.5rem; }
    .vcard-title { font-weight: 700; font-size: 0.95rem; color: #1a237e; }
    .vcard-amount { font-weight: 700; font-size: 1.1rem; color: #c62828; }
    .vcard-meta { font-size: 0.78rem; color: #777; margin-bottom: 0.4rem; }
    .vcard-story { font-size: 0.85rem; color: #333; line-height: 1.45; }
    .vcard-tag { display: inline-block; padding: 0.12rem 0.5rem; border-radius: 10px; font-size: 0.7rem; font-weight: 600; margin-top: 0.3rem; }
    .tag-phone { background: #e3f2fd; color: #1565c0; }
    .tag-appliance { background: #e8f5e9; color: #2e7d32; }
    .tag-gadget { background: #ede7f6; color: #6a1b9a; }
    .tag-furniture { background: #fff3e0; color: #e65100; }
    .tag-home { background: #fce4ec; color: #c62828; }
    .tag-luxury { background: #fff8e1; color: #f57f17; }
    .tag-other { background: #eceff1; color: #455a64; }

    /* ===== Sidebar ===== */
    section[data-testid="stSidebar"] { background-color: #f8f9fa; }
    section[data-testid="stSidebar"] .stHeader { background-color: transparent; }
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 { color: #1a237e !important; }

    /* Sidebar analytics cards */
    .side-stat {
        padding: 0.65rem 0.8rem; border-radius: 8px; margin-bottom: 0.5rem;
        border-left: 4px solid; background: #fff;
    }
    .side-stat .s-label { font-size: 0.72rem; color: #888; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 0.15rem; }
    .side-stat .s-value { font-size: 1.15rem; font-weight: 700; margin: 0; }
    .ss-red    { border-color: #c62828; } .ss-red .s-value    { color: #c62828; }
    .ss-green  { border-color: #2e7d32; } .ss-green .s-value  { color: #2e7d32; }
    .ss-blue   { border-color: #1565c0; } .ss-blue .s-value   { color: #1565c0; }
    .ss-orange { border-color: #e65100; } .ss-orange .s-value  { color: #e65100; }
    .ss-purple { border-color: #6a1b9a; } .ss-purple .s-value  { color: #6a1b9a; }
    .ss-teal   { border-color: #00796b; } .ss-teal .s-value   { color: #00796b; }

    .side-divider { border: none; border-top: 1.5px solid #e0e0e0; margin: 0.8rem 0; }
    .side-section-title { font-size: 0.82rem; font-weight: 700; color: #1a237e; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 0.5rem; }

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

    # --- Sidebar: model selector + analytics ---
    with st.sidebar:
        api_base = os.getenv("OLLAMA_API")
        cloud_base = os.getenv("OLLAMA_CLOUD_BASE_URL")
        cloud_key = os.getenv("OLLAMA_CLOUD_API_KEY")

        if "use_local" not in st.session_state:
            st.session_state.use_local = False if (cloud_base and cloud_key) else (True if api_base else False)

        # ── Model Selector (compact, top of sidebar) ──
        available_endpoints = []
        if api_base:
            available_endpoints.append("Local Ollama")
        if cloud_base and cloud_key:
            available_endpoints.append("Ollama Cloud")

        if len(available_endpoints) > 1:
            pref = st.selectbox(
                "Model",
                available_endpoints,
                index=0 if st.session_state.use_local else 1,
            )
            st.session_state.use_local = (pref == "Local Ollama")
        elif len(available_endpoints) == 1:
            st.markdown(f"**Model:** {available_endpoints[0]}")
        else:
            st.error("No Ollama endpoint configured")

        selected_model = (
            os.getenv("OLLAMA_LOCAL_MODEL", "llama3.2:latest")
            if st.session_state.use_local and api_base
            else os.getenv("OLLAMA_CLOUD_MODEL", "gpt-oss:120b-cloud")
            if cloud_base and cloud_key
            else "Unconfigured"
        )
        st.caption(f"Active: `{selected_model}`")

        st.markdown('<hr class="side-divider">', unsafe_allow_html=True)

        # ── Quick Analytics ──────────────────────────────────────────
        st.markdown('<hr class="side-divider">', unsafe_allow_html=True)
        st.markdown('<div class="side-section-title">Quick Analytics</div>', unsafe_allow_html=True)

        _sdf = load_records()
        if _sdf.empty:
            st.info("No transactions yet. Add some to see analytics here.")
        else:
            # build a small combined frame for stats
            def _safe_json(x):
                if isinstance(x, dict): return x
                if isinstance(x, str) and x.strip():
                    try:
                        obj = json.loads(x)
                        return obj if isinstance(obj, dict) else {}
                    except Exception: pass
                return {}

            _parsed = _sdf["parsed"].apply(_safe_json) if "parsed" in _sdf.columns else pd.Series([{}]*len(_sdf))
            _p_df = pd.json_normalize(_parsed.fillna({}).tolist(), errors="ignore") if len(_parsed) else pd.DataFrame()
            for c in ["amount", "date", "merchant", "category", "status", "currency"]:
                if c not in _p_df.columns:
                    _p_df[c] = _sdf[c].values if c in _sdf.columns else None
            _p_df["amount"] = pd.to_numeric(_p_df["amount"], errors="coerce")
            if "date" in _p_df.columns:
                _p_df["date"] = pd.to_datetime(_p_df["date"], errors="coerce", dayfirst=True)
            _p_df["merchant"] = _p_df["merchant"].fillna("Unknown")
            _p_df["category"] = _p_df["category"].fillna("Uncategorized")

            # classify debit / credit
            def _cls(row):
                txt = str(row.get("raw_text", _sdf.iloc[row.name].get("raw_text", ""))).lower() if row.name < len(_sdf) else ""
                if any(k in txt for k in ["debited", "spent", "payment", "withdrawal", "debit", "paid"]):
                    return "Debit"
                if any(k in txt for k in ["credited", "deposited", "credit", "salary", "income"]):
                    return "Credit"
                return "Unknown"
            _p_df["txn_type"] = _p_df.apply(_cls, axis=1) if len(_p_df) else pd.Series(dtype=str)

            _total_spend = _p_df.loc[_p_df["txn_type"] == "Debit", "amount"].sum()
            _total_income = _p_df.loc[_p_df["txn_type"] == "Credit", "amount"].sum()
            _txn_count = len(_p_df)
            _net = _total_income - _total_spend

            # ── Summary cards ──
            def _scard(label, value, css):
                return f'<div class="side-stat {css}"><div class="s-label">{label}</div><div class="s-value">{value}</div></div>'

            st.markdown(_scard("Total Spent", f"₹{_total_spend:,.0f}", "ss-red"), unsafe_allow_html=True)
            st.markdown(_scard("Total Received", f"₹{_total_income:,.0f}", "ss-green"), unsafe_allow_html=True)
            st.markdown(_scard("Net", f"₹{_net:,.0f}", "ss-blue"), unsafe_allow_html=True)
            st.markdown(_scard("Transactions", str(_txn_count), "ss-purple"), unsafe_allow_html=True)

            # ── Top spending categories ──
            _debits = _p_df.loc[_p_df["txn_type"] == "Debit"]
            if not _debits.empty and "category" in _debits.columns:
                st.markdown('<hr class="side-divider">', unsafe_allow_html=True)
                st.markdown('<div class="side-section-title">Top Spending Categories</div>', unsafe_allow_html=True)
                _cat = _debits.groupby("category")["amount"].sum().sort_values(ascending=False).head(5)
                _max_cat = _cat.max() if len(_cat) else 1
                for _cat_name, _cat_val in _cat.items():
                    _pct = int((_cat_val / _max_cat) * 100) if _max_cat else 0
                    st.markdown(
                        f'<div style="margin-bottom:0.35rem;">'
                        f'<div style="font-size:0.78rem;color:#444;font-weight:600;">{_cat_name}</div>'
                        f'<div style="background:#e8eaf6;border-radius:4px;height:8px;">'
                        f'<div style="background:linear-gradient(90deg,#3949ab,#5c6bc0);width:{_pct}%;height:100%;border-radius:4px;"></div></div>'
                        f'<div style="font-size:0.72rem;color:#888;">₹{_cat_val:,.0f}</div></div>',
                        unsafe_allow_html=True,
                    )

            # ── Top merchants ──
            if not _debits.empty and "merchant" in _debits.columns:
                st.markdown('<hr class="side-divider">', unsafe_allow_html=True)
                st.markdown('<div class="side-section-title">Top Merchants</div>', unsafe_allow_html=True)
                _merch = _debits.groupby("merchant")["amount"].sum().sort_values(ascending=False).head(5)
                _max_m = _merch.max() if len(_merch) else 1
                for _m_name, _m_val in _merch.items():
                    _pct = int((_m_val / _max_m) * 100) if _max_m else 0
                    st.markdown(
                        f'<div style="margin-bottom:0.35rem;">'
                        f'<div style="font-size:0.78rem;color:#444;font-weight:600;">{_m_name}</div>'
                        f'<div style="background:#e8f5e9;border-radius:4px;height:8px;">'
                        f'<div style="background:linear-gradient(90deg,#2e7d32,#66bb6a);width:{_pct}%;height:100%;border-radius:4px;"></div></div>'
                        f'<div style="font-size:0.72rem;color:#888;">₹{_m_val:,.0f}</div></div>',
                        unsafe_allow_html=True,
                    )

            # ── Monthly spending mini chart ──
            if "date" in _p_df.columns and not _p_df["date"].isna().all() and not _debits.empty:
                st.markdown('<hr class="side-divider">', unsafe_allow_html=True)
                st.markdown('<div class="side-section-title">Monthly Spending</div>', unsafe_allow_html=True)
                _monthly = _debits.copy()
                _monthly["month"] = _monthly["date"].dt.to_period("M")
                _m_spend = _monthly.groupby("month")["amount"].sum().sort_index().tail(6)
                _max_ms = _m_spend.max() if len(_m_spend) else 1
                for _m_label, _m_val in _m_spend.items():
                    _pct = int((_m_val / _max_ms) * 100) if _max_ms else 0
                    _bar_color = "#c62828" if _pct > 80 else "#e65100" if _pct > 50 else "#2e7d32"
                    st.markdown(
                        f'<div style="margin-bottom:0.35rem;">'
                        f'<div style="font-size:0.78rem;color:#444;font-weight:600;">{str(_m_label)}</div>'
                        f'<div style="background:#fce4ec;border-radius:4px;height:8px;">'
                        f'<div style="background:{_bar_color};width:{_pct}%;height:100%;border-radius:4px;"></div></div>'
                        f'<div style="font-size:0.72rem;color:#888;">₹{_m_val:,.0f}</div></div>',
                        unsafe_allow_html=True,
                    )

            # ── Fun story / insights ──
            if not _debits.empty:
                st.markdown('<hr class="side-divider">', unsafe_allow_html=True)
                st.markdown('<div class="side-section-title">Spending Story</div>', unsafe_allow_html=True)
                _top_cat = _debits.groupby("category")["amount"].idxmax()
                _top_cat_name = _debits.loc[_top_cat.iloc[0], "category"] if len(_top_cat) else "N/A"
                _food = _debits[_debits["category"].str.contains("Food|Restaurant|Cafe|Swiggy|Zomato", case=False, na=False)]
                _travel = _debits[_debits["category"].str.contains("Fuel|Taxi|Uber|Ola|Flight|Hotel|Train", case=False, na=False)]
                _ent = _debits[_debits["category"].str.contains("Movie|OTT|Gaming|Music|Sports|Event", case=False, na=False)]
                _food_total = _food["amount"].sum()
                _travel_total = _travel["amount"].sum()
                _ent_total = _ent["amount"].sum()
                _total_all = _debits["amount"].sum() or 1

                _story_parts = []
                _story_parts.append(f"Your biggest spending bucket is <b>{_top_cat_name}</b>.")
                if _food_total > 0:
                    _p = int((_food_total / _total_all) * 100)
                    _story_parts.append(f"Food & dining eats up <b>{_p}%</b> of your wallet — ₹{_food_total:,.0f}{'  Time to cook more?' if _p > 20 else ''}")
                if _travel_total > 0:
                    _p = int((_travel_total / _total_all) * 100)
                    _story_parts.append(f"Travel & fuel cost <b>{_p}%</b> — ₹{_travel_total:,.0f}{'  Roads are expensive!' if _p > 15 else ''}")
                if _ent_total > 0:
                    _p = int((_ent_total / _total_all) * 100)
                    _story_parts.append(f"Entertainment spend: <b>{_p}%</b> — ₹{_ent_total:,.0f}{'  Netflix binges showing up here?' if _p > 10 else ''}")

                _avg_txn = _debits["amount"].mean()
                _story_parts.append(f"Average transaction: <b>₹{_avg_txn:,.0f}</b>.")

                for _s in _story_parts:
                    st.markdown(
                        f'<div style="font-size:0.78rem;color:#333;padding:0.2rem 0;">{_s}</div>',
                        unsafe_allow_html=True,
                    )

    selected_model = (
        os.getenv("OLLAMA_LOCAL_MODEL", "llama3.2:latest")
        if st.session_state.use_local and api_base
        else os.getenv("OLLAMA_CLOUD_MODEL", "gpt-oss:120b-cloud")
        if cloud_base and cloud_key
        else "Unconfigured"
    )
    styled_title(selected_model)

    tabs = st.tabs(["Add Transactions", "Chatbot", "Analysis"])

    # --- Input & Parse ---
    with tabs[0]:
        raw = st.text_area(
            "Paste one or more SMS / bank transaction texts",
            height=150,
            placeholder="Single SMS or multiple transactions (numbered or separated by blank lines)...",
        )

        btn_cols = st.columns([6, 1])
        with btn_cols[1]:
            parse_clicked = st.button("Parse and Save", type="primary", use_container_width=True)

        if parse_clicked:
            if not raw.strip():
                st.error("Please provide some text to parse.")
            else:
                transactions = split_bulk_sms(raw)
                if not transactions:
                    st.error("Could not detect any transactions in the input.")
                else:
                    total_txns = len(transactions)
                    category_guide = get_category_suggestions()
                    progress_text = st.empty()
                    progress_bar = st.progress(0)
                    results = []
                    for i, txn in enumerate(transactions, 1):
                        progress_text.write(f"Processing transaction {i}/{total_txns}...")
                        progress_bar.progress(i / total_txns)
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

                    st.success(f"Saved {len(results)} transaction(s) to Excel")
                    section_header("Parsed Results")
                    for r in results:
                        with st.expander(f"Transaction {r['txn_num']}"):
                            st.write(f"**Raw:** {r['raw'][:100]}...")
                            st.json(r['parsed'] or {"warning": "Could not parse"})

        # ── Current Month Spending Story ──────────────────────────────
        _sdf = load_records()
        if not _sdf.empty:
            def _sj(x):
                if isinstance(x, dict): return x
                if isinstance(x, str) and x.strip():
                    try:
                        o = json.loads(x)
                        return o if isinstance(o, dict) else {}
                    except Exception: pass
                return {}

            _sp = _sdf["parsed"].apply(_sj) if "parsed" in _sdf.columns else pd.Series([{}]*len(_sdf))
            _spdf = pd.json_normalize(_sp.fillna({}).tolist(), errors="ignore") if len(_sp) else pd.DataFrame()
            for c in ["amount", "date", "merchant", "category", "currency", "raw_text"]:
                if c not in _spdf.columns:
                    _spdf[c] = _sdf[c].values if c in _sdf.columns else None
            _spdf["amount"] = pd.to_numeric(_spdf["amount"], errors="coerce")
            if "date" in _spdf.columns:
                _spdf["date"] = pd.to_datetime(_spdf["date"], errors="coerce", dayfirst=True)
            _spdf["merchant"] = _spdf["merchant"].fillna("Unknown")
            _spdf["category"] = _spdf["category"].fillna("Uncategorized")

            now = pd.Timestamp.now()
            cur_month = now.to_period("M")
            if "date" in _spdf.columns and not _spdf["date"].isna().all():
                _month_df = _spdf[_spdf["date"].dt.to_period("M") == cur_month].copy()
            else:
                _month_df = _spdf.copy()

            if not _month_df.empty:
                def _cls2(row):
                    raw = str(row.get("raw_text", "")).lower()
                    if not raw and row.name < len(_sdf):
                        raw = str(_sdf.iloc[row.name].get("raw_text", "")).lower()
                    # Credit card = money going OUT, treat as debit
                    _is_credit_card = "credit card" in raw or "credit card" in str(row.get("category", "")).lower()
                    if _is_credit_card or any(k in raw for k in ["debited", "spent", "payment", "withdrawal", "debit", "paid", "used for", "charged", "purchase at", "card used", "pos purchase", "txn"]):
                        return "Debit"
                    if any(k in raw for k in ["credited", "deposited", "salary", "income", "refund"]):
                        return "Credit"
                    # If amount is negative or has debit hints, classify as debit
                    amt = row.get("amount", 0)
                    if pd.notna(amt) and float(amt) < 0:
                        return "Debit"
                    return "Unknown"
                _month_df["txn_type"] = _month_df.apply(_cls2, axis=1)

                _debits = _month_df[_month_df["txn_type"] == "Debit"]
                _credits = _month_df[_month_df["txn_type"] == "Credit"]
                _total_spent = _debits["amount"].sum()
                _total_recv = _credits["amount"].sum()
                _total_all = _total_spent or 1
                _txn_count = len(_debits)
                _avg = _debits["amount"].mean() if _txn_count else 0
                _max_txn = _debits["amount"].max() if _txn_count else 0
                _max_merchant = _debits.loc[_debits["amount"].idxmax(), "merchant"] if _txn_count else "N/A"
                _max_category = _debits.loc[_debits["amount"].idxmax(), "category"] if _txn_count else "N/A"

                # category buckets
                _food_kw = "Food|Restaurant|Cafe|Swiggy|Zomato|Fast Food|Bakery|Fine Dining"
                _travel_kw = "Fuel|Diesel|Petrol|Taxi|Uber|Ola|Flight|Hotel|Train|Bus|Metro|Parking|FASTag"
                _gadget_kw = "Amazon|Flipkart|Electronics|Mobile|Gadget|Gaming|Laptop|Apple|Samsung"
                _ent_kw = "Movie|OTT|Netflix|Prime|Disney|Gaming|Music|Sports|Event|SonyLIV|JioCinema"
                _health_kw = "Hospital|Doctor|Pharmacy|Medical|Health Insurance|Apollo|Fortis"
                _shop_kw = "Supermarket|Department|Furniture|Jewelry|Fashion|Myntra|Ajio|Nykaa|Lifestyle"
                _bills_kw = "Electricity|Water|Gas|Broadband|Mobile Recharge|DTH|Bill"

                def _cat_total(kw):
                    m = _debits["category"].str.contains(kw, case=False, na=False) | _debits["merchant"].str.contains(kw, case=False, na=False)
                    return _debits.loc[m, "amount"].sum()

                _food_t = _cat_total(_food_kw)
                _travel_t = _cat_total(_travel_kw)
                _gadget_t = _cat_total(_gadget_kw)
                _ent_t = _cat_total(_ent_kw)
                _health_t = _cat_total(_health_kw)
                _shop_t = _cat_total(_shop_kw)
                _bills_t = _cat_total(_bills_kw)

                _month_label = cur_month.strftime("%B %Y")

                # ── Exaggerated dramatic storylines ──
                _story_items = []

                # Opening headline
                if _total_spent > _total_recv and _total_recv > 0:
                    _story_items.append(("🔥", f"<b>BREAKING:</b> Your wallet lost <span class='story-amount story-highlight'>₹{_total_spent:,.0f}</span> this month while only <span class='story-amount story-good'>₹{_total_recv:,.0f}</span> trickled back in. That's a <span class='story-highlight'>{((_total_spent/_total_recv - 1)*100):.0f}% overspend rate!</b>"))
                elif _total_recv > 0:
                    _story_items.append(("💰", f"<b>MIRACLE ALERT:</b> You somehow survived {_month_label} spending just <span class='story-amount'>₹{_total_spent:,.0f}</span> while raking in <span class='story-amount story-good'>₹{_total_recv:,.0f}</span>. Your bank account is throwing a party!"))
                else:
                    _story_items.append(("💸", f"<b>EMERGENCY BROADCAST:</b> <span class='story-amount story-highlight'>₹{_total_spent:,.0f}</span> vanished from your accounts in {_month_label}. No salary credits detected. Your money went on an adventure without you."))

                # Biggest single hit
                if _txn_count > 0:
                    _story_items.append(("🎯", f"<b>BIGGEST HEIST:</b> <span class='story-highlight'>{_max_merchant}</span> pulled off the largest robbery of <span class='story-amount story-highlight'>₹{_max_txn:,.0f}</span> in the <span class='story-badge badge-info'>{_max_category}</span> category. Truly legendary."))

                # Food saga
                if _food_t > 0:
                    _fp = int((_food_t / _total_all) * 100)
                    _food_jokes = [
                        f"Your stomach is living its best life while your wallet weeps in the corner.",
                        f"Swiggy/Zomato delivery guys probably know your address better than your own family now.",
                        f"You could've bought a small fridge with that money... which you'd fill with more food.",
                        f"Gordon Ramsay would be proud. Your bank account? Not so much.",
                    ]
                    _food_msg = random.choice(_food_jokes)
                    _badge = "badge-danger" if _fp > 25 else "badge-warn" if _fp > 15 else "badge-success"
                    _story_items.append(("🍔", f"<b>THE FOOD SAGA:</b> <span class='story-amount story-highlight'>₹{_food_t:,.0f}</span> ({_fp}% of your empire) went straight to your stomach. <span class='story-badge {_badge}'>{_fp}% SHARE</span> {_food_msg}"))

                # Travel adventure
                if _travel_t > 0:
                    _tp = int((_travel_t / _total_all) * 100)
                    _travel_jokes = [
                        f"Your car/bike thinks it's a money-shredding machine. Petrol companies send you thank-you cards.",
                        f"You've basically funded an airline pilot's vacation this month.",
                        f"Uber/Ola drivers consider you their best friend. You're basically their salary.",
                        f"Every toll booth is a mini ATM and you're the one depositing.",
                    ]
                    _travel_msg = random.choice(_travel_jokes)
                    _badge = "badge-danger" if _tp > 20 else "badge-warn" if _tp > 10 else "badge-success"
                    _story_items.append(("🚗", f"<b>ROAD WARRIOR:</b> <span class='story-amount story-highlight'>₹{_travel_t:,.0f}</span> ({_tp}%) went to fuel, tolls & rides. <span class='story-badge {_badge}'>{_tp}% SHARE</span> {_travel_msg}"))

                # Gadget spree
                if _gadget_t > 0:
                    _gp = int((_gadget_t / _total_all) * 100)
                    _gadget_jokes = [
                        f"Amazon/Flipkart's delivery person has you on speed dial. You're basically their VIP customer.",
                        f"Your house is slowly turning into a gadget showroom. Next: an electronics store?",
                        f"You bought something shiny and new! Your old gadgets are filing for emotional damages.",
                        f"Jeff Bezos just whispered 'thank you' from Seattle.",
                    ]
                    _gadget_msg = random.choice(_gadget_jokes)
                    _story_items.append(("📦", f"<b>GADGET SPREE:</b> <span class='story-amount story-highlight'>₹{_gadget_t:,.0f}</span> ({_gp}%) dropped on shiny new things. <span class='story-badge badge-danger'>{_gp}% SHARE</span> {_gadget_msg}"))

                # Entertainment binge
                if _ent_t > 0:
                    _ep = int((_ent_t / _total_all) * 100)
                    _ent_jokes = [
                        f"Netflix, Prime, Disney+ — you're basically running a personal cinema empire now.",
                        f"Your couch must be exhausted from all the binge-watching sessions.",
                        f"Movie tickets or OTT subscriptions? Either way, Hollywood thanks you for your service.",
                        f"You've watched enough content this month to write a small review blog.",
                    ]
                    _ent_msg = random.choice(_ent_jokes)
                    _badge = "badge-warn" if _ep > 15 else "badge-info"
                    _story_items.append(("🎬", f"<b>BINGE REPORT:</b> <span class='story-amount'>₹{_ent_t:,.0f}</span> ({_ep}%) on entertainment. <span class='story-badge {_badge}'>{_ep}% SHARE</span> {_ent_msg}"))

                # Shopping haul
                if _shop_t > 0:
                    _spct = int((_shop_t / _total_all) * 100)
                    _shop_jokes = [
                        f"Your wardrobe is screaming for mercy. Myntra/Ajio are your new best friends.",
                        f"Retail therapy is real and you're the champion patient. Shopping Olympics gold medalist!",
                        f"You don't have a shopping problem, you have a 'treat yourself' lifestyle.",
                        f"The shopping cart said 'add to cart' and you said 'add ALL the things.'",
                    ]
                    _shop_msg = random.choice(_shop_jokes)
                    _badge = "badge-danger" if _spct > 20 else "badge-warn"
                    _story_items.append(("🛍️", f"<b>RETAIL THERAPY:</b> <span class='story-amount story-highlight'>₹{_shop_t:,.0f}</span> ({_spct}%) on shopping. <span class='story-badge {_badge}'>{_spct}% SHARE</span> {_shop_msg}"))

                # Health is wealth
                if _health_t > 0:
                    _hp = int((_health_t / _total_all) * 100)
                    _story_items.append(("🏥", f"<b>HEALTH INVESTMENT:</b> <span class='story-amount'>₹{_health_t:,.0f}</span> ({_hp}%) on healthcare. <span class='story-badge badge-success'>WELL SPENT</span> Your body thanks you, your wallet... has mixed feelings."))

                # Bills reality check
                if _bills_t > 0:
                    _bp = int((_bills_t / _total_all) * 100)
                    _story_items.append(("📱", f"<b>BILLS CHECK:</b> <span class='story-amount'>₹{_bills_t:,.0f}</span> ({_bp}%) on utilities & bills. The boring stuff that keeps the lights on and WiFi flowing."))

                # Grand stats
                _story_items.append(("📊", f"<b>BY THE NUMBERS:</b> <span class='story-amount'>₹{_avg:,.0f}</span> avg per transaction | <b>{_txn_count}</b> total hits on your wallet | Single biggest hit: <span class='story-highlight'>₹{_max_txn:,.0f}</span>"))

                # Savings verdict
                if _total_recv > 0 and _total_spent > 0:
                    _ratio = _total_recv / _total_spent
                    if _ratio > 1.5:
                        _story_items.append(("🏆", f"<b>SAVINGS CHAMPION:</b> Income is <span class='story-good'>{_ratio:.1f}x</span> spending! You're basically a money-saving superhero. Your future self is already thanking you."))
                    elif _ratio > 1.2:
                        _story_items.append(("✅", f"<b>DOING WELL:</b> Income is <span class='story-good'>{_ratio:.1f}x</span> spending. You're building wealth while still enjoying life. Balance achieved!"))
                    elif _ratio > 0.9:
                        _story_items.append(("⚖️", f"<b>LIVING ON THE EDGE:</b> Income is just <b>{_ratio:.1f}x</b> spending. You're treading water — one big purchase away from either savings or panic."))
                    else:
                        _story_items.append(("⚠️", f"<b>RED ALERT:</b> Spending is <span class='story-highlight'>{(1/_ratio):.1f}x your income!</span> Your wallet is sending SOS signals. Time for a financial intervention!"))
                elif _total_spent > 0 and _total_recv == 0:
                    _story_items.append(("🚨", f"<b>NO INCOME DETECTED:</b> Pure spending mode activated. You spent <span class='story-amount story-highlight'>₹{_total_spent:,.0f}</span> with zero credits. Hope you have savings!"))

                # Closing dramatic line
                _closing_jokes = [
                    f"So there you have it — {_month_label}: the month your money went on an epic adventure without you. See you next month!",
                    f"TL;DR: Your wallet is tired, your bank statement is dramatic, and next month needs to be better. Or... does it? 😏",
                    f"And that's the story of {_month_label} — where every rupee had a purpose (even if that purpose was pizza at 2 AM).",
                    f"End of {_month_label} report. Your money has left the building. Mic drop. 🎤",
                ]
                _story_items.append(("🎭", f"<i>{random.choice(_closing_jokes)}</i>"))

                _story_html = f'<div class="story-card"><h4>📖 Your {_month_label} Money Story — The Dramatic Edition</h4>'
                for _icon, _text in _story_items:
                    _story_html += f'<div class="story-item"><span class="story-icon">{_icon}</span><span>{_text}</span></div>'
                _story_html += '</div>'

                # ── Value-Added Purchases Box ────────────────────────────
                # Strict keyword matching — must be actual product/brand names, NOT generic words
                _val_brands = (
                    "Samsung|Apple|iPhone|iPad|MacBook|Galaxy|Pixel|OnePlus|Nothing Phone|"
                    "Sony|LG|Whirlpool|IFB|Bosch|Voltas|Daikin|Blue Star|Crompton|Havells|"
                    "Philips|Panasonic|Midea|Carrier|Hitachi|Orient|Bajaj Electricals|"
                    "Croma|Reliance Digital|Vijay Sales|"
                    "Dyson|Nutribullet|KitchenAid|Prestige|Preethi|Havells|Borosil|"
                    "IKEA|Godrej|Sleepwell|Wakefit|"
                    "GoPro|DSLR|Canon|Nikon|Fujifilm|"
                    "Titan|Fastrack|Fossil|Casio|G-Shock|"
                    "PS5|Xbox|Nintendo|ROG|ASUS|"
                    "Boat|JBL|Bose|Sennheiser|"
                    "Apple Watch|AirPods|Galaxy Watch|Galaxy Buds"
                )
                _val_products = (
                    "Laptop|MacBook|iPhone|iPad|Galaxy S|Galaxy Z|Galaxy A|Pixel|OnePlus|"
                    "Washing Machine|Refrigerator|Fridge|Air Conditioner|AC |Split AC|Window AC|"
                    "Microwave|Oven|Dishwasher|Dryer|"
                    "Television|TV |Smart TV|OLED|QLED|4K|"
                    "Monitor|Printer|Desktop|"
                    "Speaker|Soundbar|Headphone|Earbuds|"
                    "Camera|GoPro|DSLR|Mirrorless|Lens|"
                    "Sofa|Wardrobe|Bed Frame|Mattress|IKEA|"
                    "Mixer Grinder|Food Processor|Blender|Air Purifier|Water Purifier|Geyser|Inverter|Battery|"
                    "Gaming Console|PS5|Xbox Series|Nintendo Switch|"
                    "Smart Watch|Fitness Band|"
                    "Galaxy S26|Galaxy S25|Galaxy Z|iPhone 16|iPhone 15|"
                    "ROG|Alienware|ThinkPad|MacBook Air|MacBook Pro"
                )
                # exclusion: food/restaurant/travel/financial context kills the match
                _val_exclude = "Restaurant|Cafe|Food|Swiggy|Zomato|Hotel|Flight|Train|Bus|Petrol|Diesel|Fuel|Taxi|Uber|Ola|ATM|Salary|Insurance|Bill|Recharge|Subscription|Pharmacy|Hospital|Credit Card Bill|EMI|Loan|Mutual Fund|Stock|Investment"
                # minimum amount to qualify as value-add
                _VAL_MIN = 1500

                if len(_debits):
                    _match_brand = _debits["category"].str.contains(_val_brands, case=False, na=False) | _debits["merchant"].str.contains(_val_brands, case=False, na=False)
                    _match_product = _debits["category"].str.contains(_val_products, case=False, na=False) | _debits["merchant"].str.contains(_val_products, case=False, na=False)
                    # check raw_text directly from _debits (already has it from _spdf)
                    _debits_raw = _debits["raw_text"].fillna("").str.lower() if "raw_text" in _debits.columns else pd.Series([""]*len(_debits), index=_debits.index)
                    _raw_brand = _debits_raw.str.contains(_val_brands, case=False, na=False)
                    _raw_product = _debits_raw.str.contains(_val_products, case=False, na=False)
                    _match_all = (_match_brand | _match_product | _raw_brand | _raw_product)
                    # exclude food/travel/financial context
                    _excl = _debits["category"].str.contains(_val_exclude, case=False, na=False) | _debits["merchant"].str.contains(_val_exclude, case=False, na=False)
                    # minimum amount filter
                    _min_amt = _debits["amount"] >= _VAL_MIN
                    _val_matches = _debits[_match_all & ~_excl & _min_amt].copy()
                else:
                    _val_matches = pd.DataFrame()

                _val_html = ""
                if not _val_matches.empty:
                    _val_total = _val_matches["amount"].sum()
                    _val_count = len(_val_matches)
                    _val_pct = int((_val_total / _total_all) * 100) if _total_all else 0

                    def _get_raw(row):
                        return str(row.get("raw_text", ""))

                    # tag each purchase
                    def _tag_item(row):
                        t = str(row.get("category", "") + " " + row.get("merchant", "") + " " + _get_raw(row)).lower()
                        if any(k.lower() in t for k in ["samsung", "apple", "iphone", "ipad", "galaxy", "pixel", "oneplus", "nothing phone", "mobile", "phone", "tablet"]):
                            return ("📱", "Smart Device", "tag-phone")
                        if any(k.lower() in t for k in ["laptop", "macbook", "monitor", "printer", "computer", "desktop"]):
                            return ("💻", "Computing", "tag-gadget")
                        if any(k.lower() in t for k in ["tv", "television", "speaker", "headphone", "sound", "sony", "bose", "jbl"]):
                            return ("📺", "Entertainment Tech", "tag-gadget")
                        if any(k.lower() in t for k in ["washing", "refrigerator", "fridge", "ac ", "air conditioner", "microwave", "oven", "dishwasher", "dryer"]):
                            return ("🏠", "Home Appliance", "tag-appliance")
                        if any(k.lower() in t for k in ["ac", "daikin", "voltas", "blue star", "carrier", "hitachi", "crompton", "inverter", "battery", "geyser", "purifier"]):
                            return ("❄️", "Climate & Power", "tag-appliance")
                        if any(k.lower() in t for k in ["furniture", "sofa", "bed", "wardrobe", "table", "chair", "ikea", "wood"]):
                            return ("🪑", "Furniture", "tag-furniture")
                        if any(k.lower() in t for k in ["camera", "gopro", "dslr", "lens"]):
                            return ("📷", "Camera", "tag-gadget")
                        if any(k.lower() in t for k in ["watch", "titan", "fastrack", "fossil", "g-shock"]):
                            return ("⌚", "Watch", "tag-luxury")
                        if any(k.lower() in t for k in ["ps5", "xbox", "nintendo", "rog", "gaming", "controller"]):
                            return ("🎮", "Gaming", "tag-gadget")
                        if any(k.lower() in t for k in ["kitchen", "mixer", "grinder", "blender", "prestige", "preethi", "bajaj"]):
                            return ("🍳", "Kitchen", "tag-home")
                        if any(k.lower() in t for k in ["croma", "reliance digital", "vijay sales"]):
                            return ("🏪", "Electronics Store", "tag-other")
                        return ("📦", "Purchase", "tag-other")

                    _val_stories = [
                        "A wise person once said 'treat yourself' — and you took that advice VERY seriously.",
                        "This one's going to sit in your life for years. An investment in daily happiness!",
                        "Your future self is already high-fiving your present self for this buy.",
                        "The unboxing video writes itself. Pure dopamine in a box.",
                        "This isn't spending, this is UPGRADING YOUR ENTIRE LIFESTYLE.",
                        "You didn't need it. But you DESERVED it. And that's what matters.",
                        "Quality of life: +100. Bank account: -this amount. Worth it? Absolutely.",
                        "Something tells us this won't sit in the box for long. First use incoming!",
                        "The old one was 'fine'. But this? This is NEXT LEVEL.",
                        "Science says experiences > things. But this thing ENABLES better experiences.",
                        "Peak adulting moment: getting excited about a new appliance/gadget.",
                        "This purchase has 'I'm an adult now' energy written all over it.",
                        "Not an impulse buy — this was a STRATEGIC LIFE UPGRADE.",
                        "Your home just got smarter/cooler/more comfortable. Mission accomplished.",
                        "Return on investment: measured in daily smiles, not percentages.",
                    ]

                    _val_html = (
                        f'<div class="valuebox">'
                        f'<h4>✨ Your {_month_label} Value-Upgrades — Things That Level Up Your Life</h4>'
                        f'<div class="v-subtitle">₹{_val_total:,.0f} invested across <b>{_val_count}</b> upgrades ({_val_pct}% of total spend) — every rupee working to make your daily life better</div>'
                    )

                    for _, vr in _val_matches.iterrows():
                        _emoji, _tag_name, _tag_cls = _tag_item(vr)
                        _amt = vr.get("amount", 0)
                        _merchant = vr.get("merchant", "Unknown")
                        _raw = _get_raw(vr)
                        _story = random.choice(_val_stories)

                        _product_hint = ""
                        _raw_lower = _raw.lower()
                        if "samsung" in _raw_lower and ("galaxy" in _raw_lower or "s " in _raw_lower):
                            _product_hint = " — a Samsung Galaxy upgrade! Welcome to the flagship club."
                        elif "apple" in _raw_lower or "iphone" in _raw_lower or "ipad" in _raw_lower:
                            _product_hint = " — the Apple ecosystem just got another member."
                        elif "macbook" in _raw_lower or "laptop" in _raw_lower:
                            _product_hint = " — new laptop who dis? Productivity just went through the roof."
                        elif "ps5" in _raw_lower or "xbox" in _raw_lower or "nintendo" in _raw_lower:
                            _product_hint = " — gaming just got a MASSIVE upgrade. Game on!"
                        elif "tv" in _raw_lower or "television" in _raw_lower:
                            _product_hint = " — movie nights just became cinematic. Popcorn not included."
                        elif "ac" in _raw_lower or "air conditioner" in _raw_lower or "daikin" in _raw_lower or "voltas" in _raw_lower:
                            _product_hint = " — summer is no longer your enemy. Cool breeze incoming!"
                        elif "washing machine" in _raw_lower or "refrigerator" in _raw_lower or "fridge" in _raw_lower:
                            _product_hint = " — household chores just got a serious upgrade."

                        _val_html += (
                            f'<div class="vcard">'
                            f'<div class="vcard-header">'
                            f'<span class="vcard-emoji">{_emoji}</span>'
                            f'<span class="vcard-title">{_merchant}</span>'
                            f'<span class="vcard-amount">₹{_amt:,.0f}</span>'
                            f'</div>'
                            f'<div class="vcard-meta">{vr.get("date", "N/A")} | {vr.get("category", "N/A")}</div>'
                            f'<div class="vcard-story">{_story}{_product_hint}</div>'
                            f'<span class="vcard-tag {_tag_cls}">{_tag_name}</span>'
                            f'</div>'
                        )

                    _val_close = [
                        f"That's {_val_count} upgrade(s) this month. Your lifestyle game is STRONG. 💪",
                        f"₹{_val_total:,.0f} well spent on things that matter every single day. No regrets here!",
                        f"You're not just spending, you're CURATING A BETTER LIFE. Respect. 🙌",
                    ]
                    _val_html += f'<div class="story-item" style="margin-top:0.8rem;"><span class="story-icon">🏆</span><span><b>UPGRADE VERDICT:</b> {random.choice(_val_close)}</span></div>'
                    _val_html += '</div>'

                # ── Render side by side ──
                _left, _right = st.columns(2)
                with _left:
                    st.markdown(_story_html, unsafe_allow_html=True)
                with _right:
                    if _val_html:
                        st.markdown(_val_html, unsafe_allow_html=True)
                    else:
                        st.markdown(
                            '<div class="valuebox" style="opacity:0.7;">'
                            '<h4>✨ Value-Upgrades</h4>'
                            '<div class="v-subtitle">No electronics, gadgets, appliances, or lifestyle upgrades detected this month yet.</div>'
                            '</div>',
                            unsafe_allow_html=True,
                        )

    # --- Chatbot ---
    with tabs[1]:
        st.header("Expense Assistant Chatbot")
        if "history" not in st.session_state:
            st.session_state.history = []
        if "chat_processing" not in st.session_state:
            st.session_state.chat_processing = False

        # Form ALWAYS at top — Enter-to-submit, disabled during processing
        with st.form("chat_form", clear_on_submit=False):
            user_input = st.text_input(
                "Ask about your transactions",
                placeholder="e.g. 'get all ATM transactions', 'total spent on food', 'monthly summary'",
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

        # Chat history BELOW input — always scrollable
        if st.session_state.history:
            st.markdown("---")
            for msg in reversed(st.session_state.history):
                chat_bubble("You", msg['user'])
                chat_bubble("Assistant", msg['assistant'])

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
