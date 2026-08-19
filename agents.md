# Agents: Multi-Bank SMS Transaction Parser

## Purpose
Parse bank transaction SMS from **any bank worldwide** into structured data.
Supports India (HDFC, ICICI, SBI, etc.), UK (Lloyds, Barclays, HSBC, etc.),
US (Chase, Wells Fargo, etc.), and more.

Both Streamlit buttons use this parser:
- **"Parse File"** — uploads a .txt file, calls `parse_sms_text()`, displays + exports
- **"Parse and Save"** — pastes SMS text, calls LLM for extraction, saves to records.xlsx

---

## 1. Supported Currencies

The parser auto-detects currency from SMS text. Default is configurable via `.env`:

```
DEFAULT_CURRENCY=INR      # India (default)
DEFAULT_CURRENCY=GBP      # United Kingdom
DEFAULT_CURRENCY=USD      # United States
DEFAULT_CURRENCY=EUR      # Eurozone
DEFAULT_CURRENCY=AED      # UAE
```

| Symbol | Code | Banks |
|--------|------|-------|
| ₹ / Rs. / INR | INR | HDFC, ICICI, SBI, Axis, Kotak, all Indian banks |
| £ / GBP | GBP | Lloyds, Barclays, HSBC UK, NatWest, Halifax, Nationwide |
| $ / USD | USD | Chase, Bank of America, Wells Fargo, Citibank |
| € / EUR | EUR | Deutsche Bank, BNP Paribas, ING |
| AED / د.إ | AED | Emirates NBD, ADCB, Mashreq |
| SAR | SAR | Al Rajhi, SNB, Riyad Bank |
| S$ / SGD | SGD | DBS, OCBC, UOB |
| A$ / AUD | AUD | CommBank, Westpac, ANZ, NAB |
| C$ / CAD | CAD | RBC, TD Canada, Scotiabank, BMO |
| ¥ / JPY | JPY | MUFG, SMBC, Mizuho |
| ¥ / CNY | CNY | ICBC, CCB, BoC |

---

## 2. SMS Block Splitting Rules

Bank SMS files have NO blank lines between entries. Blocks are split by **date-header lines**:

```
Day, DD Month YYYY · HH:MM       ← with year (2025)
Day, DD Month · HH:MM             ← without year (inherits previous year)
Monday · 14:16                    ← day-of-week only, no date
```

If no header exists (e.g. standalone Lloyds SMS), treat each line as a separate block.

Standalone timestamp-only lines (`18:12`, `13:22`) belong to the previous block.

---

## 3. Transaction Patterns by Region

### 3a. Indian Banks

#### Debit Transactions

| Pattern | Bank | Category |
|---------|------|----------|
| `Txn Rs.XX On Bank Card NNNN At MERCHANT by UPI REF On DD-MM` | HDFC CC | CC UPI |
| `Spent Rs.XX On Bank Card NNNN At MERCHANT On YYYY-MM-DD:HH:MM:SS` | HDFC CC | CC Spent |
| `INR XX.XX spent using Bank Card XXXX on DD-MON-YY on MERCHANT` | ICICI CC | CC Spent |
| `IMPS INR XX,XXX sent from Bank A/c XX1542 on DD-MM-YY To A/c XX Ref-XXXX` | HDFC | IMPS |
| `Sent Rs.XX From Bank A/C *XXXX To MERCHANT On DD/MM/YY Ref XXXX` | HDFC | UPI Sent |
| `Your a/c XX7503 is debited for Rs. XX on DD-MM-YYYY HH:MM and credited to a/c XX (UPI Ref no XXXX)` | SC, SBI, All | UPI Debit |
| `UPDATE: INR XX debited from Bank XX1542 on DD-MMM-YY. Info: ...` | HDFC | Account Debit |
| `Payment Successful! Rs. XX from A/c to XXXX` | HDFC | NetBanking |
| `HDFC Bank:Rs. XX debited from a/c *XXXX on DD/MM/YY to a/c **XXXX (UPI Ref No. XXXX)` | HDFC | UPI Debit |
| `Paid Rs. XX For: Credit Card payment From Bank A/c XX` | HDFC | Bill Payment |
| `Mandate Set Rs.XX For PERSON From Bank A/c xXXXX` | HDFC | Mandate |
| `FD XX of INR XX will be liquidated for INR XX` | HDFC | FD Liquidation |

#### Credit Transactions

| Pattern | Bank | Category |
|---------|------|----------|
| `Bank Cardmember, Online Payment of Rs.XX vide Ref# XXXX was credited to your card ending NNNN On DD/MON/YYYY` | HDFC CC | CC Payment |
| `Payment of Rs XX was credited to your card ending NNNN On DD/MON/YYYY` | HDFC CC | CC Payment |
| `ICICI Bank Account XX258 credited:Rs. XX on DD-MON-YY. Info ACH*...` | ICICI | Credit |
| `ICICI Bank Account XX174 is credited with Rs XX on DD-MON-YY by ... IMPS Ref. no. XXXX` | ICICI | IMPS Credit |
| `Dear SBI User, your A/c X7503-credited by Rs.XX on DDMMonYY transfer from PERSON Ref No XXXX -SBI` | SBI | Credit |
| `Update! INR XX deposited in Bank A/c XX on DD-MMM-YY for SWEEP-IN CREDIT - XXXX` | HDFC | FD Sweep-In |

### 3b. UK Banks (Lloyds, Barclays, HSBC, NatWest, etc.)

#### Debit Transactions

| Pattern | Bank | Category |
|---------|------|----------|
| `£125.00 spent at Tesco on 19/08/26 14:05` | Lloyds, any UK | CC Spent |
| `£125.00 debit at Tesco on 19/08/26 14:05` | Any UK | Debit Card |
| `£50.00 payment to Landlord on 01/09/26 09:00` | Any UK | Payment |
| `£200.00 transfer to John on 15/08/26 12:30` | Any UK | Transfer |
| `£100.00 withdrawal at ATM on 19/08/26 14:05` | Any UK | ATM Withdrawal |

#### Credit Transactions

| Pattern | Bank | Category |
|---------|------|----------|
| `£1500.00 credit from employer on 25/08/26 06:00` | Any UK | Credit |
| `£200.00 transfer from John on 15/08/26 12:30` | Any UK | Transfer |

#### Non-Transaction (SKIP)

| Pattern | Reason |
|---------|--------|
| `Avail Bal: £2,340.55` | Balance inquiry |
| `If not you, call 0345 300 0000` | Fraud alert footer |
| `Your statement is ready` | Statement notification |
| `Your card has been blocked` | Security notification |
| `New direct debit set up` | Mandate notification |
| `Direct debit for X cancelled` | Mandate cancellation |
| `Standing order has been updated` | Standing order change |
| `Your overdlimit has been changed` | Overdraft notification |

### 3c. US Banks (Chase, BofA, Wells Fargo, etc.)

| Pattern | Bank | Category |
|---------|------|----------|
| `$125.00 debit card purchase at Walmart on 08/19/26` | Any US | CC Spent |
| `$50.00 ATM withdrawal on 08/19/26` | Any US | ATM Withdrawal |
| `$1500.00 direct deposit from EMPLOYER on 08/25/26` | Any US | Credit |
| `$200.00 Zelle transfer to John on 08/15/26` | Any US | Transfer |

### 3d. Generic Multi-Currency Patterns

The parser handles ANY bank that uses these common formats:
- `£/$/€/₹ + amount + action + target + on + date`
- `amount + spent at/debit/credit/payment/transfer + target + on + date`

---

## 4. Bank Detection

The parser detects bank names from SMS text. If not found, marks as "Unknown Bank".

**Detection priority:**
1. Explicit bank name in text (e.g., "HDFC Bank", "Lloyds Bank")
2. Phone number matching (e.g., `18002586465` → Standard Chartered)
3. Card/account format patterns

**Supported banks (80+):**
- **India:** HDFC, ICICI, SBI, Axis, Kotak, PNB, BoB, Canara, Union, SC, IDBI, Federal, IndusInd, Yes, RBL, Bandhan, Indian, IOB, BOI, Central, UCO, City Union, DCB, Karur Vysya, South Indian, CSB, Karnataka, TMB, AU SFB, Equitas, Ujjivan, Fincare, J&K Bank
- **UK:** Lloyds, Barclays, HSBC, NatWest, Santander, Halifax, Nationwide, TSB, Virgin Money, First Direct, Monzo, Revolut, Starling, RBS, Clydesdale, Tesco Bank, Sainsburys Bank
- **US:** Chase, Bank of America, Wells Fargo, Citibank, Capital One, US Bank, PNC, TD Bank, Truist, Discover, American Express
- **Australia:** CommBank, Westpac, ANZ, NAB
- **Canada:** RBC, TD Canada, Scotiabank, BMO
- **Global:** Deutsche Bank

---

## 5. Date Formats

| Format | Example | Region |
|--------|---------|--------|
| `DD-MM-YYYY HH:MM` | `24-07-2026 09:32` | India |
| `DD-MON-YY` | `07-JUN-25` | India |
| `DD/MM/YYYY HH:MM` | `19/08/2026 14:05` | UK |
| `DD/MM/YY HH:MM` | `19/08/26 14:05` | UK |
| `MM/DD/YY HH:MM` | `08/19/26 14:05` | US |
| `YYYY-MM-DD:HH:MM:SS` | `2025-07-12:18:49:48` | India (ISO) |
| `Day, DD Month YYYY · HH:MM` | `Friday, 9 May 2025 · 19:14` | Header |
| `Day, DD Month · HH:MM` | `Thursday, 1 Jan · 23:05` | Header (no year) |
| `DDMonYY` | `06Jun26` | India (compact) |
| `DD-MM-YY` | `11-07-25` | India |

**Year inheritance:** If header has no year, inherit from previous block. If month sequence goes backward (Dec → Jan), increment year.

---

## 6. Amount Extraction

- Remove commas: `1,02,345.00` → `102345.00`
- Handle currency prefixes: `Rs.50`, `£125.00`, `$50.00`, `€50.00`, `INR 80,000`
- Indian number format: `1,02,345.00` (lakhs) → `102345.00`
- Western number format: `1,234.56` → `1234.56`

---

## 7. Output Schema

```json
{
  "date": "YYYY-MM-DD HH:MM",
  "txn_type": "Debit | Credit",
  "amount": 125.00,
  "currency": "GBP",
  "category": "CC Spent | Debit Card | Payment | Transfer | ATM Withdrawal | ...",
  "merchant": "Tesco",
  "bank": "Lloyds Bank",
  "card_or_account": "1234",
  "ref": "549521241794",
  "detail": "From XX1542 | EMI info | etc."
}
```

---

## 8. Non-Transaction Messages (SKIP)

### India
- Balance inquiries: `Available Bal in Bank A/c XX`
- Welcome kits: `WELCOME KIT DELIVERED`
- Form submissions: `Form Submitted!`
- Amount due notices: `Amount Due Rs.XX`
- UPI registration: `UPI Registration!`
- Card delivery: `Out for delivery`
- FD updates: `Your FD No XX maturity instruction`
- UPI requests: `has requested RsXX frm u`
- Loan offers: `Thank you for considering Bank Personal Loan`

### UK
- Balance inquiries: `Avail Bal: £X,XXX.XX`
- Statement notifications: `Your statement is ready`
- Card security: `Your card has been blocked`
- Direct debit changes: `New direct debit set up`
- Standing order changes: `Standing order has been updated`
- Overdraft changes: `Your overdraft limit has been changed`

### US
- Balance alerts: `Your available balance is $X,XXX.XX`
- Security alerts: `Unusual activity detected`
- Statement ready: `Your monthly statement is available`

---

## 9. Streamlit GUI Integration

### Button: "Parse File" (File Upload)
1. User uploads `.txt` file via `st.file_uploader`
2. File text decoded as UTF-8
3. Calls `parse_sms_text(file_text)` from `parse_transactions.py`
4. Displays parsed transactions in `st.dataframe`
5. Saves to temporary Excel + merge into `records.xlsx`
6. Provides download button for Excel file

### Button: "Parse and Save" (Text Area)
1. User pastes SMS text into text area
2. `split_bulk_sms()` splits into individual SMS
3. Each SMS sent to LLM (`call_ollama()`) for JSON extraction
4. Results normalized via `normalize_record()` and saved to `records.xlsx`
5. Uses the 200+ predefined categories from `streamlit_app.py`

### Analytics Tab
- Works with merged data from both buttons
- All charts, stories, and AI insights apply

---

## 10. File Structure

```
DailyExpenseAnalyser/
  agents.md              ← This file (master prompt)
  parse_transactions.py  ← Multi-bank, multi-currency parser
  streamlit_app.py       ← GUI with file upload + paste + analytics
  dailytran.txt          ← Sample HDFC SMS data
  records.xlsx           ← Output Excel (auto-generated)
  transactions.xlsx      ← Parser output Excel
  .env                   ← Config (DEFAULT_CURRENCY, Ollama endpoints)
```

---

## 11. Configuration (.env)

```env
# Default currency for SMS that don't specify one
DEFAULT_CURRENCY=INR

# Ollama endpoints (for "Parse and Save" button)
OLLAMA_API=http://localhost:11434
OLLAMA_LOCAL_MODEL=llama3.2:latest
OLLAMA_CLOUD_BASE_URL=https://ollama.com
OLLAMA_CLOUD_MODEL=gpt-oss:120b-cloud
OLLAMA_CLOUD_API_KEY=your-key-here
```

---

## 12. Edge Cases

1. **Multiple Txn in one block:** Split on transaction-start patterns (`Txn Rs.`, `spent at`, etc.)
2. **Standalone lines without header:** Parse individually with empty header_date
3. **Mixed currencies in one file:** Auto-detect per SMS (e.g., ₹ for HDFC, £ for Lloyds)
4. **Mixed banks in one file:** Auto-detect bank per SMS
5. **Truncated merchant names:** Use as-is (e.g., `vyapar.173291190979@hdfcbank`)
6. **Zero-amount entries:** Reward points, mandates → amount=0 is valid
7. **Date-only entries without time:** Use `00:00` as default time
8. **Card number masking:** `XX1542`, `8514` → use as-is in output
9. **Year-less headers:** Inherit year, bump on month wrap (Dec→Jan)
10. **Non-English merchant names:** Preserve original text
