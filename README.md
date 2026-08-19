# Daily Expense Analyser (Ollama + Streamlit)

Comprehensive financial transaction analyzer powered by Ollama LLM.

**Features:**
- Parse SMS/bank transaction texts into structured data (file upload or paste)
- Multi-currency support with auto-detection from SMS text
- Store transactions locally in Excel (`records.xlsx`) with deduplication
- Multi-endpoint support (Local Ollama + Ollama Cloud with automatic fallback)
- Intelligent transaction categorization using 100+ predefined categories
- Comprehensive analytics dashboard with per-currency breakdowns
- LLM-powered financial insights and recommendations
- Chat interface for expense queries

## Multi-Currency Support

The app auto-detects currency from SMS text and supports worldwide banks:

| Symbol | Code | Banks |
|--------|------|-------|
| ₹ / Rs. | INR | HDFC, ICICI, SBI, Axis, Kotak, all Indian banks |
| £ | GBP | Lloyds, Barclays, HSBC UK, NatWest, Halifax |
| $ | USD | Chase, Bank of America, Wells Fargo, Citibank |
| € | EUR | Deutsche Bank, BNP Paribas, ING |
| AED / د.إ | AED | Emirates NBD, ADCB, Mashreq |
| SAR | SAR | Al Rajhi, SNB, Riyad Bank |
| S$ | SGD | DBS, OCBC, UOB |
| A$ | AUD | CommBank, Westpac, ANZ, NAB |
| C$ | CAD | RBC, TD Canada, Scotiabank |
| ¥ | JPY | MUFG, SMBC, Mizuho |

Mixed-currency files are fully supported — each SMS is parsed independently and its currency is auto-detected.

**Analytics across currencies:**
- Sidebar **Quick Analytics** shows per-currency summary cards, top categories, top merchants, monthly spending, and spending story
- **Analysis page** has a Currency Overview with per-currency spent/received cards and a currency filter dropdown
- **Spending Story** generates a separate dramatic narrative per currency
- **Life Upgrades** (significant purchases) displays alongside the spending story

## Two Ways to Parse

### "Parse File" (File Upload)
1. Upload a `.txt` file containing bank SMS messages
2. Regex-based parser (`parse_transactions.py`) auto-detects bank, currency, and transaction type
3. Parsed transactions are appended to `records.xlsx` (deduplicated)
4. Supports 80+ banks worldwide with 100+ SMS patterns

### "Parse and Save" (Text Paste)
1. Paste SMS text into the text area
2. LLM extracts structured JSON (date, amount, merchant, category, currency)
3. Results normalized via `normalize_record()` with currency detection from raw text
4. Saved to `records.xlsx` with full audit trail

## Analytics Dashboard

- **Quick Analytics** (sidebar): Real-time per-currency stats — total spent/received, net, transaction count, top categories, top merchants, monthly spending chart, spending story
- **Spending Story** (Analysis tab): Per-currency dramatic narrative with category breakdowns, biggest transactions, savings verdict
- **Life Upgrades**: Electronics/gadget and travel booking tracker with emoji-tagged cards
- **Currency Overview**: Per-currency spent/received summary cards
- **Filters**: Currency selector, time period (daily/monthly/yearly), debit/credit toggle
- **Charts**: Spending by category, merchant breakdown, transaction trends over time
- **AI Insights**: LLM-powered financial analysis and recommendations

## Configuration

### `.env`

```env
# Default currency (used when SMS doesn't specify one)
DEFAULT_CURRENCY=INR

# Ollama local endpoint
OLLAMA_API=http://localhost:11434
OLLAMA_LOCAL_MODEL=llama3.2:latest

# Ollama cloud fallback
OLLAMA_CLOUD_BASE_URL=https://ollama.com
OLLAMA_CLOUD_MODEL=gpt-oss:120b-cloud
OLLAMA_CLOUD_API_KEY=<your-cloud-api-key>
```

The app tries local Ollama first; if it fails, it automatically falls back to the cloud endpoint.

## Quick Start

1. Create a Python environment and install dependencies:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

2. Configure `.env` with your Ollama endpoints (see above).

3. Run the Streamlit app:

```bash
streamlit run streamlit_app.py
```

## File Structure

```
DailyExpenseAnalyser/
  streamlit_app.py       ← Main Streamlit app (GUI, analytics, LLM integration)
  parse_transactions.py  ← Multi-bank, multi-currency SMS parser
  records.xlsx           ← Transaction data (auto-generated, two sheets: Transactions + Untracked)
  .env                   ← Configuration (Ollama endpoints, default currency)
  requirements.txt       ← Python dependencies
  agents.md              ← Parser documentation (bank patterns, date formats, output schema)
```

## Notes

- `records.xlsx` has two sheets: **Transactions** (tracked) and **Untracked** (non-transaction SMS)
- The app auto-detects currency from SMS symbols/text when the LLM omits it
- Deduplication is based on `(date, amount, merchant, category)` — duplicate imports are skipped
- Amount column is stored as `float64` for consistent numeric analysis
- The "Parse File" path uses a deterministic regex parser (no LLM needed)
- The "Parse and Save" path uses Ollama LLM for extraction (requires Ollama running or cloud key)
