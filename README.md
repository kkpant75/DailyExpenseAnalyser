# Daily Expense Analyser (Ollama + Streamlit)

Comprehensive financial transaction analyzer powered by Ollama LLM. 

**Features:**
- Parse SMS/bank transaction texts into structured data
- Bulk SMS processing (split and parse multiple transactions at once)
- Store transactions locally in Excel with full audit trail
- Multi-endpoint support (Local Ollama + Ollama Cloud)
- Intelligent transaction categorization using 100+ predefined categories
- Comprehensive analytics dashboard with:
  - Daily/Monthly/Yearly filtering
  - Debit vs Credit breakdown
  - Category-wise spending analysis
  - Merchant/vendor spending breakdown
  - Transaction trends over time
- LLM-powered financial insights and recommendations
- Chat interface for expense queries

**Supported Transaction Types:**
- Bank transfers (NEFT, RTGS, IMPS)
- UPI & wallet payments
- E-commerce (Amazon, Flipkart, etc.)
- Food & restaurants (Swiggy, Zomato, etc.)
- Fuel & Transportation (Petrol, FASTag, Uber, etc.)
- Shopping & retail
- Entertainment & subscriptions
- Utilities & bills
- Healthcare & education
- Investments & financial services
- ...and 50+ more categories

Quick start

1. Create a Python environment and install dependencies:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

2. Ollama configuration

- **Preferred: Local Ollama** with `llama3.2:latest` (or other model). Set `OLLAMA_LOCAL_MODEL` in `.env`:

```
OLLAMA_API=http://localhost:11434
OLLAMA_LOCAL_MODEL=llama3.2:latest
```

- **Fallback: Ollama Cloud** (if local is not available). Set the following in `.env`:

```
OLLAMA_CLOUD_BASE_URL=https://ollama.com
OLLAMA_CLOUD_MODEL=gpt-oss:120b-cloud
OLLAMA_CLOUD_API_KEY=<your-cloud-api-key>
```

The app prefers the local endpoint when `OLLAMA_API` is set; otherwise it uses the cloud config.

3. Run the Streamlit app:

```bash
streamlit run streamlit_app.py
```

Notes
- The app expects an Ollama-compatible HTTP API at `$OLLAMA_API/api/generate` that accepts a JSON body with `model` and `prompt` and returns JSON containing text output. If your Ollama setup exposes a different endpoint, update `OLLAMA_API` accordingly in `.env`.
- Parsed records are stored in `records.xlsx` in the app folder.
