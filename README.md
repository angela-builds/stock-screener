# 📊 Automated Stock Screening Tool

A Python automation tool that pulls financial data for any US-listed stock and scores it against a custom fundamental analysis framework — then writes the results directly to Google Sheets.

---

## 🔍 Why I Built This

I was manually collecting financial data from Morningstar for every stock I wanted to research. It was time-consuming and error-prone. This tool automates the entire data collection and scoring process with a single command.

---

## 📁 Two Versions Available

| | `stock.py` (FMP version) | `stock-yf.py` (yfinance version) |
|---|---|---|
| Data source | [Financial Modeling Prep API](https://financialmodelingprep.com/) | [yfinance](https://github.com/ranaroussi/yfinance) |
| API key required | ✅ Yes | ❌ No |
| Years of data | Up to 10 years | ~4 years |
| Cost | Free tier (250 calls/month) | Completely free |
| Best for | More historical depth | Quick setup, no signup |

---

## ⚙️ How It Works

1. Enter any US stock ticker (e.g. `AAPL`, `NVDA`)
2. The tool fetches financial data (income statement, balance sheet, cash flow)
3. It scores the stock across 13 financial indicators
4. Results are automatically written to a Google Sheet (overwrites existing row for the same ticker)

---

## 📋 Scoring System (Max 9.5 Points)

| Indicator | Criteria | Points |
|---|---|---|
| EPS 每股盈餘 | 10-year stable growth | +1 |
| Dividend 股息 | 10-year consistent | +0.5 |
| Dividend 股息 | Growing over time | +1 |
| Shares 流通股數 | 10-year declining | +1 |
| Book Value Per Share | 10-year stable growth | +1 |
| FCF 自由現金流 | All positive over 10 years | +1 |
| Net Margin 淨利率 | > 10% | +1 |
| Net Margin 淨利率 | 10-year stable (not declining) | +0.5 |
| ROE 股東權益率 | 15% < ROE < 40% | +1 |
| ROE 股東權益率 | < 15% but steadily rising | +0.5 |
| IC 利息保障倍數 | > 10 or no debt | +1 |
| IC 利息保障倍數 | > 5 | +0.5 |
| D/E 負債權益比 | < 0.5 | +1 |

**Stocks scoring 7 or above are shortlisted for deeper research.**

> ⚠️ **Note on yfinance version:** Due to yfinance data availability (~4 years), trend-based criteria are evaluated over a shorter window. The dividend stability threshold is also relaxed accordingly. Scores should be interpreted with this in mind.

---

## 🛠️ Tech Stack

- Python 3
- `gspread` — Google Sheets Python client
- `google-auth` — Service Account authentication
- **FMP version:** `requests` + [FMP API](https://financialmodelingprep.com/)
- **yfinance version:** [yfinance](https://github.com/ranaroussi/yfinance)

---

## 🚀 Setup

### 1. Clone this repo
```bash
git clone https://github.com/your-username/stock-screener.git
cd stock-screener
```

### 2. Install dependencies

**FMP version:**
```bash
pip install gspread google-auth requests
```

**yfinance version:**
```bash
pip install gspread google-auth yfinance
```

### 3. Set up Google Sheets API
- Go to [Google Cloud Console](https://console.cloud.google.com/)
- Create a project and enable **Google Sheets API** and **Google Drive API**
- Create a **Service Account** and download `credentials.json`
- Share your Google Sheet with the service account email (Editor access)
- Place `credentials.json` in the same folder as the script

### 4. Configure the script

**FMP version** — open `stock.py` and fill in:
```python
FMP_API_KEY = "your_fmp_api_key"
SHEET_NAME  = "your_google_sheet_name"
```

Sign up for a free FMP API key at [financialmodelingprep.com](https://financialmodelingprep.com/).

**yfinance version** — open `stock-yf.py` and fill in:
```python
SHEET_NAME = "your_google_sheet_name"
```

No API key needed.

### 5. Run
```bash
# FMP version
python stock.py

# yfinance version
python stock-yf.py
```

---

## 📁 Project Structure

```
stock-screener/
├── stock.py          # FMP version
├── stock-yf.py       # yfinance version (no API key required)
├── credentials.json  # Google Service Account key (not uploaded)
├── .gitignore        # Excludes credentials.json
└── README.md
```

---

## ⚠️ Important

- `credentials.json` is listed in `.gitignore` and will **not** be uploaded to GitHub
- Never share your `credentials.json` or FMP API key publicly
- FMP free tier allows ~250 API calls/month, sufficient for manual queries
- yfinance pulls data from Yahoo Finance; intended for personal research use only

---

## 📌 Limitations

- US-listed stocks only
- Trend analysis uses a 60% threshold (allows occasional dips)
- IC calculation uses EBITDA as a proxy for EBIT (FMP version)
- yfinance version limited to ~4 years of historical data

---

## ⚖️ Disclaimer

This tool is for personal investment research only and does not constitute financial advice.

---

*Built as a personal investment research automation project.*
