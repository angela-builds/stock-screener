import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ==============================
# 設定區
# ==============================
SHEET_NAME = "AI stock scorer"
CREDENTIALS_FILE = "credentials.json"  # 跟這個檔案放同一個資料夾

# ==============================
# Google Sheet 連線
# ==============================
def connect_sheet():
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open(SHEET_NAME).sheet1
    return sheet

# ==============================
# 初始化表頭（第一次執行時用）
# ==============================
def init_headers(sheet):
    headers = sheet.row_values(1)
    if not headers:
        sheet.update('A1:Q1', [[
            "股票代號", "現在股價",
            "EPS 10年穩定成長 +1",
            "Dividend 10年持續穩定 +0.5",
            "Dividend 越發越多 +1",
            "Shares 流通股數10年穩定減少 +1",
            "Book Value Per Share 10年穩定成長 +1",
            "FCF 10年皆為正數 +1",
            "Net Margin > 10% +1",
            "Net Margin 10年穩定不衰退 +0.5",
            "ROE 15%~40% +1",
            "ROE < 15% 且穩定上升 +0.5",
            "IC > 10 或無負債 +1",
            "IC > 5 +0.5",
            "D/E < 0.5 +1",
            "總評分",
            "最後更新時間"
        ]])
        print("✅ 表頭已建立")

# ==============================
# 抓取 yfinance 數據
# ==============================
def fetch_data(ticker):
    stock = yf.Ticker(ticker)

    # 取得財務報表（年報，最多4年，yfinance 免費版限制）
    income_stmt  = stock.financials          # 損益表（欄位 = 年份，由新到舊）
    balance_stmt = stock.balance_sheet       # 資產負債表
    cashflow_stmt = stock.cashflow           # 現金流量表
    info = stock.info                        # 基本資訊（含股價）

    return income_stmt, balance_stmt, cashflow_stmt, info

# ==============================
# 從 DataFrame 提取數列（新→舊）
# ==============================
def extract_series(df, *keys):
    """嘗試多個 key 名稱，回傳第一個找到的數列（list，新到舊）"""
    for key in keys:
        if df is not None and key in df.index:
            values = df.loc[key].tolist()
            # 過濾 NaN
            cleaned = [v for v in values if v is not None and v == v]  # v==v 過濾 NaN
            if cleaned:
                return cleaned
    return []

# ==============================
# 趨勢判斷工具
# ==============================
def is_growing(values):
    """判斷數列是否整體呈上升趨勢（允許小幅波動，60%的年份需上升）
    數據是新到舊排列：values[0]=最新年, values[-1]=最舊年
    上升 = 越舊的數字越小，即 values[i] <= values[i-1]
    """
    if len(values) < 3:
        return False
    ups = sum(1 for i in range(1, len(values)) if values[i] <= values[i-1])
    return ups >= (len(values) - 1) * 0.6

def is_declining(values):
    """判斷數列是否整體呈下降趨勢（60%的年份需下降）"""
    if len(values) < 3:
        return False
    downs = sum(1 for i in range(1, len(values)) if values[i] >= values[i-1])
    return downs >= (len(values) - 1) * 0.6

def all_positive(values):
    return bool(values) and all(v > 0 for v in values if v is not None)

# ==============================
# 評分計算
# ==============================
def score(ticker, income_stmt, balance_stmt, cashflow_stmt, info):

    # --- 數據提取 ---
    # EPS
    eps = extract_series(income_stmt, "Diluted EPS", "Basic EPS")

    # 股息（現金流量表中的 Common Stock Dividends，通常為負數）
    dividends = extract_series(cashflow_stmt,
                               "Common Stock Dividend Paid",
                               "Cash Dividends Paid",
                               "Payment Of Dividends")

    # 流通股數
    shares = extract_series(income_stmt,
                            "Diluted Average Shares",
                            "Basic Average Shares")

    # Book Value Per Share：yfinance 不直接提供，用 equity / shares 計算
    equity = extract_series(balance_stmt,
                            "Stockholders Equity",
                            "Total Stockholders Equity",
                            "Common Stock Equity")
    shares_out = extract_series(balance_stmt,
                                "Ordinary Shares Number",
                                "Common Stock Shares Outstanding")
    if equity and shares_out:
        length = min(len(equity), len(shares_out))
        bvps = [equity[i] / shares_out[i] for i in range(length) if shares_out[i] and shares_out[i] != 0]
    else:
        bvps = []

    # FCF
    fcf = extract_series(cashflow_stmt,
                         "Free Cash Flow",
                         "Capital Expenditures")  # 備用

    # Net Margin：yfinance 不直接提供，用 Net Income / Total Revenue 計算
    net_income = extract_series(income_stmt, "Net Income", "Net Income Common Stockholders")
    revenue    = extract_series(income_stmt, "Total Revenue")
    if net_income and revenue:
        length = min(len(net_income), len(revenue))
        net_margin = [net_income[i] / revenue[i] for i in range(length) if revenue[i] and revenue[i] != 0]
    else:
        net_margin = []

    # ROE：Net Income / Stockholders Equity
    if net_income and equity:
        length = min(len(net_income), len(equity))
        roe = [(net_income[i] / equity[i]) * 100 for i in range(length) if equity[i] and equity[i] != 0]
    else:
        roe = []

    # Interest Coverage：EBIT / Interest Expense
    ebit     = extract_series(income_stmt, "EBIT", "Operating Income")
    int_exp  = extract_series(income_stmt, "Interest Expense", "Interest Expense Non Operating")
    int_exp  = [abs(v) for v in int_exp]  # 確保為正數
    if ebit and int_exp:
        length = min(len(ebit), len(int_exp))
        ic_values = [ebit[i] / int_exp[i] if int_exp[i] and int_exp[i] != 0 else None for i in range(length)]
    else:
        ic_values = []

    # D/E Ratio
    total_debt = extract_series(balance_stmt, "Total Debt", "Long Term Debt And Capital Lease Obligation")
    if total_debt and equity:
        length = min(len(total_debt), len(equity))
        de_ratio = [total_debt[i] / equity[i] for i in range(length) if equity[i] and equity[i] != 0]
    else:
        de_ratio = []

    # 股價
    price = info.get("currentPrice") or info.get("regularMarketPrice") or "N/A"

    # --- 評分邏輯 ---
    results = {}

    # EPS 10年穩定成長 (+1)
    results["eps_growth"] = 1 if is_growing(eps) else 0

    # Dividend 10年持續穩定 (+0.5)
    div_abs = [abs(d) for d in dividends]
    results["div_stable"] = 0.5 if len([d for d in div_abs if d > 0]) >= 3 else 0  # yfinance 最多4年，放寬至3年

    # Dividend 越發越多 (+1)
    results["div_growing"] = 1 if is_growing(div_abs) else 0

    # Shares 流通股數10年穩定減少 (+1)
    results["shares_declining"] = 1 if is_declining(shares) else 0

    # Book Value Per Share 10年穩定成長 (+1)
    results["bvps_growth"] = 1 if is_growing(bvps) else 0

    # FCF 10年皆為正數 (+1)
    results["fcf_positive"] = 1 if all_positive(fcf) else 0

    # Net Margin > 10% (+1)
    avg_nm = sum(net_margin) / len(net_margin) if net_margin else 0
    results["nm_above10"] = 1 if avg_nm > 0.10 else 0

    # Net Margin 10年穩定不衰退 (+0.5)
    results["nm_stable"] = 0.5 if net_margin and not is_declining(net_margin) else 0

    # ROE 15%~40% (+1)
    avg_roe = sum(roe) / len(roe) if roe else 0
    results["roe_range"] = 1 if 15 <= avg_roe <= 40 else 0

    # ROE < 15% 且穩定上升 (+0.5)
    results["roe_rising"] = 0.5 if avg_roe < 15 and is_growing(roe) else 0

    # IC > 10 或無負債 (+1)
    no_debt = not int_exp or all(v == 0 for v in int_exp)
    high_ic = bool(ic_values) and all(v is None or v > 10 for v in ic_values)
    results["ic_high"] = 1 if (no_debt or high_ic) else 0

    # IC > 5 (+0.5)
    mid_ic = bool(ic_values) and all(v is None or v > 5 for v in ic_values)
    results["ic_mid"] = 0.5 if mid_ic and not high_ic else 0

    # D/E < 0.5 (+1)
    avg_de = sum(de_ratio) / len(de_ratio) if de_ratio else 99
    results["de_low"] = 1 if avg_de < 0.5 else 0

    return price, results

# ==============================
# 寫入 Google Sheet
# ==============================
def write_to_sheet(sheet, ticker, price, results):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    row_data = [
        ticker.upper(),
        price,
        results["eps_growth"],
        results["div_stable"],
        results["div_growing"],
        results["shares_declining"],
        results["bvps_growth"],
        results["fcf_positive"],
        results["nm_above10"],
        results["nm_stable"],
        results["roe_range"],
        results["roe_rising"],
        results["ic_high"],
        results["ic_mid"],
        results["de_low"],
        "",  # 總評分（下面替換成公式）
        now
    ]

    # 找看看這支股票是否已存在
    all_values = sheet.col_values(1)
    ticker_upper = ticker.upper()
    if ticker_upper in all_values:
        row_idx = all_values.index(ticker_upper) + 1
        print(f"🔄 更新第 {row_idx} 列：{ticker_upper}")
    else:
        row_idx = len(all_values) + 1
        print(f"➕ 新增第 {row_idx} 列：{ticker_upper}")

    # 替換公式中的 row index
    row_data[15] = f"=SUM(C{row_idx}:O{row_idx})"

    sheet.update(f"A{row_idx}:Q{row_idx}", [row_data])
    print(f"✅ {ticker_upper} 資料已寫入，股價：{price}")

# ==============================
# 主程式
# ==============================
def main():
    sheet = connect_sheet()
    init_headers(sheet)

    while True:
        ticker = input("\n輸入股票代號（輸入 q 離開）: ").strip().upper()
        if ticker == "Q":
            print("👋 結束程式")
            break
        if not ticker:
            continue

        print(f"📡 正在抓取 {ticker} 的資料...")
        try:
            income_stmt, balance_stmt, cashflow_stmt, info = fetch_data(ticker)
            if not info or info.get("regularMarketPrice") is None and info.get("currentPrice") is None:
                print(f"❌ 找不到 {ticker} 的資料，請確認代號是否正確")
                continue
            price, results = score(ticker, income_stmt, balance_stmt, cashflow_stmt, info)
            write_to_sheet(sheet, ticker, price, results)
        except Exception as e:
            print(f"❌ 發生錯誤：{e}")

if __name__ == "__main__":
    main()
