import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ==============================
# 設定區（回家後填入這裡）
# ==============================
FMP_API_KEY = "你的FMP_API_KEY"
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
# 抓取 FMP 數據
# ==============================
def fetch_data(ticker):
    base = "https://financialmodelingprep.com/api/v3"
    params = f"?limit=10&apikey={FMP_API_KEY}"

    def get(endpoint):
        r = requests.get(f"{base}/{endpoint}/{ticker}{params}")
        return r.json() if r.status_code == 200 else []

    income     = get("income-statement")
    balance    = get("balance-sheet-statement")
    cashflow   = get("cash-flow-statement")
    profile    = requests.get(f"{base}/profile/{ticker}?apikey={FMP_API_KEY}").json()

    return income, balance, cashflow, profile

# ==============================
# 趨勢判斷工具
# ==============================
def is_growing(values):
    """判斷數列是否整體呈上升趨勢（允許小幅波動，60%的年份需上升）
    FMP 數據是新到舊排列：values[0]=最新年, values[-1]=最舊年
    上升 = 越舊的數字越小，即 values[i] <= values[i-1]
    """
    if len(values) < 3:
        return False
    ups = sum(1 for i in range(1, len(values)) if values[i] <= values[i-1])
    return ups >= (len(values) - 1) * 0.6

def is_declining(values):
    """判斷數列是否整體呈下降趨勢（60%的年份需下降）
    下降 = 越舊的數字越大，即 values[i] >= values[i-1]
    """
    if len(values) < 3:
        return False
    downs = sum(1 for i in range(1, len(values)) if values[i] >= values[i-1])
    return downs >= (len(values) - 1) * 0.6

def all_positive(values):
    return all(v > 0 for v in values if v is not None)

# ==============================
# 評分計算
# ==============================
def score(ticker, income, balance, cashflow, profile):
    def extract(data, key):
        return [d.get(key, 0) or 0 for d in data if key in d]

    # --- 數據提取 ---
    eps         = extract(income, "eps")
    dividends   = extract(cashflow, "dividendsPaid")  # 負數代表有付
    shares      = extract(income, "weightedAverageShsOut")
    bvps        = extract(balance, "bookValuePerShare") if "bookValuePerShare" in (balance[0] if balance else {}) else []
    # BVPS 備用計算
    if not bvps and balance:
        bvps = [(d.get("totalStockholdersEquity",0) or 0) / (d.get("commonStock",1) or 1) for d in balance]
    fcf         = extract(cashflow, "freeCashFlow")
    net_margin  = extract(income, "netProfitMargin")
    roe         = extract(income, "returnOnEquity") if "returnOnEquity" in (income[0] if income else {}) else []
    # ROE 備用計算
    if not roe and income and balance:
        roe = []
        for i, d in enumerate(income):
            eq = balance[i].get("totalStockholdersEquity", 1) or 1
            ni = d.get("netIncome", 0) or 0
            roe.append(ni / eq * 100)
    interest_exp = extract(income, "interestExpense")
    ebit         = extract(income, "ebitda")  # 近似用 EBITDA
    de_ratio     = extract(balance, "debtEquityRatio") if "debtEquityRatio" in (balance[0] if balance else {}) else []
    if not de_ratio and balance:
        de_ratio = []
        for d in balance:
            eq = d.get("totalStockholdersEquity", 1) or 1
            debt = d.get("totalDebt", 0) or 0
            de_ratio.append(debt / eq)

    # --- 股價 ---
    price = profile[0].get("price", "N/A") if profile else "N/A"

    # --- 評分邏輯 ---
    results = {}

    # EPS 10年穩定成長 (+1)
    results["eps_growth"] = 1 if is_growing(eps) else 0

    # Dividend 10年持續穩定 (+0.5)
    div_abs = [abs(d) for d in dividends]
    results["div_stable"] = 0.5 if len([d for d in div_abs if d > 0]) >= 8 else 0

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
    results["nm_stable"] = 0.5 if not is_declining(net_margin) else 0

    # ROE 15%~40% (+1)
    avg_roe = sum(roe) / len(roe) if roe else 0
    results["roe_range"] = 1 if 15 <= avg_roe <= 40 else 0

    # ROE < 15% 且穩定上升 (+0.5)
    results["roe_rising"] = 0.5 if avg_roe < 15 and is_growing(roe) else 0

    # IC > 10 或 "-" (+1)
    ic_values = []
    for i in range(len(ebit)):
        ie = interest_exp[i] if i < len(interest_exp) else 0
        ic_values.append(ebit[i] / ie if ie and ie != 0 else None)
    no_debt = all(v is None for v in ic_values)
    high_ic = all(v is None or v > 10 for v in ic_values)
    results["ic_high"] = 1 if (no_debt or high_ic) else 0

    # IC > 5 (+0.5)
    results["ic_mid"] = 0.5 if all(v is None or v > 5 for v in ic_values) and not high_ic else 0

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
        f"=SUM(C{{row}}:O{{row}})",  # 總評分公式（佔位，下面替換）
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

    # 替換公式中的 {row}
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
            income, balance, cashflow, profile = fetch_data(ticker)
            if not income:
                print(f"❌ 找不到 {ticker} 的資料，請確認代號是否正確")
                continue
            price, results = score(ticker, income, balance, cashflow, profile)
            write_to_sheet(sheet, ticker, price, results)
        except Exception as e:
            print(f"❌ 發生錯誤：{e}")

if __name__ == "__main__":
    main()
