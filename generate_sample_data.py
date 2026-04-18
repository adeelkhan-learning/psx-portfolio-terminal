"""
PSX Portfolio Dashboard - Sample Data Generator
Run this once to create sample_portfolio.xlsx
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

random.seed(42)
np.random.seed(42)

# --- CONFIG ---
TICKERS = ["ENGRO", "LUCK", "HBL", "PSO", "OGDC", "MCB", "UBL", "NESTLE", "SEARL", "EFERT"]
START_DATE = datetime(2022, 1, 1)
END_DATE = datetime(2024, 12, 31)

def random_date(start=START_DATE, end=END_DATE):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))

# ── TRADES ──────────────────────────────────────────────────────────────────
trades_data = []
base_prices = {
    "ENGRO": 280, "LUCK": 680, "HBL": 140, "PSO": 220, "OGDC": 120,
    "MCB": 195, "UBL": 175, "NESTLE": 6200, "SEARL": 95, "EFERT": 105
}

for _ in range(120):
    ticker = random.choice(TICKERS)
    trade_date = random_date()
    settlement_date = trade_date + timedelta(days=2)
    txn_type = random.choice(["Buy", "Buy", "Buy", "Sell"])  # more buys

    base = base_prices[ticker]
    price = round(base * random.uniform(0.75, 1.35), 2)
    qty = random.choice([100, 200, 300, 500, 1000])
    commission = round(price * qty * 0.0015, 2)
    taxes = round(price * qty * 0.001, 2)
    gross = price * qty
    net = gross + commission + taxes if txn_type == "Buy" else gross - commission - taxes

    trades_data.append({
        "Trade Date": trade_date.strftime("%Y-%m-%d"),
        "Settlement Date": settlement_date.strftime("%Y-%m-%d"),
        "Transaction Type": txn_type,
        "Ticker": ticker,
        "Quantity": qty,
        "Price": price,
        "Commission": commission,
        "Taxes and Fees": taxes,
        "Net Total Value": round(net, 2)
    })

df_trades = pd.DataFrame(trades_data).sort_values("Trade Date").reset_index(drop=True)

# ── DIVIDENDS ────────────────────────────────────────────────────────────────
dividends_data = []
dividend_tickers = {
    "ENGRO": 12, "HBL": 6, "MCB": 9, "PSO": 5, "OGDC": 8,
    "UBL": 7, "EFERT": 10, "LUCK": 4
}
company_names = {
    "ENGRO": "Engro Corporation Ltd", "HBL": "Habib Bank Ltd",
    "MCB": "MCB Bank Ltd", "PSO": "Pakistan State Oil",
    "OGDC": "Oil & Gas Development Company", "UBL": "United Bank Ltd",
    "EFERT": "Engro Fertilizers Ltd", "LUCK": "Lucky Cement Ltd"
}

for ticker, rate in dividend_tickers.items():
    for year in [2022, 2023, 2024]:
        for quarter in [1, 2]:  # bi-annual dividends
            date = datetime(year, quarter * 6, 15)
            if date > END_DATE:
                continue
            securities = random.choice([500, 1000, 1500, 2000])
            gross = round(securities * rate * random.uniform(0.9, 1.2), 2)
            zakat = round(gross * 0.025, 2)
            tax = round(gross * 0.15, 2)
            net = round(gross - zakat - tax, 2)
            dividends_data.append({
                "Payment Date": date.strftime("%Y-%m-%d"),
                "Company Name": company_names[ticker],
                "No. of Securities": securities,
                "Rate Per Security": rate,
                "Gross Dividend": gross,
                "Zakat Deducted": zakat,
                "Tax Deducted": tax,
                "Net Amount Paid": net,
                "Ticker": ticker
            })

df_dividends = pd.DataFrame(dividends_data).sort_values("Payment Date").reset_index(drop=True)

# ── FUNDS ────────────────────────────────────────────────────────────────────
funds_data = []
for i in range(30):
    date = random_date()
    transfer_type = random.choice(["Deposit", "Deposit", "Deposit", "Withdraw"])
    amount = random.choice([50000, 100000, 150000, 200000, 250000, 300000])
    hold = round(amount * random.uniform(0, 0.05), 2) if transfer_type == "Deposit" else 0
    exposure = round(amount * random.uniform(0.7, 0.95), 2) if transfer_type == "Deposit" else 0
    funds_data.append({
        "Date": date.strftime("%Y-%m-%d"),
        "Transfer Type": transfer_type,
        "Amount Deposit": amount if transfer_type == "Deposit" else 0,
        "Amount Hold Against Charges / Dues": hold,
        "Amount Transferred To Exposure": exposure
    })

df_funds = pd.DataFrame(funds_data).sort_values("Date").reset_index(drop=True)

# ── SAVE ─────────────────────────────────────────────────────────────────────
output = "sample_portfolio.xlsx"
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_trades.to_excel(writer, sheet_name="Trades", index=False)
    df_dividends.to_excel(writer, sheet_name="Dividends", index=False)
    df_funds.to_excel(writer, sheet_name="Funds", index=False)

print(f"✅  Sample data saved → {output}")
print(f"   Trades:    {len(df_trades)} rows")
print(f"   Dividends: {len(df_dividends)} rows")
print(f"   Funds:     {len(df_funds)} rows")
