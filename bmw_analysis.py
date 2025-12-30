# ============================================================
# BMW FINANCIAL STATEMENT ANALYSIS TOOL
# Enhanced Version: Charts, Excel Export, Resume/GitHub Ready
# ============================================================

import yfinance as yf
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import os

sns.set(style="whitegrid")  # prettier charts


# ------------------------------------------------------------
# Fetch BMW Financial Statements
# ------------------------------------------------------------
def fetch_bmw_financials():
    ticker = yf.Ticker("BMW.DE")

    income = ticker.financials.T
    balance = ticker.balance_sheet.T
    cashflow = ticker.cashflow.T

    income.columns = [f"IS_{col}" for col in income.columns]
    balance.columns = [f"BS_{col}" for col in balance.columns]
    cashflow.columns = [f"CF_{col}" for col in cashflow.columns]

    financials = pd.concat([income, balance, cashflow], axis=1)
    financials.index = financials.index.year
    financials.sort_index(inplace=True)

    return financials


# ------------------------------------------------------------
# Calculate Key Financial Ratios (Robust)
# ------------------------------------------------------------
def calculate_ratios(df):
    ratios = pd.DataFrame(index=df.index)

    equity_candidates = [
        "BS_Total Stockholder Equity",
        "BS_Stockholders Equity",
        "BS_Total Equity Gross Minority Interest",
        "BS_Total Equity"
    ]
    assets_candidates = ["BS_Total Assets"]
    liabilities_candidates = [
        "BS_Total Liab",
        "BS_Total Liabilities",
        "BS_Total Liabilities Net Minority Interest"
    ]
    revenue_candidates = ["IS_Total Revenue"]
    net_income_candidates = ["IS_Net Income"]

    def find_column(candidates):
        for col in candidates:
            if col in df.columns:
                return col
        return None

    equity_col = find_column(equity_candidates)
    assets_col = find_column(assets_candidates)
    liabilities_col = find_column(liabilities_candidates)
    revenue_col = find_column(revenue_candidates)
    net_income_col = find_column(net_income_candidates)

    missing = [
        name for name, col in {
            "Equity": equity_col,
            "Assets": assets_col,
            "Liabilities": liabilities_col,
            "Revenue": revenue_col,
            "Net Income": net_income_col
        }.items() if col is None
    ]

    if missing:
        raise KeyError(f"Missing required columns: {missing}")

    ratios["ROE (%)"] = (df[net_income_col] / df[equity_col]) * 100
    ratios["ROA (%)"] = (df[net_income_col] / df[assets_col]) * 100
    ratios["Debt to Equity"] = df[liabilities_col] / df[equity_col]
    ratios["Net Profit Margin (%)"] = (df[net_income_col] / df[revenue_col]) * 100

    return ratios.round(2)


# ------------------------------------------------------------
# Plot Financial Charts
# ------------------------------------------------------------
def plot_financials(df, ratios):
    # Revenue & Net Income
    plt.figure(figsize=(12, 6))
    if "IS_Total Revenue" in df.columns:
        plt.plot(df.index, df["IS_Total Revenue"]/1e9, marker='o', label="Revenue (B EUR)")
    if "IS_Net Income" in df.columns:
        plt.plot(df.index, df["IS_Net Income"]/1e9, marker='o', label="Net Income (B EUR)")
    plt.title("BMW Revenue & Net Income Trend")
    plt.xlabel("Year")
    plt.ylabel("Billion EUR")
    plt.legend()
    plt.tight_layout()
    plt.show()

    # ROE & ROA
    plt.figure(figsize=(12, 6))
    plt.plot(ratios.index, ratios["ROE (%)"], marker='o', label="ROE (%)")
    plt.plot(ratios.index, ratios["ROA (%)"], marker='o', label="ROA (%)")
    plt.title("BMW ROE & ROA Trend")
    plt.xlabel("Year")
    plt.ylabel("Percentage (%)")
    plt.legend()
    plt.tight_layout()
    plt.show()


# ------------------------------------------------------------
# Export to Excel
# ------------------------------------------------------------
def export_to_excel(df, ratios, filename="BMW_Financial_Analysis.xlsx"):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Financials")
        ratios.to_excel(writer, sheet_name="Ratios")

        workbook = writer.book
        worksheet = writer.sheets["Ratios"]

        # Optional: add simple formatting
        format1 = workbook.add_format({'num_format': '0.00'})
        worksheet.set_column('B:E', 12, format1)

    print(f"\n📁 Exported financials & ratios to: {os.path.abspath(filename)}")


# ------------------------------------------------------------
# MAIN EXECUTION
# ------------------------------------------------------------
if __name__ == "__main__":

    print("BMW FINANCIAL STATEMENT ANALYSIS TOOL")
    print("Analyzing Real Financial Data")
    print("=" * 70)

    print("\nFetching real financial data for BMW AG (Bayerische Motoren Werke)...")
    print("=" * 70)

    try:
        df = fetch_bmw_financials()
        print("✓ Financial data fetched successfully\n")

        print("📊 Financial Statements (first 5 rows):")
        print(df.head())

        ratios = calculate_ratios(df)
        print("\n📈 Key Financial Ratios:")
        print(ratios)

        # Charts
        plot_financials(df, ratios)

        # Export to Excel
        export_to_excel(df, ratios)

        print("\n🚀 Analysis Complete! Ready for Resume / GitHub Showcase.")

    except Exception as e:
        print(f"\n✗ Error fetching data: {e}")
        print("\nTip: Make sure you have internet connection and yfinance installed:")
        print("  pip install yfinance")
