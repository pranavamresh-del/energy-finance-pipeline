import yfinance as yf
import pandas as pd
from sqlalchemy import create_engine
import os

TICKERS = ['CVX', 'XOM', 'COP']
db_filename = 'energy_financials.db'
engine = create_engine(f'sqlite:///{db_filename}')

def fetch_and_clean_financials(ticker_symbol):
    print(f"Fetching data for {ticker_symbol}...")
    ticker = yf.Ticker(ticker_symbol)

    income = ticker.financials.T
    balance = ticker.balance_sheet.T
    cash_flow = ticker.cashflow.T

    def prepare_df(df, statement_type):
        df = df.copy()
        df.index.name = 'FiscalYear'
        df.reset_index(inplace=True)
        
        # Extract clean 4-digit year
        df['FiscalYear'] = pd.to_datetime(df['FiscalYear']).dt.year
        df['Ticker'] = ticker_symbol
        df['StatementType'] = statement_type
        
        # Standardize column headers
        df.columns = [str(col).strip().lower().replace(' ', '_').replace('-', '_') for col in df.columns]
        return df

    return (
        prepare_df(income, 'income_statement'),
        prepare_df(balance, 'balance_sheet'),
        prepare_df(cash_flow, 'cash_flow')
    )

print("--- Starting Pipeline ---")

all_income = []
all_balance = []
all_cashflow = []

# Step 1: Collect and align all company data in memory first
for ticker in TICKERS:
    try:
        inc_df, bal_df, cf_df = fetch_and_clean_financials(ticker)
        all_income.append(inc_df)
        all_balance.append(bal_df)
        all_cashflow.append(cf_df)
        print(f"✓ Downloaded {ticker}")
    except Exception as e:
        print(f"✕ Failed to download {ticker}: {e}")

# Step 2: Concatenate dataframes (this automatically unions and fills missing columns with NaN)
print("\nCombining and normalizing schemas...")
combined_income = pd.concat(all_income, ignore_index=True)
combined_balance = pd.concat(all_balance, ignore_index=True)
combined_cashflow = pd.concat(all_cashflow, ignore_index=True)

# Step 3: Write cleanly to SQLite (replaces old mismatched tables)
combined_income.to_sql('income_statement', engine, if_exists='replace', index=False)
combined_balance.to_sql('balance_sheet', engine, if_exists='replace', index=False)
combined_cashflow.to_sql('cash_flow', engine, if_exists='replace', index=False)

print("\n--- Pipeline Completed Successfully ---")
print(f"Database generated at: {os.path.abspath(db_filename)}")
print(f"Rows stored - Income: {len(combined_income)}, Balance Sheet: {len(combined_balance)}, Cash Flow: {len(combined_cashflow)}")