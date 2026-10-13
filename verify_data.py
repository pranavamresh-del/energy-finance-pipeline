import pandas as pd
from sqlalchemy import create_engine

# Connect to the database created by build_pipeline.py
engine = create_engine('sqlite:///energy_financials.db')

# SQL query: Retrieve ticker, year, revenue, operating income, and operating margin %
query = """
SELECT 
    ticker,
    fiscalyear,
    total_revenue,
    operating_income,
    ROUND((operating_income * 1.0 / total_revenue) * 100, 2) AS operating_margin_pct
FROM income_statement
WHERE total_revenue IS NOT NULL AND operating_income IS NOT NULL
ORDER BY fiscalyear DESC, total_revenue DESC;
"""

# Execute query and print results
df = pd.read_sql(query, engine)
print("\n=== Multi-Company Energy Comparison ===")
print(df.to_string(index=False))