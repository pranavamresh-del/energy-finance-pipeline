# Automated Energy Sector Financial Pipeline & Peer Valuation Model

## Project Overview
This project is an end-to-end financial data engineering and corporate valuation pipeline designed to automate financial data workflows in the energy sector. In corporate finance and equity research, gathering multi-year historical financial statement data across peer companies often involves time-consuming manual entry from separate 10-K filings, which introduces human error and creates version inconsistencies. This project bridges data systems and corporate finance by building an automated extraction, transformation, and loading (ETL) architecture that extracts raw financial data for global energy leaders—specifically Chevron (CVX), ExxonMobil (XOM), and ConocoPhillips (COP)—standardizes their disparate reporting schemas, stores them in a local relational database, and dynamically exports the results into institutional-grade Excel valuation models.

## Pipeline Architecture & Engineering Workflow
The architecture begins with Python utilizing the yfinance interface to systematically download 4-year historical income statements, balance sheets, and cash flow statements for each peer company. Because different corporations use distinct line-item terminology in their SEC reporting—such as varying line items for unusual expenses, exploratory costs, and general operating expenditures—a direct database upload would fail due to schema collisions. To resolve this, the data extraction script implements dynamic dataframe restructuring and schema alignment using pandas, transposing the financial statements so fiscal dates represent uniform record indices and aligning all column metrics into a standardized tabular format before persisting the data into a local relational SQLite database (energy_financials.db) via SQLAlchemy.

## Automated Institutional Financial Modeling
Once the clean data is stored within SQLite, a specialized financial reporting script connects to the database and builds an Excel financial model (Energy_Financial_Model.xlsx) using the openpyxl library. The generated spreadsheet strictly follows professional Wall Street modeling conventions: all historical raw inputs are styled in classic blue font, while calculated outputs and margin metrics are styled in black italic font with custom background accents. Rather than exporting static numbers, the pipeline programmatically inserts live Excel mathematical formulas into each company model tab for Year-over-Year (YoY) Revenue Growth, Gross Profit, Gross Margin, Operating Margin, and Free Cash Flow Conversion, ensuring that the model functions as a dynamic analytical tool rather than a flat report.

## Peer Valuation & Multiple Analysis
In addition to individual company financial statements, the pipeline generates an executive Peer Comps dashboard on the first tab of the workbook. This summary dashboard connects to live market data feeds to extract real-time share prices, market capitalizations, and enterprise values (EV) across all target firms, pairing them with the latest twelve-month (LTM) operating figures pulled from the SQLite database. From this integrated data, the system automatically calculates key industry valuation benchmarks including EV-to-Revenue, Price-to-Earnings (P/E), and Free Cash Flow Yield percentages, complete with an automated industry peer average benchmark row at the bottom of the table.

## Setup & Execution Guide
To run this pipeline locally, open your Terminal on macOS and clone the project repository by running git clone https://github.com/your-username/corporate-finance-pipeline.git, then navigate into the project root directory with cd corporate-finance-pipeline. Create and activate an isolated Python virtual environment by entering python3 -m venv venv followed by source venv/bin/activate, which isolates your dependencies. Install the required data engineering and modeling libraries by executing pip install yfinance pandas sqlalchemy openpyxl. Once the environment is configured, execute python build_pipeline.py in your terminal to extract the raw financial filings, resolve schema differences, and generate the energy_financials.db database. Finally, run python export_to_excel.py to query the database, pull current valuation data, and generate the fully formatted Energy_Financial_Model.xlsx workbook ready for financial analysis and presentation.

## Power BI Financial Reporting Layer

The curated financial data from the pipeline feeds directly into an executive Power BI dashboard.

![Dashboard Overview](power_bi/images/dashboard_overview.png)

### Key Visual Features & DAX Measures
* **Multi-Statement Relational Model:** Joins normalized statement tables via composite keys.
* **Free Cash Flow (FCF):** Cash generated after capital reinvestment.
* **OCF Margin %:** Operational cash conversion relative to revenue.
* **Capital Returns & Solvency:** ROE % and Debt-to-Equity tracking over multi-year reporting cycles.