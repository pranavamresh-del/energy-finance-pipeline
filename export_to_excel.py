import yfinance as yf
import pandas as pd
from sqlalchemy import create_engine
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. Connect to SQLite database
engine = create_engine('sqlite:///energy_financials.db')

query = """
SELECT 
    i.ticker,
    i.fiscalyear,
    i.total_revenue,
    i.cost_of_revenue,
    i.gross_profit,
    i.operating_income,
    i.net_income,
    c.operating_cash_flow,
    c.capital_expenditure,
    c.free_cash_flow
FROM income_statement i
LEFT JOIN cash_flow c 
    ON i.ticker = c.ticker AND i.fiscalyear = c.fiscalyear
WHERE i.total_revenue IS NOT NULL
ORDER BY i.ticker ASC, i.fiscalyear ASC;
"""

df = pd.read_sql(query, engine)

wb = openpyxl.Workbook()
wb.remove(wb.active)  # Remove default blank sheet

# Institutional Styling Definitions
FONT_TITLE = Font(name='Calibri', size=14, bold=True, color='1F4E79')
FONT_HEADER = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
FONT_BLUE_INPUT = Font(name='Calibri', size=11, color='0000FF')
FONT_BLACK_FORMULA = Font(name='Calibri', size=11, color='000000', italic=True)
FONT_BOLD = Font(name='Calibri', size=11, bold=True)
FONT_REGULAR = Font(name='Calibri', size=11)

FILL_HEADER = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid') # Navy
FILL_FORMULA_BG = PatternFill(start_color='F2F4F7', end_color='F2F4F7', fill_type='solid') # Light Gray
FILL_SUMMARY_HEADER = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')

ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')

BORDER_THIN = Border(bottom=Side(style='thin', color='D9D9D9'))
BORDER_DOUBLE = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

# ==========================================
# PART 1: PEER VALUATION & COMPS SUMMARY TAB
# ==========================================
ws_summary = wb.create_sheet(title="Peer_Comps", index=0)

ws_summary['A1'] = "Energy Majors: Peer Valuation & Comparative Multiples"
ws_summary['A1'].font = FONT_TITLE
ws_summary['A2'] = "Automated Comps Dashboard | Market Cap & EV dynamically pulled from yfinance"
ws_summary['A2'].font = Font(name='Calibri', size=9, italic=True, color='595959')

headers_summary = [
    "Company / Ticker",
    "Share Price ($)",
    "Market Cap ($M)",
    "Enterprise Value ($M)",
    "LTM Revenue ($M)",
    "LTM Net Income ($M)",
    "LTM Free Cash Flow ($M)",
    "EV / Revenue",
    "P / E Ratio",
    "FCF Yield %"
]

for col_idx, h in enumerate(headers_summary, start=1):
    c = ws_summary.cell(row=4, column=col_idx, value=h)
    c.font = FONT_HEADER
    c.fill = FILL_SUMMARY_HEADER
    c.alignment = ALIGN_CENTER

tickers = df['ticker'].unique().tolist()

for row_idx, ticker in enumerate(tickers, start=5):
    print(f"Pulling live market valuation for {ticker}...")
    t_obj = yf.Ticker(ticker)
    info = t_obj.info or {}
    
    price = info.get('currentPrice') or info.get('regularMarketPrice') or 0.0
    mkt_cap = round((info.get('marketCap') or 0.0) / 1_000_000, 1)
    ev = round((info.get('enterpriseValue') or 0.0) / 1_000_000, 1)

    company_df = df[df['ticker'] == ticker].sort_values(by='fiscalyear')
    latest_row = company_df.iloc[-1]
    
    rev = round((latest_row['total_revenue'] or 0.0) / 1_000_000, 1)
    ni = round((latest_row['net_income'] or 0.0) / 1_000_000, 1)
    fcf = round((latest_row['free_cash_flow'] or 0.0) / 1_000_000, 1)

    ws_summary.cell(row=row_idx, column=1, value=ticker).font = FONT_BOLD
    ws_summary.cell(row=row_idx, column=1).alignment = ALIGN_CENTER

    c_price = ws_summary.cell(row=row_idx, column=2, value=price)
    c_price.number_format = '$#,##0.00'
    c_price.font = FONT_BLUE_INPUT
    c_price.alignment = ALIGN_RIGHT

    c_mc = ws_summary.cell(row=row_idx, column=3, value=mkt_cap)
    c_mc.number_format = '$#,##0.0;($#,##0.0);"-";@'
    c_mc.font = FONT_BLUE_INPUT
    c_mc.alignment = ALIGN_RIGHT

    c_ev = ws_summary.cell(row=row_idx, column=4, value=ev)
    c_ev.number_format = '$#,##0.0;($#,##0.0);"-";@'
    c_ev.font = FONT_BLUE_INPUT
    c_ev.alignment = ALIGN_RIGHT

    c_rev = ws_summary.cell(row=row_idx, column=5, value=rev)
    c_rev.number_format = '$#,##0.0;($#,##0.0);"-";@'
    c_rev.font = FONT_BLUE_INPUT
    c_rev.alignment = ALIGN_RIGHT

    c_ni = ws_summary.cell(row=row_idx, column=6, value=ni)
    c_ni.number_format = '$#,##0.0;($#,##0.0);"-";@'
    c_ni.font = FONT_BLUE_INPUT
    c_ni.alignment = ALIGN_RIGHT

    c_fcf = ws_summary.cell(row=row_idx, column=7, value=fcf)
    c_fcf.number_format = '$#,##0.0;($#,##0.0);"-";@'
    c_fcf.font = FONT_BLUE_INPUT
    c_fcf.alignment = ALIGN_RIGHT

    # Dynamic formulas for Multiples
    c_ev_rev = ws_summary.cell(row=row_idx, column=8, value=f"=D{row_idx}/E{row_idx}")
    c_ev_rev.number_format = '0.0x'
    c_ev_rev.font = FONT_BLACK_FORMULA
    c_ev_rev.alignment = ALIGN_RIGHT

    c_pe = ws_summary.cell(row=row_idx, column=9, value=f"=C{row_idx}/F{row_idx}")
    c_pe.number_format = '0.0x'
    c_pe.font = FONT_BLACK_FORMULA
    c_pe.alignment = ALIGN_RIGHT

    c_fcf_yield = ws_summary.cell(row=row_idx, column=10, value=f"=G{row_idx}/C{row_idx}")
    c_fcf_yield.number_format = '0.0%'
    c_fcf_yield.font = FONT_BLACK_FORMULA
    c_fcf_yield.alignment = ALIGN_RIGHT

    for col_i in range(1, 11):
        ws_summary.cell(row=row_idx, column=col_i).border = BORDER_THIN

# Industry Average Row
avg_row = len(tickers) + 5
ws_summary.cell(row=avg_row, column=1, value="Industry Average").font = FONT_BOLD
ws_summary.cell(row=avg_row, column=1).alignment = ALIGN_CENTER

for col_i in range(8, 11):
    col_let = get_column_letter(col_i)
    c_avg = ws_summary.cell(row=avg_row, column=col_i, value=f"=AVERAGE({col_let}5:{col_let}{avg_row-1})")
    c_avg.font = FONT_BOLD
    c_avg.alignment = ALIGN_RIGHT
    c_avg.border = BORDER_DOUBLE
    c_avg.number_format = '0.0x' if col_i < 10 else '0.0%'

for col_i in range(1, 8):
    ws_summary.cell(row=avg_row, column=col_i).border = BORDER_DOUBLE

for col in ws_summary.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_summary.column_dimensions[col_letter].width = max(max_len + 4, 18)


# ==========================================
# PART 2: INDIVIDUAL 3-STATEMENT TABS
# ==========================================
for ticker in tickers:
    ws = wb.create_sheet(title=f"{ticker}_Model")
    company_df = df[df['ticker'] == ticker].sort_values(by='fiscalyear')
    years = [int(y) for y in company_df['fiscalyear'].tolist()]
    
    ws['A1'] = f"{ticker} - Automated Financial Statements & KPIs"
    ws['A1'].font = FONT_TITLE
    ws['A2'] = "Source: SQLite (energy_financials.db) | Model Built with Python openpyxl"
    ws['A2'].font = Font(name='Calibri', size=9, italic=True, color='595959')

    ws.cell(row=4, column=1, value="Financial Metric / Line Item").font = FONT_HEADER
    ws.cell(row=4, column=1).fill = FILL_HEADER
    ws.cell(row=4, column=1).alignment = ALIGN_LEFT

    for col_idx, year in enumerate(years, start=2):
        header_cell = ws.cell(row=4, column=col_idx, value=f"FY {year}")
        header_cell.font = FONT_HEADER
        header_cell.fill = FILL_HEADER
        header_cell.alignment = ALIGN_CENTER

    # Row 5: Revenue
    ws.cell(row=5, column=1, value="Total Revenue ($M)").font = FONT_BOLD
    ws.cell(row=5, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        val = company_df[company_df['fiscalyear'] == year]['total_revenue'].values[0] or 0
        c = ws.cell(row=5, column=col_idx, value=round(val / 1_000_000, 1))
        c.font = FONT_BLUE_INPUT
        c.number_format = '$#,##0.0;($#,##0.0);"-";@'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 6: YoY Revenue Growth %
    ws.cell(row=6, column=1, value="  ↳ YoY Revenue Growth %").font = FONT_BLACK_FORMULA
    ws.cell(row=6, column=1).fill = FILL_FORMULA_BG
    ws.cell(row=6, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        c = ws.cell(row=6, column=col_idx)
        c.fill = FILL_FORMULA_BG
        c.font = FONT_BLACK_FORMULA
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN
        if col_idx == 2:
            c.value = "N/A"
            c.alignment = ALIGN_CENTER
        else:
            prev_letter = get_column_letter(col_idx - 1)
            curr_letter = get_column_letter(col_idx)
            c.value = f"=({curr_letter}5-{prev_letter}5)/{prev_letter}5"
            c.number_format = '0.0%'

    # Row 7: COGS
    ws.cell(row=7, column=1, value="Cost of Goods Sold (COGS) ($M)").font = FONT_REGULAR
    ws.cell(row=7, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        val = company_df[company_df['fiscalyear'] == year]['cost_of_revenue'].values[0] or 0
        c = ws.cell(row=7, column=col_idx, value=round(val / 1_000_000, 1))
        c.font = FONT_BLUE_INPUT
        c.number_format = '$#,##0.0;($#,##0.0);"-";@'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 8: Gross Profit
    ws.cell(row=8, column=1, value="Gross Profit ($M)").font = FONT_BOLD
    ws.cell(row=8, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        curr_letter = get_column_letter(col_idx)
        c = ws.cell(row=8, column=col_idx, value=f"={curr_letter}5-{curr_letter}7")
        c.font = FONT_BLACK_FORMULA
        c.number_format = '$#,##0.0;($#,##0.0);"-";@'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 9: Gross Margin %
    ws.cell(row=9, column=1, value="  ↳ Gross Margin %").font = FONT_BLACK_FORMULA
    ws.cell(row=9, column=1).fill = FILL_FORMULA_BG
    ws.cell(row=9, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        curr_letter = get_column_letter(col_idx)
        c = ws.cell(row=9, column=col_idx, value=f"={curr_letter}8/{curr_letter}5")
        c.fill = FILL_FORMULA_BG
        c.font = FONT_BLACK_FORMULA
        c.number_format = '0.0%'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 10: Operating Income (EBIT)
    ws.cell(row=10, column=1, value="Operating Income (EBIT) ($M)").font = FONT_BOLD
    ws.cell(row=10, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        val = company_df[company_df['fiscalyear'] == year]['operating_income'].values[0] or 0
        c = ws.cell(row=10, column=col_idx, value=round(val / 1_000_000, 1))
        c.font = FONT_BLUE_INPUT
        c.number_format = '$#,##0.0;($#,##0.0);"-";@'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 11: Operating Margin %
    ws.cell(row=11, column=1, value="  ↳ Operating Margin %").font = FONT_BLACK_FORMULA
    ws.cell(row=11, column=1).fill = FILL_FORMULA_BG
    ws.cell(row=11, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        curr_letter = get_column_letter(col_idx)
        c = ws.cell(row=11, column=col_idx, value=f"={curr_letter}10/{curr_letter}5")
        c.fill = FILL_FORMULA_BG
        c.font = FONT_BLACK_FORMULA
        c.number_format = '0.0%'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 12: Net Income
    ws.cell(row=12, column=1, value="Net Income ($M)").font = FONT_BOLD
    ws.cell(row=12, column=1).border = BORDER_DOUBLE
    for col_idx, year in enumerate(years, start=2):
        val = company_df[company_df['fiscalyear'] == year]['net_income'].values[0] or 0
        c = ws.cell(row=12, column=col_idx, value=round(val / 1_000_000, 1))
        c.font = FONT_BLUE_INPUT
        c.number_format = '$#,##0.0;($#,##0.0);"-";@'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_DOUBLE

    # Row 13: Free Cash Flow
    ws.cell(row=13, column=1, value="Free Cash Flow (FCF) ($M)").font = FONT_BOLD
    ws.cell(row=13, column=1).border = BORDER_THIN
    for col_idx, year in enumerate(years, start=2):
        val = company_df[company_df['fiscalyear'] == year]['free_cash_flow'].values[0] or 0
        c = ws.cell(row=13, column=col_idx, value=round(val / 1_000_000, 1))
        c.font = FONT_BLUE_INPUT
        c.number_format = '$#,##0.0;($#,##0.0);"-";@'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_THIN

    # Row 14: FCF Conversion %
    ws.cell(row=14, column=1, value="  ↳ FCF Conversion % (FCF / Net Income)").font = FONT_BLACK_FORMULA
    ws.cell(row=14, column=1).fill = FILL_FORMULA_BG
    ws.cell(row=14, column=1).border = BORDER_DOUBLE
    for col_idx, year in enumerate(years, start=2):
        curr_letter = get_column_letter(col_idx)
        c = ws.cell(row=14, column=col_idx, value=f"={curr_letter}13/{curr_letter}12")
        c.fill = FILL_FORMULA_BG
        c.font = FONT_BLACK_FORMULA
        c.number_format = '0.0%'
        c.alignment = ALIGN_RIGHT
        c.border = BORDER_DOUBLE

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

# Save
output_file = 'Energy_Financial_Model.xlsx'
wb.save(output_file)
print(f"✓ Valuation Comps & 3-Statement Model successfully written to: {output_file}")