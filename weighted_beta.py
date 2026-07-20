import akshare as ak
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import warnings

input_file = "Holdings.xlsx"
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

# Remap inaccessible listings to accessible equivalents.
# PRX (Euronext Amsterdam, EUR) → Tencent  
TICKER_REMAP = {
    ("PRX", "EUR"): ("PROSY", "USD"),
}


# FX row positions on the 'FX Rates' sheet (row 1 = header)
FX_SHEET     = "FX Rates"
FX_ROW_USD   = 3   # USD/HKD
FX_ROW_CNH   = 4   # CNH/HKD

# Styling constants
BLUE  = "FF0000FF"   # hardcoded inputs
BLACK = "FF000000"   # formulas / calculations
GREEN = "FF008000"   # cross-sheet links

HEADER_FILL = PatternFill("solid", start_color="FF1F4E79")   # dark navy
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFFFF", size=9)
DATA_FONT   = Font(name="Arial", size=9)
INPUT_FONT  = Font(name="Arial", size=9, color=BLUE)
FORMULA_FONT= Font(name="Arial", size=9, color=BLACK)
LINK_FONT   = Font(name="Arial", size=9, color=GREEN)

THIN = Side(style="thin", color="FFD9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---------------------------------------------------------------------------
# TICKER FORMATTERS
# ---------------------------------------------------------------------------

def format_hk_ticker(ticker: str) -> str:
    ticker = str(ticker).strip()
    if "." in ticker:
        ticker = ticker.split(".")[0]
    return ticker.zfill(5)

def format_a_ticker(ticker: str) -> str:
    ticker = str(ticker).strip()
    if ticker.startswith("6"):
        return f"sh{ticker}"
    elif ticker.startswith("0") or ticker.startswith("3"):
        return f"sz{ticker}"
    elif ticker.startswith("4") or ticker.startswith("8"):
        return f"bj{ticker}" if len(ticker) == 6 else f"sz{ticker}"
    return ticker

# ---------------------------------------------------------------------------
# PRICE FETCHERS  (return DataFrame with DatetimeIndex and 'Close' column)
# ---------------------------------------------------------------------------

def _normalise(hist: pd.DataFrame, date_col: str, close_col: str) -> pd.DataFrame | None:
    if hist is None or hist.empty:
        return None
    hist = hist.rename(columns={date_col: "Date", close_col: "Close"})
    hist["Date"] = pd.to_datetime(hist["Date"])
    hist = hist.set_index("Date").sort_index()
    hist["Close"] = pd.to_numeric(hist["Close"], errors="coerce")
    return hist[["Close"]]

def get_hk_prices(ticker: str) -> pd.DataFrame | None:
    try:
        return _normalise(ak.stock_hk_daily(symbol=format_hk_ticker(ticker)), "date", "close")
    except Exception as e:
        print(f"  Sina HK failed for {ticker}: {e}"); return None

def get_a_share_prices(ticker: str) -> pd.DataFrame | None:
    try:
        return _normalise(ak.stock_zh_a_daily(symbol=format_a_ticker(ticker)), "date", "close")
    except Exception as e:
        print(f"  Sina A-share failed for {ticker}: {e}"); return None

def get_us_prices(ticker: str) -> pd.DataFrame | None:
    try:
        return _normalise(ak.stock_us_daily(symbol=str(ticker).strip()), "date", "close")
    except Exception as e:
        print(f"  Sina US failed for {ticker}: {e}"); return None

def get_prices(ticker: str, currency: str) -> pd.DataFrame | None:
    ticker, currency = TICKER_REMAP.get((ticker, currency), (ticker, currency))
    if currency == "CNH" and len(ticker)==6:
        return get_a_share_prices(ticker)
    elif currency in  ['HKD', 'CNH']:
        return get_hk_prices(ticker)
    elif currency in  ['USD', 'EUR']:
        return get_us_prices(ticker)
    else:
        print(f"  {ticker} ({currency}): no data source — skipping."); return None

# ---------------------------------------------------------------------------
# TRADING-DAY ALIGNMENT
#
# Strategy: use the benchmark's trading calendar as the master index.
# On days where an asset did not trade (different holiday calendar),
# forward-fill its last known price.  This prevents US or China holidays
# from silently dropping valid HK benchmark observations, and vice-versa.
# Returns are then computed on the aligned (filled) price series so that
# every benchmark return has a matching asset return.
# ---------------------------------------------------------------------------

def align_to_benchmark(asset: pd.DataFrame, benchmark: pd.DataFrame) -> pd.DataFrame:
    """
    Reindex asset to the benchmark's trading dates, forward-filling gaps,
    then return a combined DataFrame with both Close series.
    """
    bm_idx = benchmark.index

    # Reindex asset to benchmark dates; fill forward (use last traded price)
    asset_aligned = asset.reindex(bm_idx).ffill()

    combined = pd.DataFrame({
        "asset_close": asset_aligned["Close"],
        "bm_close":    benchmark["Close"],
    }).dropna()   # drop any remaining NaNs (e.g. asset started after benchmark window)

    return combined

# ---------------------------------------------------------------------------
# EXCEL HELPERS
# ---------------------------------------------------------------------------

def col(n: int) -> str:
    """1-based column index → Excel letter(s)."""
    return get_column_letter(n)

def write_header_row(ws, headers: list[str], row: int = 1):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font   = HEADER_FONT
        cell.fill   = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER
        ws.column_dimensions[col(c)].width = max(14, len(h) + 2)

def style_data_cell(cell, font=None, num_fmt=None):
    cell.font   = font or DATA_FONT
    cell.border = BORDER
    cell.alignment = Alignment(horizontal="right")
    if num_fmt:
        cell.number_format = num_fmt

# ---------------------------------------------------------------------------
# FX RATES SHEET
# ---------------------------------------------------------------------------

def create_fx_sheet(wb: Workbook):
    ws = wb.create_sheet(FX_SHEET, 0)   # insert as first sheet
    ws.sheet_view.showGridLines = False

    # Title
    ws["A1"] = "FX Rate Assumptions (vs HKD)"
    ws["A1"].font = Font(name="Arial", bold=True, size=11)
    ws.merge_cells("A1:D1")

    headers = ["Currency Pair", "Rate (to HKD)", "Last Updated"]
    write_header_row(ws, headers, row=2)

    today_str = datetime.today().strftime("%d-%b-%Y")

    rows = [
        ("USD/HKD", 7.78,  today_str),
        ("CNH/HKD", 1.071, today_str),
    ]
    for r, (pair, rate, date) in enumerate(rows, 3):
        ws.cell(row=r, column=1, value=pair).font  = DATA_FONT
        rate_cell = ws.cell(row=r, column=2, value=rate)
        rate_cell.font         = INPUT_FONT 
        rate_cell.number_format = "0.0000"
        ws.cell(row=r, column=3, value=date).font   = DATA_FONT   
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = BORDER

    # Yellow background on the editable rate cells to flag them
    yellow = PatternFill("solid", start_color="FFFFFF00")
    ws["B3"].fill = yellow
    ws["B4"].fill = yellow

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 50

    return ws

# ---------------------------------------------------------------------------
# FX formula string for a given currency
# ---------------------------------------------------------------------------

def fx_formula(currency: str) -> str:
    """
    Return an Excel formula that looks up the appropriate FX rate from the
    'FX Rates' sheet.  HKD stocks use 1 (no conversion needed).
    """
    if currency == "USD":
        return f"='{FX_SHEET}'!$B${FX_ROW_USD}"
    elif currency == "CNH":
        return f"='{FX_SHEET}'!$B${FX_ROW_CNH}"
    else:
        return "=1"  

# ---------------------------------------------------------------------------
# WRITE STOCK SHEET
# ---------------------------------------------------------------------------

def write_stock_sheet(wb: Workbook, ticker: str, orig_currency: str,
                      combined: pd.DataFrame, beta_result: dict):
    """
    Write one stock sheet.  Columns:
      A  Date
      B  {ticker} Close (Local CCY)
      C  FX Rate (→ HKD)          ← Excel formula referencing FX Rates sheet
      D  {ticker} Close (HKD)     ← Excel formula: =B*C
      E  2800.HK Close (HKD)
      F  {ticker} Return (HKD)    ← Excel formula: % change on column D
      G  2800.HK Return           ← Excel formula: % change on column E
    """
    # Resolve final currency after remap (for FX formula)
    ticker, resolved_ccy = TICKER_REMAP.get((ticker, orig_currency), (ticker, orig_currency))
    resolved_ccy = orig_currency

    sheet_name = ticker[:31].replace("/", "_").replace("\\", "_")
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    # --- Headers ---
    headers = [
        "Date",
        f"{ticker} Close ({resolved_ccy})",
        "FX Rate (to HKD)",
        f"{ticker} Close (HKD)",
        "2800.HK Close (HKD)",
        f"{ticker} Return (HKD)",
        "2800.HK Return",
    ]
    write_header_row(ws, headers, row=1)
    ws.column_dimensions["A"].width = 13

    data_rows = combined.reset_index()  # columns: Date, asset_close, bm_close

    for r_idx, data_row in enumerate(data_rows.itertuples(index=False), start=2):
        date_val  = data_row.Date if hasattr(data_row, 'Date') else data_row[0]
        asset_val = data_row.asset_close if hasattr(data_row, 'asset_close') else data_row[1]
        bm_val    = data_row.bm_close if hasattr(data_row, 'bm_close') else data_row[2]

        # Col A: Date
        c = ws.cell(row=r_idx, column=1, value=date_val)
        c.number_format = "DD-MMM-YY"; style_data_cell(c, DATA_FONT)

        # Col B: local close price (hardcoded data value)
        c = ws.cell(row=r_idx, column=2, value=round(asset_val, 4))
        style_data_cell(c, DATA_FONT, "0.0000")

        # Col C: FX rate — Excel formula linking to FX Rates sheet (green = cross-sheet link)
        c = ws.cell(row=r_idx, column=3, value=fx_formula(resolved_ccy))
        style_data_cell(c, LINK_FONT, "0.0000")

        # Col D: Close in HKD = B * C
        c = ws.cell(row=r_idx, column=4, value=f"={col(2)}{r_idx}*{col(3)}{r_idx}")
        style_data_cell(c, FORMULA_FONT, "0.0000")

        # Col E: benchmark close (hardcoded data value)
        c = ws.cell(row=r_idx, column=5, value=round(bm_val, 4))
        style_data_cell(c, DATA_FONT, "0.0000")

        # Col F: asset return (HKD) — pct change on col D
        if r_idx == 2:
            c = ws.cell(row=r_idx, column=6, value="—")
            style_data_cell(c, DATA_FONT)
        else:
            formula = f"=({col(4)}{r_idx}-{col(4)}{r_idx-1})/{col(4)}{r_idx-1}"
            c = ws.cell(row=r_idx, column=6, value=formula)
            style_data_cell(c, FORMULA_FONT, "0.00%")

        # Col G: benchmark return — pct change on col E
        if r_idx == 2:
            c = ws.cell(row=r_idx, column=7, value="—")
            style_data_cell(c, DATA_FONT)
        else:
            formula = f"=({col(5)}{r_idx}-{col(5)}{r_idx-1})/{col(5)}{r_idx-1}"
            c = ws.cell(row=r_idx, column=7, value=formula)
            style_data_cell(c, FORMULA_FONT, "0.00%")

    # --- Beta summary block (3 rows below data) ---
    last_data_row  = len(data_rows) + 1   # 1-indexed last row with data
    summary_start  = last_data_row + 3

    # Returns start from row 3 (row 2 is "—"), so last return is at last_data_row
    last_return_row  = last_data_row
    first_return_row = last_return_row - 44  

    labels = [
        ("45-Day Covariance (HKD returns):",  beta_result["covariance"]),
        ("45-Day Benchmark Variance:",         beta_result["variance"]),
        ("FINAL 45-DAY BETA:",                 beta_result["beta"]),
    ]

    # Build native Excel beta formulas if we have enough data
    if isinstance(beta_result["beta"], float):
        f_ret  = first_return_row
        l_ret  = last_return_row
        cov_formula = (
            f"=_xlfn.COVARIANCE.S({col(6)}{f_ret}:{col(6)}{l_ret},"
            f"{col(7)}{f_ret}:{col(7)}{l_ret})"
        )
        var_formula  = f"=_xlfn.VAR.S({col(7)}{f_ret}:{col(7)}{l_ret})"
        beta_formula = f"={col(2)}{summary_start}/{col(2)}{summary_start+1}"

        excel_formulas = [cov_formula, var_formula, beta_formula]
        label_col, val_col = 1, 2
    else:
        excel_formulas = [beta_result["covariance"],
                          beta_result["variance"],
                          beta_result["beta"]]

    for i, (label, _) in enumerate(labels):
        sr = summary_start + i
        lc = ws.cell(row=sr, column=1, value=label)
        lc.font = Font(name="Arial", bold=(i == 2), size=9)
        lc.border = BORDER

        vc = ws.cell(row=sr, column=2, value=excel_formulas[i])
        vc.font   = Font(name="Arial", bold=(i == 2), size=9,
                         color=BLACK if isinstance(excel_formulas[i], str)
                               and excel_formulas[i].startswith("=") else BLACK)
        vc.border = BORDER
        if i == 2:  # Beta row: highlight
            gold = PatternFill("solid", start_color="FFFFF2CC")
            lc.fill = gold; vc.fill = gold
        if isinstance(excel_formulas[i], float):
            vc.number_format = "0.0000"

    # Freeze panes below header
    ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
# SUMMARY SHEET
# ---------------------------------------------------------------------------

def write_summary_sheet(wb: Workbook, holdings: pd.DataFrame, beta_map: dict):
    ws = wb.create_sheet("Summary", 0)
    wb.active = ws
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Beta Calculation Summary"
    ws["A1"].font = Font(name="Arial", bold=True, size=13)
    ws.merge_cells("A1:F1")

    ws["A2"] = f"Benchmark: 2800.HK (Hang Seng ETF) | Generated: {datetime.today().strftime('%d-%b-%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="FF555555")
    ws.merge_cells("A2:F2")

    headers = ["Ticker", "Currency", "Weighting", "Aligned Days", "45-Day Beta (HKD)", "Status"]
    write_header_row(ws, headers, row=4)

    widths = [12, 10, 12, 14, 20, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[col(i)].width = w

    for r_idx, (_, row) in enumerate(holdings.iterrows(), start=5):
        ticker   = str(row["Ticker"]).strip()
        currency = str(row["Local Currency"]).strip()
        ticker, currency = TICKER_REMAP.get((ticker, currency), (ticker, currency))
        weight   = row["Weighting"]
        result   = beta_map.get(ticker, {})
        n_days   = result.get("n_days", "—")
        status   = result.get("status", "Failed")
        if status == "OK" and "excel_link" in result:
            beta_val = result["excel_link"]
        else:
            beta_val = result.get("beta", "N/A")
        

        vals = [ticker, currency, weight, n_days, beta_val, status]
        fmts = [None, None, "0.00%", "0", "0.0000", None]
        fonts = [DATA_FONT, DATA_FONT, DATA_FONT, DATA_FONT,
                 Font(name="Arial", size=9, bold=True), DATA_FONT]

        for c_idx, (v, fmt, fnt) in enumerate(zip(vals, fmts, fonts), 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.font   = fnt
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center" if c_idx > 1 else "left")
            if fmt:
                if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
                    cell.number_format = fmt

        # Colour-code status
        if status == "OK":
            ws.cell(row=r_idx, column=6).fill = PatternFill("solid", start_color="FFE2EFDA")
        elif status == "Failed":
            ws.cell(row=r_idx, column=6).fill = PatternFill("solid", start_color="FFFFC7CE")
        else:
            ws.cell(row=r_idx, column=6).fill = PatternFill("solid", start_color="FFFFEB9C")

    ws.freeze_panes = "A5"

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

holdings = pd.read_excel(input_file)
holdings["Weighting"] = pd.to_numeric(holdings["Weighting"], errors="coerce")

print("\nFetching Benchmark: 2800.HK (Hang Seng ETF)...")
benchmark_raw = get_hk_prices("2800")
if benchmark_raw is None or benchmark_raw.empty:
    raise RuntimeError("Could not fetch benchmark (2800.HK). Cannot continue.")

benchmark = benchmark_raw.tail(60)

wb = Workbook()
del wb[wb.sheetnames[0]]   # remove default empty sheet

create_fx_sheet(wb)

beta_map = {}

for _, row in holdings.iterrows():
    ticker   = str(row["Ticker"]).strip()
    currency = str(row["Local Currency"]).strip()

    print(f"Processing: {ticker} ({currency})...")
    asset = get_prices(ticker, currency)

    if asset is None or asset.empty:
        print(f"  -> No data for {ticker}, skipping.")
        beta_map[ticker] = {"beta": "N/A", "status": "Failed", "n_days": 0}
        continue

    # --- Trading-day alignment ---
    # Reindex asset to benchmark's HK trading calendar, forward-filling
    # any days where the asset didn't trade (different holiday calendar).
    combined = align_to_benchmark(asset, benchmark)

    if len(combined) < 46:  
        msg = f"Insufficient aligned data ({len(combined)} days)"
        print(f"  -> {msg}")
        beta_map[ticker] = {"beta": msg, "status": "Insufficient data",
                             "n_days": len(combined)}
        continue

    # --- Beta calculation (Python side, for Summary sheet) ---
    # Compute returns from the aligned price series
    returns = combined.pct_change().dropna()
    last  = returns.tail(45)
    cov      = last[["asset_close", "bm_close"]].cov().iloc[0, 1]
    var      = last["bm_close"].var()
    beta_val = cov / var if var != 0 else float("nan")

    ticker, currency = TICKER_REMAP.get((ticker, currency), (ticker, currency))
    sheet_name = ticker[:31].replace("/", "_").replace("\\", "_")
    beta_result = {
        "covariance": round(cov, 8),
        "variance":   round(var, 8),
        "beta":       round(beta_val, 6),
        "n_days":     len(combined),
        "status":     "OK",
        "excel_link": f"='{sheet_name}'!B{len(combined) + 6}"
    }
    beta_map[ticker] = beta_result

    write_stock_sheet(wb, ticker, currency, combined, beta_result)

# Summary sheet goes first
write_summary_sheet(wb, holdings, beta_map)

# Move Summary to front, FX Rates second
sheets = wb.sheetnames
wb.move_sheet("Summary",  offset=-len(sheets))
wb.move_sheet(FX_SHEET,   offset=-(len(sheets) - 1))

output_file = "Beta_Calculation_Audit.xlsx"
wb.save(output_file)
print(f"\nDone. Audit file saved: {output_file}")