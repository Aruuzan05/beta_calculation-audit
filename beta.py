import akshare as ak
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import warnings

input_file = "Holdings.xlsx"
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

TICKER_REMAP = {
    ("PRX", "EUR"): ("PROSY", "USD"),
}

FX_SHEET    = "FX Rates"
FX_ROW_USD  = 3
FX_ROW_CNH  = 4

BLUE  = "FF0000FF"
BLACK = "FF000000"
GREEN = "FF008000"

HEADER_FILL = PatternFill("solid", start_color="FF1F4E79")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFFFF", size=9)
DATA_FONT   = Font(name="Arial", size=9)
INPUT_FONT  = Font(name="Arial", size=9, color=BLUE)
FORMULA_FONT= Font(name="Arial", size=9, color=BLACK)
LINK_FONT   = Font(name="Arial", size=9, color=GREEN)

THIN   = Side(style="thin", color="FFD9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---------------------------------------------------------------------------
# GICS SECTOR MAP  (used as fallback / override for known tickers)
# AkShare does not expose GICS directly; we derive from industry strings.
# ---------------------------------------------------------------------------

GICS_KEYWORD_MAP = [
    # (substring_lower, GICS Sector label)
    ("internet",           "Communication Services"),
    ("software",           "Information Technology"),
    ("semiconductor",      "Information Technology"),
    ("technology",         "Information Technology"),
    ("tech",               "Information Technology"),
    ("bank",               "Financials"),
    ("financ",             "Financials"),
    ("insurance",          "Financials"),
    ("securities",         "Financials"),
    ("real estate",        "Real Estate"),
    ("property",           "Real Estate"),
    ("energy",             "Energy"),
    ("oil",                "Energy"),
    ("gas",                "Energy"),
    ("health",             "Health Care"),
    ("pharma",             "Health Care"),
    ("biotech",            "Health Care"),
    ("medical",            "Health Care"),
    ("consumer",           "Consumer Discretionary"),
    ("retail",             "Consumer Discretionary"),
    ("auto",               "Consumer Discretionary"),
    ("food",               "Consumer Staples"),
    ("beverage",           "Consumer Staples"),
    ("tobacco",            "Consumer Staples"),
    ("material",           "Materials"),
    ("chemical",           "Materials"),
    ("mining",             "Materials"),
    ("metal",              "Materials"),
    ("industrial",         "Industrials"),
    ("transport",          "Industrials"),
    ("aerospace",          "Industrials"),
    ("utility",            "Utilities"),
    ("electric",           "Utilities"),
    ("telecom",            "Communication Services"),
    ("media",              "Communication Services"),
    ("entertainment",      "Communication Services"),
]


def classify_gics(industry_str: str) -> str:
    if not industry_str or pd.isna(industry_str):
        return "Unknown"
    s = str(industry_str).lower()
    for keyword, sector in GICS_KEYWORD_MAP:
        if keyword in s:
            return sector
    return "Unknown"


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
# PRICE FETCHERS
# ---------------------------------------------------------------------------

def _normalise(hist: pd.DataFrame, date_col: str, close_col: str):
    if hist is None or hist.empty:
        return None
    hist = hist.rename(columns={date_col: "Date", close_col: "Close"})
    hist["Date"] = pd.to_datetime(hist["Date"])
    hist = hist.set_index("Date").sort_index()
    hist["Close"] = pd.to_numeric(hist["Close"], errors="coerce")
    return hist[["Close"]]

def get_hk_prices(ticker: str):
    try:
        return _normalise(ak.stock_hk_daily(symbol=format_hk_ticker(ticker)), "date", "close")
    except Exception as e:
        print(f"  Sina HK failed for {ticker}: {e}"); return None

def get_a_share_prices(ticker: str):
    try:
        return _normalise(ak.stock_zh_a_daily(symbol=format_a_ticker(ticker)), "date", "close")
    except Exception as e:
        print(f"  Sina A-share failed for {ticker}: {e}"); return None

def get_us_prices(ticker: str):
    try:
        return _normalise(ak.stock_us_daily(symbol=str(ticker).strip()), "date", "close")
    except Exception as e:
        print(f"  Sina US failed for {ticker}: {e}"); return None

def get_prices(ticker: str, currency: str):
    ticker, currency = TICKER_REMAP.get((ticker, currency), (ticker, currency))
    if currency == "CNH" and len(ticker) == 6:
        return get_a_share_prices(ticker)
    elif currency in ['HKD', 'CNH']:
        return get_hk_prices(ticker)
    elif currency in ['USD', 'EUR']:
        return get_us_prices(ticker)
    else:
        print(f"  {ticker} ({currency}): no data source — skipping."); return None

# ---------------------------------------------------------------------------
# FUNDAMENTAL METRICS FETCHERS
# ---------------------------------------------------------------------------

def get_hk_fundamentals(ticker: str) -> dict:
    """Fetch PE, PB, dividend yield, market cap, industry for HK stocks."""
    result = {"pe": None, "pb": None, "div_yield": None, "mktcap_local": None, "industry": None}
    try:
        sym = format_hk_ticker(ticker)
        # Individual stock info
        df = ak.stock_hk_spot()
        # columns vary by version; try common names
        col_map = {c.lower(): c for c in df.columns}
        # filter for this ticker
        code_col = col_map.get("代码") or col_map.get("股票代码") or col_map.get("code")
        if code_col:
            row = df[df[code_col].astype(str).str.zfill(5) == sym]
            if not row.empty:
                row = row.iloc[0]
                # Market cap (usually in 亿 HKD — convert to full HKD)
                for key in ["总市值", "市值", "总市值(亿港元)", "市值(亿港元)"]:
                    if key in col_map.values() and pd.notna(row.get(key)):
                        result["mktcap_local"] = float(row[key]) * 1e8
                        break
                for key in ["市盈率(动态)", "市盈率", "pe", "PE"]:
                    if key in col_map.values() and pd.notna(row.get(key)):
                        result["pe"] = float(row[key]) if row[key] not in ["—", "-", ""] else None
                        break
                for key in ["市净率", "pb", "PB"]:
                    if key in col_map.values() and pd.notna(row.get(key)):
                        result["pb"] = float(row[key]) if row[key] not in ["—", "-", ""] else None
                        break
    except Exception as e:
        print(f"  HK spot em failed for {ticker}: {e}")

    # Try dedicated PE/PB source
    try:
        info = ak.stock_hk_valuation_baidu(symbol=format_hk_ticker(ticker), indicator="市盈率")
        if info is not None and not info.empty:
            latest = info.iloc[-1]
            val = pd.to_numeric(latest.iloc[-1], errors="coerce")
            if pd.notna(val):
                result["pe"] = float(val)
    except Exception:
        pass

    try:
        info = ak.stock_hk_valuation_baidu(symbol=format_hk_ticker(ticker), indicator="市净率")
        if info is not None and not info.empty:
            latest = info.iloc[-1]
            val = pd.to_numeric(latest.iloc[-1], errors="coerce")
            if pd.notna(val):
                result["pb"] = float(val)
    except Exception:
        pass

    # Industry classification
    try:
        profile = ak.stock_individual_info_em(symbol=format_hk_ticker(ticker), timeout=10)
        if profile is not None and not profile.empty:
            # profile is usually a 2-col df with 'item' / 'value'
            items = dict(zip(profile.iloc[:, 0], profile.iloc[:, 1]))
            for k in ["所属行业", "行业", "industry", "Industry"]:
                if k in items and items[k]:
                    result["industry"] = str(items[k])
                    break
    except Exception:
        pass

    return result


def get_a_share_fundamentals(ticker: str) -> dict:
    """Fetch fundamentals for A-share stocks via AkShare."""
    result = {"pe": None, "pb": None, "div_yield": None, "mktcap_local": None, "industry": None}
    sym = format_a_ticker(ticker)
    code = ticker  # bare code like 600519

    try:
        # East Money spot data (most reliable single source)
        df = ak.stock_zh_a_spot_em()
        col_map = {c: c for c in df.columns}
        code_col = "代码" if "代码" in df.columns else None
        if code_col:
            row = df[df[code_col].astype(str) == str(code)]
            if not row.empty:
                r = row.iloc[0]
                for k in ["总市值", "市值"]:
                    if k in r and pd.notna(r[k]):
                        result["mktcap_local"] = float(r[k])
                        break
                for k in ["市盈率-动态", "市盈率(动态)", "市盈率"]:
                    if k in r and pd.notna(r[k]):
                        try:
                            result["pe"] = float(r[k])
                        except (ValueError, TypeError):
                            pass
                        break
                for k in ["市净率"]:
                    if k in r and pd.notna(r[k]):
                        try:
                            result["pb"] = float(r[k])
                        except (ValueError, TypeError):
                            pass
                        break
    except Exception as e:
        print(f"  A-share spot em failed for {ticker}: {e}")

    # Industry
    try:
        profile = ak.stock_individual_info_em(symbol=str(code), timeout=10)
        if profile is not None and not profile.empty:
            items = dict(zip(profile.iloc[:, 0], profile.iloc[:, 1]))
            for k in ["所属行业", "行业"]:
                if k in items and items[k]:
                    result["industry"] = str(items[k])
                    break
    except Exception:
        pass

    # Dividend yield from financial data
    try:
        div_df = ak.stock_financial_abstract_ths(symbol=str(code), indicator="按年度")
        if div_df is not None and not div_df.empty:
            for col_name in div_df.columns:
                if "股息" in col_name or "分红" in col_name or "dividend" in col_name.lower():
                    val = pd.to_numeric(div_df.iloc[0][col_name], errors="coerce")
                    if pd.notna(val):
                        result["div_yield"] = float(val) / 100 if val > 1 else float(val)
                    break
    except Exception:
        pass

    return result


def get_us_fundamentals(ticker: str) -> dict:
    """Fetch fundamentals for US stocks."""
    result = {"pe": None, "pb": None, "div_yield": None, "mktcap_local": None, "industry": None}
    try:
        df = ak.stock_us_spot()
        if df is not None and not df.empty:
            col_names = list(df.columns)
            # Find code column
            for cc in ["代码", "股票代码", "Code"]:
                if cc in col_names:
                    match = df[df[cc].astype(str).str.upper() == ticker.upper()]
                    if not match.empty:
                        r = match.iloc[0]
                        for k in ["总市值", "市值"]:
                            if k in r and pd.notna(r[k]):
                                result["mktcap_local"] = float(r[k])
                                break
                        for k in ["市盈率(TTM)", "市盈率-动态", "市盈率"]:
                            if k in r and pd.notna(r[k]):
                                try:
                                    result["pe"] = float(r[k])
                                except (ValueError, TypeError):
                                    pass
                                break
                        for k in ["市净率"]:
                            if k in r and pd.notna(r[k]):
                                try:
                                    result["pb"] = float(r[k])
                                except (ValueError, TypeError):
                                    pass
                                break
                    break
    except Exception as e:
        print(f"  US spot em failed for {ticker}: {e}")

    # Industry from EastMoney stock profile
    try:
        profile = ak.stock_us_profile(symbol=ticker.upper())
        if profile is not None and not profile.empty:
            for col_name in profile.columns:
                if "industry" in col_name.lower() or "sector" in col_name.lower() or "行业" in col_name:
                    val = profile.iloc[0][col_name]
                    if val and str(val) not in ["nan", "None", ""]:
                        result["industry"] = str(val)
                    break
    except Exception:
        pass

    return result


def get_fundamentals(ticker: str, currency: str) -> dict:
    """Route to correct fundamentals fetcher and normalise output."""
    orig_ticker = ticker
    ticker, currency = TICKER_REMAP.get((ticker, currency), (ticker, currency))

    print(f"  Fetching fundamentals for {ticker} ({currency})...")

    if currency == "CNH" and len(ticker) == 6:
        raw = get_a_share_fundamentals(ticker)
    elif currency in ["HKD", "CNH"]:
        raw = get_hk_fundamentals(ticker)
    elif currency in ["USD", "EUR"]:
        raw = get_us_fundamentals(ticker)
    else:
        raw = {"pe": None, "pb": None, "div_yield": None, "mktcap_local": None, "industry": None}

    # Sanitise extreme / clearly wrong values
    if raw["pe"] is not None and (raw["pe"] < 0 or raw["pe"] > 2000):
        raw["pe"] = None
    if raw["pb"] is not None and (raw["pb"] < 0 or raw["pb"] > 500):
        raw["pb"] = None

    # Derive GICS from industry string
    raw["gics_sector"] = classify_gics(raw.get("industry"))

    return raw

# ---------------------------------------------------------------------------
# TRADING-DAY ALIGNMENT
# ---------------------------------------------------------------------------

def align_to_benchmark(asset: pd.DataFrame, benchmark: pd.DataFrame) -> pd.DataFrame:
    bm_idx = benchmark.index
    asset_aligned = asset.reindex(bm_idx).ffill()
    combined = pd.DataFrame({
        "asset_close": asset_aligned["Close"],
        "bm_close":    benchmark["Close"],
    }).dropna()
    return combined

# ---------------------------------------------------------------------------
# EXCEL HELPERS
# ---------------------------------------------------------------------------

def col(n: int) -> str:
    return get_column_letter(n)

def write_header_row(ws, headers: list, row: int = 1):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border    = BORDER
        ws.column_dimensions[col(c)].width = max(14, len(h) + 2)

def style_data_cell(cell, font=None, num_fmt=None, align="right"):
    cell.font      = font or DATA_FONT
    cell.border    = BORDER
    cell.alignment = Alignment(horizontal=align)
    if num_fmt:
        cell.number_format = num_fmt

def na_safe(val):
    """Return val if it is a valid number, else None (so Excel shows blank)."""
    if val is None:
        return None
    try:
        f = float(val)
        return None if (np.isnan(f) or np.isinf(f)) else f
    except (TypeError, ValueError):
        return None

# ---------------------------------------------------------------------------
# FX RATES SHEET
# ---------------------------------------------------------------------------

def create_fx_sheet(wb: Workbook):
    ws = wb.create_sheet(FX_SHEET, 0)
    ws.sheet_view.showGridLines = False

    ws["A1"] = "FX Rate Assumptions (vs HKD)"
    ws["A1"].font = Font(name="Arial", bold=True, size=11)
    ws.merge_cells("A1:D1")

    write_header_row(ws, ["Currency Pair", "Rate (to HKD)", "Last Updated"], row=2)

    today_str = datetime.today().strftime("%d-%b-%Y")
    rows = [("USD/HKD", 7.78, today_str), ("CNH/HKD", 1.071, today_str)]
    for r, (pair, rate, date) in enumerate(rows, 3):
        ws.cell(row=r, column=1, value=pair).font = DATA_FONT
        rc = ws.cell(row=r, column=2, value=rate)
        rc.font = INPUT_FONT; rc.number_format = "0.0000"
        ws.cell(row=r, column=3, value=date).font = DATA_FONT
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = BORDER

    yellow = PatternFill("solid", start_color="FFFFFF00")
    ws["B3"].fill = yellow
    ws["B4"].fill = yellow

    for c_ltr, w in [("A", 16), ("B", 16), ("C", 16), ("D", 50)]:
        ws.column_dimensions[c_ltr].width = w
    return ws

def fx_formula(currency: str) -> str:
    if currency == "USD":
        return f"='{FX_SHEET}'!$B${FX_ROW_USD}"
    elif currency == "CNH":
        return f"='{FX_SHEET}'!$B${FX_ROW_CNH}"
    else:
        return "=1"

# ---------------------------------------------------------------------------
# WRITE STOCK SHEET  (unchanged from original)
# ---------------------------------------------------------------------------

def write_stock_sheet(wb: Workbook, ticker: str, orig_currency: str,
                      combined: pd.DataFrame, beta_result: dict):
    ticker, resolved_ccy = TICKER_REMAP.get((ticker, orig_currency), (ticker, orig_currency))
    resolved_ccy = orig_currency

    sheet_name = ticker[:31].replace("/", "_").replace("\\", "_")
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

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

    data_rows = combined.reset_index()

    for r_idx, data_row in enumerate(data_rows.itertuples(index=False), start=2):
        date_val  = data_row[0]
        asset_val = data_row[1]
        bm_val    = data_row[2]

        c = ws.cell(row=r_idx, column=1, value=date_val)
        c.number_format = "DD-MMM-YY"; style_data_cell(c, DATA_FONT)

        c = ws.cell(row=r_idx, column=2, value=round(asset_val, 4))
        style_data_cell(c, DATA_FONT, "0.0000")

        c = ws.cell(row=r_idx, column=3, value=fx_formula(resolved_ccy))
        style_data_cell(c, LINK_FONT, "0.0000")

        c = ws.cell(row=r_idx, column=4, value=f"={col(2)}{r_idx}*{col(3)}{r_idx}")
        style_data_cell(c, FORMULA_FONT, "0.0000")

        c = ws.cell(row=r_idx, column=5, value=round(bm_val, 4))
        style_data_cell(c, DATA_FONT, "0.0000")

        if r_idx == 2:
            ws.cell(row=r_idx, column=6, value="—")
            style_data_cell(ws.cell(row=r_idx, column=6), DATA_FONT)
        else:
            c = ws.cell(row=r_idx, column=6,
                        value=f"=({col(4)}{r_idx}-{col(4)}{r_idx-1})/{col(4)}{r_idx-1}")
            style_data_cell(c, FORMULA_FONT, "0.00%")

        if r_idx == 2:
            ws.cell(row=r_idx, column=7, value="—")
            style_data_cell(ws.cell(row=r_idx, column=7), DATA_FONT)
        else:
            c = ws.cell(row=r_idx, column=7,
                        value=f"=({col(5)}{r_idx}-{col(5)}{r_idx-1})/{col(5)}{r_idx-1}")
            style_data_cell(c, FORMULA_FONT, "0.00%")

    last_data_row = len(data_rows) + 1
    summary_start = last_data_row + 3
    last_return_row  = last_data_row
    first_return_row = last_return_row - 44

    if isinstance(beta_result["beta"], float):
        f_ret = first_return_row
        l_ret = last_return_row
        excel_formulas = [
            f"=_xlfn.COVARIANCE.S({col(6)}{f_ret}:{col(6)}{l_ret},{col(7)}{f_ret}:{col(7)}{l_ret})",
            f"=_xlfn.VAR.S({col(7)}{f_ret}:{col(7)}{l_ret})",
            f"={col(2)}{summary_start}/{col(2)}{summary_start+1}",
        ]
    else:
        excel_formulas = [beta_result["covariance"], beta_result["variance"], beta_result["beta"]]

    labels = [
        ("45-Day Covariance (HKD returns):", beta_result["covariance"]),
        ("45-Day Benchmark Variance:",        beta_result["variance"]),
        ("FINAL 45-DAY BETA:",                beta_result["beta"]),
    ]
    gold = PatternFill("solid", start_color="FFFFF2CC")
    for i, (label, _) in enumerate(labels):
        sr = summary_start + i
        lc = ws.cell(row=sr, column=1, value=label)
        lc.font = Font(name="Arial", bold=(i == 2), size=9)
        lc.border = BORDER
        vc = ws.cell(row=sr, column=2, value=excel_formulas[i])
        vc.font   = Font(name="Arial", bold=(i == 2), size=9, color=BLACK)
        vc.border = BORDER
        if i == 2:
            lc.fill = gold; vc.fill = gold
        if isinstance(excel_formulas[i], float):
            vc.number_format = "0.0000"

    ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
# SUMMARY SHEET  (enhanced with fundamentals columns)
# ---------------------------------------------------------------------------

def write_summary_sheet(wb: Workbook, holdings: pd.DataFrame,
                        beta_map: dict, fund_map: dict, fx_usd: float, fx_cnh: float):
    ws = wb.create_sheet("Summary", 0)
    wb.active = ws
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Portfolio Analytics Summary"
    ws["A1"].font = Font(name="Arial", bold=True, size=13)
    ws.merge_cells("A1:K1")

    ws["A2"] = (f"Benchmark: 2800.HK (Hang Seng ETF) | "
                f"Generated: {datetime.today().strftime('%d-%b-%Y %H:%M')}")
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="FF555555")
    ws.merge_cells("A2:K2")

    headers = [
        "Ticker", "Currency", "Side", "Weighting",
        "45-Day Beta (HKD)", "Status",
        "P/E Ratio", "P/B Ratio", "Div Yield",
        "Mkt Cap (HKD bn)", "GICS Sector",
    ]
    write_header_row(ws, headers, row=4)

    col_widths = [12, 10, 8, 12, 20, 18, 12, 12, 12, 18, 28]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[col(i)].width = w

    DATA_ROW_START = 5

    for r_idx, (_, row) in enumerate(holdings.iterrows(), start=DATA_ROW_START):
        ticker   = str(row["Ticker"]).strip()
        currency = str(row["Local Currency"]).strip()
        side     = str(row.get("Side", "Long")).strip() if "Side" in row.index else "Long"
        weight   = row["Weighting"]

        ticker_r, currency_r = TICKER_REMAP.get((ticker, currency), (ticker, currency))
        result = beta_map.get(ticker_r, {})
        fund   = fund_map.get(ticker_r, {})

        n_days = result.get("n_days", "—")
        status = result.get("status", "Failed")
        beta_val = result.get("excel_link", result.get("beta", "N/A")) \
            if status == "OK" else result.get("beta", "N/A")

        # Convert mktcap to HKD using FX, then to billions
        mktcap_local = fund.get("mktcap_local")
        mktcap_hkd_bn = None
        if mktcap_local is not None:
            fx = fx_usd if currency_r == "USD" else (fx_cnh if currency_r == "CNH" else 1.0)
            mktcap_hkd_bn = mktcap_local * fx / 1e9

        pe        = na_safe(fund.get("pe"))
        pb        = na_safe(fund.get("pb"))
        div_yield = na_safe(fund.get("div_yield"))
        gics      = fund.get("gics_sector", "Unknown")

        vals  = [ticker_r, currency_r, side, weight, beta_val, status,
                 pe, pb, div_yield, mktcap_hkd_bn, gics]
        fmts  = [None, None, None, "0.00%", "0.0000", None,
                 "0.0x", "0.0x", "0.00%", "#,##0.0", None]
        aligns = ["left", "center", "center", "center", "center", "center",
                  "right", "right", "right", "right", "left"]
        bold_cols = {5}  # Beta column bold

        for c_idx, (v, fmt, aln) in enumerate(zip(vals, fmts, aligns), 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.font   = Font(name="Arial", size=9, bold=(c_idx in bold_cols))
            cell.border = BORDER
            cell.alignment = Alignment(horizontal=aln)
            if fmt and v is not None:
                if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
                    cell.number_format = fmt

        # Status colour
        if status == "OK":
            ws.cell(row=r_idx, column=6).fill = PatternFill("solid", start_color="FFE2EFDA")
        elif status == "Failed":
            ws.cell(row=r_idx, column=6).fill = PatternFill("solid", start_color="FFFFC7CE")
        else:
            ws.cell(row=r_idx, column=6).fill = PatternFill("solid", start_color="FFFFEB9C")

        # Side colour (Long = light green tint, Short = light red tint)
        side_cell = ws.cell(row=r_idx, column=3)
        if side.lower() == "long":
            side_cell.fill = PatternFill("solid", start_color="FFE2EFDA")
        elif side.lower() == "short":
            side_cell.fill = PatternFill("solid", start_color="FFFFC7CE")

    last_data_row = DATA_ROW_START + len(holdings) - 1

    # -----------------------------------------------------------------------
    # WEIGHTED AVERAGE ANALYTICS BLOCK  (Long side only)
    # -----------------------------------------------------------------------
    wa_start = last_data_row + 3
    ws.cell(row=wa_start, column=1, value="LONG SIDE — WEIGHTED AVERAGES").font = \
        Font(name="Arial", bold=True, size=10)
    ws.merge_cells(f"A{wa_start}:K{wa_start}")

    wa_headers = ["Metric", "Value", "Note"]
    for ci, h in enumerate(wa_headers, 1):
        cell = ws.cell(row=wa_start + 1, column=ci, value=h)
        cell.font   = HEADER_FONT
        cell.fill   = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    # Build weighted averages in Python (data already in hand)
    long_rows = []
    for _, row in holdings.iterrows():
        ticker   = str(row["Ticker"]).strip()
        currency = str(row["Local Currency"]).strip()
        side     = str(row.get("Side", "Long")).strip() if "Side" in row.index else "Long"
        if side.lower() != "long":
            continue
        ticker_r, _ = TICKER_REMAP.get((ticker, currency), (ticker, currency))
        fund   = fund_map.get(ticker_r, {})
        weight = float(row["Weighting"]) if pd.notna(row["Weighting"]) else 0.0
        pe     = na_safe(fund.get("pe"))
        pb     = na_safe(fund.get("pb"))
        long_rows.append({"weight": weight, "pe": pe, "pb": pb})

    total_w_pe = sum(r["weight"] for r in long_rows if r["pe"] is not None)
    total_w_pb = sum(r["weight"] for r in long_rows if r["pb"] is not None)

    wtd_pe = (sum(r["weight"] * r["pe"] for r in long_rows if r["pe"] is not None) / total_w_pe
              if total_w_pe > 0 else None)
    wtd_pb = (sum(r["weight"] * r["pb"] for r in long_rows if r["pb"] is not None) / total_w_pb
              if total_w_pb > 0 else None)

    wa_data = [
        ("Weighted Avg P/E (Long)",  wtd_pe,  "Weight-adj; excludes positions with no PE data"),
        ("Weighted Avg P/B (Long)",  wtd_pb,  "Weight-adj; excludes positions with no PB data"),
    ]
    for i, (label, value, note) in enumerate(wa_data):
        r = wa_start + 2 + i
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = Font(name="Arial", bold=True, size=9); lc.border = BORDER
        vc = ws.cell(row=r, column=2, value=value)
        vc.font = Font(name="Arial", size=9, color=BLACK); vc.border = BORDER
        vc.number_format = "0.0x"
        nc = ws.cell(row=r, column=3, value=note)
        nc.font = Font(name="Arial", size=9, color="FF555555"); nc.border = BORDER

    ws.freeze_panes = "A5"

# ---------------------------------------------------------------------------
# INDUSTRY EXPOSURE SHEET
# ---------------------------------------------------------------------------

def write_industry_sheet(wb: Workbook, holdings: pd.DataFrame, fund_map: dict):
    """Write a sheet showing long-side exposure by GICS sector."""
    ws = wb.create_sheet("Industry Exposure")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Long Side — Exposure by GICS Sector"
    ws["A1"].font = Font(name="Arial", bold=True, size=13)
    ws.merge_cells("A1:F1")

    ws["A2"] = f"Generated: {datetime.today().strftime('%d-%b-%Y %H:%M')}   |   Long positions only"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="FF555555")
    ws.merge_cells("A2:F2")

    # --- Detail table ---
    write_header_row(ws, ["Ticker", "Side", "Weight", "GICS Sector", "P/E", "P/B"], row=4)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12

    sector_weights: dict[str, float] = {}
    detail_rows = []

    for _, row in holdings.iterrows():
        ticker   = str(row["Ticker"]).strip()
        currency = str(row["Local Currency"]).strip()
        side     = str(row.get("Side", "Long")).strip() if "Side" in row.index else "Long"
        if side.lower() != "long":
            continue
        ticker_r, _ = TICKER_REMAP.get((ticker, currency), (ticker, currency))
        fund   = fund_map.get(ticker_r, {})
        weight = float(row["Weighting"]) if pd.notna(row["Weighting"]) else 0.0
        gics   = fund.get("gics_sector", "Unknown")
        pe     = na_safe(fund.get("pe"))
        pb     = na_safe(fund.get("pb"))

        sector_weights[gics] = sector_weights.get(gics, 0.0) + weight
        detail_rows.append((ticker_r, side, weight, gics, pe, pb))

    for r_idx, (t, s, w, g, pe, pb) in enumerate(detail_rows, start=5):
        vals  = [t, s, w, g, pe, pb]
        fmts  = [None, None, "0.00%", None, "0.0x", "0.0x"]
        aligns = ["left", "center", "right", "left", "right", "right"]
        for c_idx, (v, fmt, aln) in enumerate(zip(vals, fmts, aligns), 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.font   = DATA_FONT
            cell.border = BORDER
            cell.alignment = Alignment(horizontal=aln)
            if fmt and v is not None:
                cell.number_format = fmt

    # --- Sector summary table ---
    summary_start = 5 + len(detail_rows) + 3

    ws.cell(row=summary_start, column=1, value="SECTOR SUMMARY").font = \
        Font(name="Arial", bold=True, size=10)
    ws.merge_cells(f"A{summary_start}:C{summary_start}")

    write_header_row(ws, ["GICS Sector", "Gross Weight", "% of Long Book"], row=summary_start + 1)

    total_long_weight = sum(sector_weights.values())

    # Sort sectors by weight descending
    sorted_sectors = sorted(sector_weights.items(), key=lambda x: x[1], reverse=True)

    # Sector colour palette (cycle through distinct fills)
    SECTOR_FILLS = [
        "FFD6E4F0", "FFD5E8D4", "FFFFE6CC", "FFE1D5E7",
        "FFFFD7CC", "FFF8CECC", "FFDAE8FC", "FFFFE6CC",
        "FFD5E8D4", "FFD6E4F0", "FFE1D5E7", "FFF0F0F0",
    ]

    for i, (sector, weight) in enumerate(sorted_sectors):
        r = summary_start + 2 + i
        fill = PatternFill("solid", start_color=SECTOR_FILLS[i % len(SECTOR_FILLS)])

        sc = ws.cell(row=r, column=1, value=sector)
        sc.font = Font(name="Arial", bold=True, size=9)
        sc.border = BORDER; sc.fill = fill

        wc = ws.cell(row=r, column=2, value=weight)
        wc.font = DATA_FONT; wc.border = BORDER; wc.fill = fill
        wc.number_format = "0.00%"
        wc.alignment = Alignment(horizontal="right")

        pct_val = weight / total_long_weight if total_long_weight > 0 else 0
        pc = ws.cell(row=r, column=3, value=pct_val)
        pc.font = DATA_FONT; pc.border = BORDER; pc.fill = fill
        pc.number_format = "0.0%"
        pc.alignment = Alignment(horizontal="right")

    # Total row
    total_r = summary_start + 2 + len(sorted_sectors)
    gold = PatternFill("solid", start_color="FFFFF2CC")
    tc = ws.cell(row=total_r, column=1, value="TOTAL LONG")
    tc.font = Font(name="Arial", bold=True, size=9); tc.border = BORDER; tc.fill = gold
    tw = ws.cell(row=total_r, column=2, value=total_long_weight)
    tw.font = Font(name="Arial", bold=True, size=9); tw.border = BORDER; tw.fill = gold
    tw.number_format = "0.00%"; tw.alignment = Alignment(horizontal="right")
    tp = ws.cell(row=total_r, column=3, value=1.0 if total_long_weight > 0 else 0)
    tp.font = Font(name="Arial", bold=True, size=9); tp.border = BORDER; tp.fill = gold
    tp.number_format = "0.0%"; tp.alignment = Alignment(horizontal="right")

    ws.freeze_panes = "A5"

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

holdings = pd.read_excel(input_file)
holdings["Weighting"] = pd.to_numeric(holdings["Weighting"], errors="coerce")
if "Side" not in holdings.columns:
    holdings["Side"] = "Long"

FX_USD = 7.78
FX_CNH = 1.071

print("\nFetching Benchmark: 2800.HK (Hang Seng ETF)...")
benchmark_raw = get_hk_prices("2800")
if benchmark_raw is None or benchmark_raw.empty:
    raise RuntimeError("Could not fetch benchmark (2800.HK). Cannot continue.")

benchmark = benchmark_raw.tail(60)

wb = Workbook()
del wb[wb.sheetnames[0]]

create_fx_sheet(wb)

beta_map = {}
fund_map = {}

for _, row in holdings.iterrows():
    ticker   = str(row["Ticker"]).strip()
    currency = str(row["Local Currency"]).strip()

    print(f"\nProcessing: {ticker} ({currency})...")

    # --- Price data & beta ---
    asset = get_prices(ticker, currency)
    if asset is None or asset.empty:
        print(f"  -> No price data for {ticker}, skipping.")
        beta_map[ticker] = {"beta": "N/A", "status": "Failed", "n_days": 0}
    else:
        combined = align_to_benchmark(asset, benchmark)
        if len(combined) < 46:
            msg = f"Insufficient aligned data ({len(combined)} days)"
            print(f"  -> {msg}")
            beta_map[ticker] = {"beta": msg, "status": "Insufficient data", "n_days": len(combined)}
        else:
            returns  = combined.pct_change().dropna()
            last     = returns.tail(45)
            cov      = last[["asset_close", "bm_close"]].cov().iloc[0, 1]
            var      = last["bm_close"].var()
            beta_val = cov / var if var != 0 else float("nan")

            tk_r, _ = TICKER_REMAP.get((ticker, currency), (ticker, currency))
            sheet_name = tk_r[:31].replace("/", "_").replace("\\", "_")
            beta_result = {
                "covariance": round(cov, 8),
                "variance":   round(var, 8),
                "beta":       round(beta_val, 6),
                "n_days":     len(combined),
                "status":     "OK",
                "excel_link": f"='{sheet_name}'!B{len(combined) + 6}",
            }
            beta_map[tk_r] = beta_result
            write_stock_sheet(wb, ticker, currency, combined, beta_result)

    # --- Fundamentals ---
    fund = get_fundamentals(ticker, currency)
    tk_r, _ = TICKER_REMAP.get((ticker, currency), (ticker, currency))
    fund_map[tk_r] = fund
    print(f"  -> PE={fund.get('pe')}, PB={fund.get('pb')}, "
          f"DivYield={fund.get('div_yield')}, "
          f"MktCap={fund.get('mktcap_local')}, "
          f"Industry={fund.get('industry')}, "
          f"GICS={fund.get('gics_sector')}")

write_summary_sheet(wb, holdings, beta_map, fund_map, FX_USD, FX_CNH)
write_industry_sheet(wb, holdings, fund_map)

# Sheet ordering: Summary → FX Rates → Industry Exposure → individual stocks
sheets = wb.sheetnames
wb.move_sheet("Industry Exposure", offset=-len(sheets))
wb.move_sheet(FX_SHEET,           offset=-len(sheets))
wb.move_sheet("Summary",          offset=-len(sheets))

output_file = "Beta_Calculation_Audit.xlsx"
wb.save(output_file)
print(f"\nDone. Audit file saved: {output_file}")