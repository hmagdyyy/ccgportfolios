
import re
import numpy as np
import pandas as pd
from portfolio_utils import (
    normalize_str, coerce_number, find_cell, choose_nearest_numeric_threshold,
    clean_ticker, safe_percent_to_ratio, trim_empty_rows
)

def _value_to_right(df: pd.DataFrame, r: int, c: int, max_steps: int = 10):
    for dc in range(1, max_steps + 1):
        if c + dc >= df.shape[1]:
            break
        v = coerce_number(df.iat[r, c + dc])
        if v == v:
            return v
    return np.nan

# -----------------------------
# New Portfolios (single group)
# -----------------------------
def extract_new_portfolios(path: str, group_name: str = "Arqaam", fx_usd_to_egp: float | None = None):
    """
    Rules:
      - Cash from B2
      - Total NAV from B5
      - Holdings stop at row where ticker == 'Total'
    """
    with pd.ExcelFile(path) as xl:
        sheet = xl.sheet_names[0]

    df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object).replace({np.nan: None})

    cash = coerce_number(df.iat[1, 1])  # B2 (USD)
    nav  = coerce_number(df.iat[4, 1])  # B5 (USD)

    # Convert USD->EGP if rate provided
    if fx_usd_to_egp is not None and fx_usd_to_egp == fx_usd_to_egp:
        if cash == cash:
            cash = cash * float(fx_usd_to_egp)
        if nav == nav:
            nav = nav * float(fx_usd_to_egp)

    # Find holdings header row
    header_row = None
    for r in range(df.shape[0]):
        row = [normalize_str(x).lower() for x in df.iloc[r].tolist() if x is not None]
        if any("shares" in v for v in row) and any("%" in v and "nav" in v for v in row):
            header_row = r
            break

    holdings = pd.DataFrame(columns=["group", "portfolio", "ticker", "weight_ratio"])
    if header_row is not None:
        sub = pd.read_excel(path, sheet_name=sheet, header=header_row)
        sub = trim_empty_rows(sub)
        cols = list(sub.columns)
        ticker_col = cols[0]
        weight_col = None
        for c in cols:
            cl = str(c).lower()
            if "% nav" in cl or "%" in cl:
                weight_col = c
                break

        if weight_col is not None:
            tmp = sub[[ticker_col, weight_col]].copy()
            tmp.columns = ["ticker_raw", "weight_raw"]

            raw_series = tmp["ticker_raw"].astype(str).str.strip().str.lower()
            if (raw_series == "total").any():
                first_total = raw_series[raw_series == "total"].index[0]
                tmp = tmp.loc[:first_total - 1].copy()

            tmp["ticker"] = tmp["ticker_raw"].apply(clean_ticker)
            tmp["weight_ratio"] = tmp["weight_raw"].apply(lambda x: safe_percent_to_ratio(coerce_number(x)))
            tmp = tmp.dropna(subset=["ticker"])
            tmp = tmp[tmp["ticker"].astype(str).str.len() > 0]
            tmp["group"] = group_name
            tmp["portfolio"] = group_name
            holdings = tmp[["group", "portfolio", "ticker", "weight_ratio"]].copy()

    summary = pd.DataFrame([{
        "group": group_name,
        "portfolio": group_name,
        "nav": nav,
        "cash": cash,
        "purchasing_power": np.nan,
        "source": path,
        "as_of": None
    }])

    totals = pd.DataFrame([{
        "group": group_name,
        "portfolio": group_name,
        "total_nav": nav,
        "total_cash": cash,
        "total_purchasing_power": np.nan
    }])

    return summary, holdings, totals

# ---------------------------------
# Consolidated table extractor (CFH/Yasser/R&R)
# ---------------------------------
def _extract_consolidated_from_sheet(path: str, sheet_name: str, group: str, portfolio: str):
    """
    Rules:
      - Anchor on 'Consolidated'
      - Purchasing Power: label 'Purchasing Power' below table, value to the right
      - NAV: find 'Net' in COLUMN J (index 9), below 'Consolidated' row, NAV is value to the right of Net
      - Holdings: inside consolidated table (Stock/Ticker + Weight/%)
    """
    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object).replace({np.nan: None})

    cons = find_cell(df, r"\bConsolidated\b")
    if cons is None:
        raise ValueError(f"Consolidated anchor not found: {path} / {sheet_name}")
    r0, c0 = cons

    # Find holdings header row under consolidated
    header_row = None
    for r in range(r0, min(df.shape[0], r0 + 100)):
        row_vals = [normalize_str(v).lower() for v in df.iloc[r].tolist() if v is not None]
        if any(("stock" in v or v.strip() in ("ticker", "symbol")) for v in row_vals) and any(("weight" in v or "%" in v) for v in row_vals):
            header_row = r
            break

    holdings = pd.DataFrame(columns=["group", "portfolio", "ticker", "weight_ratio"])
    table_end_guess = r0 + 40
    if header_row is not None:
        sub = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
        sub = trim_empty_rows(sub)
        table_end_guess = header_row + len(sub) + 10

        stock_col = None
        weight_col = None
        for c in sub.columns:
            cl = str(c).strip().lower()
            if stock_col is None and any(k in cl for k in ["stock", "ticker", "symbol", "code", "isin"]):
                stock_col = c
            if weight_col is None and ("weight" in cl or "%" in cl):
                weight_col = c

        if stock_col is not None and weight_col is not None:
            tmp = sub[[stock_col, weight_col]].copy()
            tmp.columns = ["ticker_raw", "weight_raw"]
            tmp["ticker"] = tmp["ticker_raw"].apply(clean_ticker)
            tmp["weight_ratio"] = tmp["weight_raw"].apply(lambda x: safe_percent_to_ratio(coerce_number(x)))
            tmp = tmp.dropna(subset=["ticker"])
            tmp = tmp[tmp["ticker"].astype(str).str.len() > 0]
            tmp["group"] = group
            tmp["portfolio"] = portfolio
            holdings = tmp[["group", "portfolio", "ticker", "weight_ratio"]].copy()

    # Purchasing Power label below the table
    purchasing_power = np.nan
    pp_search_start = header_row if header_row is not None else r0
    win_pp = df.iloc[pp_search_start:min(df.shape[0], pp_search_start + 400), :]
    pp_cell = find_cell(win_pp, r"Purchasing\s*Power|Buying\s*Power|Purchase\s*Power")
    if pp_cell is not None:
        rr, cc = pp_cell
        rr = pp_search_start + rr
        purchasing_power = _value_to_right(df, rr, cc)
        if purchasing_power != purchasing_power:
            purchasing_power = choose_nearest_numeric_threshold(df, rr, cc, min_value=1000.0)

    # NAV: Net in column J (index 9)
    nav = np.nan
    net_col = 9
    for rr in range(r0 + 1, df.shape[0]):
        cell = normalize_str(df.iat[rr, net_col]).strip().lower()
        if "net" in cell and cell != "":
            nav = _value_to_right(df, rr, net_col)
            if nav != nav:
                nav = choose_nearest_numeric_threshold(df, rr, net_col, min_value=1000.0)
            break

    summary = {
        "group": group,
        "portfolio": portfolio,
        "nav": nav,
        "cash": np.nan,
        "purchasing_power": purchasing_power,
        "source": path,
        "as_of": None
    }
    return summary, holdings

def extract_yasser(path: str):
    with pd.ExcelFile(path) as xl:
        sheet_names = xl.sheet_names

    out_s, out_h = [], []
    for s in sheet_names:
        sl = s.lower()
        if "yasser" in sl:
            summ, hold = _extract_consolidated_from_sheet(path, s, "Yasser", "Yasser")
            out_s.append(summ); out_h.append(hold)
        if "r&r" in sl or "r & r" in sl or "rnr" in sl:
            summ, hold = _extract_consolidated_from_sheet(path, s, "Yasser", "R&R")
            out_s.append(summ); out_h.append(hold)

    summ_df = pd.DataFrame(out_s)
    hold_df = pd.concat(out_h, ignore_index=True) if out_h else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])

    totals_df = summ_df.groupby(["group","portfolio"], as_index=False).agg(
        total_nav=("nav","sum"),
        total_cash=("cash","sum"),
        total_purchasing_power=("purchasing_power","sum"),
    )
    return summ_df, hold_df, totals_df

def extract_cfh(path: str):
    with pd.ExcelFile(path) as xl:
        sheet_names = xl.sheet_names

    out_s, out_h = [], []
    if len(sheet_names) == 1:
        s = sheet_names[0]
        summ, hold = _extract_consolidated_from_sheet(path, s, "CFH", "CFH")
        out_s.append(summ); out_h.append(hold)
    else:
        for s in sheet_names:
            sl = s.lower().strip()
            portfolio_name = "CFH" if "raw data" in sl else s
            summ, hold = _extract_consolidated_from_sheet(path, s, "CFH", portfolio_name)
            out_s.append(summ); out_h.append(hold)

    summ_df = pd.DataFrame(out_s)
    hold_df = pd.concat(out_h, ignore_index=True) if out_h else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])

    totals_df = summ_df.groupby(["group","portfolio"], as_index=False).agg(
        total_nav=("nav","sum"),
        total_cash=("cash","sum"),
        total_purchasing_power=("purchasing_power","sum"),
    )
    return summ_df, hold_df, totals_df

# -----------------------------
# Positions by Group
# -----------------------------
def extract_positions_by_group(path: str):
    """
    One group per sheet.
    ONLY use 'Group Summary' block:
      - Total Cash / Total NAV from label rows near Group Summary
      - Holdings: ticker column under 'Stock' and weight from 3 columns after (e.g., K when H is stock)
    Skip 'Ungrouped'
    """
    with pd.ExcelFile(path) as xl:
        sheet_names = xl.sheet_names

    all_s, all_h, all_t = [], [], []
    for sheet in sheet_names:
        if sheet.strip().lower() == "ungrouped":
            continue

        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object).replace({np.nan: None})
        gs = find_cell(df, r"Group\s*Summary")
        if gs is None:
            continue
        r0, c0 = gs

        total_cash = np.nan
        total_nav = np.nan
        for r in range(r0, min(df.shape[0], r0 + 40)):
            label = normalize_str(df.iat[r, c0]).strip().lower()
            val = coerce_number(df.iat[r, c0 + 1])
            if label == "total cash":
                total_cash = val
            elif label == "total nav":
                total_nav = val

        header_row = None
        for r in range(r0, df.shape[0]):
            if normalize_str(df.iat[r, c0]).strip().lower() == "stock":
                header_row = r
                break

        rows=[]
        if header_row is not None:
            for rr in range(header_row + 1, df.shape[0]):
                ticker = clean_ticker(df.iat[rr, c0])
                if not ticker:
                    break
                weight_raw = coerce_number(df.iat[rr, c0 + 3])
                rows.append({"group":sheet,"portfolio":sheet,"ticker":ticker,"weight_ratio": safe_percent_to_ratio(weight_raw)})

        all_s.append({"group":sheet,"portfolio":sheet,"nav":total_nav,"cash":total_cash,"purchasing_power":np.nan,"source":path,"as_of":None})
        all_t.append({"group":sheet,"portfolio":sheet,"total_nav":total_nav,"total_cash":total_cash,"total_purchasing_power":np.nan})
        if rows:
            all_h.append(pd.DataFrame(rows))

    summ_df = pd.DataFrame(all_s)
    hold_df = pd.concat(all_h, ignore_index=True) if all_h else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    totals_df = pd.DataFrame(all_t)
    return summ_df, hold_df, totals_df

# -----------------------------
# Emad (Mode B)
# -----------------------------
def extract_customer_position_mode_b(path: str):
    """
    Rules (per user):
      - Sheet: Retail>5M
      - Find 'Emad Farah'
      - Cash = value in column after his name (same row)
      - NAV = value in column after 'Net' near his name block
      - Stocks are in the column right after cash value (stock_col = name_col + 2)
      - Stock Value is 2 columns after stock name (value_col = stock_col + 2)
      - weight% = Stock Value / NAV * 100  (stored as ratio)
    """
    sheet_name = "Retail>5M"
    client_name = "Emad Farah"

    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object).replace({np.nan: None})
    anchor = find_cell(df, re.escape(client_name))
    if anchor is None:
        raise ValueError(f"Client not found: {client_name} in sheet {sheet_name}")
    r, c = anchor

    cash = coerce_number(df.iat[r, c+1]) if c+1 < df.shape[1] else np.nan

    nav = np.nan
    net_r = None
    net_c = None
    for rr in range(r, min(df.shape[0], r+25)):
        for cc in range(c, min(df.shape[1], c+15)):
            s = normalize_str(df.iat[rr, cc]).strip().lower()
            if s == "net" or (s and "net" in s and len(s) <= 12):
                net_r, net_c = rr, cc
                break
        if net_c is not None:
            break
    if net_r is not None and net_c is not None and net_c+1 < df.shape[1]:
        nav = coerce_number(df.iat[net_r, net_c+1])

    stock_col = c + 2
    value_col = stock_col + 2

    rows=[]
    if stock_col < df.shape[1]:
        for rr in range(r+1, df.shape[0]):
            ticker = clean_ticker(df.iat[rr, stock_col])
            if not ticker:
                break
            stock_value = coerce_number(df.iat[rr, value_col]) if value_col < df.shape[1] else np.nan
            weight_ratio = (stock_value / nav) if (nav == nav and nav != 0 and stock_value == stock_value) else np.nan
            rows.append({"group":"Emad","portfolio":"Emad","ticker":ticker,"weight_ratio":weight_ratio})

    holdings = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    summary = pd.DataFrame([{"group":"Emad","portfolio":"Emad","nav":nav,"cash":cash,"purchasing_power":np.nan,"source":path,"as_of":None}])
    totals = pd.DataFrame([{"group":"Emad","portfolio":"Emad","total_nav":nav,"total_cash":cash,"total_purchasing_power":np.nan}])
    return summary, holdings, totals
