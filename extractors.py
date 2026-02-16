
import re
import numpy as np
import pandas as pd
from portfolio_utils import (
    normalize_str, coerce_number, find_cell, choose_nearest_numeric_threshold,
    clean_ticker, safe_percent_to_ratio, trim_empty_rows
)

def _value_to_right(df: pd.DataFrame, r: int, c: int, max_steps: int = 8):
    for dc in range(1, max_steps + 1):
        if c + dc >= df.shape[1]:
            break
        v = coerce_number(df.iat[r, c + dc])
        if v == v:
            return v
    return np.nan

def extract_new_portfolios(path: str, group_name: str = "New Portfolios"):
    """
    Updated rule:
      - Cash from B2
      - Total NAV from B5
      - Holdings stop at row whose ticker == 'Total'
    """
    with pd.ExcelFile(path) as xl:
        sheet = xl.sheet_names[0]

    df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object).replace({np.nan: None})

    cash = coerce_number(df.iat[1, 1])  # B2
    nav  = coerce_number(df.iat[4, 1])  # B5

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
        "cash_or_pp": cash,
        "source": path,
        "as_of": None
    }])
    gt = pd.DataFrame([{
        "group": group_name,
        "portfolio": group_name,
        "total_nav": nav,
        "total_cash_or_pp": cash
    }])
    return summary, holdings, gt

def _extract_consolidated_from_sheet(path: str, sheet_name: str, group: str, portfolio: str):
    """
    Updated rules (per user):
      - 'Consolidated' table anchor exists
      - Purchasing Power is the value to the right of the cell labeled 'Purchasing Power' BELOW the consolidated table
      - NAV is the value to the right of the cell 'Net' in COLUMN J (index 9), and this row must be BELOW the 'Consolidated' row
      - Holdings weights come from inside the consolidated table (Stock/Ticker + Weight)
    """
    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object).replace({np.nan: None})

    cons = find_cell(df, r"\bConsolidated\b")
    if cons is None:
        raise ValueError(f"Consolidated anchor not found: {path} / {sheet_name}")
    r0, c0 = cons

    # Find holdings header row under consolidated
    header_row = None
    for r in range(r0, min(df.shape[0], r0 + 80)):
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

    # Purchasing Power label below the consolidated table
    cash = np.nan
    pp_search_start = header_row if header_row is not None else r0
    pp_search_start = max(pp_search_start, r0)
    win_pp = df.iloc[pp_search_start:min(df.shape[0], pp_search_start + 300), :]
    pp_cell = find_cell(win_pp, r"Purchasing\s*Power|Buying\s*Power|Purchase\s*Power")
    if pp_cell is not None:
        rr, cc = pp_cell
        rr = pp_search_start + rr
        cash = _value_to_right(df, rr, cc)
        if cash != cash:
            cash = choose_nearest_numeric_threshold(df, rr, cc, min_value=1000.0)

        # NAV: 'Net' must be in column J (index 9), and must be BELOW the 'Consolidated' row.
    nav = np.nan
    net_col = 9
    for rr in range(r0 + 1, df.shape[0]):
        cell = normalize_str(df.iat[rr, net_col]).strip().lower()
        if "net" in cell and cell != "":
            nav = _value_to_right(df, rr, net_col, max_steps=10)
            if nav != nav:
                nav = choose_nearest_numeric_threshold(df, rr, net_col, min_value=1000.0)
            break

    summary = {
        "group": group,
        "portfolio": portfolio,
        "nav": nav,
        "cash_or_pp": cash,
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

    if not out_s and sheet_names:
        for i, s in enumerate(sheet_names[:2]):
            summ, hold = _extract_consolidated_from_sheet(path, s, "Yasser", f"Sheet{i+1}")
            out_s.append(summ); out_h.append(hold)

    summ_df = pd.DataFrame(out_s)
    hold_df = pd.concat(out_h, ignore_index=True) if out_h else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    gt = summ_df.groupby(["group","portfolio"], as_index=False).agg(total_nav=("nav","sum"), total_cash_or_pp=("cash_or_pp","sum"))
    return summ_df, hold_df, gt


def extract_cfh(path: str):
    with pd.ExcelFile(path) as xl:
        sheet_names = xl.sheet_names

    out_s, out_h = [], []

    # If CFH file is basically one consolidated sheet (often named Raw Data-1), just name portfolio "CFH"
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
    gt = summ_df.groupby(["group","portfolio"], as_index=False).agg(total_nav=("nav","sum"), total_cash_or_pp=("cash_or_pp","sum"))
    return summ_df, hold_df, gt


def extract_positions_by_group(path: str):
    """
    One group per sheet. ONLY use 'Group Summary' block.
    Skip 'Ungrouped'
    """
    with pd.ExcelFile(path) as xl:
        sheet_names = xl.sheet_names

    all_s, all_h, all_gt = [], [], []
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
        for r in range(r0, min(df.shape[0], r0 + 30)):
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

        h_rows = []
        if header_row is not None:
            for r in range(header_row + 1, df.shape[0]):
                ticker = clean_ticker(df.iat[r, c0])
                if not ticker:
                    break
                weight_raw = coerce_number(df.iat[r, c0 + 3])
                h_rows.append({
                    "group": sheet,
                    "portfolio": sheet,
                    "ticker": ticker,
                    "weight_ratio": safe_percent_to_ratio(weight_raw)
                })

        all_s.append({"group": sheet, "portfolio": sheet, "nav": total_nav, "cash_or_pp": total_cash, "source": path, "as_of": None})
        all_gt.append({"group": sheet, "portfolio": sheet, "total_nav": total_nav, "total_cash_or_pp": total_cash})
        if h_rows:
            all_h.append(pd.DataFrame(h_rows))

    summ_df = pd.DataFrame(all_s)
    hold_df = pd.concat(all_h, ignore_index=True) if all_h else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    gt_df = pd.DataFrame(all_gt)
    return summ_df, hold_df, gt_df


def extract_customer_position_mode_b(path: str):
    """
    Mode B (updated per user):
      - Sheet: Retail>5M
      - Find 'Emad Farah' (single cell)
      - Cash = value in the column immediately AFTER his name (same row)
      - 'Net' cell is next to his name (same row, typically after cash); NAV = value in the column AFTER 'Net'
      - Stocks list starts in the column right AFTER the cash column (i.e., same row+1 downward), and continues downward until blank
      - Each stock weight% = (Stock Value / Total NAV) * 100
        where Stock Value is located 2 columns to the right of the stock name (i.e., stock_col + 2)
      - Store weights as ratios (0-1) in weight_ratio
    """
    sheet_name = "Retail>5M"
    client_name = "Emad Farah"

    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object).replace({np.nan: None})

    # locate name
    anchor = find_cell(df, re.escape(client_name))
    if anchor is None:
        raise ValueError(f"Client not found: {client_name} in sheet {sheet_name}")
    r, c = anchor

    # cash: next column
    cash = coerce_number(df.iat[r, c+1]) if c+1 < df.shape[1] else np.nan

    # net: search for 'Net' near the name block (may be a few rows below)
    nav = np.nan
    net_r = None
    net_c = None
    for rr in range(r, min(df.shape[0], r+25)):
        for cc in range(c, min(df.shape[1], c+15)):
            s = normalize_str(df.iat[rr, cc]).strip().lower()
            if s == "net" or (s and "net" in s and len(s) <= 10):
                net_r, net_c = rr, cc
                break
        if net_c is not None:
            break
    if net_r is not None and net_c is not None and net_c + 1 < df.shape[1]:
        nav = coerce_number(df.iat[net_r, net_c + 1])

    # holdings: stock names column is right after cash column => c+2
    stock_col = c + 2
    value_col = stock_col + 2  # "Stock Value" column is 2 cols after stock name

    rows = []
    if stock_col < df.shape[1]:
        for rr in range(r+1, df.shape[0]):
            ticker = clean_ticker(df.iat[rr, stock_col])
            if not ticker:
                break
            stock_value = coerce_number(df.iat[rr, value_col]) if value_col < df.shape[1] else np.nan
            weight_ratio = (stock_value / nav) if (nav == nav and nav != 0 and stock_value == stock_value) else np.nan
            rows.append({"group":"Emad","portfolio":"Emad","ticker":ticker,"weight_ratio":weight_ratio})

    holdings = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    summary = pd.DataFrame([{"group":"Emad","portfolio":"Emad","nav":nav,"cash_or_pp":cash,"source":path,"as_of":None}])
    gt = pd.DataFrame([{"group":"Emad","portfolio":"Emad","total_nav":nav,"total_cash_or_pp":cash}])
    return summary, holdings, gt

