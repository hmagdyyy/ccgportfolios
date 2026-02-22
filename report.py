
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

def build_unified_table(totals_df: pd.DataFrame, matrix_df: pd.DataFrame):
    if matrix_df is None:
        matrix_df = pd.DataFrame()

    portfolio_cols = [c for c in matrix_df.columns if c not in ("ticker","presence","presence_count")]
    if not portfolio_cols and totals_df is not None and not totals_df.empty:
        portfolio_cols = totals_df["portfolio"].dropna().astype(str).unique().tolist()

    nav_map, cash_map, pp_map = {}, {}, {}
    if totals_df is not None and not totals_df.empty:
        agg = totals_df.groupby("portfolio", as_index=False).agg(
            total_nav=("total_nav","sum"),
            total_cash=("total_cash","sum"),
            total_purchasing_power=("total_purchasing_power","sum"),
        )
        for _, r in agg.iterrows():
            p = str(r["portfolio"])
            nav_map[p] = r.get("total_nav", np.nan)
            cash_map[p] = r.get("total_cash", np.nan)
            pp_map[p] = r.get("total_purchasing_power", np.nan)

    def fmt(x):
        return f"{x:,.2f}" if pd.notna(x) else ""

    def pct_for_portfolio(p: str):
        nav = nav_map.get(p, np.nan)
        cash = cash_map.get(p, np.nan)
        pp = pp_map.get(p, np.nan)

        # if cash exists => Cash/NAV
        if pd.notna(cash) and pd.notna(nav) and nav != 0:
            return (cash / nav) * 100.0

        # if PP exists instead => PP/(NAV+PP)
        if pd.notna(pp) and pd.notna(nav) and (nav + pp) != 0:
            return (pp / (nav + pp)) * 100.0

        return np.nan

    def fmt_pct(v):
        return f"{v:.2f}%" if pd.notna(v) else ""

    total_nav = np.nansum([nav_map.get(p, np.nan) for p in portfolio_cols]) if portfolio_cols else np.nan
    total_cash = np.nansum([cash_map.get(p, np.nan) for p in portfolio_cols]) if portfolio_cols else np.nan
    total_pp = np.nansum([pp_map.get(p, np.nan) for p in portfolio_cols]) if portfolio_cols else np.nan

    total_pct = np.nan
    if pd.notna(total_cash) and pd.notna(total_nav) and total_nav != 0:
        total_pct = (total_cash / total_nav) * 100.0
    elif pd.notna(total_pp) and pd.notna(total_nav) and (total_nav + total_pp) != 0:
        total_pct = (total_pp / (total_nav + total_pp)) * 100.0

    cols_with_total = portfolio_cols + ["TOTAL"]

    data=[]
    data.append(["TOTALS"] + [""]*len(cols_with_total))
    data.append(["Metric"] + cols_with_total)
    data.append(["Total NAV"] + [fmt(nav_map.get(p, np.nan)) for p in portfolio_cols] + [fmt(total_nav)])
    data.append(["Total Cash"] + [fmt(cash_map.get(p, np.nan)) for p in portfolio_cols] + [fmt(total_cash)])
    data.append(["Purchasing Power"] + [fmt(pp_map.get(p, np.nan)) for p in portfolio_cols] + [fmt(total_pp)])
    data.append(["%Cash"] + [fmt_pct(pct_for_portfolio(p)) for p in portfolio_cols] + [fmt_pct(total_pct)])

    data.append([""] + [""]*len(cols_with_total))

    m = matrix_df.copy()
    for c in portfolio_cols:
        if c in m.columns:
            m[c] = m[c].map(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
        else:
            m[c] = ""
    if "presence" not in m.columns:
        m["presence"] = ""

    data.append(["HOLDINGS"] + [""]*len(cols_with_total))
    data.append(["Ticker"] + portfolio_cols + ["Presence"] + [""])

    if not m.empty:
        rows = m[["ticker"] + portfolio_cols + ["presence"]].values.tolist()
        rows = [r + [""] for r in rows]
        data.extend(rows)

    max_cols = max(len(r) for r in data) if data else 1
    data = [list(r)+[""]*(max_cols-len(r)) for r in data]
    return data

def export_excel(out_path: str, summary_df: pd.DataFrame, holdings_df: pd.DataFrame, totals_df: pd.DataFrame, matrix_df: pd.DataFrame):
    unified_df = pd.DataFrame(build_unified_table(totals_df, matrix_df))
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        unified_df.to_excel(w, index=False, header=False, sheet_name="Consolidated")
        totals_df.to_excel(w, index=False, sheet_name="Totals")
        summary_df.to_excel(w, index=False, sheet_name="PortfolioSummary")
        holdings_df.to_excel(w, index=False, sheet_name="Holdings")
        matrix_df.to_excel(w, index=False, sheet_name="PresenceMatrix")

def export_pdf(out_path: str, totals_df: pd.DataFrame, matrix_df: pd.DataFrame, title: str = "Master Allocation Comparison"):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(out_path, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=18, bottomMargin=18)
    elems = [Paragraph(title, styles["Title"]), Spacer(1,8)]
    data = build_unified_table(totals_df, matrix_df)

    max_cols = max(len(r) for r in data) if data else 1
    data = [list(r)+[""]*(max_cols-len(r)) for r in data]

    holdings_row=None
    for i,r in enumerate(data):
        if str(r[0]).strip().upper()=="HOLDINGS":
            holdings_row=i; break
    holdings_header = holdings_row+1 if holdings_row is not None else None

    t = Table(data)
    style=[
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("BACKGROUND",(0,1),(-1,1),colors.lightgrey),
        ("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),
    ]
    if holdings_row is not None:
        style += [
            ("BACKGROUND",(0,holdings_row),(-1,holdings_row),colors.whitesmoke),
            ("FONTNAME",(0,holdings_row),(-1,holdings_row),"Helvetica-Bold"),
            ("BACKGROUND",(0,holdings_header),(-1,holdings_header),colors.lightgrey),
            ("FONTNAME",(0,holdings_header),(-1,holdings_header),"Helvetica-Bold"),
        ]
    t.setStyle(TableStyle(style))
    elems.append(t)
    doc.build(elems)
