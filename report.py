
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

def build_unified_table(totals_df: pd.DataFrame, matrix_df: pd.DataFrame):
    """
    Unified table layout:
      - TOTALS section (rows) before holdings:
          Row: Total NAV (portfolios as columns)
          Row: Total Cash/PP (portfolios as columns)
      - HOLDINGS section:
          Columns: Ticker, <portfolios...>, Presence (presence is last)
    """
    if matrix_df is None:
        matrix_df = pd.DataFrame()

    portfolio_cols = [c for c in matrix_df.columns if c not in ("ticker","presence","presence_count")]
    if not portfolio_cols and totals_df is not None and not totals_df.empty:
        portfolio_cols = sorted(totals_df["portfolio"].dropna().unique().tolist())

    nav_map = {}
    cash_map = {}
    if totals_df is not None and not totals_df.empty:
        agg = totals_df.groupby("portfolio", as_index=False).agg(
            total_nav=("total_nav","sum"),
            total_cash_or_pp=("total_cash_or_pp","sum")
        )
        for _, r in agg.iterrows():
            nav_map[str(r["portfolio"])] = r["total_nav"]
            cash_map[str(r["portfolio"])] = r["total_cash_or_pp"]

    def fmt(x):
        return f"{x:,.2f}" if pd.notna(x) else ""

    data = []
    data.append(["TOTALS"] + [""] * (len(portfolio_cols)))
    data.append(["Metric"] + portfolio_cols)
    data.append(["Total NAV"] + [fmt(nav_map.get(p, np.nan)) for p in portfolio_cols])
    data.append(["Total Cash/PP"] + [fmt(cash_map.get(p, np.nan)) for p in portfolio_cols])

    # spacer
    data.append([""] + [""] * (len(portfolio_cols)))

    # Holdings
    m = matrix_df.copy()
    for c in portfolio_cols:
        if c in m.columns:
            m[c] = m[c].map(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
        else:
            m[c] = ""
    if "presence" not in m.columns:
        m["presence"] = ""

    data.append(["HOLDINGS"] + [""] * (len(portfolio_cols)))
    data.append(["Ticker"] + portfolio_cols + ["Presence"])

    if not m.empty:
        rows = m[["ticker"] + portfolio_cols + ["presence"]].values.tolist()
        data.extend(rows)

    max_cols = max(len(r) for r in data) if data else 1
    data = [list(r) + [""]*(max_cols-len(r)) for r in data]
    return data

def export_excel(out_path: str, *args):
    if len(args)==1:
        unified=args[0]
        with pd.ExcelWriter(out_path, engine="openpyxl") as w:
            (unified if isinstance(unified, pd.DataFrame) else pd.DataFrame(unified)).to_excel(w, index=False, header=False, sheet_name="Consolidated")
        return
    if len(args)!=4:
        raise TypeError("export_excel expected either (out_path, unified_df) or (out_path, summary_df, holdings_df, totals_df, matrix_df)")
    summary_df, holdings_df, totals_df, matrix_df = args
    unified_df = pd.DataFrame(build_unified_table(totals_df, matrix_df))
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        unified_df.to_excel(w, index=False, header=False, sheet_name="Consolidated")
        totals_df.to_excel(w, index=False, sheet_name="Totals")
        summary_df.to_excel(w, index=False, sheet_name="PortfolioSummary")
        holdings_df.to_excel(w, index=False, sheet_name="Holdings")
        matrix_df.to_excel(w, index=False, sheet_name="PresenceMatrix")

def export_pdf(out_path: str, *args, title: str = "Master Allocation Comparison"):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(out_path, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=18, bottomMargin=18)
    elems = [Paragraph(title, styles["Title"]), Spacer(1,8)]
    if len(args)==1:
        unified=args[0]
        data = unified.values.tolist() if isinstance(unified, pd.DataFrame) else list(unified)
    elif len(args)==2:
        totals_df, matrix_df = args
        data = build_unified_table(totals_df, matrix_df)
    else:
        raise TypeError("export_pdf expected either (out_path, unified_df, title=...) or (out_path, totals_df, matrix_df, title=...)")

    max_cols = max(len(r) for r in data) if data else 1
    data = [list(r)+[""]*(max_cols-len(r)) for r in data]

    # find section header rows
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
