
import pandas as pd
import numpy as np
from extractors import extract_new_portfolios, extract_yasser, extract_cfh, extract_positions_by_group, extract_customer_position_mode_b


def build_presence_matrix(holdings_df: pd.DataFrame, portfolios: list[str]) -> pd.DataFrame:
    """
    Output columns order:
      ticker, <portfolio1>, <portfolio2>, ..., presence, presence_count
    presence shows "k/N"
    """
    if holdings_df is None or holdings_df.empty:
        return pd.DataFrame(columns=["ticker"] + portfolios + ["presence","presence_count"])

    h = holdings_df.copy()
    h["present"] = 1
    h["weight_pct"] = h["weight_ratio"] * 100.0

    weight_pivot = h.pivot_table(index="ticker", columns="portfolio", values="weight_pct", aggfunc="sum").reindex(columns=portfolios)
    pres_pivot = h.pivot_table(index="ticker", columns="portfolio", values="present", aggfunc="max").reindex(columns=portfolios)

    weight_df = weight_pivot.reset_index()
    pres_df = pres_pivot.reset_index()

    pres_only = pres_df[portfolios].fillna(0).astype(int) if portfolios else pd.DataFrame()
    presence_count = pres_only.sum(axis=1) if portfolios else pd.Series([0]*len(weight_df))

    weight_df["presence_count"] = presence_count.astype(int)
    weight_df["presence"] = weight_df["presence_count"].astype(str) + "/" + str(len(portfolios))

    # sort by presence then sum weight
    sum_weight = weight_df[portfolios].fillna(0).sum(axis=1) if portfolios else 0
    weight_df["_sum_weight"] = sum_weight
    weight_df = weight_df.sort_values(["presence_count","_sum_weight"], ascending=[False, False]).drop(columns=["_sum_weight"])

    # reorder columns
    cols = ["ticker"] + portfolios + ["presence","presence_count"]
    return weight_df[cols]


def build_master(new_portfolios_path: str, yasser_path: str, cfh_path: str, positions_by_group_path: str, customers_position_path: str):
    parts=[]
    parts.append(extract_new_portfolios(new_portfolios_path, "New Portfolios"))
    parts.append(extract_yasser(yasser_path))
    parts.append(extract_cfh(cfh_path))
    parts.append(extract_positions_by_group(positions_by_group_path))
    parts.append(extract_customer_position_mode_b(customers_position_path))

    summary_df = pd.concat([p[0] for p in parts], ignore_index=True)
    holdings_df = pd.concat([p[1] for p in parts if p[1] is not None and not p[1].empty], ignore_index=True) if any(p[1] is not None and not p[1].empty for p in parts) else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    totals_df = pd.concat([p[2] for p in parts], ignore_index=True)

    for _, row in summary_df[["group","portfolio"]].drop_duplicates().iterrows():
        g,p = row["group"], row["portfolio"]
        if not ((totals_df["group"]==g) & (totals_df["portfolio"]==p)).any():
            s = summary_df[(summary_df["group"]==g) & (summary_df["portfolio"]==p)]
            totals_df = pd.concat([totals_df, pd.DataFrame([{
                "group":g,"portfolio":p,
                "total_nav": s["nav"].sum(),
                "total_cash_or_pp": s["cash_or_pp"].sum()
            }])], ignore_index=True)

    portfolios = sorted(summary_df["portfolio"].dropna().unique().tolist())
    matrix_df = build_presence_matrix(holdings_df, portfolios)
    return summary_df, holdings_df, totals_df, matrix_df
