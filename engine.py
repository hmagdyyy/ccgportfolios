
import numpy as np
import pandas as pd
from extractors import (
    extract_new_portfolios,
    extract_yasser,
    extract_cfh,
    extract_positions_by_group,
    extract_customer_position_mode_b,
)

def build_presence_matrix(holdings_df: pd.DataFrame, portfolios: list[str]) -> pd.DataFrame:
    """
    Columns:
      ticker, <portfolios...>, presence, presence_count
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

    sum_weight = weight_df[portfolios].fillna(0).sum(axis=1) if portfolios else 0
    weight_df["_sum_weight"] = sum_weight
    weight_df = weight_df.sort_values(["presence_count","_sum_weight"], ascending=[False, False]).drop(columns=["_sum_weight"])

    cols = ["ticker"] + portfolios + ["presence","presence_count"]
    return weight_df[cols]

def build_master(
    new_portfolios_path: str | None = None,
    fx_usd_to_egp: float | None = None,
    yasser_path: str | None = None,
    cfh_path: str | None = None,
    positions_by_group_path: str | None = None,
    customers_position_path: str | None = None,
):
    parts=[]

    if new_portfolios_path:
        parts.append(extract_new_portfolios(new_portfolios_path, "Arqaam", fx_usd_to_egp=fx_usd_to_egp))
    if yasser_path:
        parts.append(extract_yasser(yasser_path))
    if cfh_path:
        parts.append(extract_cfh(cfh_path))
    if positions_by_group_path:
        parts.append(extract_positions_by_group(positions_by_group_path))
    if customers_position_path:
        parts.append(extract_customer_position_mode_b(customers_position_path))

    if not parts:
        summary_df = pd.DataFrame(columns=["group","portfolio","nav","cash","purchasing_power","source","as_of"])
        holdings_df = pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
        totals_df = pd.DataFrame(columns=["group","portfolio","total_nav","total_cash","total_purchasing_power"])
        matrix_df = pd.DataFrame(columns=["ticker","presence","presence_count"])
        return summary_df, holdings_df, totals_df, matrix_df

    summary_df = pd.concat([p[0] for p in parts], ignore_index=True)
    for col in ["cash","purchasing_power"]:
        if col not in summary_df.columns:
            summary_df[col] = np.nan

    holdings_df = (
        pd.concat([p[1] for p in parts if p[1] is not None and not p[1].empty], ignore_index=True)
        if any(p[1] is not None and not p[1].empty for p in parts)
        else pd.DataFrame(columns=["group","portfolio","ticker","weight_ratio"])
    )

    totals_df = pd.concat([p[2] for p in parts], ignore_index=True)
    for col in ["total_nav","total_cash","total_purchasing_power"]:
        if col not in totals_df.columns:
            totals_df[col] = np.nan

    portfolios = summary_df["portfolio"].dropna().astype(str).unique().tolist()

    def port_key(x: str):
        if x == "CFH": return (0, x)
        if x == "Yasser": return (1, x)
        if x == "R&R": return (2, x)
        return (3, x)
    portfolios = sorted(portfolios, key=port_key)

    matrix_df = build_presence_matrix(holdings_df, portfolios)
    return summary_df, holdings_df, totals_df, matrix_df
