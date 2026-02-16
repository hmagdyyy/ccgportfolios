
import re
from typing import Any, Optional, Tuple
import numpy as np
import pandas as pd

def normalize_str(x: Any) -> str:
    if x is None:
        return ""
    return str(x).replace("\u00a0", " ").strip()

def coerce_number(x: Any) -> float:
    if x is None:
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = normalize_str(x)
    if not s:
        return np.nan
    s = s.replace(",", "")
    if s.endswith("%"):
        try:
            return float(s[:-1])
        except Exception:
            return np.nan
    s = re.sub(r"[^\d\.\-\(\)]", "", s)
    if not s:
        return np.nan
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except Exception:
        return np.nan

def find_cell(df: pd.DataFrame, pattern: str) -> Optional[Tuple[int,int]]:
    rx = re.compile(pattern, flags=re.IGNORECASE)
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            v = df.iat[r,c]
            if v is None:
                continue
            s = normalize_str(v)
            if s and rx.search(s):
                return (r,c)
    return None

def choose_nearest_numeric_threshold(df: pd.DataFrame, r: int, c: int, min_value: float = 1000.0, max_radius: int = 8) -> float:
    best = None
    best_d = 10**9
    fallback = None
    fallback_d = 10**9
    for dr in range(-max_radius, max_radius+1):
        for dc in range(-max_radius, max_radius+1):
            rr, cc = r+dr, c+dc
            if rr < 0 or cc < 0 or rr >= df.shape[0] or cc >= df.shape[1]:
                continue
            v = coerce_number(df.iat[rr,cc])
            if np.isnan(v):
                continue
            d = abs(dr) + abs(dc)
            if d < fallback_d:
                fallback = v; fallback_d = d
            if abs(v) >= min_value and d < best_d:
                best = v; best_d = d
    return best if best is not None else (fallback if fallback is not None else np.nan)

def safe_percent_to_ratio(v: float) -> float:
    if v is None or np.isnan(v):
        return np.nan
    if -1.2 <= v <= 1.2:
        return float(v)
    return float(v)/100.0

def trim_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    mask = df.apply(lambda r: any(normalize_str(x) != "" and normalize_str(x).lower() != "nan" for x in r.tolist()), axis=1)
    return df.loc[mask].copy()

def clean_ticker(x: Any) -> str:
    s = normalize_str(x)
    if not s or s.strip().lower() in ("nan","none","null"):
        return ""
    s = s.replace(".CA","").upper()
    s = re.sub(r"[^A-Z0-9\-_]", "", s)
    junk = {
        "NAN","GROUP","NAME","TOTAL","SUMMARY","GROUPSUMMARY","NAV","STOCKS","TOTALCASH","TOTALNAV",
        "CASH","PURCHASINGPOWER","MV","WEIGHT","QUANTITY","NET"
    }
    if s in junk:
        return ""
    return s
