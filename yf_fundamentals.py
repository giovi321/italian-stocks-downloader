#!/usr/bin/env python3
"""
Annual financials only, single Excel workbook, FY-strict.

What this enforces:
- Detect each company's fiscal year-end month from yfinance .info (when available) or
  infer the mode of statement column months.
- Keep ONLY columns whose end-month matches the inferred fiscal year-end month.
  This removes interim/half-year columns like 30 June when FY-end is December.
- Skip tickers with no annual data after FY filtering.
- Add a per-ticker Metrics sheet computed from FY-only values:
    market_cap, enterprise_value, ev_to_revenue, ev_to_ebitda,
    total_cash, total_debt, current_ratio, debt_to_equity, roa, roe,
    plus raw FY values used.

Usage:
  python yf_fundamentals.py --tickers tickers.txt --outfile fundamentals.xlsx --sleep 1.0
"""

import argparse
import time
import sys
import os
from typing import List, Dict, Any, Optional, Iterable
import pandas as pd

try:
    import yfinance as yf
except ImportError:
    print("Missing dependency: yfinance. Install with: pip install yfinance pandas openpyxl", file=sys.stderr)
    sys.exit(1)


def read_tickers(path: str) -> List[str]:
    with open(path, "r", encoding="utf-8") as f:
        lines = [ln.strip() for ln in f.readlines()]
    tickers = []
    for ln in lines:
        if not ln or ln.startswith("#"):
            continue
        tickers.extend([t.strip() for t in ln.replace(";", ",").split(",") if t.strip()])
    seen = set()
    unique = []
    for t in tickers:
        t_up = t.upper()
        if t_up not in seen:
            seen.add(t_up)
            unique.append(t_up)
    return unique


def safe_df(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    try:
        df.index = df.index.astype(str)
    except Exception:
        pass
    return df


def fetch_annual_only(ticker: str):
    t = yf.Ticker(ticker)
    # Info
    info = {}
    try:
        info = t.get_info() if hasattr(t, "get_info") else (t.info or {})
    except Exception:
        info = {}

    # Annual statements only
    try:
        is_annual = safe_df(t.financials)  # Annual Income Statement
    except Exception:
        is_annual = pd.DataFrame()
    try:
        bs_annual = safe_df(t.balance_sheet)  # Annual Balance Sheet
    except Exception:
        bs_annual = pd.DataFrame()
    try:
        cf_annual = safe_df(t.cashflow)  # Annual Cash Flow
    except Exception:
        cf_annual = pd.DataFrame()

    return info, is_annual, bs_annual, cf_annual


def _to_datetime_cols(cols: Iterable) -> Optional[pd.DatetimeIndex]:
    try:
        return pd.to_datetime(list(cols))
    except Exception:
        return None


def _infer_fy_month_from_info(info: Dict[str, Any]) -> Optional[int]:
    keys = ["fiscalYearEnd", "fiscalYearEndDate", "fiscalYearEnds"]
    for k in keys:
        v = info.get(k)
        if v is None:
            continue
        try:
            if isinstance(v, (int, float)):
                s = str(int(v))
                if len(s) >= 4:
                    return int(s[-4:-2])
            if isinstance(v, str):
                try:
                    dt = pd.to_datetime(v, errors="raise")
                    return int(dt.month)
                except Exception:
                    months = {m.lower(): i for i, m in enumerate(
                        ["January","February","March","April","May","June","July","August","September","October","November","December"], start=1)}
                    lo = v.strip().lower()
                    if lo in months:
                        return months[lo]
            if isinstance(v, dict):
                for sub in ["raw", "fmt", "longFmt"]:
                    if sub in v:
                        try:
                            dt = pd.to_datetime(v[sub], errors="raise")
                            return int(dt.month)
                        except Exception:
                            pass
        except Exception:
            continue
    return None


def _infer_fy_month_from_statements(dfs: List[pd.DataFrame]) -> Optional[int]:
    months = []
    for df in dfs:
        if df is None or df.empty:
            continue
        dt_cols = _to_datetime_cols(df.columns)
        if dt_cols is not None:
            months.extend(list(dt_cols.month))
    if not months:
        return None
    s = pd.Series(months)
    mode_vals = s.mode()
    if len(mode_vals) == 0:
        return None
    if len(mode_vals) > 1 and 12 in set(mode_vals):
        return 12
    return int(mode_vals.iloc[0])


def _filter_fy_columns(df: pd.DataFrame, fy_month: Optional[int]) -> pd.DataFrame:
    if df is None or df.empty or fy_month is None:
        return df
    dt_cols = _to_datetime_cols(df.columns)
    if dt_cols is None:
        return df
    mask = dt_cols.month == fy_month
    if mask.any():
        return df.loc[:, list(df.columns[mask])]
    return df


def _find_row_value(df: pd.DataFrame, aliases: List[str]) -> Optional[float]:
    if df is None or df.empty:
        return None
    idx_map = {str(ix).strip().lower(): ix for ix in df.index}
    for a in aliases:
        key = a.strip().lower()
        if key in idx_map:
            row = df.loc[idx_map[key]]
            try:
                for col in row.index:
                    val = row[col]
                    if pd.notna(val):
                        try:
                            return float(val)
                        except Exception:
                            continue
            except Exception:
                continue
    return None


def _get_info_first(info: Dict[str, Any], keys: List[str]) -> Optional[float]:
    for k in keys:
        v = info.get(k)
        if v is None:
            continue
        try:
            return float(v)
        except Exception:
            try:
                if isinstance(v, dict) and "raw" in v:
                    return float(v["raw"])
            except Exception:
                pass
    return None


def _safe_div(a: Optional[float], b: Optional[float]) -> Optional[float]:
    try:
        if a is None or b is None or b == 0:
            return None
        return float(a) / float(b)
    except Exception:
        return None


def compute_metrics(ticker: str, info: Dict[str, Any], is_fy: pd.DataFrame, bs_fy: pd.DataFrame) -> pd.DataFrame:
    mc = _get_info_first(info, ["marketCap", "marketcap"])
    if mc is None:
        try:
            mc = float(yf.Ticker(ticker).fast_info.get("market_cap"))
        except Exception:
            pass

    total_cash = _get_info_first(info, ["totalCash", "totalcash"])
    if total_cash is None:
        total_cash = _find_row_value(bs_fy, [
            "Cash And Cash Equivalents",
            "Cash And Cash Equivalents, At Carrying Value",
            "CashCashEquivalentsAndShortTermInvestments",
            "Cash Cash Equivalents And Short Term Investments",
            "Cash",
        ])

    total_debt = _get_info_first(info, ["totalDebt", "totaldebt"])
    if total_debt is None:
        debt_candidates = [
            "Total Debt",
            "Short Long Term Debt",
            "Short Term Debt",
            "Current Debt",
            "Current Portion Of Long Term Debt",
            "Long Term Debt",
            "Long Term Debt Noncurrent",
            "Long Term Debt And Capital Lease Obligations",
        ]
        s = 0.0
        found_any = False
        for alias in debt_candidates:
            v = _find_row_value(bs_fy, [alias])
            if v is not None:
                s += float(v)
                found_any = True
        total_debt = s if found_any else None

    revenue = _find_row_value(is_fy, ["Total Revenue", "TotalRevenue", "Revenue"])
    ebitda = _find_row_value(is_fy, ["EBITDA", "Ebitda"])

    equity = _find_row_value(bs_fy, [
        "Total Stockholder Equity",
        "Total Stockholders' Equity",
        "Stockholders Equity",
        "Shareholders Equity",
        "Total Equity Gross Minority Interest",
        "Total Equity",
    ])
    total_assets = _find_row_value(bs_fy, ["Total Assets"])
    current_assets = _find_row_value(bs_fy, ["Total Current Assets", "Current Assets"])
    current_liabilities = _find_row_value(bs_fy, ["Total Current Liabilities", "Current Liabilities"])

    current_ratio = _get_info_first(info, ["currentRatio"])
    if current_ratio is None:
        current_ratio = _safe_div(current_assets, current_liabilities)

    debt_to_equity = _get_info_first(info, ["debtToEquity"])
    if debt_to_equity is None:
        debt_to_equity = _safe_div(total_debt, equity)

    roa = _get_info_first(info, ["returnOnAssets"])
    roe = _get_info_first(info, ["returnOnEquity"])
    if roa is None or roe is None:
        net_income = _find_row_value(is_fy, ["Net Income", "NetIncome", "Net Income Common Stockholders"])
        if roa is None:
            roa = _safe_div(net_income, total_assets)
        if roe is None:
            roe = _safe_div(net_income, equity)

    ev = _get_info_first(info, ["enterpriseValue"])
    if ev is None:
        if mc is not None or total_debt is not None or total_cash is not None:
            ev = (mc or 0.0) + (total_debt or 0.0) - (total_cash or 0.0)

    ev_to_revenue = _safe_div(ev, revenue)
    ev_to_ebitda = _safe_div(ev, ebitda)

    data = {
        "ticker": ticker,
        "market_cap": mc,
        "enterprise_value": ev,
        "ev_to_revenue": ev_to_revenue,
        "ev_to_ebitda": ev_to_ebitda,
        "total_cash": total_cash,
        "total_debt": total_debt,
        "current_ratio": current_ratio,
        "debt_to_equity": debt_to_equity,
        "roa": roa,
        "roe": roe,
        "revenue_fy": revenue,
        "ebitda_fy": ebitda,
        "equity_fy": equity,
        "total_assets_fy": total_assets,
        "current_assets_fy": current_assets,
        "current_liabilities_fy": current_liabilities,
    }
    return pd.DataFrame([data])


def main():
    ap = argparse.ArgumentParser(description="Annual statements via yfinance, FY-only columns, single Excel, with Metrics.")
    ap.add_argument("--tickers", default="tickers.txt", help="Path to tickers.txt (one per line, or comma-separated).")
    ap.add_argument("--outfile", default="fundamentals.xlsx", help="Output Excel workbook path.")
    ap.add_argument("--sleep", type=float, default=1.0, help="Seconds to sleep between tickers to avoid throttling.")
    ap.add_argument("--no-index", action="store_true", help="Do not create an Index sheet.")
    args = ap.parse_args()

    if not os.path.exists(args.tickers):
        print(f"Tickers file not found: {args.tickers}", file=sys.stderr)
        sys.exit(2)

    tickers = read_tickers(args.tickers)
    if not tickers:
        print("No tickers found in the provided file.", file=sys.stderr)
        sys.exit(3)

    pd.options.display.float_format = '{:,.4f}'.format

    included = []
    with pd.ExcelWriter(args.outfile, engine="openpyxl") as writer:
        for i, tk in enumerate(tickers, 1):
            print(f"[{i}/{len(tickers)}] {tk} ...", flush=True)
            try:
                info, is_annual, bs_annual, cf_annual = fetch_annual_only(tk)
            except Exception:
                time.sleep(max(args.sleep, 0.0))
                continue

            fy_month = _infer_fy_month_from_info(info)
            if fy_month is None:
                fy_month = _infer_fy_month_from_statements([is_annual, bs_annual, cf_annual])

            is_fy = _filter_fy_columns(is_annual, fy_month)
            bs_fy = _filter_fy_columns(bs_annual, fy_month)
            cf_fy = _filter_fy_columns(cf_annual, fy_month)

            if all(df is None or df.empty for df in (is_fy, bs_fy, cf_fy)):
                time.sleep(max(args.sleep, 0.0))
                continue

            if isinstance(is_fy, pd.DataFrame) and not is_fy.empty:
                is_fy.to_excel(writer, sheet_name=f"{tk}_IS_Annual")
                included.append({"ticker": tk, "sheet": f"{tk}_IS_Annual"})
            if isinstance(bs_fy, pd.DataFrame) and not bs_fy.empty:
                bs_fy.to_excel(writer, sheet_name=f"{tk}_BS_Annual")
                included.append({"ticker": tk, "sheet": f"{tk}_BS_Annual"})
            if isinstance(cf_fy, pd.DataFrame) and not cf_fy.empty:
                cf_fy.to_excel(writer, sheet_name=f"{tk}_CF_Annual")
                included.append({"ticker": tk, "sheet": f"{tk}_CF_Annual"})

            try:
                metrics_df = compute_metrics(tk, info, is_fy, bs_fy)
                metrics_df.to_excel(writer, sheet_name=f"{tk}_Metrics", index=False)
                included.append({"ticker": tk, "sheet": f"{tk}_Metrics"})
            except Exception:
                pass

            time.sleep(max(args.sleep, 0.0))

        if not args.no_index:
            if included:
                pd.DataFrame(included).to_excel(writer, sheet_name="Index", index=False)
            else:
                pd.DataFrame({"note": ["No tickers produced FY statements."]}).to_excel(writer, sheet_name="Index", index=False)

    print(f"Wrote: {args.outfile}")


if __name__ == "__main__":
    main()
