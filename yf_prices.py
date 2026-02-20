#!/usr/bin/env python3
"""
yf_prices.py
Download prices, volumes, and shares outstanding via yfinance into ONE Excel worksheet.

- Adjusted prices for splits/dividends (default).
- Includes SharesOutstanding using historical get_shares_full() if available,
  else constant fast_info.shares_outstanding or info['sharesOutstanding'].
- Optional raw prices with --no-adjusted.
- Optional currency filter.

Columns:
Ticker, Date, Price, Volume, SharesOutstanding, Currency, [RawClose], [RawAdjClose]
"""

import argparse
import sys
import os
from typing import List, Optional
import pandas as pd
import yfinance as yf

FREQ_TO_INTERVAL = {
    "daily": "1d",
    "weekly": "1wk",
    "monthly": "1mo",
    "quarterly": "3mo",
}

def _clean_ticker(raw: str) -> str:
    t = raw.strip().upper()
    if t.startswith("$"):
        t = t[1:]
    return t

def parse_tickers_file(path: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Tickers file not found: {path}")
    raw = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.replace(",", " ").split() if p.strip()]
            raw.extend(parts)
    seen, out = set(), []
    for t in raw:
        key = _clean_ticker(t)
        if key and key not in seen:
            seen.add(key)
            out.append(key)
    return out

def _slice_to_overlap(df: pd.DataFrame, start: Optional[str], end: Optional[str]) -> pd.DataFrame:
    if df.empty:
        return df
    idx = pd.to_datetime(df.index).tz_localize(None)
    left = pd.to_datetime(start) if start else idx.min()
    right = pd.to_datetime(end) if end else idx.max()
    return df.loc[(idx >= left) & (idx <= right)]

def _get_history_with_fallback(tkr: yf.Ticker, start: Optional[str], end: Optional[str], interval: str, auto_adjust: bool) -> pd.DataFrame:
    hist_kwargs = {"interval": interval, "auto_adjust": auto_adjust}
    if start or end:
        hist_kwargs["start"] = start
        hist_kwargs["end"] = end
        df = tkr.history(**hist_kwargs)
        if df is not None and not df.empty:
            return df
    df_all = tkr.history(period="max", interval=interval, auto_adjust=auto_adjust)
    if df_all is None or df_all.empty:
        return pd.DataFrame()
    if start or end:
        return _slice_to_overlap(df_all, start, end)
    return df_all

def _get_shares_series(tkr: yf.Ticker, start: Optional[str], end: Optional[str], index: pd.DatetimeIndex) -> pd.Series:
    try:
        s = tkr.get_shares_full(start=start, end=end)
        if isinstance(s, pd.Series) and len(s) > 0:
            s = s.sort_index().astype(float)
            return s.reindex(index, method="ffill")
    except Exception:
        pass
    val = None
    try:
        val = getattr(tkr, "fast_info", {}).get("shares_outstanding", None)
    except Exception:
        pass
    if val is None:
        try:
            val = (tkr.info or {}).get("sharesOutstanding", None)
        except Exception:
            val = None
    if val is None:
        return pd.Series(index=index, dtype=float)
    return pd.Series(float(val), index=index)

def fetch_one(ticker: str, start: Optional[str], end: Optional[str], interval: str,
              currency_filter: Optional[str], adjusted: bool) -> pd.DataFrame:
    tkr = yf.Ticker(ticker)
    fi = getattr(tkr, "fast_info", {}) or {}
    try:
        info = tkr.info or {}
    except Exception:
        info = {}
    currency = fi.get("currency") or info.get("currency")
    if currency_filter and currency and currency.upper() != currency_filter.upper():
        raise ValueError(f"Currency {currency} != filter {currency_filter}")
    data = _get_history_with_fallback(tkr, start, end, interval, auto_adjust=adjusted)
    if data is None or data.empty:
        raise ValueError(f"No data for {ticker}")
    data.index = pd.to_datetime(data.index).tz_localize(None)
    shares = _get_shares_series(tkr, start, end, data.index)
    out = pd.DataFrame(index=data.index)
    out["Ticker"] = ticker
    out["Date"] = out.index
    out["Currency"] = currency
    out["Volume"] = data["Volume"].astype(float)
    out["SharesOutstanding"] = shares
    if adjusted:
        out["Price"] = data["Close"].astype(float)
    else:
        out["RawClose"] = data["Close"].astype(float)
        if "Adj Close" in data.columns:
            out["RawAdjClose"] = data["Adj Close"].astype(float)
        out["Price"] = out["RawClose"]
    return out.reset_index(drop=True)

def write_combined_one_sheet(path: str, frames: list) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined.sort_values(["Ticker", "Date"], inplace=True)
    cols = ["Ticker", "Date", "Price", "Volume", "SharesOutstanding", "Currency", "RawClose", "RawAdjClose"]
    cols = [c for c in cols if c in combined.columns] + [c for c in combined.columns if c not in cols]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        combined[cols].to_excel(xw, sheet_name="data", index=False)

def main():
    ap = argparse.ArgumentParser(description="Download prices, volumes, and shares outstanding via yfinance")
    ap.add_argument("--tickers-file", default="tickers.txt", help="Path to tickers file")
    ap.add_argument("--start", type=str, default=None, help="Start date YYYY-MM-DD")
    ap.add_argument("--end", type=str, default=None, help="End date YYYY-MM-DD")
    ap.add_argument("--freq", type=str, choices=list(FREQ_TO_INTERVAL.keys()), default="daily")
    ap.add_argument("--outfile", type=str, default="prices.xlsx")
    ap.add_argument("--currency", type=str, default=None, help="Filter by currency (e.g., EUR)")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--adjusted", dest="adjusted", action="store_true")
    g.add_argument("--no-adjusted", dest="adjusted", action="store_false")
    ap.set_defaults(adjusted=True)
    args = ap.parse_args()

    tickers = parse_tickers_file(args.tickers_file)
    if not tickers:
        print("No tickers found", file=sys.stderr)
        sys.exit(1)
    interval = FREQ_TO_INTERVAL[args.freq]
    frames, errors = [], {}
    for tk in tickers:
        try:
            df = fetch_one(tk, args.start, args.end, interval, args.currency, args.adjusted)
            frames.append(df)
            print(f"[OK] {tk}: {len(df)} rows")
        except Exception as e:
            errors[tk] = str(e)
            print(f"[ERR] {tk}: {e}", file=sys.stderr)
    if not frames:
        print("No data retrieved", file=sys.stderr)
        sys.exit(1)
    write_combined_one_sheet(args.outfile, frames)
    print(f"Wrote combined workbook: {args.outfile}")
    if errors:
        print("Errors:", file=sys.stderr)
        for tk, msg in errors.items():
            print(f"- {tk}: {msg}", file=sys.stderr)

if __name__ == "__main__":
    main()
