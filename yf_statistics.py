#!/usr/bin/env python3
"""
Yahoo Finance Key Statistics Downloader

Downloads all available key statistics from Yahoo Finance for multiple tickers.
Extracts data from the key-statistics page including valuation metrics,
financial ratios, profitability metrics, and other fundamental indicators.

Usage:
  python yf_statistics.py --tickers tickers.txt --outfile statistics.xlsx --sleep 1.0
"""

import argparse
import time
import sys
import os
import re
from typing import List, Dict, Any, Optional
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


def safe_float(value: Any) -> Optional[float]:
    """Convert value to float safely, handling various formats."""
    if value is None or pd.isna(value):
        return None
    try:
        if isinstance(value, str):
            # Remove common formatting characters
            cleaned = re.sub(r'[,%$€£¥]', '', value.strip())
            if cleaned in ('N/A', 'NA', '', '-'):
                return None
            return float(cleaned)
        return float(value)
    except (ValueError, TypeError):
        return None


def compute_ttm(ticker_obj: yf.Ticker, row_aliases: List[str]) -> Optional[float]:
    """Sum the 4 most-recent quarters for a given row to produce a TTM figure.

    Returns None if fewer than 4 quarters are available or the row is not found.
    """
    try:
        qf = ticker_obj.quarterly_financials
        if qf is None or qf.empty:
            return None
        # Normalise index to lowercase strings for alias matching
        idx_map = {str(ix).strip().lower(): ix for ix in qf.index}
        matched_row = None
        for alias in row_aliases:
            key = alias.strip().lower()
            if key in idx_map:
                matched_row = qf.loc[idx_map[key]]
                break
        if matched_row is None:
            return None
        # Sort columns descending (most recent first) and take 4
        dt_cols = pd.to_datetime(matched_row.index, errors='coerce')
        order = dt_cols.argsort()[::-1]
        sorted_vals = matched_row.iloc[order]
        recent_4 = sorted_vals.iloc[:4]
        if len(recent_4) < 4:
            return None
        numeric = pd.to_numeric(recent_4, errors='coerce')
        if numeric.isna().any():
            return None
        return float(numeric.sum())
    except Exception:
        return None


def extract_info_stats(ticker_obj: yf.Ticker) -> Dict[str, Any]:
    """Extract statistics that are available in the main info dictionary."""
    info = ticker_obj.info or {}
    
    stats = {}
    
    # Market Cap & Enterprise Value
    stats['market_cap'] = safe_float(info.get('marketCap'))
    ev = safe_float(info.get('enterpriseValue'))
    stats['enterprise_value'] = ev

    ttm_revenue = compute_ttm(ticker_obj, ["Total Revenue", "TotalRevenue", "Revenue"])
    ttm_ebitda = compute_ttm(ticker_obj, ["EBITDA", "Ebitda"])
    stats['ttm_revenue'] = ttm_revenue
    stats['ttm_ebitda'] = ttm_ebitda

    if ev is not None and ttm_revenue is not None and ttm_revenue != 0:
        stats['enterprise_to_revenue'] = ev / ttm_revenue
    else:
        stats['enterprise_to_revenue'] = safe_float(info.get('enterpriseToRevenue'))

    if ev is not None and ttm_ebitda is not None and ttm_ebitda != 0:
        stats['enterprise_to_ebitda'] = ev / ttm_ebitda
    else:
        stats['enterprise_to_ebitda'] = safe_float(info.get('enterpriseToEbitda'))
    
    # Valuation Ratios
    stats['trailing_pe'] = safe_float(info.get('trailingPE'))
    stats['forward_pe'] = safe_float(info.get('forwardPE'))
    stats['price_to_sales'] = safe_float(info.get('priceToSalesTrailing12Months'))
    stats['price_to_book'] = safe_float(info.get('priceToBook'))
    stats['peg_ratio'] = safe_float(info.get('pegRatio'))
    
    # Dividend Information
    stats['dividend_yield'] = safe_float(info.get('dividendYield'))
    stats['dividend_rate'] = safe_float(info.get('dividendRate'))
    stats['payout_ratio'] = safe_float(info.get('payoutRatio'))
    
    # Financial Health
    stats['total_debt_to_equity'] = safe_float(info.get('debtToEquity'))
    stats['current_ratio'] = safe_float(info.get('currentRatio'))
    stats['quick_ratio'] = safe_float(info.get('quickRatio'))
    stats['cash_ratio'] = safe_float(info.get('cashRatio'))
    
    # Profitability
    stats['return_on_assets'] = safe_float(info.get('returnOnAssets'))
    stats['return_on_equity'] = safe_float(info.get('returnOnEquity'))
    stats['gross_margins'] = safe_float(info.get('grossMargins'))
    stats['operating_margins'] = safe_float(info.get('operatingMargins'))
    stats['profit_margins'] = safe_float(info.get('profitMargins'))
    
    # Efficiency
    stats['revenue_per_share'] = safe_float(info.get('revenuePerShare'))
    stats['earnings_per_share'] = safe_float(info.get('trailingEps'))
    stats['book_value_per_share'] = safe_float(info.get('bookValue'))
    
    # Growth Rates
    stats['revenue_growth'] = safe_float(info.get('revenueGrowth'))
    stats['earnings_growth'] = safe_float(info.get('earningsGrowth'))
    stats['earnings_quarterly_growth'] = safe_float(info.get('earningsQuarterlyGrowth'))
    
    # Balance Sheet Items
    stats['total_cash'] = safe_float(info.get('totalCash'))
    stats['total_debt'] = safe_float(info.get('totalDebt'))
    stats['total_revenue'] = ttm_revenue if ttm_revenue is not None else safe_float(info.get('totalRevenue'))
    stats['total_assets'] = safe_float(info.get('totalAssets'))
    stats['total_stockholder_equity'] = safe_float(info.get('totalStockholderEquity'))
    
    # Other Metrics
    stats['beta'] = safe_float(info.get('beta'))
    stats['52_week_change'] = safe_float(info.get('52WeekChange'))
    stats['shares_outstanding'] = safe_float(info.get('sharesOutstanding'))
    stats['float_shares'] = safe_float(info.get('floatShares'))
    stats['shares_short'] = safe_float(info.get('sharesShort'))
    stats['short_ratio'] = safe_float(info.get('shortRatio'))
    
    return stats


def extract_financial_ratios(ticker_obj: yf.Ticker) -> Dict[str, Any]:
    """Extract additional financial ratios from financial statements."""
    ratios = {}
    
    try:
        # Get most recent annual and quarterly data
        bs_annual = ticker_obj.balance_sheet
        bs_quarterly = ticker_obj.quarterly_balance_sheet
        is_annual = ticker_obj.financials
        is_quarterly = ticker_obj.quarterly_financials
        
        # Calculate additional ratios if data is available
        if bs_annual is not None and not bs_annual.empty:
            # Use the most recent column (first column)
            recent_col = bs_annual.columns[0]
            
            # Get key balance sheet items
            total_assets = safe_float(bs_annual.loc['Total Assets', recent_col] if 'Total Assets' in bs_annual.index else None)
            total_liabilities = safe_float(bs_annual.loc['Total Liabilities', recent_col] if 'Total Liabilities' in bs_annual.index else None)
            current_assets = safe_float(bs_annual.loc['Total Current Assets', recent_col] if 'Total Current Assets' in bs_annual.index else None)
            current_liabilities = safe_float(bs_annual.loc['Total Current Liabilities', recent_col] if 'Total Current Liabilities' in bs_annual.index else None)
            
            # Calculate ratios
            if total_assets and total_liabilities:
                ratios['debt_to_assets'] = total_liabilities / total_assets
            if current_assets and current_liabilities:
                ratios['working_capital'] = current_assets - current_liabilities
        
        if is_annual is not None and not is_annual.empty:
            recent_col = is_annual.columns[0]
            
            # Get income statement items
            revenue = safe_float(is_annual.loc['Total Revenue', recent_col] if 'Total Revenue' in is_annual.index else None)
            net_income = safe_float(is_annual.loc['Net Income', recent_col] if 'Net Income' in is_annual.index else None)
            ebitda = safe_float(is_annual.loc['EBITDA', recent_col] if 'EBITDA' in is_annual.index else None)
            
            # Calculate profitability ratios
            if revenue and net_income:
                ratios['net_profit_margin'] = net_income / revenue
            if revenue and ebitda:
                ratios['ebitda_margin'] = ebitda / revenue
                
    except Exception as e:
        # If calculation fails, continue without these ratios
        pass
    
    return ratios


def fetch_statistics(ticker: str) -> Dict[str, Any]:
    """Fetch all available statistics for a ticker."""
    try:
        ticker_obj = yf.Ticker(ticker)
        
        # Extract statistics from info
        stats = extract_info_stats(ticker_obj)
        
        # Add ticker symbol
        stats['ticker'] = ticker
        
        # Extract additional ratios from financial statements
        additional_ratios = extract_financial_ratios(ticker_obj)
        stats.update(additional_ratios)
        
        # Add timestamp
        stats['fetch_timestamp'] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        
        return stats
        
    except Exception as e:
        print(f"Error fetching statistics for {ticker}: {e}", file=sys.stderr)
        return {'ticker': ticker, 'error': str(e)}


def main():
    ap = argparse.ArgumentParser(description="Download key statistics from Yahoo Finance for multiple tickers.")
    ap.add_argument("--tickers", default="tickers.txt", help="Path to tickers.txt (one per line, or comma-separated).")
    ap.add_argument("--outfile", default="statistics.xlsx", help="Output Excel workbook path.")
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

    all_stats = []
    failed_tickers = []

    print(f"Fetching statistics for {len(tickers)} tickers...")
    
    for i, ticker in enumerate(tickers, 1):
        print(f"[{i}/{len(tickers)}] {ticker} ...", flush=True)
        
        try:
            stats = fetch_statistics(ticker)
            if 'error' not in stats:
                all_stats.append(stats)
            else:
                failed_tickers.append(ticker)
        except Exception as e:
            print(f"Failed to fetch {ticker}: {e}", file=sys.stderr)
            failed_tickers.append(ticker)
        
        # Sleep to avoid rate limiting
        time.sleep(max(args.sleep, 0.0))

    # Create output Excel file
    with pd.ExcelWriter(args.outfile, engine="openpyxl") as writer:
        if all_stats:
            # Main statistics sheet
            df_stats = pd.DataFrame(all_stats)
            
            # Reorder columns to put ticker first
            if 'ticker' in df_stats.columns:
                cols = ['ticker'] + [col for col in df_stats.columns if col != 'ticker']
                df_stats = df_stats[cols]
            
            df_stats.to_excel(writer, sheet_name="Statistics", index=False)
            
            # Create summary statistics sheet
            summary_data = {
                'Metric': [],
                'Count': [],
                'Mean': [],
                'Median': [],
                'Std Dev': [],
                'Min': [],
                'Max': []
            }
            
            numeric_cols = [col for col in df_stats.columns if col not in ['ticker', 'fetch_timestamp', 'error']]
            
            for col in numeric_cols:
                series = df_stats[col].dropna()
                if len(series) > 0:
                    summary_data['Metric'].append(col)
                    summary_data['Count'].append(len(series))
                    summary_data['Mean'].append(series.mean())
                    summary_data['Median'].append(series.median())
                    summary_data['Std Dev'].append(series.std())
                    summary_data['Min'].append(series.min())
                    summary_data['Max'].append(series.max())
            
            if summary_data['Metric']:
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name="Summary", index=False)
        
        # Failed tickers sheet
        if failed_tickers:
            df_failed = pd.DataFrame({'ticker': failed_tickers})
            df_failed.to_excel(writer, sheet_name="Failed", index=False)
        
        # Index sheet
        if not args.no_index:
            index_data = []
            if all_stats:
                index_data.append({'sheet': 'Statistics', 'description': f'Statistics for {len(all_stats)} tickers'})
            if len(all_stats) > 0:
                index_data.append({'sheet': 'Summary', 'description': 'Summary statistics across all tickers'})
            if failed_tickers:
                index_data.append({'sheet': 'Failed', 'description': f'{len(failed_tickers)} failed tickers'})
            
            if index_data:
                pd.DataFrame(index_data).to_excel(writer, sheet_name="Index", index=False)
            else:
                pd.DataFrame({"note": ["No data fetched."]}).to_excel(writer, sheet_name="Index", index=False)

    print(f"Completed: {len(all_stats)} successful, {len(failed_tickers)} failed")
    print(f"Output written to: {args.outfile}")


if __name__ == "__main__":
    main()
