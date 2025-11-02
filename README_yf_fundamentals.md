# Manual: Yahoo Finance Fundamentals Downloader

## Overview
This script downloads **fiscal-year-only financial statements** and **key metrics** from Yahoo Finance via `yfinance`, saving everything in a single Excel workbook.  
It filters out half-year or interim data and computes ratios based only on fiscal-year-end values.

## Features
- Annual Income Statement, Balance Sheet, Cash Flow
- FY-only columns (no 30 June if FY ends in December)
- Skips tickers with no FY data
- Metrics per ticker:
  - market_cap, enterprise_value
  - ev_to_revenue, ev_to_ebitda
  - total_cash, total_debt
  - current_ratio, debt_to_equity
  - roa, roe
  - revenue_fy, ebitda_fy, equity_fy, total_assets_fy, current_assets_fy, current_liabilities_fy
- Optional `Index` sheet listing all sheets

## Requirements
Python 3.9+  
Dependencies:
```bash
pip install yfinance pandas openpyxl
```

## Usage
```bash
python yf_fundamentals.py --tickers tickers.txt --outfile fundamentals.xlsx --sleep 1.0
```

**Arguments**
| Option | Description | Default |
|---------|-------------|----------|
| `--tickers` | Path to tickers list | `tickers.txt` |
| `--outfile` | Output Excel file | `fundamentals.xlsx` |
| `--sleep` | Seconds between requests | `1.0` |
| `--no-index` | Skip Index sheet | false |

**tickers.txt** can be formatted as:
```
AAPL
MSFT
ENI.MI, ISP.MI
```
Comma or newline separation is accepted.

## FY Filtering Logic
1. Detects fiscal year-end month from Yahoo Finance metadata (`fiscalYearEnd`).
2. If missing, infers FY-end from the mode of column months.
3. Retains only columns ending in that month (e.g., only December if FY = Dec).
4. Excludes all other columns (half-year or quarterly data).

## Output Example
```
Index
AAPL_IS_Annual
AAPL_BS_Annual
AAPL_CF_Annual
AAPL_Metrics
MSFT_IS_Annual
MSFT_BS_Annual
MSFT_CF_Annual
MSFT_Metrics
```

## Notes
- If no FY data exists, the ticker is skipped entirely.
- If none succeed, a file with a single “No data” sheet is produced.
- For many tickers, increase sleep time to 2–3 seconds to avoid throttling.
- Fiscal-year inference is heuristic; verify results on the first run.
