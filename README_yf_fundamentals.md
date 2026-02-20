# Manual: Yahoo Finance Fundamentals Downloader

## Overview
This script downloads **all financial statements** (annual and quarterly) from Yahoo Finance via `yfinance`, saving everything in a single Excel workbook.  
It applies fiscal-year filtering only to annual data and includes all quarterly data without filtering, then computes comprehensive ratios.

## Features
- **Annual Statements**: Income Statement, Balance Sheet, Cash Flow (FY-only columns)
- **Quarterly Statements**: Income Statement, Balance Sheet, Cash Flow (all data)
- FY filtering for annual data only (no 30 June if FY ends in December)
- Skips tickers with no financial data
- Metrics per ticker:
  - market_cap, enterprise_value
  - ev_to_revenue, ev_to_ebitda
  - total_cash, total_debt
  - current_ratio, debt_to_equity
  - roa, roe
  - **Annual values**: revenue_fy, ebitda_fy, net_income_fy, equity_fy, total_assets_fy, current_assets_fy, current_liabilities_fy
  - **Quarterly values**: revenue_q, ebitda_q, net_income_q, equity_q, total_assets_q, current_assets_q, current_liabilities_q
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
Comma or newline separation is accepted.

## FY Filtering Logic
1. Detects fiscal year-end month from Yahoo Finance metadata (`fiscalYearEnd`).
2. If missing, infers FY-end from the mode of annual statement column months.
3. Retains only annual columns ending in that month (e.g., only December if FY = Dec).
4. **Quarterly data is included without any FY filtering**.
5. Excludes all other annual columns (half-year or quarterly data).

## Output Example
```
Index
AAPL_IS_Annual
AAPL_BS_Annual
AAPL_CF_Annual
AAPL_IS_Quarterly
AAPL_BS_Quarterly
AAPL_CF_Quarterly
AAPL_Metrics
MSFT_IS_Annual
MSFT_BS_Annual
MSFT_CF_Annual
MSFT_IS_Quarterly
MSFT_BS_Quarterly
MSFT_CF_Quarterly
MSFT_Metrics
```

## Notes
- If no financial data exists, the ticker is skipped entirely.
- If none succeed, a file with a single "No data" sheet is produced.
- For many tickers, increase sleep time to 2–3 seconds to avoid throttling.
- Fiscal-year inference is heuristic; verify results on the first run.
- Metrics prioritize annual values but fall back to quarterly if annual is unavailable.
