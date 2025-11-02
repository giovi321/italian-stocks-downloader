# yf_prices.py Manual
## Overview
`yf_prices.py` downloads **historical prices**, **volumes**, and **shares outstanding** for multiple tickers using [yfinance](https://github.com/ranaroussi/yfinance).  
It outputs a **single Excel workbook** (`.xlsx`) containing one worksheet named `data`.

By default, prices are **adjusted** for stock splits and dividends.  
Optionally, you can request raw (unadjusted) prices.

## Features
- One combined Excel sheet for all tickers.
- Adjusted or raw prices.
- Shares outstanding (historical if available).
- Volume.
- Currency filter.
- Robust fallback when historical data are partially missing.

## Installation
### Requirements
Install dependencies:
```bash
pip install yfinance pandas openpyxl
```

## Usage
### Command-line Syntax
```bash
python yf_prices.py [OPTIONS]
```

### Main Options
| Option | Description |
|--|-|
| `--tickers-file PATH` | Path to a text file listing tickers (default: `tickers.txt`) |
| `--start YYYY-MM-DD` | Start date for data |
| `--end YYYY-MM-DD` | End date (exclusive) |
| `--freq {daily,weekly,monthly,quarterly}` | Data frequency (default: `daily`) |
| `--outfile PATH` | Output Excel path (default: `output/prices_combined.xlsx`) |
| `--currency CODE` | Filter to tickers of a given currency (e.g. `EUR`) |
| `--adjusted` | Use adjusted prices (default) |
| `--no-adjusted` | Use raw prices (no split/dividend adjustment) |

## Examples
### 1. Adjusted monthly prices for EUR stocks
```bash
python yf_prices.py --tickers-file tickers.txt \
  --start 2022-01-01 --end 2025-10-31 \
  --freq monthly --currency EUR \
  --outfile output/prices_adjusted.xlsx
```

### 2. Raw weekly prices (no adjustment)
```bash
python yf_prices.py --tickers-file tickers.txt \
  --freq weekly --no-adjusted \
  --outfile output/prices_raw.xlsx
```

## Input Format
`tickers.txt` should list one or more Yahoo Finance tickers.  
Commas, spaces, or line breaks are accepted.  
Lines starting with `#` are ignored.

Example:
```
# Example ticker list
AAPL
MSFT, NVDA
ENEL.MI
```

## Output Format
The generated Excel file has one worksheet named `data`.

Columns:

| Column | Description |
|---------|-------------|
| `Ticker` | Ticker symbol |
| `Date` | Observation date |
| `Price` | Adjusted or raw closing price |
| `Volume` | Trading volume |
| `SharesOutstanding` | Number of shares outstanding |
| `Currency` | Trading currency |
| `RawClose` | Raw close (if `--no-adjusted`) |
| `RawAdjClose` | Adjusted close (Yahoo value, if `--no-adjusted`) |

## How Adjustments Work
- With `--adjusted` (default):  
  Uses `auto_adjust=True` in `yfinance`, which adjusts OHLC for **splits and dividends**.  
  `Price` equals adjusted close.

- With `--no-adjusted`:  
  Fetches raw prices.  
  Includes both Yahoo’s raw `Close` and `Adj Close` columns for comparison.

Volume is not adjusted by Yahoo Finance and always represents raw traded volume.

## Shares Outstanding
- Tries to fetch **historical shares outstanding** with `Ticker.get_shares_full()`.  
- If unavailable, uses the **current** `shares_outstanding` from Yahoo’s `fast_info` or `info`.
- When neither is available, the field is left blank (`NaN`).

## Troubleshooting
- If some tickers show `[ERR]`, they may have been delisted or renamed.
- For newly listed stocks, data before the IPO will be empty.
- Check currency mismatches if values seem inconsistent.
- If Excel output is incomplete, verify `--start` and `--end` cover valid trading dates.
