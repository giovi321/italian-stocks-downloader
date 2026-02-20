# Italian stocks downloader

This repository contains four Python scripts for end-to-end financial data collection from **Borsa Italiana** and **Yahoo Finance**:

1. [bi_scraper.py](bi_scraper.py) — Scrapes Italian market tickers and company data (CEO, shareholders) [[Detailed readme](README_bi_scraper.md)]
2. [yf_fundamentals.py](yf_fundamentals.py) — Downloads all financial statements (annual + quarterly) and metrics [[Detailed readme](README_yf_fundamentals.md)]
3. [yf_statistics.py](yf_statistics.py) — Downloads key statistics and financial ratios [[Detailed readme](README_yf_statistics.md)]
4. [yf_prices.py](yf_prices.py) — Downloads historical prices, volumes, and shares outstanding [[Detailed readme](README_yf_prices.md)]

Each script outputs clean, analysis-ready Excel or CSV files.

## 1. Borsa Italiana Ticker Scraper (`bi_scraper.py`)
Scrapes all tickers and company data from:
- Euronext Growth Milan  
- Euronext STAR Milan  
- Mid Cap  
- Small Cap  

**Extracted fields:**  
Ticker, ISIN, company name, CEO, and shareholding details.

### Requirements
```bash
pip install requests beautifulsoup4 lxml pandas yfinance openpyxl
```

### Usage
```bash
python bi_scraper.py
```

**Output:**
- `borsaitaliana_tickers.csv`
- `borsaitaliana_tickers.json`

**Example columns:**
| segment | ticker | isin | company | ceo_name | shareholding |
|----------|--------|------|----------|-----------|--------------|
| mid-cap | ABCD | IT0001234567 | ABC S.p.A. | Mario Rossi | Free float: 45.00 %; XYZ Holding: 30.00 % |

## 2. Yahoo Finance Fundamentals Downloader (`yf_fundamentals.py`)
Downloads **all financial statements** (annual + quarterly) with fiscal-year filtering for annual data and comprehensive metrics.

### Features
- **Annual Statements**: Income Statement, Balance Sheet, Cash Flow (FY-only columns)
- **Quarterly Statements**: Income Statement, Balance Sheet, Cash Flow (all data)
- FY filtering for annual data only (excludes half-year data)
- Auto-inference of fiscal year-end
- Comprehensive metrics: market cap, EV, EV/revenue, EV/EBITDA, liquidity ratios, ROA, ROE
- Separate annual and quarterly values in metrics sheet
- One Excel workbook per run, each ticker has 6 worksheets (3 annual + 3 quarterly + metrics)

### Requirements
```bash
pip install yfinance pandas openpyxl
```

### Usage
```bash
python yf_fundamentals.py --tickers tickers.txt --outfile fundamentals.xlsx --sleep 1.0
```

**tickers.txt**
```
AAPL
MSFT
ENI.MI, ISP.MI
```

**Example workbook:**
```
Index
AAPL_IS_Annual
AAPL_BS_Annual
AAPL_CF_Annual
AAPL_IS_Quarterly
AAPL_BS_Quarterly
AAPL_CF_Quarterly
AAPL_Metrics
...
```
Skipped tickers produce no sheets.

## 3. Yahoo Finance Key Statistics Downloader (`yf_statistics.py`)
Downloads **comprehensive key statistics** and financial ratios from Yahoo Finance for multiple tickers.

### Features
- **Valuation Metrics**: Market cap, enterprise value, P/E ratios, price-to-sales/book, PEG ratio
- **Financial Health**: Debt ratios, liquidity ratios (current, quick, cash), profitability metrics
- **Growth Metrics**: Revenue growth, earnings growth, quarterly growth rates
- **Market Data**: Beta, 52-week change, shares outstanding, short interest
- **Balance Sheet Data**: Total assets, debt, cash, equity
- **Summary Statistics**: Cross-ticker analysis with mean, median, standard deviation
- **Error Handling**: Tracks failed tickers with separate logging

### Requirements
```bash
pip install yfinance pandas openpyxl
```

### Usage
```bash
python yf_statistics.py --tickers tickers.txt --outfile statistics.xlsx --sleep 1.0
```

**Example workbook:**
```
Index
Statistics
Summary
Failed (if any)
```

## 4. Yahoo Finance Prices Downloader (`yf_prices.py`)
Downloads **historical adjusted or raw prices**, volumes, and shares outstanding for multiple tickers into one Excel file.

### Features
- Single worksheet combining all tickers
- Adjusted or raw prices (splits/dividends)
- Currency filter
- Robust to missing data

### Requirements
```bash
pip install yfinance pandas openpyxl
```

### Usage
```bash
python yf_prices.py --tickers-file tickers.txt --start 2022-01-01 --end 2025-10-31 --freq monthly --currency EUR --outfile output/prices_combined.xlsx
```

**Excel columns:**
| Ticker | Date | Price | Volume | SharesOutstanding | Currency | RawClose | RawAdjClose |

### Notes
- Uses `auto_adjust=True` for split/dividend correction.
- Historical shares outstanding where available, otherwise latest.

## Integration Guide

The scripts are designed to work together for comprehensive financial analysis:

1. **Start with bi_scraper.py** to get Italian market tickers
2. **Use yf_fundamentals.py** for raw financial statements (IS, BS, CF)
3. **Use yf_statistics.py** for key ratios and market metrics
4. **Use yf_prices.py** for historical price data

### Example Workflow
```bash
# 1. Get Italian tickers
python bi_scraper.py

# 2. Download all financial statements
python yf_fundamentals.py --tickers borsaitaliana_tickers.csv --outfile fundamentals.xlsx

# 3. Download key statistics
python yf_statistics.py --tickers borsaitaliana_tickers.csv --outfile statistics.xlsx

# 4. Download historical prices
python yf_prices.py --tickers-file borsaitaliana_tickers.csv --start 2022-01-01 --end 2025-10-31 --outfile prices.xlsx
```

## License
MIT License.  
Data from Yahoo Finance and Borsa Italiana are subject to their respective terms of service.
