# Italian stocks downloader

This repository contains three Python scripts for end-to-end financial data collection from **Borsa Italiana** and **Yahoo Finance**:

1. [bi_scraper.py](bi_scraper.py) — Scrapes Italian market tickers and company data (CEO, shareholders) [[Detailed readme](README_bi_scraper.py)]
2. [yf_fundamentals.py](yf_fundamentals.py) — Downloads fiscal-year-only financial statements and metrics [[Detailed readme](README_yf_fundamentals.py)]
3. [yf_prices.py](yf_fundamentals.py) — Downloads historical prices, volumes, and shares outstanding [[Detailed readme](README_yf_fundamentals.py)]

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
Downloads **fiscal-year-only** Income Statement, Balance Sheet, Cash Flow, and computed metrics for each ticker.

### Features
- FY-only filtering (excludes half-year data)
- Auto-inference of fiscal year-end
- Metrics: market cap, EV, EV/revenue, EV/EBITDA, liquidity ratios, ROA, ROE
- One Excel workbook per run, each ticker has its 4 woksheets (balancesheet, income statement, cashflow and metrics)

### Requirements
```bash
pip install yfinance pandas openpyxl pandas yfinance openpyxl
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
AAPL_Metrics
...
```
Skipped tickers produce no sheets.

## 3. Yahoo Finance Prices Downloader (`yf_prices.py`)
Downloads **historical adjusted or raw prices**, volumes, and shares outstanding for multiple tickers into one Excel file.

### Features
- Single worksheet combining all tickers
- Adjusted or raw prices (splits/dividends)
- Currency filter
- Robust to missing data

### Requirements
```bash
pip install yfinance pandas openpyxl pandas yfinance openpyxl
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

## License
MIT License.  
Data from Yahoo Finance and Borsa Italiana are subject to their respective terms of service.
