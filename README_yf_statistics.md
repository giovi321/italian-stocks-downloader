# Manual: Yahoo Finance Key Statistics Downloader

## Overview
This script downloads **key statistics and financial ratios** from Yahoo Finance for multiple tickers, extracting comprehensive fundamental data similar to the "Statistics" tab on Yahoo Finance pages. It saves all data in a single Excel workbook with multiple sheets for detailed analysis.

## Features
- **Comprehensive Statistics**: Market cap, enterprise value, valuation ratios, dividend metrics
- **Financial Health**: Debt ratios, liquidity ratios, profitability metrics
- **Growth Metrics**: Revenue growth, earnings growth, quarterly growth rates
- **Balance Sheet Data**: Total assets, debt, cash, equity
- **Efficiency Ratios**: Revenue per share, book value per share, EPS
- **Market Metrics**: Beta, 52-week change, shares outstanding, short interest
- **Additional Calculated Ratios**: Debt-to-assets, working capital, profit margins
- **Summary Statistics**: Cross-ticker analysis with mean, median, std dev
- **Error Handling**: Tracks failed tickers with separate logging

## Requirements
Python 3.9+  
Dependencies:
```bash
pip install yfinance pandas openpyxl
```

## Usage
```bash
python yf_statistics.py --tickers tickers.txt --outfile statistics.xlsx --sleep 1.0
```

**Arguments**
| Option | Description | Default |
|---------|-------------|----------|
| `--tickers` | Path to tickers list | `tickers.txt` |
| `--outfile` | Output Excel file | `statistics.xlsx` |
| `--sleep` | Seconds between requests | `1.0` |
| `--no-index` | Skip Index sheet | false |

**tickers.txt** can be formatted as:
```
AAPL
MSFT
ENI.MI, ISP.MI
```
Comma or newline separation is accepted.

## Output Structure
The script creates an Excel workbook with the following sheets:

### Statistics Sheet
Contains all extracted statistics for each ticker:
- **Valuation**: market_cap, enterprise_value, enterprise_to_revenue, enterprise_to_ebitda, trailing_pe, forward_pe, price_to_sales, price_to_book, peg_ratio
- **TTM Denominators**: ttm_revenue, ttm_ebitda (sum of 4 most-recent quarters; used for enterprise_to_revenue and enterprise_to_ebitda)
- **Dividends**: dividend_yield, dividend_rate, payout_ratio
- **Financial Health**: total_debt_to_equity, current_ratio, quick_ratio, cash_ratio
- **Profitability**: return_on_assets, return_on_equity, gross_margins, operating_margins, profit_margins
- **Efficiency**: revenue_per_share, earnings_per_share, book_value_per_share
- **Growth**: revenue_growth, earnings_growth, earnings_quarterly_growth
- **Balance Sheet**: total_cash, total_debt, total_revenue (TTM), total_assets, total_stockholder_equity
- **Market Metrics**: beta, 52_week_change, shares_outstanding, float_shares, shares_short, short_ratio
- **Calculated Ratios**: debt_to_assets, working_capital, net_profit_margin, ebitda_margin

### Summary Sheet
Statistical summary across all successful tickers:
- Count, Mean, Median, Standard Deviation, Min, Max for each numeric metric
- Useful for quick market analysis and benchmarking

### Failed Sheet
Lists tickers that failed to fetch data (if any)

### Index Sheet
Overview of all sheets and their contents

## Data Sources
The script extracts data from multiple Yahoo Finance sources:
1. **Primary**: `ticker.info` dictionary (most comprehensive)
2. **TTM**: `ticker.quarterly_financials` — 4 most-recent quarters summed for revenue and EBITDA
3. **Secondary**: Financial statements for calculated ratios
4. **Calculated**: Derived metrics using multiple data points

## Key Metrics Explained

### Valuation Metrics
- **Market Cap**: Total market value of all outstanding shares
- **Enterprise Value**: Market cap + debt - cash (total company value)
- **EV/Revenue**: Enterprise value ÷ TTM revenue (sum of 4 most-recent quarters). Falls back to `enterpriseToRevenue` from `ticker.info` only if TTM data is unavailable.
- **EV/EBITDA**: Enterprise value ÷ TTM EBITDA (sum of 4 most-recent quarters). Falls back to `enterpriseToEbitda` from `ticker.info` only if TTM data is unavailable.
- **ttm_revenue / ttm_ebitda**: Raw TTM denominators, exposed as output columns so you can verify the ratio calculations.
- **total_revenue**: Populated from TTM revenue; falls back to `totalRevenue` from `ticker.info` only if TTM data is unavailable.
- **P/E Ratios**: Price-to-earnings (trailing 12 months vs forward estimates)
- **Price-to-Sales**: Market cap relative to revenue
- **Price-to-Book**: Market cap relative to book value
- **PEG Ratio**: P/E ratio relative to earnings growth

### Financial Health
- **Debt-to-Equity**: Total debt divided by shareholders' equity
- **Current Ratio**: Current assets ÷ current liabilities (liquidity)
- **Quick Ratio**: (Current assets - inventory) ÷ current liabilities
- **Cash Ratio**: Cash ÷ current liabilities (most conservative liquidity measure)

### Profitability
- **ROA**: Return on Assets (net income ÷ total assets)
- **ROE**: Return on Equity (net income ÷ shareholders' equity)
- **Gross Margin**: (Revenue - COGS) ÷ revenue
- **Operating Margin**: Operating income ÷ revenue
- **Profit Margin**: Net income ÷ revenue

### Growth Metrics
- **Revenue Growth**: Year-over-year revenue growth rate
- **Earnings Growth**: Year-over-year earnings growth rate
- **Quarterly Earnings Growth**: Most recent quarter vs same quarter last year

## Example Output
```
Index
Statistics
Summary
Failed (if any)
```

Sample Statistics sheet rows:
| ticker | market_cap | trailing_pe | dividend_yield | debt_to_equity | return_on_equity |
|--------|------------|-------------|----------------|----------------|------------------|
| AAPL   | 2.8T       | 28.5        | 0.52           | 1.73           | 147.3            |
| MSFT   | 2.1T       | 31.2        | 0.75           | 0.47           | 39.7             |

## Notes
- **Data Availability**: Not all metrics are available for all tickers (especially international stocks)
- **Data Freshness**: Yahoo Finance data may have delays; timestamps are recorded
- **Rate Limiting**: Use sleep delays (1-3 seconds) to avoid throttling
- **Currency**: Values are in the stock's native currency
- **Error Handling**: Failed tickers are logged but don't stop the process
- **EV Ratios**: `enterprise_to_revenue` and `enterprise_to_ebitda` are computed from TTM (trailing twelve months) figures by summing the 4 most-recent quarters from `quarterly_financials`. This corrects a known issue where `ticker.info` returns single-quarter denominators for some non-US tickers.
- **Calculated Ratios**: Some ratios are calculated from statement data and may differ from Yahoo's displayed values

## Troubleshooting
- **Missing Data**: Some metrics may be None/N/A for certain stocks or exchanges
- **Rate Limits**: Increase `--sleep` value if experiencing throttling
- **International Stocks**: Some metrics may not be available for non-US exchanges
- **Currency Conversion**: All values remain in native currency (no automatic conversion)

## Integration
This script complements the `yf_fundamentals.py` script:
- `yf_fundamentals.py`: Raw financial statements (IS, BS, CF)
- `yf_statistics.py`: Key ratios and market metrics
- Use both for comprehensive fundamental analysis
