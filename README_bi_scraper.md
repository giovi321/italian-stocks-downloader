# Borsa Italiana Ticker Scraper
This script scrapes **all tickers** from the following Borsa Italiana pages and extracts key company data:

- [Euronext Growth Milan](https://www.borsaitaliana.it/borsa/azioni/euronext-growth-milan/lista.html)
- [Mid Cap](https://www.borsaitaliana.it/borsa/azioni/mid-cap/lista.html)
- [Small Cap](https://www.borsaitaliana.it/borsa/azioni/small-cap/lista.html)
- [Euronext STAR Milan](https://www.borsaitaliana.it/borsa/azioni/euronext-star-milan/lista.html)

For each company, the script retrieves:
- **Ticker (Codice Alfanumerico)**
- **ISIN**
- **Company name**
- **CEO name (Nome Amministratore Delegato)**
- **Shareholding information (Informazioni sull'azionariato)**

## 1. Requirements
Tested on **Python 3.9+**.
Install required packages:
```bash
pip install requests beautifulsoup4 lxml
```

## 2. Usage
Save the script as `scraper.py` and run:

```bash
python3 scraper.py
```
The script automatically paginates through all list pages until it finds repeated pages (the last page is repeated indefinitely).

### Output files
Two files are created in the same directory:
- `borsaitaliana_tickers.csv`
- `borsaitaliana_tickers.json`
Both contain one row per ticker **per segment** (a company may appear in multiple indexes).

### CSV Columns
| Column | Description |
| ------- | ------------ |
| segment | The index or market segment (e.g. mid-cap) |
| ticker | Codice Alfanumerico |
| isin | ISIN code |
| company | Company name |
| detail_url | URL of the “scheda” page |
| profile_url | URL of the “profilo-societa-dettaglio” page |
| ceo_name | CEO / Amministratore Delegato |
| shareholding | Shareholding breakdown (“Informazioni sull’azionariato”) |

## 3. How It Works
1. The script iterates over the 4 base URLs listed above.  
2. It appends `?page=n` to fetch each paginated list until the next page has identical links.  
3. From each list page, it extracts every **“scheda”** link using a regular expression that supports both path styles:  
   - `/borsa/azioni/<segment>/scheda/ISIN.html`  
   - `/borsa/azioni/scheda/ISIN.html?lang=it`  
4. Each “scheda” page is opened to extract the **Codice Alfanumerico** and **ISIN**.  
5. For each ISIN, a **profile page** is queried:  
   `https://www.borsaitaliana.it/borsa/azioni/profilo-societa-dettaglio.html?isin=<ISIN>&lang=it`  
6. From the profile page, the script parses two sections:
   - “Dirigenti principali” → finds the row with “Amministratore Delegato” to extract the CEO name.
   - “Informazioni sull’azionariato” → reads the shareholder table (name + percentage).

## 4. Known Limitations
- Borsa Italiana pages sometimes change HTML layout; if a section’s structure changes, parsing rules may need updating.
- Some companies do not list a CEO or shareholders publicly, leaving those fields empty.
- Script does not handle JavaScript-rendered content (Borsa pages are static HTML).

## 5. Troubleshooting
| Issue | Possible Fix |
| ------ | ------------- |
| No tickers found | Ensure site structure hasn’t changed; check regex in `SCHEda_RE`. |
| CEO or shareholders empty | Verify if the information actually exists on the page. Some are missing or listed differently. |
| Requests blocked | Add longer delays between requests or use a different User-Agent. |

## 6. Output Example
| segment | ticker | isin | company | ceo_name | shareholding |
|----------|--------|------|----------|-----------|--------------|
| mid-cap | ABCD | IT0001234567 | ABC S.p.A. | Mario Rossi | Free float: 45.00 %; XYZ Holding: 30.00 % |



