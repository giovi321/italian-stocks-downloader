#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import json
import re
import time
from urllib.parse import urljoin, urlsplit, urlunsplit, urlencode

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter, Retry

BASE = "https://www.borsaitaliana.it"

SEGMENTS = {
    "euronext-growth-milan": "https://www.borsaitaliana.it/borsa/azioni/euronext-growth-milan/lista.html",
    "mid-cap":               "https://www.borsaitaliana.it/borsa/azioni/mid-cap/lista.html",
    "small-cap":             "https://www.borsaitaliana.it/borsa/azioni/small-cap/lista.html",
    "euronext-star-milan":   "https://www.borsaitaliana.it/borsa/azioni/euronext-star-milan/lista.html",
}

PROFILE_PATH = "/borsa/azioni/profilo-societa-dettaglio.html"

OUT_CSV  = "borsaitaliana_tickers.csv"
OUT_JSON = "borsaitaliana_tickers.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "it,en;q=0.8",
}

# match both styles, with optional ?lang=it
SCHEda_RE = re.compile(
    r"/borsa/azioni/(?:[^/]+/)?scheda/[A-Z0-9]{10,12}\.html(?:\?[^#\"']*)?",
    re.I
)

# ---------- HTTP session ----------

def session_with_retries():
    s = requests.Session()
    retries = Retry(
        total=5, connect=5, read=5,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.headers.update(HEADERS)
    return s

def get_html(s, url):
    r = s.get(url, timeout=20)
    return r.status_code, r.text

def add_page_param(url, page):
    return url + ("&page=" if "?" in url else "?page=") + str(page)

# ---------- parsing helpers ----------

def text_norm(s):
    return re.sub(r"\s+", " ", (s or "").strip())

def canonicalize(url):
    u = urlsplit(urljoin(BASE, url))
    return urlunsplit((u.scheme, u.netloc, u.path, "", ""))

def extract_scheda_links_from_html(html):
    soup = BeautifulSoup(html, "lxml")
    full_urls = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if SCHEda_RE.search(href):
            full_urls.append(urljoin(BASE, href))
    # unique by canonical path (to ignore ?lang variants)
    by_canon = {}
    for u in full_urls:
        by_canon.setdefault(canonicalize(u), u)
    return sorted(by_canon.values()), set(by_canon.keys())

def list_all_segment_links(s, list_url):
    """
    Iterate page=1,2,... and stop when canonical set(page n+1) == set(page n).
    """
    page = 1
    prev_canon = None
    canon_to_full = {}
    SAFETY_MAX = 200

    while page <= SAFETY_MAX:
        url = add_page_param(list_url, page)
        code, html = get_html(s, url)
        print(f"[LIST] {url} -> {code}")
        if code != 200:
            print(f"  skip page {page}: HTTP {code}")
            page += 1
            continue

        full_urls, canon_set = extract_scheda_links_from_html(html)
        print(f"  page {page}: {len(canon_set)} links")
        if prev_canon is not None and canon_set == prev_canon:
            print(f"  stop: page {page} equals page {page-1}")
            break

        for u in full_urls:
            canon_to_full.setdefault(canonicalize(u), u)

        prev_canon = canon_set
        page += 1
        time.sleep(0.2)

    if page > SAFETY_MAX:
        print("  warn: hit safety ceiling without repeat; data may be incomplete.")

    return sorted(canon_to_full.values())

# ---------- detail page parsing ----------

def extract_label_value_pairs(soup):
    d = {}
    # tables
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if len(cells) >= 2:
            d[text_norm(cells[0].get_text())] = text_norm(cells[1].get_text())
    # dl/dt/dd
    for dl in soup.find_all("dl"):
        for dt, dd in zip(dl.find_all("dt"), dl.find_all("dd")):
            d[text_norm(dt.get_text())] = text_norm(dd.get_text())
    # text fallback
    page_text = soup.get_text("\n")
    for m in re.finditer(r"(Codice\s+Alfanumerico|Codice\s+Isin)\s*[:,]\s*([A-Z0-9\.\-]+)", page_text, re.I):
        d[text_norm(m.group(1))] = text_norm(m.group(2))
    return d

def parse_company_name(soup):
    for tag in ["h1", "h2", "title"]:
        el = soup.find(tag)
        if el:
            name = text_norm(el.get_text())
            return re.sub(r"\s*\|\s*Borsa Italiana.*$", "", name, flags=re.I)
    return ""

def fetch_detail(s, url):
    code, html = get_html(s, url)
    if code != 200:
        print(f"[DETAIL] {url} -> {code}")
        return {"detail_url": url, "status": code, "ticker": "", "isin": "", "company": ""}
    soup = BeautifulSoup(html, "lxml")
    labels = extract_label_value_pairs(soup)

    def get_label(kvs, keys):
        for k, v in kvs.items():
            if text_norm(k).lower() in keys:
                return v
        return ""

    ticker = get_label(labels, ["codice alfanumerico"])
    isin = get_label(labels, ["codice isin", "isin"])
    if not isin:
        m = re.search(r"/scheda/([A-Z0-9]{10,12})\.html", url, re.I)
        if m:
            isin = m.group(1)
    company = parse_company_name(soup)
    print(f"[TICKER] {ticker or '??'} ({isin or '??'}) - {company or '??'}")
    return {"detail_url": url, "status": code, "ticker": ticker, "isin": isin, "company": company}

# ---------- profile page parsing ----------

def build_profile_url(isin):
    return urljoin(BASE, PROFILE_PATH) + "?" + urlencode({"isin": isin, "lang": "it"})

def _find_section_root(soup, header_regex):
    """
    Find first heading/strong/bold whose text matches header_regex,
    then return a lightweight container of its following siblings up to next heading.
    """
    for tag in soup.find_all(["h1","h2","h3","h4","strong","b"]):
        if re.search(header_regex, text_norm(tag.get_text()), re.I):
            holder = soup.new_tag("div")
            for sib in tag.next_siblings:
                if getattr(sib, "name", None) in ("h1","h2","h3","h4","strong","b"):
                    break
                holder.append(sib if getattr(sib, "name", None) else soup.new_string(str(sib)))
            return holder
    return None

def _looks_like_percent(s):
    return "%" in s or re.search(r"\d+[.,]?\d*\s*%", s)

def _normalize_percent(s):
    s = s.replace(",", ".")
    s = re.sub(r"\s*%\s*", " %", s)
    return text_norm(s)

def _best_name_guess(cells):
    """
    Pick the cell most likely to be a person or shareholder name:
    - few digits
    - several letters with capitalization
    """
    best = ""
    best_score = -1
    for c in cells:
        digits = sum(ch.isdigit() for ch in c)
        caps = sum(1 for ch in c if ch.isalpha() and ch.upper() == ch)
        letters = sum(1 for ch in c if ch.isalpha())
        score = letters - 5*digits + 0.1*caps
        if score > best_score:
            best_score = score
            best = c
    return best

def extract_ceo_from_profile(soup):
    # Narrow to "Dirigenti principali" if present
    sec = _find_section_root(soup, r"Dirigenti\s+principali")
    root = sec if sec else soup

    # Try tables in the section
    for tbl in root.find_all("table"):
        # header mapping
        headers = [text_norm(th.get_text()) for th in tbl.find_all("th")]
        role_idx = None
        name_idx = None
        if headers:
            for i, h in enumerate(headers):
                if re.search(r"(incarico|ruolo|carica|posizione)", h, re.I):
                    role_idx = i
                if re.search(r"(nome|cognome|dirigente|componente|manager)", h, re.I):
                    name_idx = i
        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if not tds:
                continue
            cells = [text_norm(td.get_text(" ")) for td in tds]
            role = ""
            name = ""
            if role_idx is not None and role_idx < len(cells):
                role = cells[role_idx]
            if name_idx is not None and name_idx < len(cells):
                name = cells[name_idx]
            if not role:  # guess: role often first
                role = cells[0]
            if not name:
                name = _best_name_guess(cells)

            if re.search(r"amministratore\s+delegato|chief\s+executive", role, re.I):
                return name

    # Fallback: scan lines in the section
    for node in root.find_all(["li","p","div","span"], recursive=True):
        line = text_norm(node.get_text(" "))
        if re.search(r"amministratore\s+delegato|chief\s+executive", line, re.I):
            # strip role words, keep trailing name tokens
            line2 = re.sub(r".*amministratore\s+delegato", "", line, flags=re.I)
            line2 = re.sub(r"presidente\s+e\s*", "", line2, flags=re.I)
            guess = _best_name_guess([line2])
            if guess:
                return guess

    # Whole page fallback for explicit label
    page_text = soup.get_text("\n", strip=True)
    m = re.search(r"Nome\s+amministratore\s+delegato\s*:\s*([^\n]+)", page_text, re.I)
    return text_norm(m.group(1)) if m else ""

def extract_shareholding_block(soup):
    # Narrow to “Informazioni sull'azionariato” if present, else “Azionariato”
    sec = _find_section_root(soup, r"Informazioni\s+sull[’']?azionariato|Azionariato")
    root = sec if sec else soup

    # Table-first strategy
    items = []
    for tbl in root.find_all("table"):
        # detect header
        head = []
        thead = tbl.find("thead")
        if thead:
            head = [text_norm(th.get_text()) for th in thead.find_all("th")]
        else:
            first_tr = tbl.find("tr")
            if first_tr and first_tr.find_all("th"):
                head = [text_norm(th.get_text()) for th in first_tr.find_all("th")]

        name_idx = None
        pct_idx = None
        if head:
            for i, h in enumerate(head):
                if re.search(r"(azionista|soggetto|titolare|free\s*float|flottante|azionariato)", h, re.I):
                    if name_idx is None:
                        name_idx = i
                if re.search(r"(quota|percentuale|%)", h, re.I):
                    pct_idx = i

        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if not tds:
                continue
            cells = [text_norm(td.get_text(" ")) for td in tds]
            if not cells:
                continue

            # guess indices when headers missing
            if name_idx is None or name_idx >= len(cells):
                # pick the cell with least digits and most letters
                name = _best_name_guess(cells)
            else:
                name = cells[name_idx]

            if pct_idx is None or pct_idx >= len(cells):
                pct = ""
                for c in cells:
                    if _looks_like_percent(c):
                        pct = c
                        break
            else:
                pct = cells[pct_idx]

            if not name and not pct:
                continue

            # keep only plausible shareholder/free float rows
            if (_looks_like_percent(pct) or re.search(r"(free\s*float|flottante)", name, re.I)):
                items.append(f"{name}: {_normalize_percent(pct) if pct else ''}".strip(" :"))

        if items:
            break

    if not items:
        # Fallback: lines inside the section that look like "Name 12,34 %"
        raw = root.get_text("\n", strip=True)
        for line in raw.split("\n"):
            line = text_norm(line)
            m = re.search(r"(.+?)\s+(\d+[.,]?\d*)\s*%", line)
            if m:
                items.append(f"{text_norm(m.group(1))}: {_normalize_percent(m.group(2)+'%')}")

    # Cleanup and dedupe
    cleaned = []
    seen = set()
    for x in items:
        y = re.sub(r"\s{2,}", " ", x).strip(" -•")
        if y and y not in seen:
            seen.add(y)
            cleaned.append(y)
    return "; ".join(cleaned)

def fetch_profile_fields(s, isin):
    url = build_profile_url(isin)
    code, html = get_html(s, url)
    print(f"[PROFILE] {url} -> {code}")
    if code != 200:
        return {"profile_url": url, "ceo_name": "", "shareholding": ""}
    soup = BeautifulSoup(html, "lxml")
    ceo = extract_ceo_from_profile(soup)
    azion = extract_shareholding_block(soup)
    if not ceo:
        print(f"  [WARN] CEO not found for {isin}")
    if not azion:
        print(f"  [WARN] Shareholding not found for {isin}")
    return {"profile_url": url, "ceo_name": ceo, "shareholding": azion}

# ---------- main ----------

def main():
    s = session_with_retries()
    results = []

    for segment, list_url in SEGMENTS.items():
        print(f"\n=== {segment.upper()} ===")
        seg_links = list_all_segment_links(s, list_url)
        print(f" total links in {segment}: {len(seg_links)}")
        if not seg_links:
            raise RuntimeError(f"No scheda links found for segment '{segment}'.")

        # dedup only within the segment
        seen = set()
        for link in seg_links:
            if link in seen:
                continue
            seen.add(link)

            # detail page
            rec = fetch_detail(s, link)
            rec["segment"] = segment

            # profile page
            if rec.get("isin"):
                prof = fetch_profile_fields(s, rec["isin"])
            else:
                prof = {"profile_url": "", "ceo_name": "", "shareholding": ""}

            rec.update(prof)
            results.append(rec)
            time.sleep(0.15)

        # enforce ticker presence per segment
        missing = [r for r in results if r["segment"] == segment and not r["ticker"]]
        if missing:
            ex = "\n".join(f"- {m['detail_url']} (ISIN={m['isin']}, company={m['company']})" for m in missing[:10])
            raise RuntimeError(
                f"Missing 'Codice Alfanumerico' in {len(missing)} pages for '{segment}'. Examples:\n{ex}"
            )

    # outputs
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "segment", "ticker", "isin", "company", "detail_url",
            "profile_url", "ceo_name", "shareholding"
        ])
        for r in results:
            w.writerow([
                r.get("segment",""), r.get("ticker",""), r.get("isin",""),
                r.get("company",""), r.get("detail_url",""),
                r.get("profile_url",""), r.get("ceo_name",""), r.get("shareholding","")
            ])

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    counts = {}
    for r in results:
        counts[r["segment"]] = counts.get(r["segment"], 0) + 1
    print("\nDone.")
    print("Counts:", counts)
    print(f"Saved {OUT_CSV} and {OUT_JSON}")

if __name__ == "__main__":
    main()
