"""
scrape_atlas.py — Haalt alle 27 hondenras-pagina's op van yoggies.cz
==================================================================
Plaats dit bestand + yoggiesCZ_HONDATLAS.xlsx in de projectmap (yoggies.be)
en draai het van daar.

EERST INSTALLEREN (eenmalig):
    pip install pandas openpyxl beautifulsoup4 cloudscraper requests

DRAAIEN:
    python scrape_atlas.py

OUTPUT:
    atlas_raw/<slug>.html   -> ruwe HTML van elke pagina (back-up voor analyse)
    atlas_scraped.json      -> uitgelezen tekst per ras

Stuur daarna atlas_scraped.json (of 1 .html voorbeeld) terug zodat ik
de juiste extractie + vertaling kan bouwen.
"""

import os
import sys
import json
import time
import re

# ── Config ───────────────────────────────────────────────────────────────────
XLSX     = "yoggiesCZ_HONDATLAS.xlsx"   # zet dit bestand in deze map (naam mag afwijken)
OUT_DIR  = "atlas_raw"
OUT_JSON = "atlas_scraped.json"
DELAY    = 2.5                          # seconden pauze tussen requests (beleefd)


def find_xlsx():
    """Vind het Hondatlas Excel-bestand, ook als de naam afwijkt (spatie/underscore)."""
    import glob
    if os.path.exists(XLSX):
        return XLSX
    # Zoek elk .xlsx-bestand met 'hondatlas' of 'atlas' in de naam
    for pat in ("*HONDATLAS*.xlsx", "*hondatlas*.xlsx", "*HONDATLAS*.xls",
                "*atlas*.xlsx", "*ATLAS*.xlsx"):
        hits = glob.glob(pat)
        if hits:
            return hits[0]
    return None

# ── Dependencies ──────────────────────────────────────────────────────────────
try:
    import pandas as pd
except ImportError:
    print("FOUT: pandas ontbreekt. Draai eerst:")
    print("    pip install pandas openpyxl beautifulsoup4 cloudscraper requests")
    sys.exit(1)


def make_session():
    """Geef een HTTP-sessie terug. cloudscraper omzeilt Cloudflare beter."""
    try:
        import cloudscraper
        print("→ cloudscraper gevonden (Cloudflare-aware). Goed.")
        return cloudscraper.create_scraper(
            browser={"browser": "chrome", "platform": "windows", "mobile": False}
        )
    except ImportError:
        import requests
        print("→ cloudscraper niet gevonden — gebruik gewone requests (kan geblokkeerd worden).")
        print("  Tip: 'pip install cloudscraper' voor betere kans.")
        s = requests.Session()
        s.headers.update({
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/120.0.0.0 Safari/537.36"),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "cs-CZ,cs;q=0.9,nl;q=0.8,en;q=0.7",
            "Referer": "https://yoggies.cz/pages/atlas-psu",
        })
        return s


def find_col(df, *needles):
    """Vind een kolom waarvan de naam een van de needles bevat."""
    for c in df.columns:
        cl = str(c).lower()
        if any(n in cl for n in needles):
            return c
    return None


def slug_from_url(url):
    return url.rstrip("/").split("/")[-1]


def extract_text(html):
    """Haal de zichtbare hoofdtekst uit de HTML (ruwe eerste poging)."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "nav", "footer", "header", "form"]):
        tag.decompose()
    main = (soup.find("main")
            or soup.find("div", class_=re.compile("rte|page|content|shopify-section"))
            or soup.body)
    text = main.get_text("\n", strip=True) if main else soup.get_text("\n", strip=True)
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    return "\n".join(lines)


def main():
    xlsx = find_xlsx()
    if not xlsx:
        print(f"FOUT: geen Hondatlas Excel-bestand gevonden in deze map.")
        print("Zet het bestand hier neer (bv. 'yoggiesCZ HONDATLAS.xlsx') en draai opnieuw.")
        sys.exit(1)
    print(f"Excel gevonden: {xlsx}")
    df = pd.read_excel(xlsx)
    href_col  = find_col(df, "href", "link", "url")
    title_col = find_col(df, "title", "naam", "name")
    if not href_col or not title_col:
        print("FOUT: kon de kolommen voor link/titel niet vinden.")
        print("Gevonden kolommen:", list(df.columns))
        sys.exit(1)

    os.makedirs(OUT_DIR, exist_ok=True)
    session = make_session()
    results = []
    total = len(df)

    print(f"\nStart: {total} rassen ophalen van yoggies.cz\n" + "-" * 50)
    for i, row in df.iterrows():
        url   = str(row[href_col]).strip()
        title = str(row[title_col]).strip()
        slug  = slug_from_url(url)
        print(f"[{i+1}/{total}] {title}")
        print(f"        {url}")
        try:
            r = session.get(url, timeout=30)
            status = r.status_code
            html = r.text
            with open(os.path.join(OUT_DIR, f"{slug}.html"), "w", encoding="utf-8") as f:
                f.write(html)
            text = extract_text(html)
            results.append({
                "title": title, "slug": slug, "url": url,
                "status": status, "text": text, "text_len": len(text),
            })
            flag = "OK" if status == 200 else f"LET OP status={status}"
            print(f"        {flag} | tekst: {len(text)} tekens")
        except Exception as e:
            print(f"        MISLUKT: {e}")
            results.append({
                "title": title, "slug": slug, "url": url,
                "status": "ERROR", "error": str(e),
            })
        time.sleep(DELAY)

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    ok = sum(1 for r in results if r.get("status") == 200)
    print("-" * 50)
    print(f"Klaar. {ok}/{total} pagina's succesvol opgehaald.")
    print(f"Ruwe HTML  -> {OUT_DIR}/")
    print(f"Tekst      -> {OUT_JSON}")
    if ok == 0:
        print("\n!! Niets opgehaald — waarschijnlijk geblokkeerd.")
        print("   Installeer cloudscraper en probeer opnieuw:")
        print("       pip install cloudscraper")
    else:
        print(f"\nStuur '{OUT_JSON}' terug (of 1 bestand uit {OUT_DIR}/) "
              "zodat ik de extractie + vertaling kan bouwen.")


if __name__ == "__main__":
    main()
