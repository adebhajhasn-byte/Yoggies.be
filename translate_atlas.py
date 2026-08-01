"""
translate_atlas.py — Vertaalt de gescrapte hondenras-info (Tsjechisch -> Nederlands)
===================================================================================
Leest atlas_scraped.json, knipt elke pagina in secties, vertaalt via OpenRouter
(zelfde sleutel als de pipeline), en schrijft atlas_translated.json.

VEREIST (zoals de pipeline al gebruikt):
    een omgevingsvariabele OPENROUTER_API_KEY

DRAAIEN (vanuit de projectmap):
    python translate_atlas.py

OUTPUT:
    atlas_translated.json          -> vertaalde secties per ras (klaar voor pagina-bouw)
    atlas_translate_cache.json     -> cache (zodat opnieuw draaien gratis is)
"""

import os
import re
import sys
import json
import time
import urllib.request

# ── Config ───────────────────────────────────────────────────────────────────
IN_JSON    = "atlas_scraped.json"
OUT_JSON   = "atlas_translated.json"
CACHE_JSON = "atlas_translate_cache.json"
API_KEY    = os.environ.get("OPENROUTER_API_KEY", "")
MODEL      = "deepseek/deepseek-v4-flash:free"
DELAY      = 1.5

# Sectie-markers (regex) in documentvolgorde. key -> Tsjechisch kopje
SECTION_PATTERNS = [
    ("intro",   r'– vše o plemenu'),
    ("povaha",  r'\bPovaha\b'),
    ("hodi",    r'Pro koho se hodí'),
    ("zaklad",  r'Základní informace[^\n]*'),
    ("vychova", r'Výchova a výcvik'),
    ("nemoci",  r'Nemoci[^\n]*'),
    ("krmit",   r'Čím a jak[^\n]*krmit'),
    ("krmivo",  r'Krmivo pro'),   # alleen als grens (producttabel -> weglaten)
]

# Nederlandse kopjes voor de uiteindelijke pagina
NL_HEADINGS = {
    "intro":   "Oorsprong",
    "povaha":  "Karakter",
    "hodi":    "Voor wie is dit ras geschikt",
    "zaklad":  "Algemene informatie",
    "vychova": "Opvoeding en training",
    "nemoci":  "Veelvoorkomende aandoeningen",
    "krmit":   "Voeding",
}


def parse_breed(text):
    """Knip de ruwe tekst in secties op basis van Tsjechische kopjes."""
    t = text.replace("\xa0", " ")
    found = []
    for key, pat in SECTION_PATTERNS:
        m = re.search(pat, t)
        if m:
            found.append((key, m.start(), m.end()))
    found.sort(key=lambda x: x[1])
    sections = {}
    for idx, (key, start, end) in enumerate(found):
        nxt = found[idx + 1][1] if idx + 1 < len(found) else len(t)
        chunk = t[end:nxt].strip()
        # opschonen: losse restjes van kopjes
        chunk = re.sub(r'^\s*(o plemeni|[a-záčďéěíňóřšťúůýž ]{0,30})\n', '', chunk, count=1, flags=re.IGNORECASE)
        sections[key] = chunk.strip()
    sections.pop("krmivo", None)  # producttabel weglaten
    return sections


def extract_size(zaklad_text):
    """Haal hoogte/gewicht/levensduur uit de (rommelige) tabel in 'zaklad'."""
    res = {}
    def grab(label_after):
        m = re.search(label_after, zaklad_text)
        return m.group(1).strip() if m else ""
    # Verzamel alle 'cm', 'kg' en 'let/jaar' waarden
    cm  = re.findall(r'(\d{1,3}\s*[-–]\s*\d{1,3}\s*cm)', zaklad_text)
    kg  = re.findall(r'(\d{1,3}\s*[-–]\s*\d{1,3}\s*kg)', zaklad_text)
    life = re.findall(r'(\d{1,2}\s*[-–]\s*\d{1,2})\s*let', zaklad_text)
    if cm:  res["hoogte"] = " / ".join(dict.fromkeys(cm))
    if kg:  res["gewicht"] = " / ".join(dict.fromkeys(kg))
    if life: res["levensduur"] = dict.fromkeys(life).__iter__().__next__() + " jaar"
    return res


def clean_zaklad(zaklad_text):
    """Houd alleen de beschrijvende eerste alinea van 'zaklad' (zonder de tabel)."""
    # Stop bij de eerste regel die op de maattabel lijkt
    stop = re.search(r'(Fena|Pes)\s+(štěně|dospěl|dosp)', zaklad_text)
    if stop:
        return zaklad_text[:stop.start()].strip()
    return zaklad_text.strip()


# ── Translation via OpenRouter ────────────────────────────────────────────────
_cache = {}

def cache_load():
    global _cache
    if os.path.exists(CACHE_JSON):
        _cache = json.load(open(CACHE_JSON, encoding="utf-8"))

def cache_save():
    json.dump(_cache, open(CACHE_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

def translate_text(cz_text):
    """Vertaal één sectie Tsjechisch -> Nederlands. Gecachet per sectie."""
    cz_text = cz_text.strip()
    if not cz_text:
        return ""
    if cz_text in _cache:
        return _cache[cz_text]
    sys_prompt = (
        "Je bent een professionele vertaler Tsjechisch -> Nederlands voor een "
        "hondenvoeding-website (yoggies.be). Vertaal naar vlot, natuurlijk Nederlands. "
        "Gebruik 'uw hond'. Geef ALLEEN de Nederlandse vertaling terug, geen uitleg, "
        "geen Tsjechische woorden, geen markdown. Behoud productnamen die met "
        "'Yoggies' beginnen letterlijk."
    )
    payload = json.dumps({
        "model": MODEL,
        "messages": [
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": cz_text},
        ],
    }).encode()
    req = urllib.request.Request(
        "https://openrouter.ai/api/v1/chat/completions",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {API_KEY}",
            "HTTP-Referer": "https://yoggies.be",
        },
    )
    with urllib.request.urlopen(req, timeout=90) as r:
        out = json.loads(r.read())["choices"][0]["message"]["content"].strip()
    _cache[cz_text] = out
    cache_save()
    return out


def main():
    if not API_KEY:
        print("FOUT: OPENROUTER_API_KEY ontbreekt in de omgeving.")
        print("Zet hem zoals voor de pipeline en draai opnieuw.")
        sys.exit(1)
    if not os.path.exists(IN_JSON):
        print(f"FOUT: {IN_JSON} niet gevonden. Draai eerst scrape_atlas.py.")
        sys.exit(1)

    cache_load()
    data = json.load(open(IN_JSON, encoding="utf-8"))
    out = []
    total = len(data)
    print(f"Vertalen van {total} rassen (Tsjechisch -> Nederlands)\n" + "-" * 50)

    for i, d in enumerate(data):
        title = d["title"]
        print(f"[{i+1}/{total}] {title}")
        secs = parse_breed(d["text"])
        size = extract_size(secs.get("zaklad", ""))
        secs["zaklad"] = clean_zaklad(secs.get("zaklad", ""))

        order = ["intro", "povaha", "hodi", "zaklad", "vychova", "nemoci", "krmit"]
        nl_secs = {}
        for k in order:
            cz = secs.get(k, "")
            if not cz:
                continue
            try:
                nl = translate_text(cz)
                if nl:
                    nl_secs[k] = nl
            except Exception as e:
                print(f"        [{k}] MISLUKT: {e}")
            time.sleep(DELAY)
        print(f"        OK — {len(nl_secs)}/{len([k for k in order if secs.get(k)])} secties vertaald")

        out.append({
            "title_cz": title,
            "slug": d["slug"],
            "url": d["url"],
            "size": size,
            "sections": nl_secs,
            "headings": {k: NL_HEADINGS[k] for k in order if k in nl_secs},
        })

    json.dump(out, open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    ok = sum(1 for o in out if o["sections"])
    print("-" * 50)
    print(f"Klaar. {ok}/{total} rassen vertaald.")
    print(f"Output -> {OUT_JSON}")
    print(f"Stuur '{OUT_JSON}' terug zodat ik de 27 HTML-pagina's kan bouwen.")


if __name__ == "__main__":
    main()
