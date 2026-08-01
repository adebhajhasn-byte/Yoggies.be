"""
build_site_final.py — Yoggies.be Complete Site Builder
=======================================================
Stap 1: Vertaalt alle CZ teksten naar NL via Anthropic API
Stap 2: Bouwt het volledige website (339 productpagina's + 7 categoriepagina's + homepage)
+ Voedingscalculator voor producten met has_calculator=TRUE

Gebruik:
  pip install openpyxl anthropic
  set ANTHROPIC_API_KEY=sk-ant-api03-...
  python build_site_final.py
"""

import os, re, html, unicodedata, json, shutil, time
import openpyxl

# ── Configuratie ───────────────────────────────────────────────────────────────
DATA_FILE   = "READY_FOR_BUILD_FINAL.xlsx"
PROD_TMPL   = "site/product.html"
CAT_TMPL    = "site/hoofdpageproduct.html"
HOME_TMPL   = "site/home.html"
OUT_DIR     = "site"
PROD_DIR    = os.path.join(OUT_DIR, "products")
CACHE_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "translation_cache_v2.json")

os.makedirs(PROD_DIR, exist_ok=True)

# ── Vertalingen ────────────────────────────────────────────────────────────────
ICON_TRANS = {
    "Bez chemie": "Zonder chemie",
    "Lisované za studena": "Koudgeperst",
    "S probiotiky": "Met probiotica",
    "Farmářské suroviny": "Boerderijgrondstoffen",
    "Pro zdravé zuby": "Voor gezonde tanden",
    "Pro správné zažívání": "Voor gezonde spijsvertering",
    "Vhodné pro alergiky": "Geschikt voor allergische honden",
    "EKO přátelský": "Milieuvriendelijk",
    "Organické produkty": "Biologische producten",
    "Ideální na trénink": "Ideaal voor training",
    "Pro hedvábnou srst": "Voor een zijdezachte vacht",
    "Pro zdravý mikrobiom": "Voor een gezond microbioom",
    "Pro čilý mozek": "Voor een actieve geest",
    "Pro všechna plemena i velikosti": "Voor alle rassen en maten",
}

INGR_TRANS = {
    "sušené mleté maso": "gedroogd gemalen vlees",
    "sušené maso": "gedroogd vlees",
    "sušené rybí maso": "gedroogd visvlees",
    "hovězí maso": "rundvlees", "hovězí": "rund",
    "kuřecí maso": "kippenvlees", "kuřecí": "kip",
    "krůtí maso": "kalkoenvlees", "krůtí": "kalkoen",
    "jehněčí maso": "lamsvlees", "jehněčí": "lam",
    "kachní maso": "eendenvlees", "kachní": "eend",
    "kachny Barbarie": "Barbarie-eend",
    "jelení maso": "hertenvlees", "jelení": "hert",
    "lososové maso": "zalmvlees", "losos": "zalm",
    "vepřové maso": "varkensvlees", "vepřové": "varken",
    "zvěřinové maso": "wildvlees", "zvěřina": "wild",
    "svalovina s kostí": "spiervlees met bot",
    "svalovina": "spiervlees",
    "vnitřnosti": "organen",
    "srdce": "hart", "játra": "lever", "žaludky": "maag",
    "hnědá rýže": "bruine rijst", "rýže": "rijst",
    "bramborové vločky": "aardappelvlokken",
    "brambory": "aardappelen",
    "batáty": "zoete aardappelen",
    "proso hrubomleté": "grof gemalen gierst",
    "proso": "gierst", "pohanka": "boekweit",
    "oves": "haver", "ovesné vločky": "havermout",
    "rýžové sušené klíčky": "gedroogde rijstkiemen",
    "želatinový prášek": "gelatinepoeder",
    "pivovarské kvasnice": "biergist",
    "konopný olej": "hennepolie",
    "lněný olej": "lijnzaadolie",
    "lososový olej": "zalmolie",
    "rybí olej": "visolie",
    "probiotika": "probiotica",
    "Lithothamnium": "Lithothamnium",
    "polyfenoly": "polyfenolen",
    "quinoa": "quinoa",
    "chlorella": "chlorella",
    "mrkev": "wortel",
    "hrášek": "erwten",
    "špenát": "spinazie",
    "dýně": "pompoen",
    "brusinky": "veenbessen",
    "šípky sušené": "gedroogde rozenbottels",
    "šípky": "rozenbottels",
    "vývar": "bouillon",
    "sůl": "zout",
    "voda": "water",
    "Neobsahuje": "Bevat geen",
    "červená řepa": "rode biet",
    "cuketa": "courgette",
    "topinambury": "topinamboer",
    "Analytické složky": "Analytische bestanddelen",
    "hrubý protein": "ruw eiwit",
    "hrubý tuk": "ruw vet",
    "hrubá vláknina": "ruwe celstof",
    "hrubý popel": "ruw as",
    "vlhkost": "vocht",
    "vápník": "calcium",
    "fosfor": "fosfor",
    "sodík": "natrium",
    "červená řepa sušená mletá": "gedroogde gemalen rode biet",
    "sušené mleté": "gedroogd gemalen",
    "sušená mletá": "gedroogd gemalen",
    "sušené": "gedroogd",
    "mleté": "gemalen",
    "zvěřinové kosti mleté": "gemalen wildbeenderen",
    "zvěřinové kosti": "wildbeenderen",
    "kosti mleté": "gemalen beenderen",
    "kosti": "beenderen",
    "jelen": "hert",
    "směs mořských řas": "zeewier mix",
    "směs sušených bylinek": "mix gedroogde kruiden",
    "směs zeleniny": "groentemix",
    "směs": "mix",
}

DOSE_TRANS = {
    "Doporučené krmné dávky:": "Aanbevolen dagelijkse porties:",
    "Doporučená krmná dávka": "Aanbevolen dagelijkse portie",
    "doporučené množství krmiva na den odpovídá zhruba": "de aanbevolen hoeveelheid voer per dag bedraagt ongeveer",
    "tělesné hmotnosti zdravého psa": "van het lichaamsgewicht van een gezonde hond",
    "tělesné hmotnosti": "van het lichaamsgewicht",
    "dospělého psa": "van de volwassen hond",
    "Přizpůsobte denní krmnou dávku stáří a aktivitě vašeho psa.": "Pas de dagelijkse portie aan op de leeftijd en activiteit van uw hond.",
    "U štěňat může množství krmiva přesahovat až": "Bij pups kan de hoeveelheid voer oplopen tot",
    "U štěňat může množství krmiva tvořit až do": "Bij pups kan de hoeveelheid voer oplopen tot",
    "Podávejte při pokojové teplotě.": "Serveer op kamertemperatuur.",
    "Zajistěte, aby měl váš pes vždy přístup k čerstvé pitné vodě.": "Zorg dat uw hond altijd toegang heeft tot vers drinkwater.",
    "Krmivo nenahrazuje kompletní krmnou dávku.": "Vervangt geen volledig dieet.",
    "Individuální potřeba závisí na věku, kondici, aktivitě i zdravotním stavu psa.": "De individuele behoefte hangt af van leeftijd, conditie, activiteit en gezondheid van de hond.",
    "Skladujte v suchu.": "Bewaar op een droge plaats.",
    "Po rozbalení chraňte před": "Na opening beschermen tegen",
    "Vždy přihlížejte k potřebám vašeho psa (zdravotní stav, kondice, zátěž).": "Houd altijd rekening met de behoeften van uw hond (gezondheidstoestand, conditie, inspanning).",
    "Váš pes by měl mít vždy k dispozici dostatek čerstvé vody v čisté misce.": "Uw hond moet altijd toegang hebben tot voldoende vers water in een schone bak.",
    "Ideální krmná dávka se určuje podle věku, tělesné váhy a fyzické aktivity psa.": "De ideale portie wordt bepaald op basis van leeftijd, lichaamsgewicht en lichamelijke activiteit van de hond.",
    "Tyto informace jsou pouze orientační.": "Deze informatie is slechts indicatief.",
    "denně": "per dag",
    "krmiva na den": "voer per dag",
    "krmivo": "voer",
}

# ── Calculator helper ──────────────────────────────────────────────────────────
def extract_pct_range(dose_text):
    """Extract low/high percentage from dosage text. Returns (lo, hi) floats."""
    if not dose_text:
        return (2.0, 3.0)
    # Range: 2–3 % or 2-3 %
    m = re.search(r'(\d+[,.]?\d*)\s*[–\-]\s*(\d+[,.]?\d*)\s*%', dose_text)
    if m:
        lo = float(m.group(1).replace(',', '.'))
        hi = float(m.group(2).replace(',', '.'))
        return (lo, hi)
    # Single: 1,2 % or 2 %
    m = re.search(r'(\d+[,.]?\d*)\s*%', dose_text)
    if m:
        v = float(m.group(1).replace(',', '.'))
        return (v, v)
    return (2.0, 3.0)

def build_calculator_html(dose_raw):
    """Build the voedingscalculator HTML block — injects into accordion."""
    pct_lo, pct_hi = extract_pct_range(dose_raw)

    if pct_lo == pct_hi:
        pct_label = f"{pct_lo:.1f}%".replace('.', ',')
        pct_desc  = f"<strong>{pct_label}</strong> van het lichaamsgewicht per dag"
        js_calc   = f"var gram = Math.round(weight * 1000 * {pct_lo} / 100) + ' g';"
        result_label = "Dagelijkse portie"
    else:
        plo = f"{pct_lo:.1f}".replace('.', ',')
        phi = f"{pct_hi:.1f}".replace('.', ',')
        pct_label = f"{plo}\u2013{phi}%"
        pct_desc  = f"<strong>{pct_label}</strong> van het lichaamsgewicht per dag"
        js_calc   = (
            f"var gramLo = Math.round(weight * 1000 * {pct_lo} / 100);"
            f"var gramHi = Math.round(weight * 1000 * {pct_hi} / 100);"
            f"var gram = gramLo + '\u2013' + gramHi + ' g';"
        )
        result_label = "Dagelijkse portie"

    return f"""
    <details class="accordion" id="calc-accordion">
      <summary class="accordion-summary">🧮 Voedingscalculator</summary>
      <div class="accordion-body">
        <div class="calc-wrap">
          <p class="calc-intro">Bereken de aanbevolen dagelijkse portie op basis van het gewicht van uw hond.<br>
          <span class="calc-pct-note">Richtlijn: {pct_desc}</span></p>

          <div class="calc-row">
            <label class="calc-label" for="calc-weight">Gewicht van uw hond</label>
            <div class="calc-input-group">
              <input
                type="number"
                id="calc-weight"
                class="calc-input"
                placeholder="bijv. 25"
                min="0.5"
                max="100"
                step="0.5"
                oninput="yoggiesCalc(this.value)"
              >
              <span class="calc-unit">kg</span>
            </div>
          </div>

          <div class="calc-result" id="calc-result" style="display:none">
            <div class="calc-result-label">{result_label}</div>
            <div class="calc-result-value"><span id="calc-grams">—</span> <span class="calc-result-unit">gram</span></div>
            <p class="calc-result-note">Pas de portie aan op conditie, leeftijd en activiteit van uw hond. Zorg altijd voor vers water.</p>
          </div>
        </div>

        <script>
        function yoggiesCalc(val) {{
          var weight = parseFloat(val);
          var result = document.getElementById('calc-result');
          var gramsEl = document.getElementById('calc-grams');
          if (!weight || weight <= 0) {{ result.style.display='none'; return; }}
          {js_calc}
          gramsEl.textContent = gram;
          result.style.display = 'block';
        }}
        </script>
      </div>
    </details>"""

# ── Calculator CSS (embedded in <style> in product.html) ───────────────────────
CALC_CSS = """
    /* ── Voedingscalculator ── */
    .calc-wrap {
      padding: 4px 0 8px;
    }
    .calc-intro {
      font-size: 14px;
      color: var(--text-muted);
      margin-bottom: 20px;
      line-height: 1.55;
    }
    .calc-pct-note {
      display: inline-block;
      margin-top: 4px;
      color: var(--brown-mid);
      font-size: 13px;
    }
    .calc-row {
      display: flex;
      align-items: center;
      gap: 16px;
      flex-wrap: wrap;
    }
    .calc-label {
      font-size: 14px;
      font-weight: 700;
      color: var(--brown);
      min-width: 180px;
    }
    .calc-input-group {
      display: flex;
      align-items: center;
      gap: 0;
      border: 2px solid var(--cream-deep);
      border-radius: var(--radius-sm);
      overflow: hidden;
      transition: border-color .2s;
    }
    .calc-input-group:focus-within {
      border-color: var(--green-dark);
      box-shadow: 0 0 0 3px rgba(125,196,0,.13);
    }
    .calc-input {
      border: none;
      outline: none;
      padding: 10px 14px;
      font-family: 'Nunito', sans-serif;
      font-size: 16px;
      font-weight: 700;
      width: 110px;
      background: var(--white);
      color: var(--brown);
    }
    .calc-unit {
      padding: 10px 14px;
      background: var(--cream-deep);
      font-size: 13px;
      font-weight: 800;
      color: var(--brown-mid);
      letter-spacing: .3px;
    }
    .calc-result {
      margin-top: 20px;
      background: var(--cream);
      border: 2px solid var(--green);
      border-radius: var(--radius);
      padding: 18px 22px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .calc-result-label {
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 1px;
      text-transform: uppercase;
      color: var(--green-dark);
    }
    .calc-result-value {
      font-size: 32px;
      font-weight: 900;
      color: var(--brown);
      font-family: 'Playfair Display', serif;
      line-height: 1.1;
    }
    .calc-result-unit {
      font-size: 18px;
      font-weight: 700;
      color: var(--brown-mid);
    }
    .calc-result-note {
      font-size: 12px;
      color: var(--text-muted);
      margin-top: 4px;
      line-height: 1.5;
    }
    @media (max-width: 560px) {
      .calc-row { flex-direction: column; align-items: flex-start; gap: 10px; }
      .calc-input { width: 90px; }
      .calc-result-value { font-size: 26px; }
    }
"""

# ── Translate helpers ──────────────────────────────────────────────────────────
def translate_ingr(text):
    if not text: return ""
    clean = text.strip().replace('\xa0', ' ').replace('\u202f', ' ')
    # Check translation cache first
    for k in [clean[:120], text[:120]]:
        if k in _cache:
            return _cache[k]
    # Fallback: strip footer + dictionary
    for stop in ['Analytick', 'Vitamíny', 'Vitaminy', 'Kontakt', 'Contact\n', '©', 'sales@',
                 'Gratis verzending', 'Přirozené krmivo', 'Natuurlijk voer voor']:
        if stop in clean:
            clean = clean.split(stop)[0]
    clean = clean.strip()
    for cz, nl in sorted(INGR_TRANS.items(), key=lambda x: -len(x[0])):
        clean = clean.replace(cz, nl)
    return clean

def translate_dose(text):
    if not text: return ""
    clean = text.strip().replace('\xa0', ' ').replace('\u202f', ' ')
    # Check translation cache first (populated by translate_all.py)
    for k in [clean[:120], text[:120]]:
        if k in _cache:
            return _cache[k]
    # Fallback: dictionary-based translation
    clean = re.split(r'Kontakt|Contact\n|©|sales@', clean)[0].strip()
    for cz, nl in sorted(DOSE_TRANS.items(), key=lambda x: -len(x[0])):
        clean = clean.replace(cz, nl)
    return clean

# ── Anthropic vertaling ────────────────────────────────────────────────────────
_cache = {}
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, encoding="utf-8") as f:
        _cache = json.load(f)
    print(f"   Cache geladen: {len(_cache)} vertalingen")

CZECH_WORDS = ["jsou", "pro psy", "krmivo", "granule", "pamlsky",
               "vašeho", "psích", "doplněk", "potrava", "sušené mleté",
               "lisované", "přirozené", "šetrným", "surovin"]

def is_czech(text):
    if not text: return True
    if text.startswith('[CZ]'): return True
    return any(w in text for w in CZECH_WORDS)

def translate_via_api(text):
    if not text or len(text) < 20:
        return text
    import hashlib, re as _re
    _TR = _re.compile(r"\[(CZ|NL|EN|DE|FR|SK|PL)\]\s*", _re.IGNORECASE)
    _NB = _re.compile(r"[\xa0\u202f\u2009\u200b\t]+")
    def _norm(t):
        t = _TR.sub("", t); t = _NB.sub(" ", t)
        t = _re.sub(r"\n{2,}", "\n", t); t = _re.sub(r" *\n *", "\n", t)
        return _re.sub(r" {2,}", " ", t).strip()
    def _key(t): return hashlib.sha256(_norm(t).encode()).hexdigest()
    # Try SHA256 key first (v2 cache)
    sha_key = _key(text)
    if sha_key in _cache:
        return _cache[sha_key]
    # Fallback: legacy key formats
    clean = text.strip().replace("\xa0", " ").replace("\u202f", " ")
    for k in [clean[:120], text[:120], text[:100], clean[:100]]:
        if k in _cache:
            return _cache[k]
    api_key = os.environ.get("OPENROUTER_API_KEY", "") or os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return text
    import urllib.request
    prompt = (
            "Vertaal deze tekst naar professioneel Nederlands voor een Belgische honden- en kattenvoer webshop. "
            "Geef ALLEEN de Nederlandse vertaling terug, geen uitleg of extra tekst. "
            "Toon: vriendelijk, professioneel.\n\n" + clean[:1500]
    )
    payload = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 800,
        "messages": [{"role": "user", "content": prompt}]
    }).encode()
    import hashlib as _hl, re as _re2
    _TR2 = _re2.compile(r"\[(CZ|NL|EN|DE|FR|SK|PL)\]\s*", _re2.IGNORECASE)
    _NB2 = _re2.compile(r"[\xa0\u202f\u2009\u200b\t]+")
    def _norm2(t):
        t = _TR2.sub("", t); t = _NB2.sub(" ", t)
        t = _re2.sub(r"\n{2,}", "\n", t); t = _re2.sub(r" *\n *", "\n", t)
        return _re2.sub(r" {2,}", " ", t).strip()
    sha_key = _hl.sha256(_norm2(text).encode()).hexdigest()

    for attempt in range(3):
        try:
            payload = json.dumps({
                "model": "deepseek/deepseek-v4-flash:free",
                "messages": [
                    {"role": "system", "content": "Vertaal van Tsjechisch naar professioneel Nederlands voor yoggies.be. Gebruik 'uw hond'. ALLEEN Nederlandse tekst. Geen tags. Geen Tsjechische woorden."},
                    {"role": "user", "content": clean[:2000]}
                ]
            }).encode()
            req = urllib.request.Request(
                "https://openrouter.ai/api/v1/chat/completions",
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {api_key}",
                    "HTTP-Referer": "https://yoggies.be",
                    "X-Title": "Yoggies NL Builder"
                }
            )
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read())
                result = data["choices"][0]["message"]["content"].strip()
                result = _TR2.sub("", result).strip()
                _cache[sha_key] = result
                _cache[clean[:120]] = result
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(_cache, f, ensure_ascii=False)
                return result
        except Exception as e:
            if attempt < 2:
                time.sleep(2)
            else:
                return text
    return text

# ── Lees data ──────────────────────────────────────────────────────────────────
print("📂 Data lezen...")
wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
ws = wb.active
headers = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
products = [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]
wb.close()
print(f"   {len(products)} producten geladen")

# ── Lees templates ─────────────────────────────────────────────────────────────
with open(PROD_TMPL, encoding="utf-8") as f: PROD_HTML = f.read()
with open(CAT_TMPL,  encoding="utf-8") as f: CAT_HTML  = f.read()
with open(HOME_TMPL, encoding="utf-8") as f: HOME_HTML = f.read()

# Inject calculator CSS into product template <style> block
if "calc-wrap" not in PROD_HTML:
    PROD_HTML = PROD_HTML.replace("</style>", CALC_CSS + "\n  </style>", 1)

# ── Category mapping ───────────────────────────────────────────────────────────
CAT_INFO = {
    "barf":        {"nl": "BARF",                   "file": "barf.html",                  "nav": "BARF"},
    "blikvoer":    {"nl": "Blikvoer",                "file": "blikvoer.html",              "nav": "Blikvoer"},
    "brokken":     {"nl": "Brokken met probiotica",  "file": "brokken-met-probiotica.html","nav": "Brokken met probiotica"},
    "snacks":      {"nl": "Snacks",                  "file": "snacks.html",                "nav": "Snacks"},
    "gekookt":     {"nl": "Vers gekookt voer",       "file": "vers-gekookt-voer.html",     "nav": "Vers gekookt voer"},
    "supplementen":{"nl": "Supplementen",            "file": "supplementen.html",          "nav": "Supplementen"},
    "speelgoed":   {"nl": "Speelgoed & Accessoires", "file": "speelgoed-accessoires.html", "nav": "Speelgoed &amp; Accessoires"},
    "overig":      {"nl": "Overig",                  "file": "barf.html",                  "nav": "BARF"},
}

# ── Groepeer per categorie ─────────────────────────────────────────────────────
by_cat = {k: [] for k in CAT_INFO}
for p in products:
    cat = str(p.get("category") or "overig").strip()
    if cat not in by_cat: cat = "overig"
    by_cat[cat].append(p)

# ── Slugify ────────────────────────────────────────────────────────────────────
def slugify(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^\w\s-]", "", s.lower())
    s = re.sub(r"[\s_]+", "-", s).strip("-")
    return s[:80] or "product"

# ── Nav ────────────────────────────────────────────────────────────────────────
def build_nav(current_cat="", depth="../"):
    parts = [f'<li class="nav-item"><a href="{depth}index.html" class="nav-link nav-home">HOME</a></li>']
    for cat_key, info in CAT_INFO.items():
        if cat_key == "overig": continue
        active = ' style="color:var(--green-dark)"' if cat_key == current_cat else ""
        cat_file = f"{depth}{info['file']}"
        items = by_cat.get(cat_key, [])[:10]
        dd_items = []
        for p in items:
            name = str(p.get("name_nl") or "Product").strip()
            img  = str(p.get("image_main") or p.get("doglife_image") or "").strip()
            slug = slugify(name)
            href = f"{depth}products/{slug}.html"
            img_tag = f'<img class="dd-thumb" src="{html.escape(img)}" alt="" loading="lazy">' if img else '<span class="dd-thumb" style="width:38px;height:38px;background:var(--cream-deep);display:inline-block;border-radius:5px;"></span>'
            dd_items.append(f'<a href="{href}" class="dd-item">{img_tag}{html.escape(name[:40])}</a>')
        dd_html = "\n        ".join(dd_items)
        parts.append(f'<li class="nav-item"><a href="{cat_file}" class="nav-link"{active}>{info["nav"]}</a><div class="dropdown"><div class="dropdown-inner">{dd_html}</div></div></li>')
    return "\n".join(parts)

def inject_nav(template, current_cat="", depth="../"):
    nav_content = build_nav(current_cat, depth)
    new_nav = f'<ul class="nav-list">\n{nav_content}\n  </ul>'
    return re.sub(r'<ul class="nav-list">.*?</ul>', new_nav, template, flags=re.DOTALL, count=1)

# ── Bouw productpagina ─────────────────────────────────────────────────────────
def build_product_page(p, slug):
    page = PROD_HTML

    name_nl  = str(p.get("name_nl") or "Product").strip()
    desc_nl  = str(p.get("description_nl") or "").strip()
    ingr_raw = str(p.get("ingredients_cz") or "").strip()
    dose_raw = str(p.get("dosage_cz") or "").strip()
    img_main = str(p.get("image_main") or "").strip()
    img2     = str(p.get("image_2") or "").strip()
    img3     = str(p.get("image_3") or "").strip()
    dl_url   = str(p.get("doglife_url") or "").strip()
    dl_img   = str(p.get("doglife_image") or "").strip()
    icons_t  = str(p.get("icons_text") or "").strip()
    icons_i  = str(p.get("icons_images") or "").strip()
    cat      = str(p.get("category") or "overig").strip()
    has_calc = str(p.get("has_calculator") or "").strip().upper() == "TRUE"

    # Strip footer uit beschrijving
    for _stop in ['Gratis verzending', 'Přirozené krmivo', 'Natuurlijk voer voor', '©', 'Contact\n\n*']:
        if _stop in desc_nl:
            desc_nl = desc_nl.split(_stop)[0]
    desc_nl = desc_nl.strip()

    # Strip [CZ] prefix, try cache with both original and stripped key
    desc_original = desc_nl
    desc_stripped = desc_nl[4:].strip() if desc_nl.startswith("[CZ]") else desc_nl.strip()

    if is_czech(desc_nl):
        # Cache stores by text[:120] — try original first, then stripped
        result = translate_via_api(desc_original)
        if result == desc_original:
            result = translate_via_api(desc_stripped)
        desc_nl = result
    else:
        desc_nl = desc_stripped

    ingr_nl = translate_ingr(ingr_raw)
    dose_nl = translate_dose(dose_raw)

    # Hoofdafbeelding
    main_img = img_main or dl_img or ""
    if main_img:
        main_img_html = f'<img src="{html.escape(main_img)}" alt="{html.escape(name_nl)}" loading="lazy">'
    else:
        main_img_html = '<div style="display:flex;align-items:center;justify-content:center;height:100%;font-size:56px">📦</div>'

    page = re.sub(r'<div class="ph-main-img"[^>]*>.*?</div>',
                  f'<div class="ph-main-img" id="phMainImg">{main_img_html}</div>',
                  page, flags=re.DOTALL, count=1)

    # Thumbnails
    thumb_imgs = [i for i in [main_img, img2, img3, dl_img] if i and i.startswith('http')][:4]
    thumbs_html = ""
    for i, src in enumerate(thumb_imgs):
        active = " active" if i == 0 else ""
        thumbs_html += f'<div class="ph-thumb{active}" onclick="phSwap(this)"><img src="{html.escape(src)}" alt="" loading="lazy"></div>\n'
    if not thumbs_html:
        thumbs_html = '<div class="ph-thumb active" onclick="phSwap(this)"><span style="display:block;width:60px;height:60px;background:var(--cream-deep);border-radius:5px"></span></div>'
    page = re.sub(r'<div class="ph-thumbs">.*?</div>',
                  f'<div class="ph-thumbs">{thumbs_html}</div>',
                  page, flags=re.DOTALL, count=1)

    # Badges
    icon_labels = [ICON_TRANS.get(i.strip(), i.strip()) for i in icons_t.split("|") if i.strip()][:3]
    icon_imgs   = [i.strip() for i in icons_i.split("|") if i.strip()][:3]
    badges_html = ""
    for idx in range(min(len(icon_labels), 3)):
        label = icon_labels[idx]
        iurl  = icon_imgs[idx] if idx < len(icon_imgs) else ""
        img_part = f'<img src="{html.escape(iurl)}" alt="{html.escape(label)}" style="width:24px;height:24px">' if iurl else "🌿"
        badges_html += f'<div class="ph-badge"><div class="ph-badge-icon">{img_part}</div><span class="ph-badge-label badge-label-{idx+1}">{html.escape(label)}</span></div>\n'
    page = re.sub(r'<div class="ph-badges">.*?</div>',
                  f'<div class="ph-badges">{badges_html}</div>',
                  page, flags=re.DOTALL, count=1)

    # Bestelknop
    if dl_url and dl_url.startswith("http"):
        btn_html = f'<a href="{html.escape(dl_url)}" class="ph-order-btn" target="_blank" rel="noopener">Bestel bij DogLife.be →</a>'
    else:
        btn_html = '<button class="ph-order-btn" disabled style="opacity:.5;cursor:default">Niet beschikbaar bij DogLife</button>'
    page = re.sub(r'<button class="ph-order-btn"[^>]*>.*?</button>', btn_html, page, flags=re.DOTALL, count=1)

    # Beschrijving
    if desc_nl and len(desc_nl) > 20:
        desc_paras = [l.strip() for l in desc_nl.split("\n") if l.strip()][:6]
        desc_body = "\n".join(f"<p>{html.escape(l)}</p>" for l in desc_paras)
    else:
        desc_body = "<p>Beschrijving binnenkort beschikbaar.</p>"
    page = re.sub(
        r'(accordion-summary\">Beschrijving</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + desc_body + r'\n      \2', page, flags=re.DOTALL, count=1)

    # Ingrediënten
    if ingr_nl:
        ingr_lines = [l.strip() for l in ingr_nl.split("\n") if l.strip()]
        ingr_body = "\n".join(f"<p>{html.escape(l)}</p>" for l in ingr_lines[:12])
    else:
        ingr_body = "<p>Zie de verpakking voor ingrediënteninformatie.</p>"
    page = re.sub(
        r'(accordion-summary\">Ingrediënten[^<]*</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + ingr_body + r'\n      \2', page, flags=re.DOTALL, count=1)

    # Dosering
    if dose_nl:
        dose_lines = [l.strip() for l in dose_nl.split("\n") if l.strip()]
        dose_body = "\n".join(f"<p>{html.escape(l)}</p>" for l in dose_lines[:8])
    else:
        dose_body = "<p>Raadpleeg de verpakking voor voedingsadvies.</p>"
    page = re.sub(
        r'(accordion-summary\">Voedingsadvies[^<]*</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + dose_body + r'\n      \2', page, flags=re.DOTALL, count=1)

    # ── Calculator accordion invoegen (na Voedingsadvies, voor Verzending) ──
    if has_calc and dose_raw:
        calc_html = build_calculator_html(dose_raw)
        page = re.sub(
            r'(<details class="accordion">\s*<summary class="accordion-summary">Verzending)',
            calc_html + r'\n\n    \1',
            page, count=1
        )

    # Productnaam in h1
    page = page.replace(
        '<h1 class="ph-product-title"><!-- TITLE HERE --></h1>',
        f'<h1 class="ph-product-title">{html.escape(name_nl)}</h1>', 1
    )

    # Paginatitel
    page = page.replace("<title>", f"<title>{html.escape(name_nl)} | Yoggies.be — ", 1)

    # Nav + links
    page = inject_nav(page, current_cat=cat, depth="../")
    page = page.replace('href="index.html"', 'href="../index.html"')
    for ci in CAT_INFO.values():
        page = page.replace(f'href="{ci["file"]}"', f'href="../{ci["file"]}"')

    return page

# ── Bouw categoriepagina ───────────────────────────────────────────────────────
def build_cat_card(p, slug):
    name = str(p.get("name_nl") or "Product").strip()
    img  = str(p.get("image_main") or p.get("doglife_image") or "").strip()
    desc = str(p.get("description_nl") or "").strip()
    if desc.startswith("[CZ]"): desc = desc[4:].strip()
    desc_short = desc[:120] + "…" if len(desc) > 120 else desc
    img_tag = f'<img src="{html.escape(img)}" alt="{html.escape(name)}" loading="lazy">' if img else '<div style="font-size:48px;display:flex;align-items:center;justify-content:center;height:100%">📦</div>'
    dl_url = str(p.get("doglife_url") or "").strip()
    cart_onclick = f'event.stopPropagation();window.open("{html.escape(dl_url)}","_blank")' if dl_url else ""
    return f'''    <div class="pg-card" onclick="window.location='products/{slug}.html'">
      <div class="pg-img-wrap">{img_tag}</div>
      <div class="pg-body">
        <h2 class="pg-name">{html.escape(name)}</h2>
        <p class="pg-desc">{html.escape(desc_short)}</p>
        <div class="pg-footer">
          <button class="pg-cart-btn" onclick="{cart_onclick}" title="Bestel bij DogLife">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 01-8 0"/>
            </svg>
          </button>
        </div>
      </div>
    </div>'''

def build_category_page(cat_key):
    info  = CAT_INFO[cat_key]
    items = by_cat.get(cat_key, [])
    page  = CAT_HTML
    page = re.sub(r'<title>[^<]*</title>', f'<title>{info["nl"]} | Yoggies.be</title>', page, count=1)
    page = re.sub(r'class="pg-title"[^>]*>[^<]*', f'class="pg-title">{html.escape(info["nl"])}', page, count=1)
    page = re.sub(r'class="pg-count"[^>]*>[^<]*', f'class="pg-count">({len(items)} producten)', page, count=1)

    slugs_used = {}
    cards_html = []
    for p in items:
        name = str(p.get("name_nl") or "Product").strip()
        slug_base = slugify(name)
        if slug_base in slugs_used:
            slugs_used[slug_base] += 1
            slug = f"{slug_base}-{slugs_used[slug_base]}"
        else:
            slugs_used[slug_base] = 0
            slug = slug_base
        cards_html.append(build_cat_card(p, slug))

    cards_block = "\n".join(cards_html)
    page = re.sub(
        r'<div class="pg-grid">.*?</div>\s*(?=\s*<!--|<\/section|<div class="pg-pagination)',
        f'<div class="pg-grid">\n{cards_block}\n  </div>\n  ',
        page, flags=re.DOTALL, count=1)
    page = inject_nav(page, current_cat=cat_key, depth="")
    return page

# ── Genereer alle pagina's ─────────────────────────────────────────────────────
print("\n📄 Productpagina's schrijven...")
slug_map   = {}
prod_count = 0
calc_count = 0

api_key = os.environ.get("ANTHROPIC_API_KEY", "")
cache_size = len(_cache)
if cache_size > 0:
    print(f"   ✅ Cache actief — {cache_size} vertalingen geladen")
elif api_key:
    print(f"   ✅ Anthropic API actief — CZ beschrijvingen worden vertaald")
else:
    print(f"   ⚠️  Geen API key en geen cache — CZ tekst wordt overgeslagen")

for p in products:
    name = str(p.get("name_nl") or "Product").strip()
    slug_base = slugify(name)
    if slug_base in slug_map.values():
        existing = [k for k, v in slug_map.items() if v == slug_base]
        slug = f"{slug_base}-{len(existing)+1}"
    else:
        slug = slug_base
    slug_map[name] = slug

    has_calc = str(p.get("has_calculator") or "").strip().upper() == "TRUE"
    if has_calc:
        calc_count += 1

    page_html = build_product_page(p, slug)
    with open(os.path.join(PROD_DIR, f"{slug}.html"), "w", encoding="utf-8") as f:
        f.write(page_html)
    prod_count += 1
    if prod_count % 50 == 0:
        print(f"   ... {prod_count}/{len(products)}")

print(f"   ✅ {prod_count} productpagina's klaar ({calc_count} met calculator)")

# Categoriepagina's
print("\n📁 Categoriepagina's schrijven...")
for cat_key, info in CAT_INFO.items():
    if cat_key == "overig": continue
    page_html = build_category_page(cat_key)
    with open(os.path.join(OUT_DIR, info["file"]), "w", encoding="utf-8") as f:
        f.write(page_html)
    print(f"   ✅ {info['file']} ({len(by_cat.get(cat_key,[]))} producten)")

# Homepage
print("\n🏠 Homepage schrijven...")
home_html = inject_nav(HOME_HTML, current_cat="", depth="")
with open(os.path.join(OUT_DIR, "index.html"), "w", encoding="utf-8") as f:
    f.write(home_html)
print("   ✅ index.html")

# CSS/JS kopiëren
for fname in ["product.css","product.js","hoofdpageproduct.css","hoofdpageproduct.js","home.css"]:
    if os.path.exists(fname):
        shutil.copy(fname, os.path.join(OUT_DIR, fname))

print(f"\n🎉 Klaar! {prod_count} pagina's in ./{OUT_DIR}/")
print(f"\nPer categorie:")
for cat_key, info in CAT_INFO.items():
    if cat_key != "overig":
        print(f"   {info['nl']:<30}: {len(by_cat.get(cat_key,[]))}")