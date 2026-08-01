"""
build_site.py — Yoggies.be Site Builder
========================================
Leest READY_FOR_BUILD_FINAL.xlsx en genereert:
  - site/products/{slug}.html   → 339 productpagina's
  - site/barf.html              → categoriepagina
  - site/blikvoer.html
  - site/brokken-met-probiotica.html
  - site/snacks.html
  - site/vers-gekookt-voer.html
  - site/supplementen.html
  - site/voor-katten.html
  - site/speelgoed-accessoires.html
  - site/yoghurt.html
  - site/index.html             → homepage

Gebruik:
  pip install openpyxl
  python build_site.py
"""

import os, re, html, unicodedata, json, shutil
import openpyxl
import httpx, time

# ── Translation cache (CZ → NL) ─────────────────────────────────────────────────
_CACHE_FILE = "translation_cache.json"
try:
    with open(_CACHE_FILE, encoding="utf-8") as _f:
        _trans_cache = json.load(_f)
except FileNotFoundError:
    _trans_cache = {}

_CZECH_WORDS = frozenset(["jsou", "pro", "psy", "pes", "krmivo", "obsahuje"])

def is_czech(text):
    if not text:
        return False
    if text.startswith("[CZ]"):
        return True
    words = set(re.findall(r'\b\w+\b', text.lower()))
    return bool(words & _CZECH_WORDS)

def translate_nl(cz_text):
    if not cz_text:
        return ""
    if cz_text in _trans_cache:
        return _trans_cache[cz_text]
    api_key = os.environ.get("OPENROUTER_API_KEY", "")
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    prompt = (
            "Vertaal deze Tsjechische productbeschrijving voor hondenvoer naar het Nederlands. "
            "Geef alleen de Nederlandse vertaling terug, geen uitleg of extra tekst.\n\n" + cz_text
    )
    models = ["google/gemma-4-31b-it:free", "openai/gpt-oss-20b:free"]
    data = None
    for model in models:
        payload = {"model": model, "max_tokens": 500,
                   "messages": [{"role": "user", "content": prompt}]}
        for attempt in range(5):
            resp = httpx.post("https://openrouter.ai/api/v1/chat/completions",
                              headers=headers, json=payload, timeout=40)
            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", min(2 ** attempt * 3, 60)))
                print(f"   ⏳ {model} rate-limit, wacht {wait}s…")
                time.sleep(wait)
                continue
            if resp.status_code != 200:
                print(f"   ⚠ {model} HTTP {resp.status_code}, probeer volgend model")
                break
            data = resp.json()
            if "choices" not in data:
                print(f"   ⚠ {model} geen choices: {data.get('error',{}).get('message','?')[:60]}")
                data = None
                break
            break
        if data:
            break
    if not data:
        return ""
    result = data["choices"][0]["message"]["content"].strip()
    _trans_cache[cz_text] = result
    with open(_CACHE_FILE, "w", encoding="utf-8") as _f:
        json.dump(_trans_cache, _f, ensure_ascii=False, indent=2)
    time.sleep(1)  # stay under free-tier rate limit
    return result

def clean_ingredients(text):
    for marker in ["Analytické složky", "Vitamíny"]:
        idx = text.find(marker)
        if idx != -1:
            text = text[:idx].strip().rstrip(",").strip()
    return text

# ── Paths ──────────────────────────────────────────────────────────────────────
# Resolved from this file's location so the script runs from any working
# directory (it used to hardcode an absolute path on one developer's machine).
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent

DATA_FILE = str(ROOT / "YOGGIES_FINAL_NL.xlsx")
PROD_TMPL    = str(ROOT / "templates" / "product.html")
CAT_TMPL     = str(ROOT / "templates" / "hoofdpageproduct.html")
HOME_TMPL    = str(ROOT / "templates" / "home.html")
OUT_DIR      = str(ROOT / "site")
PROD_DIR     = os.path.join(OUT_DIR, "products")

os.makedirs(PROD_DIR, exist_ok=True)
CAT_INFO = {
    "barf":        {"nl": "BARF",                   "file": "barf.html",                   "nav": "BARF"},
    "blikvoer":    {"nl": "Blikvoer",                "file": "blikvoer.html",               "nav": "Blikvoer"},
    "brokken":     {"nl": "Brokken met probiotica",  "file": "brokken-met-probiotica.html",  "nav": "Brokken met probiotica"},
    "snacks":      {"nl": "Snacks",                  "file": "snacks.html",                  "nav": "Snacks"},
    "gekookt":     {"nl": "Vers gekookt voer",       "file": "vers-gekookt-voer.html",       "nav": "Vers gekookt voer"},
    "supplementen":{"nl": "Supplementen",            "file": "supplementen.html",            "nav": "Supplementen"},
    "speelgoed":   {"nl": "Speelgoed & Accessoires", "file": "speelgoed-accessoires.html",   "nav": "Speelgoed &amp; Accessoires"},
    "overig":      {"nl": "Overig",                  "file": "barf.html",                    "nav": "BARF"},
}

# ── Slugify ─────────────────────────────────────────────────────────────────────
def slugify(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^\w\s-]", "", s.lower())
    s = re.sub(r"[\s_]+", "-", s).strip("-")
    return s[:80] or "product"

# ── Lees data ───────────────────────────────────────────────────────────────────
print("📂 Data lezen...")
wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
ws = wb.active
headers = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
products = []
for row in ws.iter_rows(min_row=2, values_only=True):
    r = dict(zip(headers, row))
    products.append(r)
wb.close()
print(f"   {len(products)} producten geladen")

# ── Lees templates ──────────────────────────────────────────────────────────────
with open(PROD_TMPL, encoding="utf-8") as f:
    PROD_HTML = f.read()
with open(CAT_TMPL, encoding="utf-8") as f:
    CAT_HTML = f.read()
with open(HOME_TMPL, encoding="utf-8") as f:
    HOME_HTML = f.read()

# ── Groepeer producten per categorie ────────────────────────────────────────────
by_cat = {k: [] for k in CAT_INFO}
for p in products:
    cat = str(p.get("category") or "overig").strip()
    if cat not in by_cat:
        cat = "overig"
    by_cat[cat].append(p)

# ── Build nav dropdown HTML ─────────────────────────────────────────────────────
def build_nav(current_cat="", depth="../"):
    """Genereer de <ul class='nav-list'> met echte producten in dropdowns."""
    parts = []
    parts.append(f'<li class="nav-item"><a href="{depth}index.html" class="nav-link nav-home">HOME</a></li>')

    for cat_key, info in CAT_INFO.items():
        if cat_key == "overig":
            continue
        active = ' style="color:var(--green-dark)"' if cat_key == current_cat else ""
        cat_file = f"{depth}{info['file']}"
        items = by_cat.get(cat_key, [])[:10]

        dd_items = []
        for p in items:
            name = str(p.get("name_nl") or p.get("name_cz") or "Product").strip()
            img  = str(p.get("image_main") or p.get("doglife_image") or "").strip()
            slug = slugify(name)
            href = f"{depth}products/{slug}.html"
            img_tag = f'<img class="dd-thumb" src="{html.escape(img)}" alt="{html.escape(name)}" loading="lazy">' if img else '<span class="dd-thumb" style="width:38px;height:38px;background:var(--cream-deep);display:inline-block;border-radius:5px;"></span>'
            dd_items.append(f'<a href="{href}" class="dd-item">{img_tag}{html.escape(name[:40])}</a>')

        dd_html = "\n        ".join(dd_items) if dd_items else ""
        parts.append(f"""    <li class="nav-item">
      <a href="{cat_file}" class="nav-link"{active}>{info['nav']}</a>
      <div class="dropdown"><div class="dropdown-inner">
        {dd_html}
      </div></div>
    </li>""")

    return "\n".join(parts)

def inject_nav(template, current_cat="", depth="../"):
    """Vervang de nav-list in template door echte nav."""
    nav_content = build_nav(current_cat, depth)
    # Vervang alles tussen <ul class="nav-list"> en </ul>
    new_nav = f'<ul class="nav-list">\n{nav_content}\n  </ul>'
    result = re.sub(
        r'<ul class="nav-list">.*?</ul>',
        new_nav,
        template,
        flags=re.DOTALL,
        count=1
    )
    return result

# ── Build product page ──────────────────────────────────────────────────────────
def build_product_page(p, slug):
    page = PROD_HTML

    name_nl  = str(p.get("name_nl")       or p.get("name_cz") or "Product").strip()
    desc_nl  = str(p.get("description_nl") or "").strip()
    ingr_cz  = clean_ingredients(str(p.get("ingredients_cz") or "").strip())
    dose_cz  = str(p.get("dosage_cz")      or "").strip()
    img_main = str(p.get("image_main")     or "").strip()
    img2     = str(p.get("image_2")        or "").strip()
    img3     = str(p.get("image_3")        or "").strip()
    dl_url   = str(p.get("doglife_url")    or "").strip()
    dl_img   = str(p.get("doglife_image")  or "").strip()
    icons_t  = str(p.get("icons_text")     or "").strip()
    icons_i  = str(p.get("icons_images")   or "").strip()
    has_calc = str(p.get("has_calculator") or "FALSE").strip().upper() == "TRUE"
    cat      = str(p.get("category")       or "overig").strip()

    # Gebruik Yoggies-foto als hoofdafbeelding, anders Doglife-foto
    main_img = img_main or dl_img or ""

    # ── Title ──────────────────────────────────────────────────────────────────
    page = page.replace("<title>", f"<title>{html.escape(name_nl)} | Yoggies.be — ")
    page = re.sub(
        r'class="ph-product-title"[^>]*>.*?</h1>',
        f'class="ph-product-title">{html.escape(name_nl)}</h1>',
        page, flags=re.DOTALL, count=1
    )

    # ── Hoofdafbeelding ────────────────────────────────────────────────────────
    if main_img:
        main_img_html = f'<img src="{html.escape(main_img)}" alt="{html.escape(name_nl)}" id="phMainImg" loading="lazy">'
    else:
        main_img_html = f'<div style="display:flex;align-items:center;justify-content:center;height:100%;font-size:56px">📦</div>'

    page = re.sub(
        r'<div class="ph-main-img"[^>]*>.*?</div>',
        f'<div class="ph-main-img" id="phMainImg">{main_img_html}</div>',
        page, flags=re.DOTALL, count=1
    )

    # ── Thumbnails ─────────────────────────────────────────────────────────────
    thumb_imgs = [i for i in [main_img, img2, img3, dl_img] if i][:4]
    thumbs_html = ""
    for i, src in enumerate(thumb_imgs):
        active = " active" if i == 0 else ""
        thumbs_html += f'<div class="ph-thumb{active}" onclick="phSwap(this)"><img src="{html.escape(src)}" alt="{html.escape(name_nl)}" loading="lazy"></div>\n'
    if not thumbs_html:
        thumbs_html = f'<div class="ph-thumb active" onclick="phSwap(this)"><span style="display:block;width:60px;height:60px;background:var(--cream-deep);border-radius:5px"></span></div>'

    page = re.sub(
        r'<div class="ph-thumbs">.*?</div>',
        f'<div class="ph-thumbs">{thumbs_html}</div>',
        page, flags=re.DOTALL, count=1
    )

    # ── Badges / Iconen ────────────────────────────────────────────────────────
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
    icon_labels = [ICON_TRANS.get(i.strip(), i.strip()) for i in icons_t.split("|") if i.strip()][:3]
    icon_imgs   = [i.strip() for i in icons_i.split("|") if i.strip()][:3]

    badges_html = ""
    for idx in range(3):
        label = icon_labels[idx] if idx < len(icon_labels) else ""
        iurl  = icon_imgs[idx]   if idx < len(icon_imgs)   else ""
        if label or iurl:
            img_part = f'<img src="{html.escape(iurl)}" alt="{html.escape(label)}" class="badge-img">' if iurl else "🌿"
            badges_html += f'''<div class="ph-badge">
          <div class="ph-badge-icon">{img_part}</div>
          <span class="ph-badge-label badge-label-{idx+1}">{html.escape(label)}</span>
        </div>\n'''

    page = re.sub(
        r'<div class="ph-badges">.*?</div>',
        f'<div class="ph-badges">{badges_html}</div>',
        page, flags=re.DOTALL, count=1
    )

    # ── Bestelknop ─────────────────────────────────────────────────────────────
    if dl_url and dl_url.startswith("http"):
        btn_html = f'<a href="{html.escape(dl_url)}" class="ph-order-btn" target="_blank" rel="noopener">Bestel bij DogLife.be →</a>'
    else:
        btn_html = '<button class="ph-order-btn" disabled style="opacity:.5;cursor:default">Niet beschikbaar bij DogLife</button>'

    page = re.sub(
        r'<button class="ph-order-btn"[^>]*>.*?</button>',
        btn_html,
        page, flags=re.DOTALL, count=1
    )

    # ── Beschrijving accordion ─────────────────────────────────────────────────
    desc_text = desc_nl
    if desc_text:
        desc_paras = [l.strip() for l in desc_text.split("\n") if l.strip()][:6]
        desc_body = "\n".join(f"<p>{html.escape(p)}</p>" for p in desc_paras)
    else:
        desc_body = "<p>Beschrijving binnenkort beschikbaar.</p>"

    page = re.sub(
        r'(accordion-summary\">Beschrijving</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + desc_body + r'\n      \2',
        page, flags=re.DOTALL, count=1
    )

    # ── Ingrediënten accordion ─────────────────────────────────────────────────
    INGR_WORD_TRANS = {
        "sušené mleté maso": "gedroogd gemalen vlees",
        "sušené rybí maso": "gedroogd visvlees",
        "sušené maso": "gedroogd vlees",
        "hovězí maso": "rundvlees", "hovězí": "rund",
        "kuřecí maso": "kippenvlees", "kuřecí": "kip",
        "krůtí maso": "kalkoenvlees", "krůtí": "kalkoen",
        "jehněčí maso": "lamsvlees", "jehněčí": "lam",
        "kachní maso": "eendenvlees", "kachní": "eend",
        "jelení maso": "hertenvlees", "jelení": "hert",
        "lososové maso": "zalmvlees", "losos": "zalm",
        "vepřové maso": "varkensvlees", "vepřové": "varken",
        "zvěřinové maso": "wildvlees", "zvěřina": "wild",
        "srdce": "hart", "játra": "lever", "žaludky": "maag",
        "hnědá rýže": "bruine rijst", "rýže": "rijst",
        "bramborové vločky": "aardappelvlokken", "brambory": "aardappelen",
        "proso hrubomleté": "grof gemalen gierst", "proso": "gierst",
        "pohanka": "boekweit", "oves": "haver", "ovesné": "haver",
        "rýžové sušené klíčky": "gedroogde rijstkiemen",
        "želatinový prášek": "gelatinepoeder",
        "pivovarské kvasnice": "biergist", "kvasnice": "gist",
        "konopný olej": "hennepolie", "lněný olej": "lijnzaadolie",
        "lososový olej": "zalmolie", "rybí olej": "visolie",
        "probiotika": "probiotica", "prebiotika": "prebiotica",
        "enzymy": "enzymen",
        "batáty": "zoete aardappelen",
        "mrkev": "wortel", "hrášek": "erwten",
        "špenát": "spinazie", "dýně": "pompoen",
        "brusinky": "veenbessen", "šípky": "rozenbottels",
        "chlorella": "chlorella", "kelp": "kelp",
        "Lithothamnium": "Lithothamnium",
        "polyfenoly": "polyfenolen",
        "quinoa": "quinoa",
        "sůl": "zout", "voda": "water", "vývar": "bouillon",
        "Neobsahuje": "Bevat geen",
        "pářáty, hlavy a peří": "poten, koppen en veren",
        "Většina mas použitých pro výrobu": "Het meeste vlees gebruikt voor de productie van",
        "je v potravinářské kvalitě": "is van voedingskwaliteit",
        "Bez konzervantů": "Zonder conserveringsmiddelen",
    }

    def translate_ingr(text):
        if not text: return ""
        for cz, nl in sorted(INGR_WORD_TRANS.items(), key=lambda x: -len(x[0])):
            text = text.replace(cz, nl)
        return text

    if ingr_cz:
        ingr_lines = [l.strip() for l in ingr_cz.split("\n") if l.strip()]
        ingr_body = "\n".join(f"<p>{html.escape(translate_ingr(l))}</p>" for l in ingr_lines[:10])
    else:
        ingr_body = "<p>Zie de verpakking voor ingrediënteninformatie.</p>"

    page = re.sub(
        r'(accordion-summary\">Ingrediënten[^<]*</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + ingr_body + r'\n      \2',
        page, flags=re.DOTALL, count=1
    )

    # ── Dosering accordion ─────────────────────────────────────────────────────
    DOSE_TRANS = {
        "Doporučené krmné dávky:": "Aanbevolen dagelijkse porties:",
        "Doporučená krmná dávka": "Aanbevolen dagelijkse portie",
        "doporučené množství krmiva na den odpovídá zhruba": "de aanbevolen hoeveelheid voer per dag komt overeen met ongeveer",
        "tělesné hmotnosti": "van het lichaamsgewicht",
        "dospělého psa": "van de volwassen hond",
        "Přizpůsobte denní krmnou dávku stáří a aktivitě vašeho psa.": "Pas de dagelijkse portie aan op de leeftijd en activiteit van uw hond.",
        "U štěňat může množství krmiva": "Bij pups kan de hoeveelheid voer",
        "přesahovat až": "oplopen tot",
        "denně": "per dag",
        "Podávejte při pokojové teplotě.": "Serveer op kamertemperatuur.",
        "Váš pes b": "Uw hond",
        "denní dávku": "dagelijkse portie",
        "tělesné hmotnosti zdravého psa": "lichaamsgewicht van een gezonde hond",
        "krmiva na den": "voer per dag",
        "krmivo": "voer",
    }

    def translate_dose(text):
        if not text: return ""
        for cz, nl in sorted(DOSE_TRANS.items(), key=lambda x: -len(x[0])):
            text = text.replace(cz, nl)
        return text

    if dose_cz:
        dose_lines = [l.strip() for l in dose_cz.split("\n") if l.strip()]
        dose_body = "\n".join(f"<p>{html.escape(translate_dose(l))}</p>" for l in dose_lines[:8])
    else:
        dose_body = "<p>Raadpleeg de verpakking voor voedingsadvies.</p>"

    page = re.sub(
        r'(accordion-summary\">Voedingsadvies[^<]*</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + dose_body + r'\n      \2',
        page, flags=re.DOTALL, count=1
    )

    # ── Nav injecteren ─────────────────────────────────────────────────────────
    page = inject_nav(page, current_cat=cat, depth="../")

    # ── Fix relatieve paden ────────────────────────────────────────────────────
    page = page.replace('href="index.html"', 'href="../index.html"')
    for c_info in CAT_INFO.values():
        page = page.replace(f'href="{c_info["file"]}"', f'href="../{c_info["file"]}"')

    return page

# ── Build category page ─────────────────────────────────────────────────────────
def build_cat_card(p, slug):
    name = str(p.get("name_nl") or p.get("name_cz") or "Product").strip()
    img  = str(p.get("image_main") or p.get("doglife_image") or "").strip()
    desc = str(p.get("description_nl") or p.get("description_cz") or "").strip()
    desc_short = desc[:120] + "…" if len(desc) > 120 else desc

    img_tag = f'<img src="{html.escape(img)}" alt="{html.escape(name)}" loading="lazy">' if img else \
        f'<div style="display:flex;align-items:center;justify-content:center;height:100%;font-size:48px">📦</div>'

    return f'''    <div class="pg-card" onclick="window.location='products/{slug}.html'">
      <div class="pg-img-wrap">{img_tag}</div>
      <div class="pg-body">
        <h2 class="pg-name">{html.escape(name)}</h2>
        <p class="pg-desc">{html.escape(desc_short)}</p>
        <div class="pg-footer">
          <button class="pg-cart-btn" onclick="event.stopPropagation();window.open('{html.escape(str(p.get('doglife_url') or ''))}','_blank')" title="Bestel bij DogLife">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/>
              <line x1="3" y1="6" x2="21" y2="6"/>
              <path d="M16 10a4 4 0 01-8 0"/>
            </svg>
          </button>
        </div>
      </div>
    </div>'''

def build_category_page(cat_key):
    info  = CAT_INFO[cat_key]
    items = by_cat.get(cat_key, [])
    page  = CAT_HTML

    # Titel
    page = re.sub(r'<title>[^<]*</title>', f'<title>{info["nl"]} | Yoggies.be</title>', page, count=1)
    page = re.sub(r'class="pg-title"[^>]*>[^<]*', f'class="pg-title">{html.escape(info["nl"])}', page, count=1)
    page = re.sub(r'class="pg-count"[^>]*>[^<]*', f'class="pg-count">({len(items)} producten)', page, count=1)

    # Product cards
    slugs_used = {}
    cards_html = []
    for p in items:
        name = str(p.get("name_nl") or p.get("name_cz") or "Product").strip()
        slug_base = slugify(name)
        # Zorg voor unieke slugs
        if slug_base in slugs_used:
            slugs_used[slug_base] += 1
            slug = f"{slug_base}-{slugs_used[slug_base]}"
        else:
            slugs_used[slug_base] = 0
            slug = slug_base
        cards_html.append(build_cat_card(p, slug))

    cards_block = "\n".join(cards_html)
    page = re.sub(
        r'<div class="pg-grid">.*?</div>\s*(?=\s*<!--|\s*</section|\s*<div class="pg-pagination)',
        f'<div class="pg-grid">\n{cards_block}\n  </div>\n  ',
        page, flags=re.DOTALL, count=1
    )

    # Nav
    page = inject_nav(page, current_cat=cat_key, depth="")

    return page

# ── Build homepage ──────────────────────────────────────────────────────────────
def build_homepage():
    page = HOME_HTML
    page = inject_nav(page, current_cat="", depth="")
    return page

# ── Schrijf alle files ──────────────────────────────────────────────────────────
print("\n🔨 Productpagina's schrijven...")

# Maak slug-index aan (voor consistentie tussen product en categorie pagina's)
slug_index = {}
for p in products:
    name = str(p.get("name_nl") or p.get("name_cz") or "Product").strip()
    slug_base = slugify(name)
    if slug_base not in slug_index:
        slug_index[slug_base] = {"count": 0, "product": p}
    else:
        slug_index[slug_base]["count"] += 1

# Genereer productpagina's
prod_count = 0
slug_map = {}  # name → slug (voor category pages)

for p in products:
    name = str(p.get("name_nl") or p.get("name_cz") or "Product").strip()
    slug_base = slugify(name)

    # Unieke slug
    if slug_base in slug_map.values():
        existing = [k for k, v in slug_map.items() if v == slug_base]
        slug = f"{slug_base}-{len(existing)+1}"
    else:
        slug = slug_base
    slug_map[name] = slug

    page_html = build_product_page(p, slug)
    out_path = os.path.join(PROD_DIR, f"{slug}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(page_html)
    prod_count += 1
    if prod_count % 50 == 0:
        print(f"   … {prod_count}/{len(products)}")

print(f"   ✅ {prod_count} productpagina's klaar")

# Genereer categoriepagina's
print("\n📁 Categoriepagina's schrijven...")
for cat_key, info in CAT_INFO.items():
    if cat_key == "overig":
        continue
    page_html = build_category_page(cat_key)
    out_path = os.path.join(OUT_DIR, info["file"])
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(page_html)
    print(f"   ✅ {info['file']} ({len(by_cat.get(cat_key,[]))} producten)")

# Genereer homepage
print("\n🏠 Homepage schrijven...")
home_html = build_homepage()
with open(os.path.join(OUT_DIR, "index.html"), "w", encoding="utf-8") as f:
    f.write(home_html)
print("   ✅ index.html")

# Kopieer CSS/JS
print("\n📎 CSS/JS kopiëren...")
for fname in ["product.css", "product.js", "hoofdpageproduct.css", "hoofdpageproduct.js", "home.css"]:
    src = fname
    if os.path.exists(src):
        shutil.copy(src, os.path.join(OUT_DIR, fname))
        print(f"   ✅ {fname}")

# ── Statistieken ────────────────────────────────────────────────────────────────
print(f"\n🎉 Klaar!")
print(f"   Productpagina's : {prod_count}")
print(f"   Categoriepagina's: {len(CAT_INFO)-1}")
print(f"   Output map      : ./{OUT_DIR}/")
print()
print("📊 Per categorie:")
for cat_key, info in CAT_INFO.items():
    if cat_key != "overig":
        n = len(by_cat.get(cat_key, []))
        print(f"   {info['nl']:<30} : {n}")