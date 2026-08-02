"""
yoggies_pipeline.py  —  Cache-Only Build Pipeline  v3.0
=========================================================
USAGE:
  python yoggies_pipeline.py                # MODE 1: CACHE_ONLY build
  python yoggies_pipeline.py --fill-cache   # MODE 2: translate missing via API
  python yoggies_pipeline.py --audit        # show missing without building

ARCHITECTURE:
  Cache is the SINGLE SOURCE OF TRUTH.
  Every translation: normalize(text) → SHA256 → cache.get(key) → return | FAIL
  No fallback. No dictionary. No partial logic. No legacy keys.

NORMALIZATION (applied before every SHA256 key):
  1. strip [CZ]/[NL]/[EN]/[DE]/[FR]/[SK]/[PL] tags
  2. normalize all unicode whitespace → space
  3. strip markdown prefixes (# ## ###)
  4. collapse consecutive newlines → single
  5. strip spaces around newlines
  6. collapse multiple spaces → single
  7. strip, then lowercase

REMOVED LOGIC (was: fallback paths):
  - INGR_TRANS dictionary          → REMOVED
  - DOSE_TRANS dictionary          → REMOVED
  - translate_ingr() function      → REMOVED
  - translate_dose() function      → REMOVED
  - legacy 120-char key lookup     → REMOVED
  - text[:120] key lookup          → REMOVED
  - any non-cache translation path → REMOVED
"""

import os, sys, re, json, hashlib, html, unicodedata
import urllib.request, time
from urllib.parse import quote
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# Dutch hondenatlas data (handmatig vertaald) — atlas_nl_data.py in projectmap
try:
    from atlas_nl_data import BREEDS_NL as _BREEDS_NL
except Exception:
    _BREEDS_NL = []

# Manual variant overrides (e.g. Brokken sizes that aren't in the Excel)
# See manual_variants.py for the format.
try:
    from manual_variants import MANUAL_VARIANTS as _MANUAL_VARIANTS
except Exception:
    _MANUAL_VARIANTS = {}

# Markers used to make the nav CSS / JS injection idempotent across rebuilds
_NAV_MARK_S = "/* YGS_NAV_CSS_START */"
_NAV_MARK_E = "/* YGS_NAV_CSS_END */"
_NAV_JS_MARK_S = "<!-- YGS_NAV_JS_START -->"
_NAV_JS_MARK_E = "<!-- YGS_NAV_JS_END -->"

def _strip_block(content, start, end):
    """Remove a previously injected marked block.

    Also eats the indentation in front of it and the newline after it. Without
    that, every rebuild left the old whitespace behind and added a fresh blank
    line, so pages the pipeline edits in place grew a little on each run.
    """
    return re.sub(
        r'[ \t]*' + re.escape(start) + r'.*?' + re.escape(end) + r'[ \t]*\n?',
        '', content, flags=re.DOTALL)


def _inject_nav_css(content):
    """Idempotently inject _NAV_CSS into HTML content.
    - Strips any previously-injected nav CSS (between markers).
    - Appends to an existing <style> block if present, otherwise creates a
      standalone <style> block before </head>.
    """
    # Strip previous injection (works inside <style> or standalone)
    content = _strip_block(content, _NAV_MARK_S, _NAV_MARK_E)
    block = _NAV_MARK_S + "\n" + _NAV_CSS.strip() + "\n" + _NAV_MARK_E
    if "</style>" in content:
        content = content.replace("</style>", block + "\n  </style>", 1)
    elif "</head>" in content:
        content = content.replace("</head>", f"<style>{block}</style>\n</head>", 1)
    return content

def _inject_nav_js(content):
    """Idempotently inject _NAV_JS (which already includes <script>...</script>)
    into HTML content before </body>. Strips any previous injection first."""
    content = _strip_block(content, _NAV_JS_MARK_S, _NAV_JS_MARK_E)
    block = _NAV_JS_MARK_S + _NAV_JS + _NAV_JS_MARK_E
    if "</body>" in content:
        content = content.replace("</body>", block + "\n</body>", 1)
    return content

# ─── Config ────────────────────────────────────────────────────────────────────
DATA_FILE   = "READY_FOR_BUILD_FINAL.xlsx"
CACHE_FILE  = "translation_cache_v2.json"
OUTPUT_XLSX = "READY_FOR_BUILD_NL.xlsx"
MISSING_RPT = "missing_report.txt"
SITE_DIR    = Path("site")
PROD_DIR    = SITE_DIR / "products"
# Source templates live outside the output folder: everything under site/ is
# build output and is overwritten on every run.
TPL_DIR     = Path("templates")

# Footer legal links. Both used to point at the privacy policy.
# NOTE: verify TERMS_URL with DogLife — this is the conventional Dutch slug,
# not a URL that has been confirmed to exist.
PRIVACY_URL = "https://doglife.be/privacy/"
TERMS_URL   = "https://doglife.be/algemene-voorwaarden/"

# Static assets that ship with the templates and must be copied into the build.
TPL_ASSETS  = ["home.css", "product.css", "product.js",
               "hoofdpageproduct.css", "hoofdpageproduct.js"]


def _copy_template_assets():
    """Copy the templates' CSS/JS into site/ so the generated pages resolve."""
    import shutil
    SITE_DIR.mkdir(parents=True, exist_ok=True)
    for name in TPL_ASSETS:
        src = TPL_DIR / name
        if src.exists():
            shutil.copyfile(src, SITE_DIR / name)
    # favicon lives under images/ because that is what every page links to
    fav = TPL_DIR / "favicon.svg"
    if fav.exists():
        (SITE_DIR / "images").mkdir(parents=True, exist_ok=True)
        shutil.copyfile(fav, SITE_DIR / "images" / "favicon.svg")


# ─── Localized images ──────────────────────────────────────────────────────────
# localize_images.py downloads every external image into site/images/ and writes
# image_map.json: {url: {thumb, md, lg, orig, w, h}}. Everything the build emits
# resolves through img_src() so no page hotlinks yoggies.cz or doglife.be.
_IMG_MAP_FILE = Path("image_map.json")
_IMG_MAP = json.loads(_IMG_MAP_FILE.read_text(encoding="utf-8")) \
    if _IMG_MAP_FILE.exists() else {}


def _img_canon(url: str) -> str:
    """Canonical identity of an image URL: scheme + host + path, no query.

    Mirrors canonical_url() in localize_images.py so the keys line up — the CDN
    serves the same file under many ?width=/&v= query strings.
    """
    from urllib.parse import urlsplit, urlunsplit
    parts = urlsplit(url.strip())
    scheme = "https" if parts.scheme in ("http", "https", "") else parts.scheme
    netloc = parts.netloc.lower()
    if netloc.startswith("www."):
        netloc = netloc[4:]
    return urlunsplit((scheme, netloc, parts.path, "", ""))


def _img_entry(url):
    """Look up an image_map entry by any spelling of the URL."""
    if not url or not str(url).startswith("http"):
        return None
    clean = str(url).replace("&amp;", "&").strip()
    return (_IMG_MAP.get(url)
            or _IMG_MAP.get(clean)
            or _IMG_MAP.get(_img_canon(clean)))


def img_src(url, size="md", depth=""):
    """Return a local path for an external image URL, falling back to the URL.

    size: "thumb" (200px) | "md" (600px) | "lg" (1200px) | "orig".
    depth: path prefix for pages below site/ (e.g. "../" for products/).
    SVGs have no raster variants, so every size resolves to the original.
    """
    entry = _img_entry(url)
    if not entry:
        return url
    return depth + entry.get(size, entry.get("orig", url))


# ─── <head> metadata ───────────────────────────────────────────────────────────
BASE_URL = "https://yoggies.be"

_META_MARK_S = "<!-- YGS_META_START -->"
_META_MARK_E = "<!-- YGS_META_END -->"
_META_BLOCK_RE = re.compile(
    re.escape(_META_MARK_S) + r'.*?' + re.escape(_META_MARK_E), re.DOTALL)
_TAG_STRIP_RE = re.compile(r'<(script|style)\b.*?</\1>|<[^>]+>', re.DOTALL | re.IGNORECASE)


def _plain_text(markup: str) -> str:
    """Collapse HTML to a single line of readable text."""
    t = _TAG_STRIP_RE.sub(" ", str(markup or ""))
    return " ".join(html.unescape(t).split())


def _meta_desc(text: str, limit: int = 155) -> str:
    """First `limit` characters of text, cut on a word boundary, escaped."""
    return html.escape(_truncate_words(_plain_text(text), limit), quote=True)


def _abs_url(path: str) -> str:
    """Site-root-relative path -> absolute URL (og: tags require absolute)."""
    return f"{BASE_URL}/{str(path or '').lstrip('/')}"


def _head_block(title, description, url_path, image=None,
                depth="", og_type="website", robots=None) -> str:
    """Build the meta/canonical/OG/Twitter block for one page.

    url_path and image are site-root-relative ("products/x.html",
    "images/y-1200.webp"); depth only prefixes the local favicon link.
    """
    t = html.escape(_plain_text(title), quote=True)
    d = description if description else t
    canonical = _abs_url(url_path)
    # An unmapped image_src() falls back to the original URL, which is already
    # absolute — don't prefix it a second time.
    image = image or "images/favicon.svg"
    og_img = image if str(image).startswith("http") else _abs_url(image)
    parts = [
        _META_MARK_S,
        f'<meta name="description" content="{d}">',
        f'<link rel="canonical" href="{canonical}">',
        f'<link rel="icon" type="image/svg+xml" href="{depth}images/favicon.svg">',
        f'<meta property="og:type" content="{og_type}">',
        f'<meta property="og:site_name" content="Yoggies.be">',
        f'<meta property="og:locale" content="nl_BE">',
        f'<meta property="og:title" content="{t}">',
        f'<meta property="og:description" content="{d}">',
        f'<meta property="og:url" content="{canonical}">',
        f'<meta property="og:image" content="{og_img}">',
        f'<meta name="twitter:card" content="summary_large_image">',
        f'<meta name="twitter:title" content="{t}">',
        f'<meta name="twitter:description" content="{d}">',
        f'<meta name="twitter:image" content="{og_img}">',
    ]
    if robots:
        parts.insert(1, f'<meta name="robots" content="{robots}">')
    parts.append(_META_MARK_E)
    return "\n  ".join(parts)


def _redirect_stub(target: str, title: str, canonical_path: str,
                   note: str = "Deze pagina is verplaatst.") -> str:
    """A noindex meta-refresh stub that keeps an old URL working.

    canonical_path points at the *destination* so link equity consolidates
    there rather than on the retired URL.
    """
    t = html.escape(title, quote=True)
    return (
        '<!DOCTYPE html><html lang="nl"><head>'
        '<meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<meta name="robots" content="noindex,follow">'
        f'<meta http-equiv="refresh" content="0; url={target}">'
        f'<meta name="description" content="{t}">'
        f'<link rel="canonical" href="{_abs_url(canonical_path)}">'
        f'<title>{t}</title>'
        '</head><body>'
        f'<p>{html.escape(note)} <a href="{target}">Ga verder</a>.</p>'
        f'<script>window.location.replace("{target}");</script>'
        '</body></html>'
    )


def _inject_head(page: str, **kw) -> str:
    """Idempotently place the meta block just before </head>."""
    page = _strip_block(page, _META_MARK_S, _META_MARK_E)
    block = _head_block(**kw)
    if "</head>" in page:
        return page.replace("</head>", f"  {block}\n</head>", 1)
    return page


# ─── Post-build content fixups ─────────────────────────────────────────────────
_TERMS_LINK_RE = re.compile(
    r'<a\s+href="' + re.escape(PRIVACY_URL) + r'"([^>]*)>(\s*Algemene\s+voorwaarden\s*)</a>',
    re.IGNORECASE)
# "€49." -> "€ 49,00." (Dutch/Belgian uses a comma decimal separator).
# The lookahead skips amounts that already carry decimals — "€ 49,00" and
# "€ 49.50" are left alone — but a trailing sentence period is not decimals,
# so "€49. Levertijd" still matches.
_PRICE_RE = re.compile(r'€\s?(\d{1,4})(?![\d,]|\.\d)')

_SVG_BLOCK_RE = re.compile(r'<svg\b.*?</svg>', re.DOTALL | re.IGNORECASE)
_ID_ATTR_RE = re.compile(r'\s+id="([^"]+)"')
_ID_REF_RE = re.compile(
    r'url\(#([^)\s]+)\)|(?:xlink:)?href="#([^"]+)"'
    r'|aria-(?:labelledby|describedby)="([^"]+)"|for="([^"]+)"')


def _dedupe_svg_ids(content: str):
    """Strip duplicate ids from inline SVGs.

    The logo SVG is emitted twice per page (header + footer), duplicating
    yoggies_logo_on_light_bg_RGB and Path_7..Path_15. The ids are Illustrator
    export artifacts that nothing references, so the later copies just lose
    them; anything actually referenced is left alone.
    """
    referenced = set()
    for m in _ID_REF_RE.finditer(content):
        for g in m.groups():
            if g:
                referenced.update(g.split())

    # ids living outside any <svg> still claim the namespace
    seen = set(_ID_ATTR_RE.findall(_SVG_BLOCK_RE.sub(" ", content)))
    removed = [0]

    def fix_block(bm):
        def fix_id(im):
            name = im.group(1)
            if name in referenced:
                return im.group(0)
            if name in seen:
                removed[0] += 1
                return ""
            seen.add(name)
            return im.group(0)
        return _ID_ATTR_RE.sub(fix_id, bm.group(0))

    return _SVG_BLOCK_RE.sub(fix_block, content), removed[0]


def _content_fixups(content: str):
    """Cross-cutting text fixes applied to every generated page."""
    n = 0
    content, k = _TERMS_LINK_RE.subn(
        lambda m: f'<a href="{TERMS_URL}"{m.group(1)}>{m.group(2)}</a>', content)
    n += k
    content, k = _PRICE_RE.subn(lambda m: f'€ {m.group(1)},00', content)
    n += k
    content, k = _dedupe_svg_ids(content)
    n += k
    return content, n


_IMG_TAG_RE = re.compile(r'<img\b[^>]*>', re.IGNORECASE)
_IMG_SRC_RE = re.compile(r'(\bsrc=")(https?://[^"]+)(")', re.IGNORECASE)
_HAS_ALT_RE = re.compile(r'\balt\s*=', re.IGNORECASE)
_HAS_LOAD_RE = re.compile(r'\bloading\s*=', re.IGNORECASE)
_HAS_DEC_RE = re.compile(r'\bdecoding\s*=', re.IGNORECASE)
_EAGER_RE = re.compile(r'\b(fetchpriority\s*=\s*"high"|class="[^"]*\bhero\b)', re.IGNORECASE)


def _append_attr(tag: str, attr: str) -> str:
    """Add an attribute to an <img> tag, preserving XHTML self-closing form."""
    body = tag[:-1].rstrip()          # drop '>'
    if body.endswith("/"):
        return f"{body[:-1].rstrip()} {attr} />"
    return f"{body} {attr}>"


def _localize_img_tags(content: str, depth: str = "", unmapped=None):
    """Point leftover external <img> tags at site/images/ and tidy them up.

    Only rewrites tags whose URL is in image_map.json; anything unmapped is
    counted (so the build reports it) and left exactly as it was.
    """
    count = [0]

    def fix_tag(m):
        tag = m.group(0)
        src = _IMG_SRC_RE.search(tag)
        if src:
            url = src.group(2)
            local = img_src(url.replace("&amp;", "&"), "md", depth)
            if local != url and not local.startswith("http"):
                tag = tag[:src.start()] + f'src="{local}"' + tag[src.end():]
                count[0] += 1
            elif unmapped is not None:
                unmapped[url] += 1
        # Backfill accessibility/perf attributes the templates omit. These are
        # decorative template images, so an empty alt is the correct choice —
        # content images get a real alt at their call site.
        if not _HAS_ALT_RE.search(tag):
            tag = _append_attr(tag, 'alt=""')
        if not _HAS_DEC_RE.search(tag):
            tag = _append_attr(tag, 'decoding="async"')
        if not _HAS_LOAD_RE.search(tag) and not _EAGER_RE.search(tag):
            tag = _append_attr(tag, 'loading="lazy"')
        return tag

    return _IMG_TAG_RE.sub(fix_tag, content), count[0]


def _truncate_words(text: str, limit: int) -> str:
    """Cut to `limit` chars on a word boundary, appending an ellipsis.

    Plain slicing produced card descriptions ending mid-word ("milieu",
    "sint-janskrui"), so back off to the last space before the limit.
    """
    text = " ".join(str(text or "").split())
    if len(text) <= limit:
        return text
    cut = text[:limit]
    space = cut.rfind(" ")
    if space > limit * 0.5:      # only back off if it doesn't gut the string
        cut = cut[:space]
    return cut.rstrip(" ,.;:-") + "…"


def img_srcset(url, depth=""):
    """Responsive srcset+sizes for a hero image. Empty string if unmapped."""
    entry = _img_entry(url)
    if not entry or "md" not in entry or "lg" not in entry:
        return ""
    return (f' srcset="{depth}{entry["md"]} 600w, {depth}{entry["lg"]} 1200w"'
            f' sizes="(max-width: 700px) 100vw, 600px"')

COLS_TRANSLATE = ["description_nl", "dosage_cz", "ingredients_cz"]

# Products excluded from the Belgian catalogue. The Kč gift vouchers are priced
# in Czech crowns, are not redeemable at DogLife (match_status NOT_ON_DOGLIFE,
# pointing at an empty search), and cannot be converted to euro without a
# business decision on face value. Their URLs redirect to the category page
# rather than 404. Remove an entry here to bring a product back.
EXCLUDE_PRODUCTS = re.compile(r'\bK[čc]\b', re.IGNORECASE)
EXCLUDE_REDIRECT = "speelgoed-accessoires.html"

# API keys — used ONLY in --fill-cache mode
_API_OR = os.environ.get("OPENROUTER_API_KEY", "")
_API_AN = os.environ.get("ANTHROPIC_API_KEY", "")
_MODEL_OR = "deepseek/deepseek-v4-flash:free"
_MODEL_AN = "claude-haiku-4-5-20251001"

# ─── Normalization ─────────────────────────────────────────────────────────────
_RE_TAG = re.compile(r'\[(CZ|NL|EN|DE|FR|SK|PL)\]\s*', re.IGNORECASE)
_RE_NB  = re.compile(r'[\xa0\u202f\u2009\u200b\u00a0\t]+')
_RE_MD  = re.compile(r'^#+\s*', re.MULTILINE)
_RE_SP  = re.compile(r' {2,}')

def normalize(text: str) -> str:
    """
    Deterministic normalization.
    Identical output for any whitespace/tag variant of the same semantic text.
    """
    t = _RE_TAG.sub("", str(text))   # strip language tags
    t = _RE_NB.sub(" ", t)           # normalize unicode whitespace
    t = _RE_MD.sub("", t)            # strip markdown headers
    t = re.sub(r'\n{2,}', '\n', t)   # collapse multi-newlines
    t = re.sub(r' *\n *', '\n', t)   # strip spaces around newlines
    t = _RE_SP.sub(" ", t)           # collapse spaces
    return t.strip().lower()

def make_key(text: str) -> str:
    """SHA256 of normalized text. The ONLY valid cache key format."""
    return hashlib.sha256(normalize(text).encode("utf-8")).hexdigest()

def sanitize_output(text: str) -> str:
    """Remove tags/markdown from stored/returned translation."""
    t = _RE_TAG.sub("", str(text))
    t = _RE_MD.sub("", t)
    t = _RE_NB.sub(" ", t)
    t = _RE_SP.sub(" ", t)
    lines = [l.strip() for l in t.splitlines() if l.strip()]
    return "\n".join(lines).strip()

# ─── Czech detection ───────────────────────────────────────────────────────────
_CZ_CHARS = re.compile(r'[čďěňřšťůžČĎĚŇŘŠŤŮŽ]')


# ─── Size variant detection (e.g. "1 kg Yoggies BARF X" vs "5 kg Yoggies BARF X") ──
_SIZE_PATTERNS = [
    # At start: "1 kg Xxx" or "1kg Xxx" or "1,5 kg Xxx"
    re.compile(r'^(\d+(?:[.,]\d+)?)\s*(kg|g|ml|l)\s+', re.IGNORECASE),
    # At end: "Xxx 1 kg" or "Xxx 1kg"
    re.compile(r'\s+(\d+(?:[.,]\d+)?)\s*(kg|g|ml|l)\s*$', re.IGNORECASE),
]

def _extract_size(name):
    """Return (stripped_base_name, pretty_size_label) or (name, None)."""
    if not name: return name, None
    for pat in _SIZE_PATTERNS:
        m = pat.search(name)
        if m:
            qty, unit = m.group(1).strip(), m.group(2).lower()
            pretty = f"{qty} {unit}"
            stripped = (name[:m.start()] + " " + name[m.end():]).strip()
            stripped = re.sub(r'\s+', ' ', stripped)
            return stripped, pretty
    return name, None

def _size_to_grams(size_label):
    """Sortable numeric value for a size label like '1 kg' or '500 g'."""
    if not size_label: return 0
    m = re.match(r'(\d+(?:[.,]\d+)?)\s*(kg|g|ml|l)', str(size_label), re.IGNORECASE)
    if not m: return 0
    qty = float(m.group(1).replace(',', '.'))
    unit = m.group(2).lower()
    return int(qty * (1000 if unit in ("kg", "l") else 1))

def _merge_size_variants(out_rows):
    """Merge products that are size variants of the same base product.

    Returns (merged_rows, n_groups, retired) where `retired` lists the
    per-size products that no longer get their own page:
    [{"old_name", "base_name", "size"}]. The caller turns those into redirect
    stubs so the old URLs keep working and stay reachable from search.
    """
    by_base = defaultdict(list)
    for p in out_rows:
        name = str(p.get("name_nl") or "").strip()
        if not name: continue
        base, size = _extract_size(name)
        if size:
            by_base[base.lower()].append((p, name, size, base))

    real_groups = {k: v for k, v in by_base.items() if len(v) >= 2}
    if not real_groups:
        return out_rows, 0, []

    seen = set()
    merged = []
    for p in out_rows:
        name = str(p.get("name_nl") or "").strip()
        if not name:
            merged.append(p)
            continue
        base, size = _extract_size(name)
        key = base.lower() if size else None
        if key and key in real_groups:
            if key in seen:
                continue
            seen.add(key)
            sorted_grp = sorted(real_groups[key], key=lambda x: _size_to_grams(x[2]))
            canonical = dict(sorted_grp[0][0])
            canonical["name_nl"] = sorted_grp[0][3]   # base name without size
            canonical["_variants"] = [
                {"size": s, "dl_url": str(gp.get("doglife_url") or "").strip()}
                for gp, _, s, _ in sorted_grp
            ]
            # Use the smallest variant's dl_url as the page's default doglife_url
            canonical["doglife_url"] = canonical["_variants"][0]["dl_url"]
            merged.append(canonical)
        else:
            merged.append(p)

    # Every per-size name in a merged group loses its own page — including the
    # smallest, whose page is now served under the base name.
    retired = [
        {"old_name": old_name, "base_name": base, "size": size}
        for grp in real_groups.values()
        for _gp, old_name, size, base in grp
    ]
    return merged, len(real_groups), retired


def _is_real_product_url(url):
    """A URL counts as a real DogLife product page if it's an http(s) link that
    isn't a search-results page (WordPress ?s=, search=) and isn't empty."""
    if not url:
        return False
    u = str(url).strip().lower()
    if not (u.startswith("http://") or u.startswith("https://")):
        return False
    # WordPress search URLs like https://doglife.be/?s=yoggies+cbd are NOT product pages
    if "?s=" in u or "&s=" in u or "/?search=" in u or "&search=" in u:
        return False
    return True


def _norm_product_name(s):
    """Normalize a product name for lookup: lowercase + strip accents + collapse spaces.
    Lets manual_variants.py keys match Excel rows even if exotic diacritics differ
    ('é' vs 'Ú', 'ï' vs '´', etc.)."""
    import unicodedata
    if not s:
        return ""
    # Step 1: replace known mojibake / corrupted characters explicitly
    s = s.replace('\u00b4', 'i').replace('\u00da', 'e')   # ´ → i, Ú → e
    # Step 2: NFKD decomposition + drop combining marks (handles é, ï, etc.)
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    # Step 3: lowercase, collapse whitespace
    s = re.sub(r'\s+', ' ', s.lower().strip())
    return s


def _apply_manual_variants(out_rows):
    """Overlay variants from manual_variants.py onto products by normalized name match.
    Returns count of products that got manual variants applied."""
    if not _MANUAL_VARIANTS:
        return 0
    # Build lookup with normalized keys
    lookup = {_norm_product_name(k): v for k, v in _MANUAL_VARIANTS.items()}
    applied = 0
    seen_keys = set()
    for p in out_rows:
        name = str(p.get("name_nl") or "").strip()
        if not name:
            continue
        norm = _norm_product_name(name)
        if norm in lookup and norm not in seen_keys:
            variants = lookup[norm]
            if isinstance(variants, list) and len(variants) >= 2:
                cleaned = []
                for v in variants:
                    if isinstance(v, dict) and v.get("size"):
                        cleaned.append({
                            "size":   str(v.get("size", "")).strip(),
                            "dl_url": str(v.get("dl_url", "")).strip(),
                        })
                if len(cleaned) >= 2:
                    p["_variants"] = cleaned
                    p["doglife_url"] = cleaned[0]["dl_url"]
                    applied += 1
                    seen_keys.add(norm)
    return applied

# Unambiguous Czech words — never appear in Dutch
_CZ_WORDS = [
    "doporučen", "krmivo", "tělesn", "hmotnost",
    "skladujte", "vašeho psa", "krmnou dávk", "krmné dávk",
    "zamrazujte", "spotřebujte", "sušené mleté",
    "vnitřnosti", "svalovinu", "jsou gran",
    "nenechávejte psa", "hračkou bez dozoru",
]

# Proper nouns / scientific names / brands that legitimately contain Czech chars
_PROPER_NOUNS = {
    "anička", "oškerová", "jeřábek", "pospíšilová", "kraußová", "tomáš",
    "lithothamnium", "cordyceps", "hericium", "coriolus", "maitake",
    "chaga", "reishi", "agaricus", "nosek",
    "kč", "mycotherapeut", "rugaas",
}

# ─── Dutch wording overrides ───────────────────────────────────────────────────
# Applied to every translated string. "granule" is the Czech word for kibble and
# survived into the Dutch copy in a handful of descriptions; the Dutch product
# line is called "brokken".
# No trailing \b: the word also shows up inside compounds
# ("granuleassortiment" -> "brokkenassortiment").
_NL_OVERRIDES = [
    (re.compile(r'\bgranule[ns]?'), 'brokken'),
    (re.compile(r'\bGranule[ns]?'), 'Brokken'),
    (re.compile(r'\bGRANULE[NS]?'), 'BROKKEN'),
]


def apply_nl_overrides(text: str) -> str:
    """Fix Czech loanwords that the translation cache left in the Dutch copy."""
    for pat, repl in _NL_OVERRIDES:
        text = pat.sub(repl, text)
    return text


# ─── Icon label translations ───────────────────────────────────────────────────
_ICON_TRANS = {
    'Bez chemie': 'Zonder chemie',
    'Bez\u2028chemie': 'Zonder chemie',
    'EKO\u2028přátelský': 'Milieuvriendelijk',
    'EKO přátelský': 'Milieuvriendelijk',
    'Pro správné zažívání': 'Voor gezonde spijsvertering',
    'Pro správné\u2028zažívání': 'Voor gezonde spijsvertering',
    'Pro všechna plemena i velikosti': 'Voor alle rassen en maten',
    'Pro všechna\u2028plemena i velikosti': 'Voor alle rassen en maten',
    'Vhodné pro alergiky': 'Geschikt voor allergische honden',
    'Vhodné pro\u2028alergiky': 'Geschikt voor allergische honden',
    'Farmářské suroviny': 'Boerderijgrondstoffen',
    'Farmářské\u2028suroviny': 'Boerderijgrondstoffen',
    'Ideální na trénink': 'Ideaal voor training',
    'Ideální na\u2028trénink': 'Ideaal voor training',
    'Organické produkty': 'Biologische producten',
    'Organické\u2028produkty': 'Biologische producten',
    'Pro čilý mozek': 'Voor een actieve geest',
    'Pro čilý\u2028mozek': 'Voor een actieve geest',
    'Pro hedvábnou srst': 'Voor een zijdezachte vacht',
    'Pro hedvábnou\u2028srst': 'Voor een zijdezachte vacht',
    'Pro zdravý mikrobiom': 'Voor een gezond microbioom',
    'Pro zdravý\u2028mikrobiom': 'Voor een gezond microbioom',
    'Pro zdravé zuby': 'Voor gezonde tanden',
    'S probiotiky': 'Met probiotica',
    'Lisované za studena': 'Koudgeperst',
}

def translate_icon(label: str) -> str:
    s = label.strip()
    return _ICON_TRANS.get(s, s)


def _is_proper_noun_context(text: str, pos: int) -> bool:
    """True if Czech char at pos is inside a proper noun."""
    ctx = text[max(0, pos-20):pos+20].lower()
    if any(p in ctx for p in _PROPER_NOUNS):
        return True
    # Capitalized word = likely proper noun
    word = re.search(r'\b[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ][a-záčďéěíňóřšťúůýž]{2,}\b',
                     text[max(0, pos-5):pos+25])
    return bool(word)

def has_contamination(text: str) -> bool:
    """
    Returns True if text has Czech contamination.
    False if only proper nouns / product names trigger Czech chars.
    """
    if not text:
        return False
    # Hard checks: tags and markdown always = contamination
    if _RE_TAG.search(text):
        return True
    if re.search(r'^#+\s+\S', text, re.MULTILINE):
        return True
    # Czech words (unambiguous)
    tl = text.lower()
    for w in _CZ_WORDS:
        if w in tl:
            return True
    # Czech chars — skip proper nouns
    for m in _CZ_CHARS.finditer(text):
        if not _is_proper_noun_context(text, m.start()):
            return True
    return False

def needs_translation(text: str) -> bool:
    """True if text contains Czech and needs cache lookup."""
    n = normalize(text)
    if not n or len(n) < 5:
        return False
    if _RE_TAG.search(text):
        return True
    nl = n.lower()
    for w in _CZ_WORDS:
        if w in nl:
            return True
    for m in _CZ_CHARS.finditer(n):
        if not _is_proper_noun_context(n, m.start()):
            return True
    return False

# ─── Cache ──────────────────────────────────────────────────────────────────────
_cache: dict = {}

def cache_load() -> int:
    global _cache
    p = Path(CACHE_FILE)
    if p.exists():
        _cache = json.loads(p.read_text(encoding="utf-8"))
    return len(_cache)

def cache_save():
    Path(CACHE_FILE).write_text(
        json.dumps(_cache, ensure_ascii=False, indent=2), encoding="utf-8"
    )

def cache_get(text: str):
    """
    THE ONLY translation path.
    Returns sanitized translation string, or None if not found.
    """
    k = make_key(text)
    if k not in _cache:
        return None
    return sanitize_output(_cache[k])

def cache_set(text: str, translation: str):
    k = make_key(text)
    _cache[k] = sanitize_output(translation)
    cache_save()

# ─── Translation function (cache-only) ────────────────────────────────────────
def translate(raw) -> tuple[str, str]:
    """
    Returns (result, status).
    status: 'cache' | 'passthrough' | 'missing'

    STRICT CONTRACT:
    - 'cache'       → translation found in cache
    - 'passthrough' → text needs no translation (URL, number, NL already)
    - 'missing'     → needs translation but NOT in cache → caller must FAIL
    """
    if not raw:
        return "", "passthrough"

    text = str(raw)

    # Passthrough: non-translatable content
    if re.match(r'^https?://', text.strip()):
        return text.strip(), "passthrough"
    if re.match(r'^\d+(\.\d+)?$', text.strip()):
        return text.strip(), "passthrough"

    # Does this text need translation?
    if not needs_translation(text):
        return apply_nl_overrides(sanitize_output(text)), "passthrough"

    # Cache lookup — the ONLY path
    result = cache_get(text)
    if result is not None:
        return apply_nl_overrides(result), "cache"

    # Missing: return original text flagged as missing
    # Caller decides to fail or log
    return text, "missing"

# ─── Data ──────────────────────────────────────────────────────────────────────
def load_xlsx(path: str):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    hdrs = [str(c.value or "").strip()
            for c in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = [dict(zip(hdrs, row))
            for row in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return hdrs, rows

# ─── MODE 2: fill cache ─────────────────────────────────────────────────────────
def _api_call(text: str) -> str:
    """Single API call. Raises on failure. Used ONLY by fill_cache_mode."""
    norm_text = normalize(text)
    sys_prompt = (
        "Vertaal van Tsjechisch naar professioneel Nederlands voor yoggies.be. "
        "Gebruik 'uw hond'. ALLEEN Nederlandse tekst. "
        "Geen [CZ]/[NL] tags. Geen Tsjechische woorden. Geen markdown (#)."
    )
    if _API_OR:
        payload = json.dumps({
            "model": _MODEL_OR,
            "messages": [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": norm_text}
            ]
        }).encode()
        req = urllib.request.Request(
            "https://openrouter.ai/api/v1/chat/completions",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {_API_OR}",
                "HTTP-Referer": "https://yoggies.be",
            }
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())["choices"][0]["message"]["content"].strip()

    if _API_AN:
        payload = json.dumps({
            "model": _MODEL_AN,
            "max_tokens": 1500,
            "system": sys_prompt,
            "messages": [{"role": "user", "content": norm_text}]
        }).encode()
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": _API_AN,
                "anthropic-version": "2023-06-01"
            }
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())["content"][0]["text"].strip()

    raise RuntimeError("No API key. Set OPENROUTER_API_KEY or ANTHROPIC_API_KEY")


def fill_cache_mode():
    print("MODE: FILL_CACHE")
    print(f"Cache: {len(_cache)} entries")

    _, rows = load_xlsx(DATA_FILE)

    # Collect unique missing texts
    todo: dict[str, str] = {}
    for col in COLS_TRANSLATE:
        for row in rows:
            raw = row.get(col)
            if raw and needs_translation(str(raw)):
                k = make_key(str(raw))
                if k not in _cache and k not in todo:
                    todo[k] = str(raw)

    print(f"Missing unique: {len(todo)}")
    if not todo:
        print("Cache complete.")
        return

    done = errors = 0
    for k, text in todo.items():
        try:
            result = sanitize_output(_api_call(text))
            # Retry once if contamination detected
            if has_contamination(result):
                time.sleep(1)
                result = sanitize_output(_api_call(text))
            _cache[k] = result
            cache_save()
            done += 1
            print(f"  [{done}/{len(todo)}] {normalize(text)[:55]}")
            time.sleep(0.3)
        except Exception as e:
            print(f"  ERR: {e}")
            errors += 1
            time.sleep(2)

    print(f"Done: {done} | Errors: {errors} | Cache: {len(_cache)}")

# ─── MODE 3: audit ─────────────────────────────────────────────────────────────
def audit_mode():
    print("MODE: AUDIT")
    _, rows = load_xlsx(DATA_FILE)
    missing = []
    for col in COLS_TRANSLATE:
        for i, row in enumerate(rows):
            raw = row.get(col)
            if raw and needs_translation(str(raw)):
                if cache_get(str(raw)) is None:
                    missing.append({
                        "row": i + 2, "col": col,
                        "name": str(row.get("name_nl") or "")[:40],
                        "text": str(raw)[:80]
                    })
    unique = len({make_key(str(m["text"])) for m in missing})
    print(f"Cache: {len(_cache)} | Missing cells: {len(missing)} ({unique} unique)")
    by_col = defaultdict(int)
    for m in missing:
        by_col[m["col"]] += 1
    for col, n in by_col.items():
        print(f"  {col}: {n}")
    if missing:
        for m in missing[:5]:
            print(f"  row {m['row']} {m['col']}: {m['text'][:60]}")

# ─── MODE 1: build ─────────────────────────────────────────────────────────────
def build_mode():
    print("MODE: CACHE_ONLY BUILD")
    hdrs, rows = load_xlsx(DATA_FILE)
    print(f"Rows: {len(rows)} | Cache: {len(_cache)}")

    # ── PASS 1: check for missing translations ──────────────────────────────────
    missing_cells = []
    for i, row in enumerate(rows):
        for col in COLS_TRANSLATE:
            raw = row.get(col)
            if raw and needs_translation(str(raw)):
                if cache_get(str(raw)) is None:
                    missing_cells.append({
                        "row": i + 2, "col": col,
                        "name": str(row.get("name_nl") or "")[:50],
                        "text": str(raw)[:120]
                    })

    # Always write missing report
    rpt_lines = [
        f"MISSING REPORT — {datetime.now().isoformat()}",
        f"Total missing cells: {len(missing_cells)}",
        "",
    ]
    if missing_cells:
        rpt_lines.append("Run:  python yoggies_pipeline.py --fill-cache")
        rpt_lines.append("=" * 60)
        for m in missing_cells:
            rpt_lines.append(f"Row {m['row']} | {m['col']} | {m['name']}")
            rpt_lines.append(f"  {m['text']}")
    else:
        rpt_lines.append("All texts found in cache. Build successful.")

    Path(MISSING_RPT).write_text("\n".join(rpt_lines), encoding="utf-8")

    if missing_cells:
        print(f"FAIL: {len(missing_cells)} missing translations.")
        print(f"Run:  python yoggies_pipeline.py --fill-cache")
        print(f"Report: {MISSING_RPT}")
        sys.exit(1)

    # ── PASS 2: translate all cells ─────────────────────────────────────────────
    import openpyxl
    wb_in  = openpyxl.load_workbook(DATA_FILE)
    ws_in  = wb_in.active
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active

    col_idxs = {col: hdrs.index(col) + 1
                for col in COLS_TRANSLATE if col in hdrs}
    stats = defaultdict(int)
    validation_fails = []

    # Copy header
    for c in range(1, ws_in.max_column + 1):
        ws_out.cell(1, c).value = ws_in.cell(1, c).value

    for r in range(2, ws_in.max_row + 1):
        for c in range(1, ws_in.max_column + 1):
            raw      = ws_in.cell(r, c).value
            col_name = hdrs[c - 1]
            if col_name in col_idxs and raw:
                result, status = translate(raw)
                ws_out.cell(r, c).value = result
                stats[status] += 1
                if status == "missing":
                    # Should not happen after pass 1 — defensive
                    validation_fails.append(
                        f"Row {r} | {col_name} | MISSING (defensive)"
                    )
                elif status == "cache" and has_contamination(result):
                    validation_fails.append(
                        f"Row {r} | {col_name} | CONTAMINATION: {result[:60]}"
                    )
            else:
                ws_out.cell(r, c).value = raw

    # ── PASS 3: HARD FAIL if any validation error ───────────────────────────────
    if validation_fails:
        print(f"HARD FAIL: {len(validation_fails)} validation errors")
        for e in validation_fails:
            print(f"  {e}")
        print("Fix cache entries and re-run.")
        sys.exit(1)

    wb_out.save(OUTPUT_XLSX)
    print(f"Saved: {OUTPUT_XLSX}")
    print(f"Stats: cache={stats['cache']} passthrough={stats['passthrough']} missing={stats['missing']}")
    print("Validation: PASS — zero contamination")

    # ── PASS 4: build HTML ──────────────────────────────────────────────────────
    _build_html(hdrs)
    print("Build complete.")

# ─── HTML builder ──────────────────────────────────────────────────────────────
def _slugify(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^\w\s-]", "", s.lower())
    s = re.sub(r"[\s_]+", "-", s).strip("-")
    return s[:80] or "product"

# ─── Pet-size filter pages ─────────────────────────────────────────────────────
_PET_FILTERS = [
    ("grote-honden",       "Grote honden",        "grote-honden.html"),
    ("middelgrote-honden", "Middelgrote honden",  "middelgrote-honden.html"),
    ("kleine-honden",      "Kleine honden",       "kleine-honden.html"),
    ("pups",               "Pups",                "pups.html"),
    # "katten" hergebruikt voor-katten.html (geen aparte filterpagina)
]

def _get_pet_tags(p):
    """
    Determine which pet-filter tags apply to this product.
    Returns a set of tag keys from _PET_FILTERS.
    Heuristic — analyzes name + description for breed-size hints.
    """
    name_nl = str(p.get("name_nl") or "").lower()
    name_cz = str(p.get("name_cz") or "").lower()
    desc_cz = str(p.get("description_cz") or "").lower()
    desc_nl = str(p.get("description_nl") or "").lower()
    cat     = str(p.get("category") or "").lower()

    # Cat products: belong to voor-katten.html only
    if 'katt' in name_nl or cat == 'voor_katten':
        return set()

    # Toys, snacks, supplements: usable by all dog sizes
    if cat in ('speelgoed', 'snacks', 'supplementen'):
        return {'grote-honden', 'middelgrote-honden', 'kleine-honden', 'pups'}

    # Default for dog food: medium + large dogs
    tags = {'grote-honden', 'middelgrote-honden'}

    # Small breeds — name contains "mini" or description mentions
    if ('mini' in name_nl or 'mini' in name_cz or
            'malá plemena' in desc_cz or 'malých plemen' in desc_cz or
            'kleine rassen' in desc_nl or 'small breed' in desc_nl):
        tags.add('kleine-honden')

    # Puppies
    if ('štěňata' in desc_cz or 'štěnata' in desc_cz or 'štěně' in desc_cz or
            'pups' in desc_nl or 'puppy' in desc_nl or 'puppies' in desc_nl or
            'kojící' in desc_cz):
        tags.add('pups')

    return tags

# ─── Hondenatlas breeds (Dutch name, image, source detail URL) ─────────────────
_ATLAS_BREEDS = [
    ('Akita Inu', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Akita_inu.webp?v=1740737433&width=767', 'https://yoggies.cz/pages/atlas-psu/akita-inu'),
    ('Australische herdershond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Australsky_ovcak.webp?v=1740737436&width=767', 'https://yoggies.cz/pages/atlas-psu/australsky-ovcak'),
    ('Belgische herdershond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Belgicky_ovcak.webp?v=1740737431&width=767', 'https://yoggies.cz/pages/atlas-psu/belgicky-ovcak'),
    ('Berner sennenhond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Bernsky_salasnicky.webp?v=1740737439&width=767', 'https://yoggies.cz/pages/atlas-psu/bernsky-salasnicky-pes'),
    ('Border collie', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Border_Kolie.webp?v=1740737459&width=767', 'https://yoggies.cz/pages/atlas-psu/border-kolie'),
    ('Chihuahua', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/civava.webp?v=1740737430&width=767', 'https://yoggies.cz/pages/atlas-psu/civava'),
    ('Dobermann', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/dobrman.webp?v=1740737438&width=767', 'https://yoggies.cz/pages/atlas-psu/dobrman'),
    ('Franse bulldog', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/francouzsky_buldocek.webp?v=1740737426&width=767', 'https://yoggies.cz/pages/atlas-psu/francouzsky-buldocek'),
    ('Havanezer', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Havansky_psik.jpg?v=1740737446&width=767', 'https://yoggies.cz/pages/atlas-psu/havansky-psik'),
    ('Hovawart', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Hovawart.jpg?v=1740737441&width=767', 'https://yoggies.cz/pages/atlas-psu/hovawart'),
    ('Boheemse herdershond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/chodsky_pes.webp?v=1740737431&width=767', 'https://yoggies.cz/pages/atlas-psu/chodsky-pes'),
    ('Teckel', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Jezevcik.webp?v=1740737435&width=767', 'https://yoggies.cz/pages/atlas-psu/jazevcik'),
    ('Kaukasische herdershond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Kavkazsky_pastevecky_pes.webp?v=1740737459&width=767', 'https://yoggies.cz/pages/atlas-psu/kavkazsky-pastevecky-pes'),
    ('Labrador retriever', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/labrador.webp?v=1740737426&width=767', 'https://yoggies.cz/pages/atlas-psu/labradorsky-retrivr'),
    ('Mopshond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/mops.jpg?v=1740737445&width=767', 'https://yoggies.cz/pages/atlas-psu/mops'),
    ('Duitse bokser', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/nemecky_boxer.jpg?v=1740737449&width=767', 'https://yoggies.cz/pages/atlas-psu/nemecky-boxer'),
    ('Duitse herdershond', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/nemecky_ovcak.webp?v=1740737432&width=767', 'https://yoggies.cz/pages/atlas-psu/nemecky-ovcak'),
    ('Pitbull', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/pitbull.webp?v=1740737457&width=767', 'https://yoggies.cz/pages/atlas-psu/pitbul'),
    ('Pomeriaan', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Pomerian.webp?v=1740737431&width=767', 'https://yoggies.cz/pages/atlas-psu/pomeranian'),
    ('Rottweiler', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/rotvajler.webp?v=1740737433&width=767', 'https://yoggies.cz/pages/atlas-psu/rotvajler'),
    ('Samojeed', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/samojed.webp?v=1740737447&width=767', 'https://yoggies.cz/pages/atlas-psu/samojed'),
    ('Shiba Inu', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Shiba_inu.webp?v=1740736933&width=767', 'https://yoggies.cz/pages/atlas-psu/shiba-inu'),
    ('Staffordshire bull terriër', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Stafordsirsky_bulterier.webp?v=1740737440&width=767', 'https://yoggies.cz/pages/atlas-psu/stafordsirsky-bulterier'),
    ('Tibetaanse mastiff', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Tibetska_doga.webp?v=1740737434&width=767', 'https://yoggies.cz/pages/atlas-psu/tibetska-doga'),
    ('Welsh Corgi Pembroke', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/welsh_corgi_pembroke.webp?v=1740737429&width=767', 'https://yoggies.cz/pages/atlas-psu/welsh-corgi-pembroke'),
    ('Yorkshire terriër', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Yorksirsky_terier.webp?v=1740737430&width=767', 'https://yoggies.cz/pages/atlas-psu/yorkshirsky-terier'),
    ('Golden retriever', 'https://cdn.shopify.com/s/files/1/0908/0894/2932/files/Zlaty_retrivr_zm.jpg?v=1740737435&width=767', 'https://yoggies.cz/pages/atlas-psu/zlaty-retrivr'),
]

def _atlas_slug(href):
    """Extract breed slug from a yoggies.cz atlas href (last path segment)."""
    return href.rstrip('/').split('/')[-1]

def _atlas_img_by_slug():
    """Map slug -> image url from _ATLAS_BREEDS."""
    return {_atlas_slug(href): img for (name, img, href) in _ATLAS_BREEDS}


def _atlas_nl_slug_by_cz():
    """Map the Czech atlas slug -> the Dutch slug the page is published under."""
    return {b["slug"]: (b.get("nl_slug") or _slugify(b["nl_name"]) or b["slug"])
            for b in _BREEDS_NL if b.get("slug")}

def _build_atlas_pages(per_page=12):
    """Return HTML for the dog-grid atlas-page divs (paginated breed cards).
    Cards link to local breed pages (ras-<slug>.html)."""
    pages_html = ""
    total = len(_ATLAS_BREEDS)
    n_pages = (total + per_page - 1) // per_page
    nl_by_cz = _atlas_nl_slug_by_cz()
    for pi in range(n_pages):
        chunk = _ATLAS_BREEDS[pi * per_page:(pi + 1) * per_page]
        hidden = "" if pi == 0 else ' style="display:none;"'
        cards = ""
        for name, img, href in chunk:
            n = html.escape(name)
            slug = _atlas_slug(href)
            # Link straight to the Dutch URL, not through the redirect.
            local = f"ras-{nl_by_cz.get(slug, slug)}.html"
            cards += f'''        <div class="dog-card">
          <a href="{local}" style="display:block;text-decoration:none;color:inherit;">
            <div class="dog-card-img"><img src="{html.escape(img_src(img, "thumb"))}" alt="{n}" loading="lazy" decoding="async"></div>
            <div class="dog-card-label"><span class="card-name">{n}</span><span class="card-sub">Details</span></div>
          </a>
        </div>\n'''
        pages_html += f'''      <div class="dog-grid atlas-page" id="page-{pi+1}"{hidden}>\n{cards}      </div>\n\n'''
    return pages_html, n_pages

# ── Dutch breed detail pages (Hondenatlas) ──────────────────────────────────────
_BREED_HEADINGS = [
    ("intro",   "Oorsprong en geschiedenis"),
    ("povaha",  "Karakter"),
    ("hodi",    "Voor wie is dit ras geschikt?"),
    ("zaklad",  "Algemene informatie"),
    ("vychova", "Opvoeding en training"),
    ("nemoci",  "Veelvoorkomende aandoeningen"),
    ("krmit",   "Voeding"),
]

def _paras(text):
    """Split a section into <p> paragraphs (on newlines)."""
    out = ""
    for line in str(text).split("\n"):
        line = line.strip()
        if line:
            out += f'        <p style="margin:0 0 16px;font-size:16.5px;line-height:1.8;color:#2a1a0c;">{html.escape(line)}</p>\n'
    return out

def _build_breed_body(breed, img_url):
    """Return the <main>...</main> HTML for one breed detail page."""
    name = html.escape(breed["nl_name"])
    size = breed.get("size", {}) or {}
    secs = breed.get("sections", {}) or {}

    # size stat cards
    stat_items = [("Hoogte", size.get("hoogte")), ("Gewicht", size.get("gewicht")), ("Leeftijd", size.get("levensduur"))]
    stats = ""
    for label, val in stat_items:
        if val:
            stats += f'''        <div style="flex:1 1 180px;min-width:160px;background:#fff;border-radius:16px;padding:20px 22px;box-shadow:0 4px 16px rgba(60,30,0,.06);">
          <div style="font-size:12px;letter-spacing:.6px;text-transform:uppercase;color:#7a3b1e;font-weight:800;margin-bottom:6px;">{label}</div>
          <div style="font-size:17px;color:#3c1e00;font-weight:700;line-height:1.4;">{html.escape(str(val))}</div>
        </div>\n'''
    stats_block = ""
    if stats:
        stats_block = f'''    <div style="max-width:1000px;margin:-40px auto 0;padding:0 24px;position:relative;z-index:3;">
      <div style="display:flex;gap:18px;flex-wrap:wrap;">
{stats}      </div>
    </div>
'''

    # sections
    sections_html = ""
    for key, heading in _BREED_HEADINGS:
        body = secs.get(key)
        if not body:
            continue
        sections_html += f'''      <section style="margin:0 0 36px;">
        <h2 style="font-family:'Playfair Display','Nunito',serif;font-size:26px;color:#3c1e00;margin:0 0 14px;font-weight:700;border-left:5px solid #A1E647;padding-left:14px;">{html.escape(heading)}</h2>
{_paras(body)}      </section>
'''

    hero_img = ""
    if img_url:
        hero_img = f'''      <img src="{html.escape(img_src(img_url, "md"))}"{img_srcset(img_url)} alt="{name}"
           decoding="async" fetchpriority="high"
           style="width:100%;max-width:560px;border-radius:22px;display:block;margin:26px auto 0;box-shadow:0 12px 40px rgba(60,30,0,.18);"
           onerror="this.style.display='none'">'''

    return f'''<main style="background:#f5f1e6;padding-bottom:64px;">

  <!-- Breed hero -->
  <section style="background:linear-gradient(160deg,#3c1e00 0%,#5a3010 100%);padding:48px 24px 72px;text-align:center;">
    <div style="max-width:900px;margin:0 auto;">
      <p style="margin:0 0 10px;"><a href="hondenatlas.html" style="color:#A1E647;text-decoration:none;font-weight:700;font-size:14px;">&larr; Terug naar de Hondenatlas</a></p>
      <h1 style="font-family:'Playfair Display','Nunito',serif;font-size:40px;color:#fff;margin:0;font-weight:800;letter-spacing:.3px;">{name}</h1>
{hero_img}
    </div>
  </section>

{stats_block}
  <!-- Breed content -->
  <article style="max-width:1000px;margin:48px auto 0;padding:0 24px;">
{sections_html}
    <!-- CTA -->
    <div style="background:#fff;border-radius:22px;padding:40px 28px;text-align:center;box-shadow:0 8px 30px rgba(60,30,0,.08);margin-top:8px;">
      <h3 style="font-family:'Playfair Display','Nunito',serif;font-size:24px;color:#3c1e00;margin:0 0 14px;font-weight:700;">Op zoek naar gezonde voeding voor uw {name}?</h3>
      <p style="font-size:16px;color:#5a3010;line-height:1.7;margin:0 0 26px;">Ontdek het complete Yoggies-assortiment: natuurlijk, zonder chemische toevoegingen.</p>
      <a href="index.html" style="display:inline-block;background:#A1E647;color:#3c1e00;padding:15px 38px;border-radius:30px;text-decoration:none;font-weight:800;font-size:15px;letter-spacing:.3px;box-shadow:0 4px 14px rgba(60,30,0,.10);">Bekijk onze producten</a>
    </div>
  </article>

</main>
'''

def _write_breed_pages(home_content):
    """Generate ras-<slug>.html for every breed in _BREEDS_NL, using home nav + footer."""
    if not _BREEDS_NL:
        print("Hondenatlas: atlas_nl_data.py niet gevonden — rassenpagina's overgeslagen.")
        return 0
    img_map = _atlas_img_by_slug()
    count = 0
    redirects = 0
    for breed in _BREEDS_NL:
        slug = breed.get("slug")
        if not slug:
            continue
        # Published under the Dutch slug; the Czech one only survives as a
        # redirect and as the key for the atlas image lookup.
        nl_slug = breed.get("nl_slug") or _slugify(breed["nl_name"]) or slug
        breed_img = img_map.get(slug, "")
        body = _build_breed_body(breed, breed_img)
        nl_name = breed["nl_name"]
        secs = breed.get("sections", {}) or {}
        blurb = next((secs[k] for k, _ in _BREED_HEADINGS if secs.get(k)), "")

        page = home_content
        page = re.sub(r'<title>[^<]*</title>',
                      f'<title>{html.escape(nl_name)} | Hondenatlas | Yoggies.be</title>',
                      page, count=1)
        page = re.sub(r'(</nav>)(.*?)(<footer)',
                      lambda m: m.group(1) + body + '\n' + m.group(3),
                      page, count=1, flags=re.DOTALL)
        page = _inject_head(
            page,
            title=f"{nl_name} | Hondenatlas | Yoggies.be",
            description=_meta_desc(blurb) or _meta_desc(
                f"Alles over de {nl_name}: karakter, verzorging, veelvoorkomende "
                f"aandoeningen en de juiste voeding."),
            url_path=f"ras-{nl_slug}.html",
            image=img_src(breed_img, "lg") if breed_img else None,
            og_type="article",
        )
        (SITE_DIR / f"ras-{nl_slug}.html").write_text(page, encoding="utf-8")
        count += 1

        # Keep the old Czech URL alive.
        if slug != nl_slug:
            (SITE_DIR / f"ras-{slug}.html").write_text(
                _redirect_stub(f"ras-{nl_slug}.html",
                               f"{nl_name} | Hondenatlas | Yoggies.be",
                               f"ras-{nl_slug}.html",
                               note=f"Deze pagina staat nu op ras-{nl_slug}.html."),
                encoding="utf-8")
            redirects += 1

    print(f"Hondenatlas: {count} rassenpagina's gegenereerd (ras-<nl_slug>.html), "
          f"{redirects} Tsjechische URL's redirecten.")
    return count


def _build_html(hdrs, by_cat_outer=None, slugs_outer=None):
    import openpyxl
    wb = openpyxl.load_workbook(OUTPUT_XLSX, read_only=True)
    ws = wb.active
    out_hdrs = [str(ws.cell(1, c).value or "").strip()
                for c in range(1, ws.max_column + 1)]
    out_rows = [dict(zip(out_hdrs, row))
                for row in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()

    _copy_template_assets()

    tmpl = (TPL_DIR / "product.html").read_text(encoding="utf-8") \
        if (TPL_DIR / "product.html").exists() else _minimal_tmpl()

    # Inject calculator CSS + nav improvements (idempotent across rebuilds)
    if "calc-wrap" not in tmpl:
        tmpl = tmpl.replace("</style>", _CALC_CSS + "\n  </style>", 1)
    tmpl = _inject_nav_css(tmpl)
    tmpl = _inject_search_css(tmpl)
    tmpl = _inject_nav_js(tmpl)
    tmpl = _inject_search_js(tmpl)

    CAT = {
        "barf":         ("BARF",                   "barf.html"),
        "blikvoer":     ("Blikvoer",                "blikvoer.html"),
        "brokken":      ("Brokken met probiotica",  "brokken-met-probiotica.html"),
        "snacks":       ("Snacks",                  "snacks.html"),
        "gekookt":      ("Vers gekookt voer",        "vers-gekookt-voer.html"),
        "supplementen": ("Supplementen",            "supplementen.html"),
        "voor_katten":  ("Voor katten",             "voor-katten.html"),
        "yoghurt":      ("Yoghurt",                 "yoghurt.html"),
        "speelgoed":    ("Speelgoed & Accessoires", "speelgoed-accessoires.html"),
    }

    # ── Merge size variants (e.g. "1 kg Yoggies X" + "5 kg Yoggies X" → one page) ──
    # ── Drop products that don't belong in the Belgian catalogue ──
    excluded = [p for p in out_rows
                if EXCLUDE_PRODUCTS.search(str(p.get("name_nl") or ""))]
    if excluded:
        out_rows = [p for p in out_rows if p not in excluded]
        print(f"Excluded {len(excluded)} products priced in Czech crowns "
              f"(redirect → {EXCLUDE_REDIRECT}).")

    out_rows, n_var_groups, retired_variants = _merge_size_variants(out_rows)
    if n_var_groups:
        print(f"Size variants: merged {n_var_groups} groups, "
              f"{len(retired_variants)} per-size pages become redirects.")
    # ── Apply manual variant overrides (e.g. Brokken sizes from manual_variants.py) ──
    n_manual = _apply_manual_variants(out_rows)
    if n_manual:
        print(f"Manual variants: applied to {n_manual} products.")

    by_cat = defaultdict(list)
    slugs  = {}
    def _is_voor_katten(p) -> bool:
        """
        Deterministic rule: product belongs to voor_katten if
        'katt' appears in name_nl (case-insensitive).
        This covers purely cat products AND shared honden+katten products.
        Rule is data-driven — no hardcoded product names.
        """
        return 'katt' in str(p.get("name_nl") or "").lower()

    for p in out_rows:
        cat = str(p.get("category") or "speelgoed").strip()
        if cat not in CAT:
            cat = "speelgoed"
        by_cat[cat].append(p)
        # Virtual category: add to voor_katten if rule matches
        if _is_voor_katten(p):
            by_cat["voor_katten"].append(p)
        name = str(p.get("name_nl") or "product").strip()
        base = _slugify(name)
        slug = base
        count = sum(1 for v in slugs.values() if v.startswith(base))
        if count:
            slug = f"{base}-{count}"
        slugs[name] = slug

    PROD_DIR.mkdir(parents=True, exist_ok=True)
    written = set()
    for p in out_rows:
        name = str(p.get("name_nl") or "product").strip()
        slug = slugs.get(name, _slugify(name))
        (PROD_DIR / f"{slug}.html").write_text(
            _render_product(p, slug, tmpl, by_cat=by_cat, slugs=slugs), encoding="utf-8"
        )
        written.add(f"{slug}.html")
    print(f"{len(out_rows)} product pages → {PROD_DIR}/")

    # ── Retired size variants → redirect stubs, deep-linked to the parent ──
    # Merging left the old per-size pages orphaned on disk: still crawlable,
    # never regenerated, and absent from the search index. Replace each with a
    # stub that lands on the parent page with the right size preselected.
    variant_links = []
    for rv in retired_variants:
        old_slug = _slugify(rv["old_name"])
        parent_slug = slugs.get(rv["base_name"], _slugify(rv["base_name"]))
        if old_slug == parent_slug or f"{old_slug}.html" in written:
            continue
        target = f"{parent_slug}.html?maat={quote(rv['size'])}"
        (PROD_DIR / f"{old_slug}.html").write_text(
            _redirect_stub(target, f"{rv['old_name']} | Yoggies.be",
                           f"products/{parent_slug}.html",
                           note=f"{rv['old_name']} staat nu op de productpagina "
                                f"van {rv['base_name']}."),
            encoding="utf-8")
        written.add(f"{old_slug}.html")
        variant_links.append((rv["old_name"], f"{parent_slug}.html?maat={quote(rv['size'])}"))
    if variant_links:
        print(f"Variant redirects: {len(variant_links)} stubs → parent pages.")

    # ── Excluded products → redirect to their category ──
    for p in excluded:
        ex_slug = _slugify(str(p.get("name_nl") or "product").strip())
        if f"{ex_slug}.html" in written:
            continue
        (PROD_DIR / f"{ex_slug}.html").write_text(
            _redirect_stub(f"../{EXCLUDE_REDIRECT}",
                           f"{p.get('name_nl')} | Yoggies.be",
                           EXCLUDE_REDIRECT,
                           note="Dit product is niet beschikbaar in België."),
            encoding="utf-8")
        written.add(f"{ex_slug}.html")

    # ── Prune anything else left over from an older build ──
    stale = [f for f in PROD_DIR.glob("*.html") if f.name not in written]
    for f in stale:
        f.unlink()
    if stale:
        print(f"Pruned {len(stale)} stale product pages no longer generated.")

    # Write search index for the live site search dropdown
    import json as _json
    search_idx = []
    for p in out_rows:
        nm = str(p.get("name_nl") or "").strip()
        if not nm:
            continue
        sl = slugs.get(nm, _slugify(nm))
        cat_key = str(p.get("category") or "").strip()
        # Get pretty Dutch category label (e.g. "BARF" instead of "barf")
        cat_label = CAT.get(cat_key, (cat_key,))[0] if cat_key in CAT else cat_key
        search_idx.append({"name": nm, "slug": sl, "cat": cat_label})

    # Index the retired size variants too, pointing at the parent page with the
    # size preselected — searching "zalmolie 500 ml" must find something.
    by_name = {row["name"]: row for row in search_idx}
    for old_name, target in variant_links:
        parent = by_name.get(_extract_size(old_name)[0])
        search_idx.append({
            "name": old_name,
            "slug": parent["slug"] if parent else target.split(".html")[0],
            # "url" wins over "slug" in the dropdown: it carries the ?maat=
            # deep link, which slug + ".html" cannot express.
            "url": target,
            "cat": parent["cat"] if parent else "",
        })

    (SITE_DIR / "search-index.json").write_text(
        _json.dumps(search_idx, ensure_ascii=False, separators=(',', ':')),
        encoding="utf-8"
    )
    print(f"search-index.json: {len(search_idx)} entries "
          f"({len(out_rows)} products + {len(variant_links)} size variants).")

    # Category pages
    cat_tmpl_path = TPL_DIR / "hoofdpageproduct.html"
    if cat_tmpl_path.exists():
        ct = cat_tmpl_path.read_text(encoding="utf-8")
        # Inject nav CSS + JS (idempotent: refreshes on every rebuild)
        ct = _inject_nav_css(ct)
        ct = _inject_search_css(ct)
        ct = _inject_nav_js(ct)
        ct = _inject_search_js(ct)
        # Force logo to point to homepage (templates may have href="#")
        ct = _fix_logo_link(ct, "index.html")
        for cat_key, (cat_name, cat_file) in CAT.items():
            page = ct
            page = re.sub(r'<title>[^<]*</title>',
                          f'<title>{cat_name} | Yoggies.be</title>', page, count=1)
            page = re.sub(r'class="pg-title"[^>]*>[^<]*',
                          f'class="pg-title">{html.escape(cat_name)}', page, count=1)
            page = re.sub(r'class="pg-count"[^>]*>[^<]*',
                          f'class="pg-count">({len(by_cat[cat_key])} producten)', page, count=1)
            # yoghurt: single product → instant redirect page
            if cat_key == "yoghurt" and len(by_cat[cat_key]) == 1:
                p = by_cat[cat_key][0]
                nm = str(p.get("name_nl") or "product").strip()
                sl = slugs.get(nm, _slugify(nm))
                (SITE_DIR / cat_file).write_text(
                    _redirect_stub(
                        f"products/{sl}.html",
                        "Yoghurt | Yoggies.be",
                        f"products/{sl}.html",
                        note="Yoghurt heeft één product.",
                    ),
                    encoding="utf-8")
                continue

            n_prod = len(by_cat[cat_key])
            first_img = next((str(p.get("image_main") or p.get("doglife_image") or "").strip()
                              for p in by_cat[cat_key]
                              if p.get("image_main") or p.get("doglife_image")), "")
            page = _inject_head(
                page,
                title=f"{cat_name} | Yoggies.be",
                description=_meta_desc(
                    f"{cat_name} van Yoggies: {n_prod} producten voor je hond. "
                    f"Natuurlijke voeding met probiotica, te bestellen bij DogLife.be."),
                url_path=cat_file,
                image=img_src(first_img, "lg") if first_img else None,
            )

            # Inject nav
            nav_html = _build_nav(current_cat=cat_key, depth="", by_cat=by_cat, slugs=slugs)
            page = re.sub(r'<ul class="nav-list">.*?</ul>', nav_html, page, count=1, flags=re.DOTALL)
            cards = "".join(_cat_card(p, slugs) for p in by_cat[cat_key])
            page = re.sub(
                r'<div class="pg-grid">.*?</div>\s*(?=\s*<!--|</section|<footer)',
                f'<div class="pg-grid">\n{cards}</div>\n',
                page, count=1, flags=re.DOTALL)
            (SITE_DIR / cat_file).write_text(page, encoding="utf-8")
        print("Category pages written.")

        # === Filter pages (pet-size filters) ===
        by_filter = defaultdict(list)
        for p in out_rows:
            for tag in _get_pet_tags(p):
                by_filter[tag].append(p)

        for filter_key, filter_name, filter_file in _PET_FILTERS:
            page = ct
            page = re.sub(r'<title>[^<]*</title>',
                          f'<title>{filter_name} | Yoggies.be</title>', page, count=1)
            page = re.sub(r'class="pg-title"[^>]*>[^<]*',
                          f'class="pg-title">{html.escape(filter_name)}', page, count=1)
            page = re.sub(r'class="pg-count"[^>]*>[^<]*',
                          f'class="pg-count">({len(by_filter[filter_key])} producten)', page, count=1)
            page = _inject_head(
                page,
                title=f"{filter_name} | Yoggies.be",
                description=_meta_desc(
                    f"Yoggies voeding en snacks voor {filter_name.lower()}: "
                    f"{len(by_filter[filter_key])} producten, natuurlijk en met probiotica."),
                url_path=filter_file,
            )

            # Inject nav (no specific category active for filter pages)
            nav_html = _build_nav(current_cat="", depth="", by_cat=by_cat, slugs=slugs)
            page = re.sub(r'<ul class="nav-list">.*?</ul>', nav_html, page, count=1, flags=re.DOTALL)
            # Inject product cards
            cards = "".join(_cat_card(p, slugs) for p in by_filter[filter_key])
            page = re.sub(
                r'<div class="pg-grid">.*?</div>\s*(?=\s*<!--|</section|<footer)',
                f'<div class="pg-grid">\n{cards}</div>\n',
                page, count=1, flags=re.DOTALL)
            (SITE_DIR / filter_file).write_text(page, encoding="utf-8")
        print(f"Filter pages written: {len(_PET_FILTERS)} pages.")

    # Homepage — inject nav + CSS + JS
    home_path = TPL_DIR / "home.html"
    if home_path.exists():
        home_content = home_path.read_text(encoding="utf-8")
        # Inject nav CSS + JS (idempotent: refreshes on every rebuild)
        home_content = _inject_nav_css(home_content)
        home_content = _inject_search_css(home_content)
        home_content = _inject_nav_js(home_content)
        home_content = _inject_search_js(home_content)
        # Force logo to point to homepage (template had href="#")
        home_content = _fix_logo_link(home_content, "index.html")
        # Replace nav with data-driven version (depth="" since home is in site root)
        nav_html = _build_nav(current_cat="", depth="", by_cat=by_cat, slugs=slugs)
        import re as _re
        home_content = _re.sub(
            r'<ul class="nav-list"[^>]*>.*?</ul>',
            nav_html, home_content, count=1, flags=_re.DOTALL
        )

        # ── Update animal-icon labels (replace placeholder text like HIHIHIHII) ──
        animal_data = [
            ("Grote honden",        "grote-honden.html"),
            ("Middelgrote honden",  "middelgrote-honden.html"),
            ("Kleine honden",       "kleine-honden.html"),
            ("Pups",                "pups.html"),
            ("Katten",              "voor-katten.html"),
        ]
        _lbl_i = [0]
        def _replace_animal_label(m):
            i = _lbl_i[0]
            _lbl_i[0] += 1
            if i < len(animal_data):
                return f'<div class="animal-label">{animal_data[i][0]}</div>'
            return m.group(0)
        home_content = _re.sub(
            r'<div class="animal-label">[^<]+</div>',
            _replace_animal_label, home_content
        )

        # ── Update animal-link hrefs to point to filter pages ──
        _lnk_i = [0]
        def _replace_animal_link(m):
            i = _lnk_i[0]
            _lnk_i[0] += 1
            if i < len(animal_data):
                return f'<a class="animal-link" href="{animal_data[i][1]}">'
            return m.group(0)
        home_content = _re.sub(
            r'<a class="animal-link" href="[^"]*">',
            _replace_animal_link, home_content
        )

        # ── Update category buttons to match nav (9 buttons, same order) ──
        nav_cat_order = [
            ("BARF",                       "barf.html"),
            ("Blikvoer",                   "blikvoer.html"),
            ("Brokken met probiotica",     "brokken-met-probiotica.html"),
            ("Snacks",                     "snacks.html"),
            ("Vers gekookt voer",          "vers-gekookt-voer.html"),
            ("Supplementen",               "supplementen.html"),
            ("Voor katten",                "voor-katten.html"),
            ("Yoghurt",                    "yoghurt.html"),
            ("Speelgoed &amp; Accessoires", "speelgoed-accessoires.html"),
        ]
        new_buttons = "\n".join(
            f'    <a href="{file}" class="cat-btn">{label} <span class="arrow">&rsaquo;</span></a>'
            for label, file in nav_cat_order
        )
        home_content = _re.sub(
            r'<div class="categories">.*?</div>',
            f'<div class="categories">\n{new_buttons}\n  </div>',
            home_content, count=1, flags=_re.DOTALL
        )

        # ── Link "Bekijk advies" button to Dierenkliniek.html ──
        home_content = _re.sub(
            r'(<a\s+[^>]*?)href="[^"]*"([^>]*>\s*Bekijk\s+advies\s*</a>)',
            r'\1href="Dierenkliniek.html"\2',
            home_content, count=1
        )

        # ── Link "Over Yoggies" button to over-yoggies.html ──
        home_content = _re.sub(
            r'(<a\s+[^>]*?)href="[^"]*"([^>]*>\s*Over\s+Yoggies\s*</a>)',
            r'\1href="over-yoggies.html"\2',
            home_content, count=1
        )

        home_content = _inject_head(
            home_content,
            title="Yoggies.be — natuurlijke hondenvoeding met probiotica",
            description=_meta_desc(
                "Yoggies maakt natuurlijke hondenvoeding met probiotica: BARF, "
                "koudgeperste brokken, vers gekookt voer, snacks en supplementen. "
                "In België te bestellen bij DogLife.be."),
            url_path="",
            image="images/favicon.svg",
            og_type="website",
        )
        (SITE_DIR / "index.html").write_text(home_content, encoding="utf-8")
        # home.html used to be a byte-identical copy of index.html, which is
        # duplicate content. Emit a redirect stub instead so any old inbound
        # link still lands on the canonical homepage.
        (SITE_DIR / "home.html").write_text(
            _redirect_stub("index.html", "Yoggies.be", "index.html",
                           note="Deze pagina is verplaatst naar de homepage."),
            encoding="utf-8")

        # ── Generate over-yoggies.html (Dutch translation of yoggies.cz/o-nas) ──
        over_yoggies_body = '''
<main style="background:#fdfaf3;color:#3c1e00;font-family:'Nunito',sans-serif;">

  <!-- HERO -->
  <section style="text-align:center;padding:80px 24px 40px;max-width:900px;margin:0 auto;">
    <h1 style="font-family:'Playfair Display','Nunito',serif;font-size:48px;font-weight:700;color:#3c1e00;margin:0 0 16px;">Over Yoggies</h1>
    <p style="font-size:22px;color:#7a3b1e;font-style:italic;margin:0;">Gezondheid begint in de etensbak.</p>
  </section>

  <!-- STORY -->
  <section style="max-width:900px;margin:0 auto;padding:24px;line-height:1.75;font-size:17px;color:#1e1007;">
    <p>Het verhaal van Yoggies begon met één zieke pup.</p>
    <p>Onze hond Roxy, nog geen jaar oud, werd vergiftigd in het bos. Onze dierenarts redde zijn leven, maar daarna kreeg hij ernstige spijsverteringsproblemen — winderigheid, braken en diarree die maar niet wilden verdwijnen.</p>
    <p>We probeerden van alles. Tot we op een artikel stuitten over een Griekse immigrant die in de Verenigde Staten yoghurt was gaan maken. Op de foto bij het artikel gaf hij yoghurt aan zijn eigen herdershonden — en het werkte. Yoghurt hielp tegen spijsverteringsklachten en versterkte het immuunsysteem.</p>
    <p>Dus dachten we: als yoghurt mensen kan helpen, waarom dan niet ook Roxy?</p>
    <p>We ontwikkelden een speciale yoghurt voor honden. Voor Roxy bracht het verlichting. En het gaf ons meer dan alleen een antwoord — het gaf ons een missie.</p>
    <p>Voeding maken zonder chemische toevoegingen, die de gezondheid van viervoeters bij elke hap ondersteunt. Met levende probiotische culturen, kwaliteitsingrediënten en met liefde gemaakt.</p>
    <p>Vandaag is Yoggies een familiebedrijf met tientallen medewerkers en duizenden tevreden viervoeters. Maar het verhaal blijft hetzelfde — gezondheid begint in de etensbak.</p>
  </section>

  <!-- VIDEO -->
  <section style="max-width:900px;margin:0 auto;padding:40px 24px;">
    <div style="position:relative;padding-bottom:56.25%;height:0;overflow:hidden;border-radius:20px;box-shadow:0 8px 32px rgba(60,30,0,.15);background:#000;">
      <div id="yt-facade" onclick="ytPlay()" style="position:absolute;inset:0;top:0;left:0;width:100%;height:100%;cursor:pointer;">
        <img src="https://img.youtube.com/vi/_OzSLIDSnk0/maxresdefault.jpg"
             alt="Het verhaal van Yoggies"
             style="position:absolute;top:0;left:0;width:100%;height:100%;object-fit:cover;opacity:.88;"
             onerror="this.src='https://img.youtube.com/vi/_OzSLIDSnk0/hqdefault.jpg'">
        <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:84px;height:84px;background:rgba(161,230,71,.96);border-radius:50%;display:flex;align-items:center;justify-content:center;box-shadow:0 6px 24px rgba(0,0,0,.35);">
          <div style="width:0;height:0;border-left:26px solid #3c1e00;border-top:16px solid transparent;border-bottom:16px solid transparent;margin-left:7px;"></div>
        </div>
      </div>
    </div>
  </section>

  <script>
  function ytPlay(){
    var f=document.getElementById('yt-facade');
    if(location.protocol==='file:'){
      window.open('https://www.youtube.com/watch?v=_OzSLIDSnk0','_blank');
    }else{
      f.outerHTML='<iframe src="https://www.youtube.com/embed/_OzSLIDSnk0?autoplay=1" title="Het verhaal van Yoggies" frameborder="0" allow="autoplay; encrypted-media; picture-in-picture" allowfullscreen style="position:absolute;top:0;left:0;width:100%;height:100%;border:0;"></iframe>';
    }
  }
  </script>

  <!-- MISSION -->
  <section style="background:#f5f1e6;padding:64px 24px;margin-top:40px;">
    <div style="max-width:900px;margin:0 auto;text-align:center;">
      <h2 style="font-family:'Playfair Display','Nunito',serif;font-size:32px;color:#3c1e00;margin:0 0 24px;font-weight:700;">Onze missie</h2>
      <p style="font-size:18px;line-height:1.75;margin:0 0 18px;color:#1e1007;">
        Bij Yoggies geloven we dat een hond of kat geen gewoon dier is. Het is een maatje, onze liefde, een lid van het gezin.
        En gezondheid is de basis van geluk. Een gezonde viervoeter brengt vreugde in het hele huishouden.
      </p>
      <p style="font-size:18px;line-height:1.75;margin:0 0 32px;color:#1e1007;">
        Daarom kiezen wij voor natuurlijk voer: zonder chemische conserveringsmiddelen, zonder kunstmatige smaakstoffen,
        met ingrediënten van topkwaliteit. Hun neuzen zijn gevoeliger dan de onze — waarom zou je hen iets geven dat
        is volgestopt met chemie voor een "betere" geur of een "sterkere" smaak?
      </p>
      <a href="index.html"
         style="display:inline-block;background:#A1E647;color:#3c1e00;padding:16px 40px;border-radius:30px;text-decoration:none;font-weight:800;font-size:15px;letter-spacing:.3px;box-shadow:0 4px 14px rgba(60,30,0,.10);">
        Bekijk onze producten
      </a>
    </div>
  </section>

  <!-- THANKS TO VETS -->
  <section style="max-width:900px;margin:0 auto;padding:64px 24px;text-align:center;">
    <h3 style="font-family:'Playfair Display','Nunito',serif;font-size:22px;color:#7a3b1e;font-style:italic;margin:0 0 16px;font-weight:700;">Met dank aan</h3>
    <p style="font-size:16px;color:#5a3010;line-height:1.9;margin:0;">
      MVDr.&nbsp;Daniela Králová (Praag)<br>
      MVDr.&nbsp;Markéta Reinert (Kutná Hora)<br>
      MVDr.&nbsp;Jan Herčík (Praag)<br>
      <span style="font-size:14px;color:#7a3b1e;display:inline-block;margin-top:8px;">— die hielpen Roxy te redden.</span>
    </p>
  </section>

</main>
'''
        # Build over-yoggies.html from home_content (keeps nav + footer)
        over_content = home_content
        # Update <title>
        over_content = _re.sub(r'<title>[^<]*</title>',
                               '<title>Over Yoggies | Yoggies.be</title>',
                               over_content, count=1)
        # Replace everything between </nav> and <footer> with the over-yoggies body
        over_content = _re.sub(r'(</nav>)(.*?)(<footer)',
                               lambda m: m.group(1) + over_yoggies_body + '\n' + m.group(3),
                               over_content, count=1, flags=_re.DOTALL)
        over_content = _inject_head(
            over_content,
            title="Over Yoggies | Yoggies.be",
            description=_meta_desc(
                "Gezondheid begint in de etensbak. Lees waarom Yoggies natuurlijke "
                "hondenvoeding met probiotica maakt en hoe die wordt geproduceerd."),
            url_path="over-yoggies.html",
        )
        (SITE_DIR / "over-yoggies.html").write_text(over_content, encoding="utf-8")
        print("over-yoggies.html created.")

        # Build all Hondenatlas breed detail pages (ras-<slug>.html) from the home template
        _write_breed_pages(home_content)

    # ── calculator.html: hand-written page, never regenerated ──
    # 156 pages link to it, but no builder owned it, so it kept a nav frozen
    # from an old build and had no <head> metadata at all.
    calc_path = SITE_DIR / "calculator.html"
    if calc_path.exists():
        calc_content = calc_path.read_text(encoding="utf-8")
        calc_content = _inject_nav_css(calc_content)
        calc_content = _inject_search_css(calc_content)
        calc_content = _inject_nav_js(calc_content)
        calc_content = _inject_search_js(calc_content)
        calc_content = _fix_logo_link(calc_content, "index.html")
        calc_content = _refresh_nav(calc_content, by_cat=by_cat, slugs=slugs)
        calc_content = _inject_head(
            calc_content,
            title="Voedingscalculator | Yoggies.be",
            description=_meta_desc(
                "Bereken hoeveel Yoggies voer je hond per dag nodig heeft op basis "
                "van gewicht, leeftijd en activiteit."),
            url_path="calculator.html",
        )
        calc_path.write_text(calc_content, encoding="utf-8")
        print("calculator.html: nav + metadata refreshed.")

    # ── Update Dierenkliniek.html if exists — convert "Van hier om te berekenen" button to link → calculator.html ──
    dk_path = SITE_DIR / "Dierenkliniek.html"
    if dk_path.exists():
        dk_content = dk_path.read_text(encoding="utf-8")

        # 1. Replace entire advies-hero section with Voedingsadviseur Lucie content
        new_advies_hero = '''<section class="advies-hero">

  <!-- Full-width banner -->
  <div class="hero-full">
    <img src="https://image2url.com/r2/default/images/1772406984965-1480270f-3e78-4dd9-b21e-8e1727b0f530.jpg" alt="Hero afbeelding" />
  </div>

  <!-- Voedingsadviseur card: image left, info right (like yoggies.cz) -->
  <div class="hero-row">

    <div class="hero-dog">
      <img src="https://image2url.com/r2/default/images/1772407170221-140e4731-e4b3-4b2e-8486-ec9b6f94ea32.jpg" alt="Lucie - Voedingsadviseur" />
    </div>

    <div class="hero-text">
      <p class="hero-text__name" style="font-size:14px;color:#7a3b1e;font-weight:700;margin:0 0 6px 0;letter-spacing:.3px;">Lucie</p>
      <h2 class="hero-text__title">Voedingsadviseur</h2>
      <p class="hero-text__body">
        Wilt u dat uw pup of volwassen hond uitgroeit tot een sterke en gezonde metgezel?
        Of weet u niet zeker welke voeding het beste past? Onze voedingsadviseur helpt
        u graag bij het samenstellen van een dieet en bij vragen over winderigheid,
        braken of diarree.
      </p>
      <p class="hero-text__contact" style="margin-top:18px;font-size:15px;line-height:1.75;">
        Schrijf naar e-mail: <a href="mailto:Lucie@doglife.be" style="color:#3c1e00;text-decoration:underline;font-weight:700;">Lucie@doglife.be</a>
        of bel op werkdagen <a href="tel:+32472403406" style="color:#3c1e00;text-decoration:underline;font-weight:700;">+32 472 40 34 06</a>.
        Lucie helpt u graag.
      </p>
    </div>

  </div>
</section>'''

        dk_content_new = re.sub(
            r'<section class="advies-hero">.*?</section>',
            lambda m: new_advies_hero,
            dk_content,
            count=1,
            flags=re.DOTALL
        )
        if dk_content_new != dk_content:
            dk_content = dk_content_new
            print("Dierenkliniek.html: advies-hero section updated (Lucie + contact).")

        # 2. Convert "Van hier om te berekenen" button to link → calculator.html
        new_dk = re.sub(
            r'<button(\s+[^>]*?)>(\s*Van\s+hier\s+om\s+te\s+berekenen\s*)</button>',
            r'<a href="calculator.html"\1 style="display:inline-block;text-decoration:none">\2</a>',
            dk_content, count=1
        )
        if new_dk != dk_content:
            dk_content = new_dk
            print("Dierenkliniek.html: 'Van hier om te berekenen' link updated.")

        # 3. Insert or update Hondenatlas section (always reflect current content)
        new_atlas_section = '''<section class="advies-atlas" style="padding:64px 24px;background:#f5f1e6;">
  <div style="max-width:1100px;margin:0 auto;display:flex;align-items:center;gap:48px;flex-wrap:wrap;">
    <div style="flex:1 1 320px;min-width:280px;">
      <h2 style="font-family:'Playfair Display','Nunito',serif;font-size:32px;color:#3c1e00;margin:0 0 18px 0;font-weight:700;">Hondenatlas</h2>
      <p style="font-size:16px;line-height:1.7;color:#1e1007;margin:0 0 28px 0;">
        Ontdek in onze hondenatlas alles over het karakter, de opvoeding en de voeding van
        verschillende hondenrassen. We helpen u de juiste metgezel te vinden voor jullie
        gezamenlijke avonturen en geven advies over de beste verzorging.
      </p>
      <a href="hondenatlas.html"
         style="display:inline-block;background:#A1E647;color:#3c1e00;padding:14px 34px;border-radius:30px;text-decoration:none;font-weight:800;font-size:14px;letter-spacing:.3px;box-shadow:0 4px 14px rgba(60,30,0,.10);transition:background .18s, transform .15s;">
        Ontdek
      </a>
    </div>
    <div style="flex:1 1 320px;min-width:280px;">
      <img src="https://yoggies.cz/cdn/shop/files/Zlaty_retrivr_detail.webp?v=1740737440&amp;width=536"
           alt="Hondenatlas"
           style="width:100%;max-width:520px;border-radius:20px;display:block;margin:0 auto;"
           onerror="this.style.display='none'">
    </div>
  </div>
</section>
'''
        if 'advies-atlas' in dk_content:
            # Replace existing section (keeps content in sync with pipeline)
            dk_content_new = re.sub(
                r'<section class="advies-atlas".*?</section>',
                lambda m: new_atlas_section.rstrip(),
                dk_content, count=1, flags=re.DOTALL
            )
            if dk_content_new != dk_content:
                dk_content = dk_content_new
                print("Dierenkliniek.html: Hondenatlas section refreshed.")
        else:
            # Insert before advies-form
            dk_content_new = re.sub(
                r'(<section class="advies-form">)',
                lambda m: new_atlas_section + '\n' + m.group(1),
                dk_content, count=1
            )
            if dk_content_new != dk_content:
                dk_content = dk_content_new
                print("Dierenkliniek.html: Hondenatlas section added.")

        dk_content = _refresh_nav(dk_content, by_cat=by_cat, slugs=slugs)
        dk_content = _inject_head(
            dk_content,
            title="Dierenkliniek & voedingsadvies | Yoggies.be",
            description=_meta_desc(
                "Voedingsadvies voor je hond: welk Yoggies voer past bij welke "
                "leeftijd, maat en gezondheidsklacht. Stel je vraag aan onze adviseur."),
            url_path="Dierenkliniek.html",
        )
        dk_path.write_text(dk_content, encoding="utf-8")

    # ── Update hondenatlas.html if exists — fill breed grid with real breeds ──
    ha_path = SITE_DIR / "hondenatlas.html"
    if ha_path.exists():
        ha_content = ha_path.read_text(encoding="utf-8")
        atlas_pages_html, n_pages = _build_atlas_pages(per_page=12)

        # Replace placeholder atlas-page grids (everything from page-1 up to the pagination block)
        # [ \t]* so the replacement's own indentation replaces the existing
        # one instead of stacking on top of it every rebuild.
        ha_new = re.sub(
            r'[ \t]*<div class="dog-grid atlas-page" id="page-1">.*?(?=<!-- ===== PAGINATION CONTROLS)',
            lambda m: atlas_pages_html,
            ha_content, count=1, flags=re.DOTALL
        )
        # Keep JS totalPages in sync with the number of generated pages
        ha_new = re.sub(r'var totalPages = \d+;',
                        f'var totalPages = {n_pages};', ha_new, count=1)

        ha_new = _refresh_nav(ha_new, by_cat=by_cat, slugs=slugs)
        ha_new = _inject_head(
            ha_new,
            title="Hondenatlas | Yoggies.be",
            description=_meta_desc(
                f"Hondenatlas met {len(_ATLAS_BREEDS)} rassen: karakter, verzorging, "
                f"veelvoorkomende aandoeningen en het juiste voer per ras."),
            url_path="hondenatlas.html",
        )
        if ha_new != ha_content:
            ha_path.write_text(ha_new, encoding="utf-8")
            print(f"hondenatlas.html: breed grid updated ({len(_ATLAS_BREEDS)} breeds, {n_pages} pages).")

    # ── Post-build sweep: HTTPS hygiene + image localization ──
    # The builders above route their own <img> tags through img_src(), but the
    # static templates (home.html, hondenatlas.html, Dierenkliniek.html) carry
    # markup we never touch. This sweep catches whatever is left so no page
    # ships a hotlink, and backfills alt/loading on those template images.
    fixed_files = 0
    fixed_refs  = 0
    local_files = 0
    local_refs  = 0
    fixups      = 0
    unmapped    = defaultdict(int)

    for html_file in SITE_DIR.rglob("*.html"):
        try:
            t = html_file.read_text(encoding="utf-8")
        except Exception:
            continue
        orig = t

        n_https = t.count("http://yoggies.cz")
        t = t.replace("http://yoggies.cz", "https://yoggies.cz")

        # Pages under site/products/ need to climb one level to reach images/.
        depth = "../" * len(html_file.relative_to(SITE_DIR).parent.parts)
        t, n_local = _localize_img_tags(t, depth, unmapped)
        t, n_fix = _content_fixups(t)
        fixups += n_fix

        if t != orig:
            html_file.write_text(t, encoding="utf-8")
            if n_https:
                fixed_refs += n_https
                fixed_files += 1
            if n_local:
                local_refs += n_local
                local_files += 1

    if fixed_files:
        print(f"HTTPS sweep: fixed {fixed_refs} http://yoggies.cz refs in {fixed_files} files.")
    if local_refs:
        print(f"Image sweep: localized {local_refs} leftover refs in {local_files} files.")
    if fixups:
        print(f"Content fixups: {fixups} (terms link, € format, duplicate SVG ids).")
    if unmapped:
        top = sorted(unmapped.items(), key=lambda kv: -kv[1])[:5]
        total = sum(unmapped.values())
        print(f"Image sweep: {total} refs still external ({len(unmapped)} unique) — "
              f"not in image_map.json. Re-run localize_images.py.")
        for u, n in top:
            print(f"    {n:>5}x {u[:90]}")

    _write_sitemap_and_robots()

def _write_sitemap_and_robots():
    """Emit sitemap.xml and robots.txt from the pages actually on disk.

    Redirect stubs carry <meta name="robots" content="noindex"> and are left
    out — a sitemap listing noindex URLs is a Search Console warning.
    """
    today = datetime.now().strftime("%Y-%m-%d")
    entries = []
    skipped = 0
    for f in sorted(SITE_DIR.rglob("*.html")):
        try:
            t = f.read_text(encoding="utf-8")
        except Exception:
            continue
        if 'name="robots" content="noindex' in t:
            skipped += 1
            continue
        rel = f.relative_to(SITE_DIR).as_posix()
        loc = _abs_url("" if rel == "index.html" else rel)
        # Homepage first, then product pages, then the rest.
        prio = "1.0" if rel == "index.html" else ("0.8" if rel.startswith("products/") else "0.6")
        entries.append((loc, prio))

    body = "\n".join(
        f'  <url>\n    <loc>{html.escape(loc)}</loc>\n'
        f'    <lastmod>{today}</lastmod>\n'
        f'    <priority>{prio}</priority>\n  </url>'
        for loc, prio in entries
    )
    (SITE_DIR / "sitemap.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        f'{body}\n</urlset>\n',
        encoding="utf-8")

    (SITE_DIR / "robots.txt").write_text(
        "User-agent: *\n"
        "Allow: /\n"
        "\n"
        f"Sitemap: {_abs_url('sitemap.xml')}\n",
        encoding="utf-8")
    print(f"sitemap.xml: {len(entries)} URLs ({skipped} redirect stubs excluded). "
          f"robots.txt written.")


def _cat_card(p, slugs, depth=""):
    name = str(p.get("name_nl") or "Product").strip()
    slug = slugs.get(name, _slugify(name))
    img  = str(p.get("image_main") or p.get("doglife_image") or "").strip()
    desc = _truncate_words(str(p.get("description_nl") or "").strip(), 120)
    # Category grids render at card size — 200px is plenty.
    img_tag = (f'<img src="{html.escape(img_src(img, "thumb", depth))}" '
               f'alt="{html.escape(name)}" loading="lazy" decoding="async">'
               if img else "<div>📦</div>")
    # A real anchor, not a div+onclick: this is what makes the product pages
    # crawlable and keyboard-reachable.
    return f'''  <a class="pg-card" href="{depth}products/{slug}.html">
    <div class="pg-img-wrap">{img_tag}</div>
    <div class="pg-body"><h2 class="pg-name">{html.escape(name)}</h2>
    <p class="pg-desc">{html.escape(desc)}</p></div>
  </a>\n'''


def _refresh_nav(content, by_cat=None, slugs=None, current_cat="", depth=""):
    """Swap a page's <ul class="nav-list"> for a freshly generated one.

    Static pages (calculator/Dierenkliniek/hondenatlas) otherwise keep whatever
    nav was baked in when they were last hand-edited — stale product links and
    full-size thumbnails included.
    """
    nav_html = _build_nav(current_cat=current_cat, depth=depth,
                          by_cat=by_cat, slugs=slugs)
    return re.sub(r'<ul class="nav-list"[^>]*>.*?</ul>',
                  lambda _m: nav_html, content, count=1, flags=re.DOTALL)


def _build_nav(current_cat="", depth="", by_cat=None, slugs=None):
    """
    Build full navigation HTML with data-driven dropdowns.
    Each category gets a dropdown with up to 10 products (image + name).
    Yoghurt: single product → direct link, no dropdown.
    Voor katten: virtual category → populated from by_cat.
    Mobile: hamburger toggle + expandable items via JS.
    """
    by_cat  = by_cat  or {}
    slugs   = slugs   or {}

    CAT_NAV = [
        ("barf",         "BARF",                   "barf.html"),
        ("blikvoer",     "Blikvoer",               "blikvoer.html"),
        ("brokken",      "Brokken met probiotica",  "brokken-met-probiotica.html"),
        ("snacks",       "Snacks",                  "snacks.html"),
        ("gekookt",      "Vers gekookt voer",       "vers-gekookt-voer.html"),
        ("supplementen", "Supplementen",            "supplementen.html"),
        ("voor_katten",  "Voor katten",             "voor-katten.html"),
        ("yoghurt",      "Yoghurt",                 "yoghurt.html"),
        ("speelgoed",    "Speelgoed &amp; Accessoires", "speelgoed-accessoires.html"),
    ]

    def _dd_items(cat_key, max_items=10):
        products = by_cat.get(cat_key, [])[:max_items]
        items_html = []
        for p in products:
            nm   = str(p.get("name_nl") or "").strip()
            img  = str(p.get("image_main") or p.get("doglife_image") or "").strip()
            slug = slugs.get(nm, _slugify(nm))
            # 200px thumbs: this dropdown repeats on all ~400 pages, so it is
            # where serving 1200px originals cost the most bandwidth.
            img_tag = (f'<img class="dd-thumb" src="{html.escape(img_src(img, "thumb", depth))}" '
                       f'alt="{html.escape(nm)}" loading="lazy" decoding="async" '
                       f'width="40" height="40">'
                       if img else '<span class="dd-thumb dd-thumb-placeholder"></span>')
            items_html.append(
                f'<a href="{depth}products/{slug}.html" class="dd-item">'
                f'{img_tag}<span>{html.escape(nm[:40])}</span></a>'
            )
        return "".join(items_html)

    items = [
        f'<li class="nav-item">'
        f'<a href="{depth}index.html" class="nav-link nav-home">HOME</a>'
        f'</li>'
    ]

    for key, label, file in CAT_NAV:
        active = ' class="nav-link nav-link--active"' if key == current_cat else ' class="nav-link"'
        products = by_cat.get(key, [])

        # Yoghurt: single product → direct link, no dropdown
        if key == "yoghurt" and len(products) == 1:
            nm   = str(products[0].get("name_nl") or "").strip()
            slug = slugs.get(nm, _slugify(nm))
            items.append(
                f'<li class="nav-item">'
                f'<a href="{depth}products/{slug}.html"{active}>{label}</a>'
                f'</li>'
            )
            continue

        # Build dropdown items
        dd_html = _dd_items(key)
        if dd_html:
            dd_block = (
                f'<div class="dropdown" role="region" aria-label="{label}">'
                f'<div class="dropdown-inner">{dd_html}</div>'
                f'</div>'
            )
            items.append(
                f'<li class="nav-item" data-cat="{key}">'
                f'<a href="{depth}{file}"{active} aria-haspopup="true" aria-expanded="false">'
                f'{label}<span class="nav-chevron" aria-hidden="true"><svg width="10" height="6" viewBox="0 0 10 6" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M1 1L5 5L9 1" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg></span></a>'
                f'{dd_block}'
                f'</li>'
            )
        else:
            items.append(
                f'<li class="nav-item" data-cat="{key}">'
                f'<a href="{depth}{file}"{active}>{label}</a>'
                f'</li>'
            )

    return f'<ul class="nav-list" role="menubar">\n{"".join(items)}\n  </ul>'

def _render_product(p, slug, tmpl, by_cat=None, slugs=None):
    name  = html.escape(str(p.get("name_nl") or "Product").strip())
    desc  = str(p.get("description_nl") or "").strip()
    ingr  = str(p.get("ingredients_cz") or "").strip()
    dose  = str(p.get("dosage_cz") or "").strip()
    img1  = str(p.get("image_main") or "").strip()
    img2  = str(p.get("image_2") or "").strip()
    img3  = str(p.get("image_3") or "").strip()
    dl_url = str(p.get("doglife_url") or "").strip()
    dl_img = str(p.get("doglife_image") or "").strip()
    icons_t = str(p.get("icons_text") or "").strip()
    icons_i = str(p.get("icons_images") or "").strip()
    has_calc = str(p.get("has_calculator") or "").upper() == "TRUE"

    page = tmpl
    # Inject full nav with data-driven dropdowns
    cat = str(p.get("category") or "").strip()
    nav_html = _build_nav(current_cat=cat, depth="../", by_cat=by_cat, slugs=slugs)
    page = re.sub(r'<ul class="nav-list">.*?</ul>', nav_html, page, count=1, flags=re.DOTALL)

    # Force logo link to root homepage (products live in /products/, so use ../)
    page = re.sub(r'(<a\s+class="logo"\s+)href="[^"]*"',
                  r'\1href="../index.html"', page, count=1)
    page = re.sub(r'<h1 class="ph-product-title">[^<]*</h1>',
                  f'<h1 class="ph-product-title">{name}</h1>',
                  page, count=1)
    page = re.sub(r'<title>[^<]*</title>',
                  f'<title>{name} | Yoggies.be</title>',
                  page, count=1)

    main = img1 or dl_img or ""
    page = _inject_head(
        page,
        title=f"{name} | Yoggies.be",
        description=_meta_desc(desc) or _meta_desc(f"{name} van Yoggies — natuurlijke hondenvoeding met probiotica."),
        url_path=f"products/{slug}.html",
        image=img_src(main, "lg") if main else None,
        depth="../",
        og_type="product",
    )

    # Hero: the LCP element, so it is eager + high priority and gets a srcset.
    mi = (f'<img src="{html.escape(img_src(main, "md", "../"))}"'
          f'{img_srcset(main, "../")} alt="{name}" '
          f'decoding="async" fetchpriority="high">'
          if main else '<div style="font-size:48px;text-align:center">📦</div>')
    page = re.sub(r'<div class="ph-main-img"[^>]*>.*?</div>',
                  lambda _m: f'<div class="ph-main-img" id="phMainImg">{mi}</div>',
                  page, count=1, flags=re.DOTALL)

    thumb_srcs = [s for s in [img1, img2, img3, dl_img] if s and s.startswith("http")][:4]
    # data-full carries the 600px rendition so phSwap() can promote a thumb
    # into the hero slot without re-fetching a 1200px original.
    thumbs = "".join(
        f'<div class="ph-thumb{"  active" if i == 0 else ""}" onclick="phSwap(this)">'
        f'<img src="{html.escape(img_src(s, "thumb", "../"))}" '
        f'data-full="{html.escape(img_src(s, "md", "../"))}" '
        f'alt="{name} — afbeelding {i + 1}" loading="lazy" decoding="async"></div>'
        for i, s in enumerate(thumb_srcs)
    ) or '<div class="ph-thumb active" onclick="phSwap(this)"><span style="width:60px;height:60px;background:var(--cream-deep);display:block;border-radius:5px"></span></div>'
    page = re.sub(r'<div class="ph-thumbs">.*?</div>',
                  lambda _m: f'<div class="ph-thumbs">{thumbs}</div>',
                  page, count=1, flags=re.DOTALL)

    il = [translate_icon(l.strip()) for l in icons_t.split("|") if l.strip()][:3]
    ii = [l.strip() for l in icons_i.split("|") if l.strip()][:3]
    badges = "".join(
        f'<div class="ph-badge"><div class="ph-badge-icon">'
        + (f'<img src="{html.escape(img_src(ii[i], "thumb", "../"))}" alt="{html.escape(l)}" '
           f'loading="lazy" decoding="async" width="24" height="24" '
           f'style="width:24px;height:24px">' if i < len(ii) else "🌿")
        + f'</div><span class="ph-badge-label badge-label-{i+1}">{html.escape(l)}</span></div>\n'
        for i, l in enumerate(il)
    )
    page = re.sub(r'<div class="ph-badges">.*?</div>',
                  lambda _m: f'<div class="ph-badges">{badges}</div>',
                  page, count=1, flags=re.DOTALL)

    # Shared contact-fallback HTML (used both as standalone and inside the variant toggle)
    contact_html = (
        '<div class="ph-contact-fallback">'
        '<p class="ph-contact-note">Dit product is niet beschikbaar bij DogLife.be.<br>'
        'Neem rechtstreeks contact op om dit product te bestellen:</p>'
        '<div class="ph-contact-links">'
        '<a href="tel:+32475344868" class="ph-contact-link">📞 +32 (0)475 34 48 68</a>'
        '<a href="mailto:sales@doglife.be" class="ph-contact-link">✉️ sales@doglife.be</a>'
        '</div>'
        '</div>'
    )

    variants = p.get("_variants") or []
    if variants and len(variants) >= 2:
        # Dropdown shows ALL variants. Button/contact toggles per selection.
        # A "real" URL (not a search-results page) means Bestel; otherwise contact.
        opts = "".join(
            f'<option value="{html.escape(v.get("dl_url") or "") if _is_real_product_url(v.get("dl_url")) else ""}">{html.escape(v["size"])}</option>'
            for v in variants
        )
        first = variants[0]
        first_url = (first.get("dl_url") or "").strip()
        first_has = _is_real_product_url(first_url)
        btn_style = "" if first_has else "display:none"
        fb_style  = "display:none" if first_has else ""
        btn = (
            '<div class="ph-variant-row">'
            '<label class="ph-variant-label" for="ph-size-sel">Kies grootte:</label>'
            f'<select class="ph-variant-select" id="ph-size-sel">{opts}</select>'
            '</div>'
            f'<a id="ph-variant-btn" href="{html.escape(first_url) if first_has else "#"}" '
            f'class="ph-order-btn" target="_blank" rel="noopener" style="{btn_style}">'
            f'Bestel bij DogLife.be →</a>'
            f'<div id="ph-variant-fallback" style="{fb_style}">{contact_html}</div>'
            '<script>(function(){'
            'var s=document.getElementById("ph-size-sel");'
            'var b=document.getElementById("ph-variant-btn");'
            'var f=document.getElementById("ph-variant-fallback");'
            'if(!s||!b||!f)return;'
            'function sync(){'
            '  var u=s.value;'
            '  if(u&&u.indexOf("http")===0){b.href=u;b.style.display="";f.style.display="none";}'
            '  else{b.style.display="none";f.style.display="";}'
            '}'
            's.addEventListener("change",sync);'
            # Deep link: the retired per-size URLs redirect here with ?maat=,
            # so preselect that size instead of dumping the visitor on the
            # smallest variant.
            'try{'
            '  var want=new URLSearchParams(location.search).get("maat");'
            '  if(want){'
            '    var norm=function(x){return (x||"").replace(/\\s+/g,"").toLowerCase();};'
            '    for(var i=0;i<s.options.length;i++){'
            '      if(norm(s.options[i].text)===norm(want)){s.selectedIndex=i;break;}'
            '    }'
            '    sync();'
            '  }'
            '}catch(e){}'
            '})();</script>'
        )
    elif _is_real_product_url(dl_url):
        btn = (f'<a href="{html.escape(dl_url)}" class="ph-order-btn" '
               f'target="_blank" rel="noopener">Bestel bij DogLife.be →</a>')
    else:
        # Product not on DogLife — show contact info from footer
        btn = contact_html
    # lambda, not a plain string: the replacement carries generated markup that
    # can contain backslashes (JS regexes) and \1-looking sequences, which
    # re.sub would otherwise try to interpret as escapes/group references.
    page = re.sub(r'<button class="ph-order-btn"[^>]*>.*?</button>',
                  lambda _m: btn, page, count=1, flags=re.DOTALL)

    def _paras(text, fallback):
        if text:
            return "\n".join(f"<p>{html.escape(l)}</p>" for l in text.split("\n") if l.strip())
        return f"<p>{fallback}</p>"

    page = re.sub(
        r'(accordion-summary[^>]*>Beschrijving</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + _paras(desc, "Beschrijving binnenkort beschikbaar.") + r'\n      \2',
        page, count=1, flags=re.DOTALL)

    page = re.sub(
        r'(accordion-summary[^>]*>Ingrediënten[^<]*</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + _paras(ingr, "Zie de verpakking voor ingrediënteninformatie.") + r'\n      \2',
        page, count=1, flags=re.DOTALL)

    page = re.sub(
        r'(accordion-summary[^>]*>Voedingsadvies[^<]*</summary>\s*<div class="accordion-body">).*?(</div>\s*</details>)',
        r'\1\n        ' + _paras(dose, "Raadpleeg de verpakking voor voedingsadvies.") + r'\n      \2',
        page, count=1, flags=re.DOTALL)

    if has_calc and dose:
        calc = _calc_html(dose)
        page = re.sub(
            r'(<details class="accordion">\s*<summary class="accordion-summary">Verzending)',
            calc + r'\n\n    \1', page, count=1)

    return page

def _extract_pct(text):
    m = re.search(r'(\d+[,.]?\d*)\s*[–\-]\s*(\d+[,.]?\d*)\s*%', text)
    if m: return float(m.group(1).replace(',','.')), float(m.group(2).replace(',','.'))
    m = re.search(r'(\d+[,.]?\d*)\s*%', text)
    if m: v = float(m.group(1).replace(',','.')); return v, v
    return 2.0, 3.0

def _calc_html(dose_raw):
    """Voedingscalculator section: short blurb + button linking to calculator.html.
    (The full interactive calculator lives on calculator.html.)"""
    return """
    <details class="accordion" id="calc-accordion">
      <summary class="accordion-summary">🧮 Voedingscalculator</summary>
      <div class="accordion-body">
        <div class="calc-wrap" style="text-align:center;padding:14px 0 8px;">
          <p style="font-size:16px;line-height:1.7;color:#3c1e00;margin:0 0 22px;">
            Bereken eenvoudig de juiste dagelijkse portie op basis van het gewicht, de leeftijd en de activiteit van uw hond.
          </p>
          <a href="../calculator.html"
             style="display:inline-block;background:#A1E647;color:#3c1e00;padding:14px 36px;border-radius:30px;text-decoration:none;font-weight:800;font-size:15px;letter-spacing:.3px;box-shadow:0 4px 14px rgba(60,30,0,.10);transition:background .18s, transform .15s;">
            Open de voedingscalculator &rarr;
          </a>
        </div>
      </div>
    </details>"""


_NAV_CSS = """
  /* ── Dropdown improvements ── */
  nav.main-nav .nav-item { position: relative; }
  .nav-chevron {
    display: inline-flex;
    align-items: center;
    margin-left: 4px;
    transition: transform .2s ease;
    vertical-align: middle;
    opacity: .75;
  }
  .nav-chevron svg { display: block; }
  .nav-item:hover > a .nav-chevron,
  .nav-item--open > a .nav-chevron {
    transform: rotate(180deg);
    opacity: 1;
  }

  /* Nav stacking context — isolates dropdown from page content */
  nav.main-nav { position: relative; z-index: 1000; isolation: isolate; }

  .dropdown {
    display: none;
    position: absolute;
    top: 100%;
    left: 50%;
    transform: translateX(-50%);
    width: min(72vw, 1100px);
    min-width: 480px;
    background: var(--white);
    border-radius: var(--radius);
    box-shadow: 0 12px 48px rgba(0,0,0,.18);
    z-index: 9999;
    padding: 20px 24px;
    border: 1px solid var(--border);
    max-height: 380px;
    overflow-y: auto;
  }
  .nav-item:hover > .dropdown,
  .nav-item--open > .dropdown {
    display: block;
  }
  /* Hover bridge — prevents flicker on gap between link and dropdown */
  .nav-item::after {
    content: '';
    position: absolute;
    top: 100%;
    left: 0; right: 0;
    height: 8px;
    z-index: 1002;
  }

  .dd-thumb-placeholder {
    display: inline-block;
    width: 38px; height: 38px;
    background: var(--cream-deep);
    border-radius: 6px;
    flex-shrink: 0;
  }

  /* ── Mobile nav ── */
  .nav-hamburger {
    display: none;
    background: none;
    border: none;
    cursor: pointer;
    padding: 8px;
    flex-direction: column;
    gap: 5px;
    margin-left: auto;
  }
  .nav-hamburger span {
    display: block;
    width: 24px; height: 2px;
    background: var(--brown);
    border-radius: 2px;
    transition: all .25s;
  }
  .nav-hamburger.is-open span:nth-child(1) { transform: translateY(7px) rotate(45deg); }
  .nav-hamburger.is-open span:nth-child(2) { opacity: 0; }
  .nav-hamburger.is-open span:nth-child(3) { transform: translateY(-7px) rotate(-45deg); }

  @media (max-width: 900px) {
    .nav-hamburger { display: flex; }
    .nav-list {
      display: none;
      position: absolute;
      top: 100%; left: 0; right: 0;
      background: var(--white);
      border-top: 1px solid var(--border);
      box-shadow: 0 8px 24px rgba(0,0,0,.1);
      flex-direction: column;
      align-items: stretch;
      padding: 8px 0 16px;
      z-index: 800;
      overflow-y: auto;
      max-height: calc(100vh - 120px);
    }
    .nav-list.is-open { display: flex; }
    .nav-item { border-bottom: 1px solid var(--cream-deep); }
    .nav-link {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 14px 20px;
      width: 100%;
    }
    nav.main-nav .dropdown {
      position: static !important;
      transform: none !important;
      left: auto !important;
      right: auto !important;
      top: auto !important;
      width: 100% !important;
      min-width: 0 !important;
      max-width: none !important;
      max-height: 240px;
      box-shadow: none;
      border: none;
      border-top: 1px solid var(--cream-deep);
      border-radius: 0;
      padding: 12px 16px;
      opacity: 1;
      pointer-events: auto;
      display: none;
      animation: none !important;
      overflow-x: hidden;
    }
    nav.main-nav .nav-item--open > .dropdown { display: block; }
    nav.main-nav .dropdown-inner {
      grid-template-columns: repeat(2, 1fr) !important;
      gap: 8px;
    }
    nav.main-nav .dd-item {
      min-width: 0 !important;
      overflow: hidden;
      text-overflow: ellipsis;
    }
  }

  @media (max-width: 480px) {
    nav.main-nav .dropdown-inner { grid-template-columns: 1fr !important; }
  }

"""

_NAV_JS = """<script>
(function yoggiesNavInit() {
  var nav = document.querySelector('nav.main-nav');
  if (!nav) return;

  // ── Hamburger ────────────────────────────────────────────────
  var list = nav.querySelector('.nav-list');
  if (list && !nav.querySelector('.nav-hamburger')) {
    var btn = document.createElement('button');
    btn.className = 'nav-hamburger';
    btn.setAttribute('aria-label', 'Menu openen');
    btn.setAttribute('aria-expanded', 'false');
    btn.innerHTML = '<span></span><span></span><span></span>';
    btn.addEventListener('click', function () {
      var open = list.classList.toggle('is-open');
      btn.classList.toggle('is-open', open);
      btn.setAttribute('aria-expanded', String(open));
    });
    nav.insertBefore(btn, list);
  }

  // ── Mobile: click nav-link to toggle dropdown ────────────────
  var isMobile = function () { return window.innerWidth <= 900; };
  nav.querySelectorAll('.nav-item').forEach(function (item) {
    var link = item.querySelector(':scope > a');
    var dd   = item.querySelector(':scope > .dropdown');
    if (!link || !dd) return;
    link.addEventListener('click', function (e) {
      if (!isMobile()) return;       // desktop: CSS hover handles it
      e.preventDefault();
      var wasOpen = item.classList.contains('nav-item--open');
      // Close all others
      nav.querySelectorAll('.nav-item--open').forEach(function (o) {
        o.classList.remove('nav-item--open');
        var a = o.querySelector(':scope > a');
        if (a) a.setAttribute('aria-expanded', 'false');
      });
      item.classList.toggle('nav-item--open', !wasOpen);
      link.setAttribute('aria-expanded', String(!wasOpen));
    });
  });

  // ── Keyboard: Escape closes open dropdown ────────────────────
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      nav.querySelectorAll('.nav-item--open').forEach(function (o) {
        o.classList.remove('nav-item--open');
      });
      if (list) list.classList.remove('is-open');
    }
  });

  // ── Click outside closes menu ────────────────────────────────
  document.addEventListener('click', function (e) {
    if (!nav.contains(e.target)) {
      nav.querySelectorAll('.nav-item--open').forEach(function (o) {
        o.classList.remove('nav-item--open');
      });
      if (list) { list.classList.remove('is-open'); }
      var btn2 = nav.querySelector('.nav-hamburger');
      if (btn2) { btn2.classList.remove('is-open'); btn2.setAttribute('aria-expanded','false'); }
    }
  });
})();
</script>"""

_CALC_CSS = """
    .calc-wrap{padding:4px 0 8px}
    .calc-intro{font-size:14px;color:var(--text-muted);margin-bottom:20px;line-height:1.55}
    .calc-pct-note{display:inline-block;margin-top:4px;color:var(--brown-mid);font-size:13px}
    .calc-row{display:flex;align-items:center;gap:16px;flex-wrap:wrap}
    .calc-label{font-size:14px;font-weight:700;color:var(--brown);min-width:180px}
    .calc-input-group{display:flex;align-items:center;border:2px solid var(--cream-deep);border-radius:var(--radius-sm);overflow:hidden;transition:border-color .2s}
    .calc-input-group:focus-within{border-color:var(--green-dark);box-shadow:0 0 0 3px rgba(125,196,0,.13)}
    .calc-input{border:none;outline:none;padding:10px 14px;font-family:'Nunito',sans-serif;font-size:16px;font-weight:700;width:110px;background:var(--white);color:var(--brown)}
    .calc-unit{padding:10px 14px;background:var(--cream-deep);font-size:13px;font-weight:800;color:var(--brown-mid)}
    .calc-result{margin-top:20px;background:var(--cream);border:2px solid var(--green);border-radius:var(--radius);padding:18px 22px}
    .calc-result-label{font-size:11px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:var(--green-dark)}
    .calc-result-value{font-size:32px;font-weight:900;color:var(--brown);font-family:'Playfair Display',serif}
    .calc-result-note{font-size:12px;color:var(--text-muted);margin-top:4px;line-height:1.5}
    @media(max-width:560px){.calc-row{flex-direction:column}.calc-result-value{font-size:26px}}

    /* Variant size selector (above order button) */
    .ph-variant-row{display:flex;align-items:center;gap:12px;margin:14px 0 12px;flex-wrap:wrap}
    .ph-variant-label{font-weight:800;font-size:14px;color:var(--brown);letter-spacing:.2px}
    .ph-variant-select{padding:10px 36px 10px 14px;border:2px solid var(--cream-deep);border-radius:10px;font-family:inherit;font-size:15px;font-weight:700;background:#fff url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'><path d='M1 1l5 5 5-5' stroke='%233c1e00' stroke-width='2' fill='none' stroke-linecap='round'/></svg>") no-repeat right 12px center;color:var(--brown);cursor:pointer;appearance:none;-webkit-appearance:none;min-width:120px}
    .ph-variant-select:focus{outline:none;border-color:var(--green);box-shadow:0 0 0 3px rgba(125,196,0,.18)}

    /* Contact fallback when product not available on DogLife */
    .ph-contact-fallback{background:var(--cream,#f5f1e6);border:1px solid var(--cream-deep,#e8dcc0);border-radius:12px;padding:18px 20px;margin-top:12px}
    .ph-contact-note{margin:0 0 14px;font-size:14px;line-height:1.55;color:var(--brown,#3c1e00);font-weight:600}
    .ph-contact-links{display:flex;flex-wrap:wrap;gap:8px}
    .ph-contact-link{display:inline-flex;align-items:center;gap:6px;padding:10px 16px;background:var(--brown,#3c1e00);color:#fff !important;text-decoration:none;border-radius:10px;font-weight:700;font-size:13px;transition:background .15s}
    .ph-contact-link:hover{background:#5c2e10}
    @media(max-width:560px){.ph-contact-link{flex:1 1 100%;justify-content:center}}"""

# ── Search dropdown ────────────────────────────────────────────────────
_SEARCH_MARK_S = "/* YGS_SEARCH_CSS_START */"
_SEARCH_MARK_E = "/* YGS_SEARCH_CSS_END */"
_SEARCH_JS_MARK_S = "<!-- YGS_SEARCH_JS_START -->"
_SEARCH_JS_MARK_E = "<!-- YGS_SEARCH_JS_END -->"

_SEARCH_CSS = """
  .search-bar { position: relative; }
  .search-results {
    display: none;
    position: absolute;
    top: calc(100% + 6px); left: 0; right: 0;
    background: #fff;
    border: 1px solid var(--border, #e5e5e5);
    border-radius: 14px;
    box-shadow: 0 12px 36px rgba(0,0,0,.18);
    max-height: 480px;
    overflow-y: auto;
    z-index: 9999;
    min-width: 280px;
  }
  .search-result-item {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 11px 14px;
    color: var(--brown, #3c1e00);
    font-size: 14px;
    text-decoration: none;
    border-bottom: 1px solid var(--cream-deep, #f0e9d8);
    transition: background .12s;
  }
  .search-result-item:last-child { border-bottom: none; }
  .search-result-item:hover { background: var(--cream, #f5f1e6); }
  .search-result-info {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 2px;
    min-width: 0;
  }
  .search-result-name {
    font-weight: 700;
    line-height: 1.3;
    color: var(--brown, #3c1e00);
  }
  .search-result-cat {
    font-size: 11px;
    font-weight: 700;
    color: var(--brown-mid, #7a3b1e);
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .search-empty {
    padding: 16px;
    color: #888;
    font-size: 13px;
    font-style: italic;
    text-align: center;
  }
  @media (max-width: 700px) {
    .search-results {
      position: fixed;
      top: 130px; left: 12px; right: 12px;
      max-height: 60vh;
    }
  }
"""

_SEARCH_JS = """<script>
(function yoggiesSearchInit() {
  var bar = document.querySelector('.search-bar');
  if (!bar) return;
  var input = bar.querySelector('input');
  var btn = bar.querySelector('button, .search-btn');
  if (!input) return;

  var atProduct = location.pathname.indexOf('/products/') >= 0;
  var depth = atProduct ? '../' : '';

  var results = document.createElement('div');
  results.className = 'search-results';
  bar.appendChild(results);

  var index = null, loading = null;
  function loadIndex() {
    if (index) return Promise.resolve(index);
    if (loading) return loading;
    loading = fetch(depth + 'search-index.json')
      .then(function(r) { return r.json(); })
      .then(function(data) { index = data; return data; })
      .catch(function() { index = []; return index; });
    return loading;
  }

  function norm(s) {
    return (s || '').toLowerCase()
      .normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')
      .replace(/[^a-z0-9 ]/g, ' ');
  }

  function search(q) {
    if (!index) return [];
    q = norm(q).trim();
    if (!q) return [];
    var words = q.split(/\\s+/).filter(Boolean);
    var scored = [];
    for (var i = 0; i < index.length; i++) {
      var it = index[i];
      var n = it._n || (it._n = norm(it.name));
      var c = it._c || (it._c = norm(it.cat || ''));
      var score = 0;
      var allMatch = true;
      for (var w = 0; w < words.length; w++) {
        var word = words[w];
        var idx = n.indexOf(word);
        if (idx >= 0) {
          // Higher score if it's at start of name OR at a word boundary
          score += idx === 0 ? 10 : (n.charAt(idx - 1) === ' ' ? 7 : 3);
        } else if (c.indexOf(word) >= 0) {
          score += 2;  // matches category
        } else {
          allMatch = false; break;
        }
      }
      if (allMatch) scored.push({item: it, score: score});
    }
    scored.sort(function(a, b) { return b.score - a.score; });
    var out = [];
    for (var k = 0; k < scored.length && k < 10; k++) out.push(scored[k].item);
    return out;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function(c) {
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
    });
  }

  function render(matches) {
    if (!matches.length) {
      results.innerHTML = '<div class="search-empty">Geen producten gevonden</div>';
      results.style.display = 'block';
      return;
    }
    var html = '';
    for (var i = 0; i < matches.length; i++) {
      var m = matches[i];
      var catHtml = m.cat ? '<span class="search-result-cat">' + escapeHtml(m.cat) + '</span>' : '';
      var href = depth + 'products/' + (m.url ? m.url : m.slug + '.html');
      html += '<a class="search-result-item" href="' + escapeHtml(href) + '">'
           +    '<div class="search-result-info">'
           +      '<span class="search-result-name">' + escapeHtml(m.name) + '</span>'
           +      catHtml
           +    '</div>'
           + '</a>';
    }
    results.innerHTML = html;
    results.style.display = 'block';
  }

  function hide() { results.style.display = 'none'; }

  function doSearch() {
    var q = input.value.trim();
    if (!q) { hide(); return; }
    loadIndex().then(function() { render(search(q)); });
  }

  input.addEventListener('input', doSearch);
  input.addEventListener('focus', function() { if (input.value.trim()) doSearch(); });

  function submit() {
    var matches = search(input.value);
    if (matches.length) {
      location.href = depth + 'products/' + matches[0].slug + '.html';
    }
  }

  input.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') { e.preventDefault(); submit(); }
    if (e.key === 'Escape') { hide(); input.blur(); }
  });
  if (btn) btn.addEventListener('click', function(e) {
    e.preventDefault();
    if (input.value.trim()) submit();
    else input.focus();
  });

  document.addEventListener('click', function(e) {
    if (!bar.contains(e.target)) hide();
  });
})();
</script>"""

def _inject_search_css(content):
    """Idempotently inject search CSS."""
    content = _strip_block(content, _SEARCH_MARK_S, _SEARCH_MARK_E)
    block = _SEARCH_MARK_S + "\n" + _SEARCH_CSS.strip() + "\n" + _SEARCH_MARK_E
    if "</style>" in content:
        content = content.replace("</style>", block + "\n  </style>", 1)
    elif "</head>" in content:
        content = content.replace("</head>", f"<style>{block}</style>\n</head>", 1)
    return content

def _inject_search_js(content):
    """Idempotently inject search JS."""
    content = _strip_block(content, _SEARCH_JS_MARK_S, _SEARCH_JS_MARK_E)
    block = _SEARCH_JS_MARK_S + _SEARCH_JS + _SEARCH_JS_MARK_E
    if "</body>" in content:
        content = content.replace("</body>", block + "\n</body>", 1)
    return content

def _fix_logo_link(content, target="index.html"):
    """Force <a class='logo' href='...'> to point to a specific page."""
    return re.sub(
        r'(<a\s+class="logo"\s+)href="[^"]*"',
        r'\1href="' + target + '"',
        content, count=1
    )

def _minimal_tmpl():
    return """<!DOCTYPE html>
<html lang="nl"><head><meta charset="UTF-8"><title></title>
<link rel="stylesheet" href="../product.css"><style></style></head><body>
<div class="ph-main-img" id="phMainImg"></div>
<div class="ph-thumbs"></div>
<h1 class="ph-product-title"><!-- TITLE HERE --></h1>
<div class="ph-badges"></div>
<button class="ph-order-btn">Bestel bij DogLife.be →</button>
<details class="accordion"><summary class="accordion-summary">Beschrijving</summary><div class="accordion-body"></div></details>
<details class="accordion"><summary class="accordion-summary">Ingrediënten & Samenstelling</summary><div class="accordion-body"></div></details>
<details class="accordion"><summary class="accordion-summary">Voedingsadvies & Dosering</summary><div class="accordion-body"></div></details>
<details class="accordion"><summary class="accordion-summary">Verzending & Levering</summary><div class="accordion-body"><p>Gratis verzending bij bestellingen boven €49. Levertijd: 1–3 werkdagen.</p></div></details>
<script src="../product.js"></script></body></html>"""

# ─── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    n = cache_load()
    print(f"[CACHE] {n} entries | [DATA] {DATA_FILE}")

    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    if mode == "--fill-cache":
        fill_cache_mode()
    elif mode == "--audit":
        audit_mode()
    else:
        build_mode()