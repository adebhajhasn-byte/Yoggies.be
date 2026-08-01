"""
replace_badge_icons.py
═══════════════════════════════════════════════════════════════
Vervangt de emoji 🌿 in product-badges door duidelijke SVG-icons
(één per badge-positie: probiotica, granen, blad).

Voegt ook CSS toe aan site/product.css om de SVG's correct
te laten weergeven binnen de groene cirkel.

Gebruik:  python replace_badge_icons.py
═══════════════════════════════════════════════════════════════
"""
import re
import sys
from pathlib import Path

SITE         = Path(r'C:\Users\AdibHAJHASAN\IdeaProjects\yoggies.be\site')
PRODUCTS_DIR = SITE / 'products'
PRODUCT_CSS  = SITE / 'product.css'

if not PRODUCTS_DIR.exists():
    sys.exit(f"FOUT: {PRODUCTS_DIR} niet gevonden")
if not PRODUCT_CSS.exists():
    sys.exit(f"FOUT: {PRODUCT_CSS} niet gevonden")

# === SVG ICONS (één per positie) ===
SVG_PROBIOTIC = (
    '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" '
    'stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" '
    'aria-hidden="true">'
    '<circle cx="11" cy="13" r="4"/>'
    '<circle cx="20" cy="19" r="4"/>'
    '<circle cx="22" cy="10" r="2"/>'
    '<circle cx="9" cy="22" r="2"/>'
    '</svg>'
)

SVG_GRAIN = (
    '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" '
    'stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" '
    'aria-hidden="true">'
    '<path d="M16 28V12"/>'
    '<path d="M16 12c-3 0-5-2-5-5 3 0 5 2 5 5z"/>'
    '<path d="M16 12c3 0 5-2 5-5-3 0-5 2-5 5z"/>'
    '<path d="M16 18c-3 0-5-2-5-5 3 0 5 2 5 5z"/>'
    '<path d="M16 18c3 0 5-2 5-5-3 0-5 2-5 5z"/>'
    '<path d="M16 24c-3 0-5-2-5-5 3 0 5 2 5 5z"/>'
    '<path d="M16 24c3 0 5-2 5-5-3 0-5 2-5 5z"/>'
    '</svg>'
)

SVG_LEAF = (
    '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" '
    'stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" '
    'aria-hidden="true">'
    '<path d="M16 27A10 10 0 0 1 14 8c8-1.5 10-2.5 13-4.5C29 17 21 25 16 27z"/>'
    '<path d="M4 28c0-4 2.5-7 6.5-8"/>'
    '</svg>'
)

ICONS = {'1': SVG_PROBIOTIC, '2': SVG_GRAIN, '3': SVG_LEAF}

# === CSS toe te voegen om SVG's correct te tonen ===
CSS_BLOCK = '''

/* ── SVG icons binnen product-badges ── */
.ph-badge-icon {
  display: flex !important;
  align-items: center;
  justify-content: center;
}
.ph-badge-icon svg {
  width: 26px;
  height: 26px;
  color: var(--brown);
  stroke: var(--brown);
}
'''

# === Regex: match badge-icon div gevolgd door badge-label met positie ===
PATTERN = re.compile(
    r'<div class="ph-badge-icon">.*?</div>\s*<span class="ph-badge-label badge-label-(\d)">',
    re.DOTALL
)

def replace_badge(m):
    pos = m.group(1)
    icon = ICONS.get(pos, SVG_LEAF)
    return f'<div class="ph-badge-icon">{icon}</div><span class="ph-badge-label badge-label-{pos}">'

# === STAP 1: voeg CSS toe ===
css = PRODUCT_CSS.read_text(encoding='utf-8')
if '.ph-badge-icon svg' in css:
    print(f"OK   CSS al aanwezig in {PRODUCT_CSS.name}")
else:
    PRODUCT_CSS.write_text(css.rstrip() + CSS_BLOCK, encoding='utf-8')
    print(f"OK   CSS toegevoegd aan {PRODUCT_CSS.name}")

# === STAP 2: vervang in alle product-HTMLs ===
modified = 0
total = 0
for f in sorted(PRODUCTS_DIR.glob('*.html')):
    total += 1
    content = f.read_text(encoding='utf-8')
    new_content, count = PATTERN.subn(replace_badge, content)
    if count > 0:
        f.write_text(new_content, encoding='utf-8')
        modified += 1

print(f"\nKlaar:")
print(f"  {modified:3d} productpagina's bijgewerkt met SVG-icons")
print(f"  {total:3d} pagina's in totaal")
print(f"\nOpen een productpagina met Ctrl+Shift+R.")
