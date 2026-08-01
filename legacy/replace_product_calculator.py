"""
replace_product_calculator.py
═══════════════════════════════════════════════════════════════
Vervangt de ingebouwde mini-Voedingscalculator in elke product-
pagina (site/products/*.html) door een korte uitleg + knop die
linkt naar de volledige calculator.html.

Voegt ook de nodige CSS toe aan site/product.css.

Gebruik:  python replace_product_calculator.py
═══════════════════════════════════════════════════════════════
"""
import re
import sys
from pathlib import Path

# === PATHS ===
ROOT         = Path(__file__).resolve().parent.parent
SITE         = ROOT / 'site'
PRODUCTS_DIR = SITE / 'products'
PRODUCT_CSS  = SITE / 'product.css'

if not PRODUCTS_DIR.exists():
    sys.exit(f"FOUT: {PRODUCTS_DIR} niet gevonden")
if not PRODUCT_CSS.exists():
    sys.exit(f"FOUT: {PRODUCT_CSS} niet gevonden")

# === NEW accordion HTML (CTA met link naar volledige calculator) ===
NEW_ACCORDION = '''<details class="accordion">
      <summary class="accordion-summary">🧮 Voedingscalculator</summary>
      <div class="accordion-body">
        <div class="calc-cta">
          <p class="calc-cta-text">
            Bereken eenvoudig de juiste dagelijkse portie voor uw huisdier —
            afgestemd op gewicht, leeftijd, activiteit en voedingstype.
          </p>
          <a href="../calculator.html" class="calc-cta-btn">
            <span>Open Voedingscalculator</span>
            <span aria-hidden="true">→</span>
          </a>
        </div>
      </div>
    </details>'''

# === CSS toe te voegen aan product.css ===
CTA_CSS = '''

/* ── Voedingscalculator CTA (knop naar volledige calculator) ── */
.calc-cta {
  text-align: center;
  padding: 18px 8px 8px;
}
.calc-cta-text {
  color: var(--text-muted);
  font-size: 14px;
  line-height: 1.65;
  margin: 0 auto 20px;
  max-width: 540px;
}
.calc-cta-btn {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  background: var(--green);
  color: var(--brown);
  font-family: 'Nunito', sans-serif;
  font-size: 14px;
  font-weight: 800;
  padding: 13px 28px;
  border-radius: 30px;
  text-decoration: none;
  letter-spacing: .3px;
  transition: background .18s, transform .15s, box-shadow .18s;
  box-shadow: 0 4px 14px rgba(60,30,0,.10);
}
.calc-cta-btn:hover {
  background: var(--green-dark);
  transform: translateY(-2px);
  box-shadow: 0 6px 20px rgba(60,30,0,.18);
}
.calc-cta-btn:active { transform: translateY(0); }
'''

# === Regex die het oude calc-accordion match ===
# Gebruikt id="calc-accordion" om alleen dit specifieke accordion te raken,
# nooit andere accordions (Beschrijving, Ingrediënten, etc.)
CALC_ACCORDION_RE = re.compile(
    r'<details class="accordion" id="calc-accordion">.*?</details>',
    re.DOTALL
)

# === STAP 1: voeg CSS toe aan product.css (indien nog niet aanwezig) ===
css_content = PRODUCT_CSS.read_text(encoding='utf-8')
if '.calc-cta-btn' in css_content:
    print(f"OK   CSS al aanwezig in {PRODUCT_CSS.name}")
else:
    PRODUCT_CSS.write_text(css_content.rstrip() + CTA_CSS, encoding='utf-8')
    print(f"OK   CSS toegevoegd aan {PRODUCT_CSS.name}")

# === STAP 2: vervang accordion in elke product-HTML ===
modified = 0
skipped  = 0
total    = 0

for product_file in sorted(PRODUCTS_DIR.glob('*.html')):
    total += 1
    content = product_file.read_text(encoding='utf-8')
    new_content, count = CALC_ACCORDION_RE.subn(NEW_ACCORDION, content)
    if count > 0:
        product_file.write_text(new_content, encoding='utf-8')
        modified += 1
    else:
        skipped += 1

print(f"\nKlaar:")
print(f"  {modified:3d} productpagina's bijgewerkt")
print(f"  {skipped:3d} pagina's zonder calc-accordion (overgeslagen)")
print(f"  {total:3d} pagina's in totaal")
print(f"\nOpen een productpagina in de browser met Ctrl+Shift+R.")
