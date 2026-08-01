"""
fix_home_categories.py (v2 — meer robuust + debug-output)
═══════════════════════════════════════════════════════════════
Werkt site/index.html bij:
  1. Vervangt 5 placeholder-labels door dier-categorieën
  2. Vervangt categorie-blok met echte categorieën uit nav

Robuuste nav-parsing: vindt elke <li class="nav-item"> en
extracteert de eerste <a> ongeacht attribuut-volgorde of extra
attributen.

Gebruik:  python fix_home_categories.py
═══════════════════════════════════════════════════════════════
"""
import re
import sys
from pathlib import Path

INDEX = Path(r'C:\Users\AdibHAJHASAN\IdeaProjects\yoggies.be\site\index.html')

if not INDEX.exists():
    sys.exit(f"FOUT: {INDEX} niet gevonden")

content = INDEX.read_text(encoding='utf-8')

# === STAP 1: extract nav-items robust ===
nav_items = re.findall(
    r'<li\s+class="nav-item"[^>]*>(.*?)</li>',
    content, re.DOTALL
)
print(f"Gevonden {len(nav_items)} <li class=\"nav-item\"> blokken\n")

categories = []
for i, item in enumerate(nav_items, 1):
    # Vind eerste <a> tag — werkt met elke attribuut-volgorde
    anchor = re.search(r'<a\s([^>]*)>(.*?)</a>', item, re.DOTALL)
    if not anchor:
        print(f"  {i}. (geen <a> gevonden)")
        continue

    attrs = anchor.group(1)
    inner = anchor.group(2)

    # Skip HOME
    if 'nav-home' in attrs:
        print(f"  {i}. HOME  (overgeslagen)")
        continue

    # Extract href
    href_m = re.search(r'href="([^"]*)"', attrs)
    if not href_m:
        print(f"  {i}. (geen href in: {attrs[:60]}...)")
        continue
    href = href_m.group(1)

    # Tekst opschonen
    text = re.sub(r'<[^>]+>', '', inner).strip()

    if text:
        categories.append((href, text))
        print(f"  {i}. {text:35s} -> {href}")

if not categories:
    sys.exit("\nFOUT: geen categorieen gevonden")

print(f"\n{len(categories)} categorieen klaar.\n")

# === STAP 2: dier-iconen labels ===
ANIMAL_LABELS = ['Grote honden', 'Middelgrote honden', 'Kleine honden', 'Pups', 'Katten']

counter = [0]
def replace_animal_label(m):
    if counter[0] < len(ANIMAL_LABELS):
        new = ANIMAL_LABELS[counter[0]]
        counter[0] += 1
        return f'<div class="animal-label">{new}</div>'
    return m.group(0)

content, n_labels = re.subn(
    r'<div class="animal-label">[^<]+</div>',
    replace_animal_label,
    content
)
print(f"OK   {n_labels} dier-labels bijgewerkt")

# === STAP 3: categorie-blok ===
buttons = []
for href, name in categories:
    buttons.append(f'    <a href="{href}" class="cat-btn">{name} <span class="arrow">&rsaquo;</span></a>')

new_categories = '<div class="categories">\n' + '\n'.join(buttons) + '\n  </div>'

content, n_cats = re.subn(
    r'<div class="categories">.*?</div>',
    lambda m: new_categories,
    content,
    count=1,
    flags=re.DOTALL
)
print(f"OK   categorie-blok bijgewerkt ({len(categories)} knoppen)")

INDEX.write_text(content, encoding='utf-8')
print(f"\nKlaar - {INDEX.name} opgeslagen.")
print("Open in browser met Ctrl+Shift+R.")