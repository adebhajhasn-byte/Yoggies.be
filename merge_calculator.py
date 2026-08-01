"""
merge_calculator.py
═══════════════════════════════════════════════════════════════
Bouwt site/calculator.html opnieuw op met dezelfde header, nav
en footer als site/index.html. Behoudt alle calculator-content
(hero, Hond/Kat tabs, alle formulieren, JS).

Gebruik:  python merge_calculator.py
═══════════════════════════════════════════════════════════════
"""
import re
import sys
from pathlib import Path

# === PATHS ===
SITE  = Path(r'C:\Users\AdibHAJHASAN\IdeaProjects\yoggies.be\site')
INDEX = SITE / 'index.html'
CALC  = SITE / 'calculator.html'

if not INDEX.exists():
    sys.exit(f"FOUT: {INDEX} niet gevonden")
if not CALC.exists():
    sys.exit(f"FOUT: {CALC} niet gevonden")

# === READ ===
home     = INDEX.read_text(encoding='utf-8')
calc_old = CALC.read_text(encoding='utf-8')

# === EXTRACT calculator's main content (between </nav> and <footer>) ===
m = re.search(r'</nav>\s*(.*?)\s*(?:<!--[^>]*-->\s*)?<footer', calc_old, re.DOTALL)
if not m:
    sys.exit("FOUT: calculator-content niet gevonden (geen </nav>...<footer>)")
calc_content = m.group(1).strip()
calc_content = re.sub(r'\s*<!--\s*FOOTER\s*-->\s*$', '', calc_content, flags=re.IGNORECASE)

# === BUILD new calculator.html from index.html as template ===
new_calc = home

# Update <title>
new_calc = re.sub(
    r'<title>[^<]*</title>',
    '<title>Yoggies – Voedingscalculator</title>',
    new_calc, count=1
)

# Link calculator.css (after home.css)
if 'calculator.css' not in new_calc:
    new_calc = new_calc.replace(
        '<link rel="stylesheet" href="home.css">',
        '<link rel="stylesheet" href="home.css">\n  <link rel="stylesheet" href="calculator.css">',
        1
    )

# Replace home's main content (between </nav> and <footer>) with calculator content
new_calc = re.sub(
    r'(</nav>)\s*.*?\s*(<footer)',
    lambda m: m.group(1) + '\n\n' + calc_content + '\n\n' + m.group(2),
    new_calc,
    flags=re.DOTALL,
    count=1
)

# Add calculator.js before </body>
if 'calculator.js' not in new_calc:
    new_calc = new_calc.replace(
        '</body>',
        '  <script src="calculator.js"></script>\n</body>',
        1
    )

# === SAVE ===
CALC.write_text(new_calc, encoding='utf-8')

print(f"OK  {CALC.name} herbouwd ({len(new_calc):,} tekens)")
print(f"    Open in browser met Ctrl+Shift+R om cache te legen.")
