"""
merge_extra_pages.py
═══════════════════════════════════════════════════════════════
Bouwt site/hondenatlas.html en site/Dierenkliniek.html opnieuw op
met dezelfde header, nav en footer als site/index.html. Behoudt
de unieke content (hero, secties, formulier) van elke pagina.

Vereist:
  - site/index.html bestaat
  - site/hondenatlas.html en site/hondenatlas.css in site/ geplaatst
  - site/Dierenkliniek.html en site/Dierenkliniek.css in site/ geplaatst

Gebruik:  python merge_extra_pages.py
═══════════════════════════════════════════════════════════════
"""
import re
import sys
from pathlib import Path

ROOT  = Path(__file__).resolve().parent.parent
SITE  = ROOT / 'site'
INDEX = SITE / 'index.html'

# Pagina's om bij te werken: (HTML, CSS, titel)
PAGES = [
    ('hondenatlas.html',   'hondenatlas.css',   'Yoggies – Hondenatlas'),
    ('Dierenkliniek.html', 'Dierenkliniek.css', 'Yoggies – Dierenkliniek'),
]

if not INDEX.exists():
    sys.exit(f"FOUT: {INDEX} niet gevonden")

home = INDEX.read_text(encoding='utf-8')

print(f"Template: {INDEX.name} ({len(home):,} chars)\n")

updated = 0
skipped = 0

for page_html, page_css, title in PAGES:
    page_path = SITE / page_html
    if not page_path.exists():
        print(f"OVERSLAAN  {page_html} (niet gevonden in site/)")
        skipped += 1
        continue

    page_old = page_path.read_text(encoding='utf-8')

    # Extract main content (between </nav> and <footer>)
    m = re.search(
        r'</nav>\s*(.*?)\s*(?:<!--[^>]*-->\s*)?<footer',
        page_old, re.DOTALL
    )
    if not m:
        print(f"FOUT       {page_html}: kan content niet vinden tussen </nav> en <footer>")
        skipped += 1
        continue
    page_content = m.group(1).strip()

    # Build new from home template
    new_page = home

    # Update title
    new_page = re.sub(
        r'<title>[^<]*</title>',
        f'<title>{title}</title>',
        new_page, count=1
    )

    # Link page-specific CSS (if not already present)
    if f'href="{page_css}"' not in new_page:
        new_page = new_page.replace(
            '<link rel="stylesheet" href="home.css">',
            f'<link rel="stylesheet" href="home.css">\n  <link rel="stylesheet" href="{page_css}">',
            1
        )

    # Replace home's main content (between </nav> and <footer>) with page content
    new_page = re.sub(
        r'(</nav>)\s*.*?\s*(<footer)',
        lambda m: m.group(1) + '\n\n' + page_content + '\n\n' + m.group(2),
        new_page,
        flags=re.DOTALL,
        count=1
    )

    # Save
    page_path.write_text(new_page, encoding='utf-8')
    print(f"OK         {page_html} herbouwd ({len(new_page):,} chars)")
    updated += 1

print(f"\n{updated} bijgewerkt, {skipped} overgeslagen")
print("Open in browser met Ctrl+Shift+R om cache te legen.")
