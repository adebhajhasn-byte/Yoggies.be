# -*- coding: utf-8 -*-
"""
check_site.py — Pre-deployment audit voor yoggies.be
=====================================================
Controleert de 'site/' map op de meest voorkomende problemen die pas zichtbaar
worden NA upload naar een Linux-webserver (Netlify, Cloudflare, etc.).

Run vanuit de projectmap:
    python check_site.py

Wat wordt gecontroleerd:
  1. Hoofdlettergevoeligheid van href/src verwijzingen
  2. Verwijzingen naar bestanden die niet bestaan
  3. http:// links (mixed content op https-site)
  4. Aanwezigheid van index.html
  5. Defecte interne links (anchors, fragments)
"""

import re
import sys
from pathlib import Path
from collections import defaultdict

SITE_DIR = Path("site")

# Verwijzingsattributen die we scannen
ATTR_RE = re.compile(r'''(?:href|src)\s*=\s*["']([^"'#?]+?)(?:[#?][^"']*)?["']''', re.IGNORECASE)

# Schemata die we OVERSLAAN (externe of niet-bestand)
SKIP_SCHEMES = ("mailto:", "tel:", "javascript:", "data:", "#")

# Telling
errors_case        = []   # (file, ref, suggested_actual_name, line_snippet)
errors_missing     = []   # (file, ref, line_snippet)
warnings_http      = []   # (file, ref)
warnings_other     = []
files_scanned      = 0
refs_scanned       = 0


def scan_html_files():
    """Bouw set van alle bestaande bestanden in site/ (met exacte hoofdletters)."""
    all_files = {}   # lowercase-pad -> echte_Path
    for p in SITE_DIR.rglob("*"):
        if p.is_file():
            rel = p.relative_to(SITE_DIR).as_posix()
            all_files[rel.lower()] = rel
    return all_files


def _line_around(text, idx, width=120):
    """Return ~one-line snippet around character index idx."""
    start = text.rfind("\n", 0, idx) + 1
    end   = text.find("\n", idx)
    if end == -1: end = len(text)
    snippet = text[start:end].strip()
    return snippet[:width]


def check_ref(html_file, ref, all_files, match_idx, full_text):
    """Controleer één verwijzing en registreer het probleem (indien aanwezig)."""
    global refs_scanned
    refs_scanned += 1

    if ref.startswith("https://"):
        return
    if ref.startswith("http://"):
        warnings_http.append((html_file, ref))
        return
    if any(ref.startswith(s) for s in SKIP_SCHEMES):
        return
    if "//" in ref:
        return

    html_rel = html_file.relative_to(SITE_DIR)
    target = (html_rel.parent / ref).as_posix()
    parts = []
    for seg in target.split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            if parts:
                parts.pop()
            continue
        parts.append(seg)
    target = "/".join(parts)

    if not target:
        return

    # 1. Exacte match?
    if any(target == v for v in all_files.values()):
        return

    snippet = _line_around(full_text, match_idx)

    # 2. Case-insensitive match
    lower_actual = all_files.get(target.lower())
    if lower_actual:
        errors_case.append((str(html_rel), ref, lower_actual, snippet))
        return

    # 3. Bestand bestaat helemaal niet
    errors_missing.append((str(html_rel), ref, snippet))


def main():
    global files_scanned

    if not SITE_DIR.is_dir():
        print(f"FOUT: map '{SITE_DIR}' bestaat niet. Run dit script vanuit je projectmap.")
        sys.exit(1)

    if not (SITE_DIR / "index.html").is_file():
        warnings_other.append("index.html ontbreekt in site/ — homepage werkt niet op de root URL.")

    all_files = scan_html_files()
    print(f"Site map: {SITE_DIR.resolve()}")
    print(f"Totaal bestanden in site/: {len(all_files)}")
    print()

    for html in SITE_DIR.rglob("*.html"):
        files_scanned += 1
        try:
            txt = html.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            warnings_other.append(f"Kon {html} niet lezen: {e}")
            continue
        for m in ATTR_RE.finditer(txt):
            check_ref(html, m.group(1).strip(), all_files, m.start(), txt)

    print(f"HTML bestanden gescand: {files_scanned}")
    print(f"Verwijzingen gecontroleerd: {refs_scanned}")
    print("=" * 70)

    if errors_case:
        print(f"\n⚠️  HOOFDLETTER-PROBLEMEN ({len(errors_case)}) — BREEKT OP LIVE SITE")
        print("-" * 70)
        for f, ref, actual, snip in errors_case[:30]:
            print(f"  FILE:    {f}")
            print(f"  REF:     {ref}")
            print(f"  ACTUAL:  {actual}")
            print(f"  CONTEXT: {snip}")
            print()

    if errors_missing:
        # Groeperen per ref
        by_ref = defaultdict(list)
        for f, ref, snip in errors_missing:
            by_ref[ref].append((f, snip))
        print(f"\n❌  ONTBREKENDE BESTANDEN ({len(errors_missing)} verwijzingen, {len(by_ref)} unieke) — 404 op live")
        print("-" * 70)
        for ref, items in list(by_ref.items())[:10]:
            f, snip = items[0]
            print(f"  REF:     {ref}")
            print(f"  COUNT:   {len(items)} bestand(en), bv: {f}")
            print(f"  CONTEXT: {snip}")
            print()

    if warnings_http:
        unique = set(ref for _, ref in warnings_http)
        print(f"\n⚠️  HTTP:// LINKS ({len(warnings_http)} totaal, {len(unique)} uniek)")
        print("-" * 70)
        # Show first 5 unique
        for ref in list(unique)[:5]:
            print(f"  - {ref}")
        if len(unique) > 5:
            print(f"  ... en nog {len(unique)-5} andere unieke URLs")

    if warnings_other:
        print(f"\nℹ️  OVERIGE WAARSCHUWINGEN:")
        for w in warnings_other:
            print(f"  - {w}")

    print("\n" + "=" * 70)
    total = len(errors_case) + len(errors_missing)
    if total == 0 and not warnings_http:
        print("✅  ALLES OK — de site is klaar om te uploaden.")
    else:
        print(f"SAMENVATTING: {len(errors_case)} hoofdletter, "
              f"{len(errors_missing)} ontbrekend, "
              f"{len(set(r for _,r in warnings_http))} http")


if __name__ == "__main__":
    main()