"""
1_fetch_yoggies.py — pull the complete yoggies.cz catalogue
===========================================================
yoggies.cz runs on Shopify, which exposes the whole catalogue as JSON at
/products.json. That is the authoritative source: every product, full Czech
description, every image, every size variant, plus product_type and tags.

This replaces HTML scraping entirely. Scraping only ever saw what a page
happened to render; this sees all 526 products with complete fields.

Output
  data/yoggies_catalogue.json   normalised product records
  data/yoggies_raw.json         untouched API response (kept for auditing)

Usage
    python pipe/1_fetch_yoggies.py
    python pipe/1_fetch_yoggies.py --refresh      # ignore existing files
    python pipe/1_fetch_yoggies.py --from-raw     # re-extract from the saved raw
                                                  # response, no download
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("Missing dependency. Run: pip install requests")

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = DATA / "yoggies_catalogue.json"
RAW = DATA / "yoggies_raw.json"

BASE = "https://yoggies.cz"
PAGE_SIZE = 250
MAX_PAGES = 10
TIMEOUT = 30
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

TAG_RE = re.compile(r"<[^>]+>")
WS_RE = re.compile(r"[ \t\r\f\v]+")

# Elements whose CONTENTS are code, not prose. Stripping only their tags leaves
# the stylesheet or the script body behind as "text", which is how 111 product
# descriptions ended up opening with ".yoggies-section { max-width: 1100px; ... }".
# An unclosed tag is treated as running to the end of the document rather than
# being ignored - leaking the whole tail is safer than leaking code.
CODE_ELEMENT_RE = re.compile(
    r"(?is)<\s*(style|script|noscript)\b[^>]*>.*?(?:<\s*/\s*\1\s*>|$)"
)
COMMENT_RE = re.compile(r"(?s)<!--.*?(?:-->|$)")


def html_to_text(html: str) -> str:
    """Flatten Shopify's body_html into readable plain text.

    Block-level tags become newlines so paragraph structure survives; the
    translation step needs those boundaries to keep sentences intact.

    <style>, <script> and <noscript> are removed element AND contents first,
    together with HTML comments, so no CSS or JavaScript can survive as text.
    """
    if not html:
        return ""
    text = CODE_ELEMENT_RE.sub("\n", html)
    text = COMMENT_RE.sub(" ", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(p|div|li|h[1-6]|tr)>", "\n", text)
    text = re.sub(r"(?i)<li[^>]*>", "• ", text)
    text = TAG_RE.sub(" ", text)
    for entity, char in (
        ("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
        ("&quot;", '"'), ("&#39;", "'"), ("&scaron;", "š"), ("&eacute;", "é"),
    ):
        text = text.replace(entity, char)
    text = WS_RE.sub(" ", text)
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    return "\n".join(line.strip() for line in text.split("\n")).strip()


def fetch_page(page: int) -> list[dict]:
    url = f"{BASE}/products.json?limit={PAGE_SIZE}&page={page}"
    for attempt in range(1, 4):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            resp.raise_for_status()
            return resp.json().get("products", [])
        except Exception as exc:  # noqa: BLE001
            if attempt == 3:
                raise RuntimeError(f"page {page}: {exc}") from exc
            time.sleep(2 * attempt)
    return []


def normalise(product: dict) -> dict:
    """Reduce a Shopify product to the fields the site build actually needs."""
    variants = product.get("variants") or []
    images = [i.get("src", "") for i in (product.get("images") or []) if i.get("src")]
    body = html_to_text(product.get("body_html", ""))

    sizes = []
    for v in variants:
        title = (v.get("title") or "").strip()
        if title and title.lower() != "default title":
            sizes.append(title)

    return {
        "id": product.get("id"),
        "handle": product.get("handle", ""),
        "url": f"{BASE}/products/{product.get('handle','')}",
        "name_cz": (product.get("title") or "").strip(),
        "description_cz": body,
        "product_type": (product.get("product_type") or "").strip(),
        "tags": product.get("tags") or [],
        "vendor": (product.get("vendor") or "").strip(),
        "images": images,
        "image_main": images[0] if images else "",
        "variants": sizes,
        "variant_count": len(variants),
        "available": any(v.get("available") for v in variants),
        "published_at": product.get("published_at"),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--refresh", action="store_true")
    ap.add_argument("--from-raw", action="store_true",
                    help="re-normalise data/yoggies_raw.json instead of downloading "
                         "(use after changing html_to_text)")
    args = ap.parse_args()

    DATA.mkdir(parents=True, exist_ok=True)

    if OUT.exists() and not (args.refresh or args.from_raw):
        existing = json.loads(OUT.read_text(encoding="utf-8"))
        print(f"{OUT.name} already has {len(existing)} products. Use --refresh to re-download.")
        return 0

    raw: list[dict] = []
    if args.from_raw:
        # The API response is kept verbatim precisely so extraction can be redone
        # without hitting yoggies.cz again.
        if not RAW.exists():
            sys.exit(f"{RAW.name} not found. Run without --from-raw to download it.")
        raw = json.loads(RAW.read_text(encoding="utf-8"))
        print(f"  re-normalising {len(raw)} products from {RAW.name} (no download)")
    else:
        for page in range(1, MAX_PAGES + 1):
            batch = fetch_page(page)
            print(f"  page {page}: {len(batch)} products")
            if not batch:
                break
            raw.extend(batch)
            if len(batch) < PAGE_SIZE:
                break
            time.sleep(0.5)

        if not raw:
            sys.exit("Nothing downloaded. Check your internet connection and try again.")

        RAW.write_text(json.dumps(raw, ensure_ascii=False, indent=1), encoding="utf-8")

    catalogue = [normalise(p) for p in raw]
    # A handle is unique per shop, so it is a safe deduplication key.
    seen: dict[str, dict] = {}
    for item in catalogue:
        seen[item["handle"]] = item
    catalogue = list(seen.values())

    OUT.write_text(json.dumps(catalogue, ensure_ascii=False, indent=1), encoding="utf-8")

    no_desc = sum(1 for c in catalogue if not c["description_cz"])
    no_img = sum(1 for c in catalogue if not c["image_main"])
    types = {}
    for c in catalogue:
        types[c["product_type"] or "(none)"] = types.get(c["product_type"] or "(none)", 0) + 1

    print(f"\n{len(catalogue)} products written to {OUT.relative_to(ROOT)}")
    print(f"  without description : {no_desc}")
    print(f"  without image       : {no_img}")
    print(f"  total images        : {sum(len(c['images']) for c in catalogue)}")
    print("\n  product types:")
    for k, v in sorted(types.items(), key=lambda x: -x[1])[:15]:
        print(f"    {v:4d}  {k}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
