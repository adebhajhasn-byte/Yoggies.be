"""
download_images.py
==================
Downloads alle externe afbeeldingen (image2url.com, etc.) en host ze lokaal
in site/images/. Werkt home.html, product.html, hoofdpageproduct.html bij
zodat ze de lokale paden gebruiken.

Run dit ÉÉN keer (of telkens wanneer er nieuwe externe afbeeldingen worden
toegevoegd aan de templates).

Vereisten:
    pip install requests
"""
import re
import hashlib
import shutil
from pathlib import Path
from urllib.parse import urlparse

try:
    import requests
except ImportError:
    print("X 'requests' niet geïnstalleerd. Run: pip install requests")
    raise SystemExit(1)

PROJECT_ROOT = Path(__file__).parent
SITE_DIR     = PROJECT_ROOT / "site"
IMG_DIR      = SITE_DIR / "images"

# Bronbestanden die we doorlopen (zowel template als gebouwde site)
TARGETS = [
    PROJECT_ROOT / "home.html",
    PROJECT_ROOT / "product.html",
    PROJECT_ROOT / "hoofdpageproduct.html",
    SITE_DIR     / "home.html",
    SITE_DIR     / "index.html",
]

# Externe afbeeldings-hosts die we lokaal opslaan
HOSTS_TO_LOCALIZE = (
    "image2url.com",
    "imgtourl.com",
    "i.imgur.com",
    "imgur.com",
)

URL_RE = re.compile(r'''(?:src|href)=["'](https?://[^"']+)["']''')


def is_external_image(url):
    if not url.startswith(("http://", "https://")):
        return False
    host = urlparse(url).netloc.lower()
    return any(h in host for h in HOSTS_TO_LOCALIZE)


def local_filename_for(url):
    """Build a deterministic local filename based on URL hash + original ext."""
    parsed = urlparse(url)
    orig_name = Path(parsed.path).name or "image"
    # Keep extension, hash the rest
    ext = Path(orig_name).suffix or ".jpg"
    h = hashlib.md5(url.encode()).hexdigest()[:12]
    return f"{h}{ext}"


def collect_urls():
    urls = set()
    for f in TARGETS:
        if not f.is_file():
            continue
        text = f.read_text(encoding="utf-8", errors="replace")
        for m in URL_RE.finditer(text):
            u = m.group(1)
            if is_external_image(u):
                urls.add(u)
    return urls


def download(url, out_path):
    try:
        r = requests.get(url, timeout=20, headers={
            "User-Agent": "Mozilla/5.0 (compatible; YoggiesPipeline/1.0)",
            "Accept":     "image/*,*/*;q=0.5",
        })
        if r.status_code == 200 and len(r.content) > 100:
            out_path.write_bytes(r.content)
            return True
        return False
    except Exception as e:
        print(f"   X {e}")
        return False


def rewrite_files(mapping):
    """Vervang alle externe URLs door lokale paden in TARGETS."""
    rewritten = 0
    for f in TARGETS:
        if not f.is_file():
            continue
        text = f.read_text(encoding="utf-8", errors="replace")
        original = text
        for url, local_name in mapping.items():
            # Bepaal pad relatief aan dit bestand
            if f.parent == SITE_DIR:
                local_path = f"images/{local_name}"
            elif "products" in f.parts:
                local_path = f"../images/{local_name}"
            else:
                # Templates in projectroot: gebruik relative naar site/
                local_path = f"site/images/{local_name}"
            text = text.replace(url, local_path)
        if text != original:
            f.write_text(text, encoding="utf-8")
            print(f"   . Updated {f.relative_to(PROJECT_ROOT)}")
            rewritten += 1
    return rewritten


def main():
    if not SITE_DIR.is_dir():
        print(f"X Site directory '{SITE_DIR}' bestaat niet. Run eerst yoggies_pipeline.py.")
        return

    IMG_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print("Externe afbeeldingen lokaliseren")
    print("=" * 70)

    urls = collect_urls()
    print(f"\nGevonden: {len(urls)} unieke externe afbeeldingsverwijzingen")
    if not urls:
        print("Niets te doen.")
        return

    mapping = {}
    successes = 0
    failures = []

    for i, url in enumerate(sorted(urls), 1):
        local_name = local_filename_for(url)
        local_file = IMG_DIR / local_name
        print(f"\n[{i}/{len(urls)}] {url[:80]}")
        if local_file.exists() and local_file.stat().st_size > 100:
            print(f"   * already downloaded ({local_file.name})")
            mapping[url] = local_name
            successes += 1
            continue
        if download(url, local_file):
            print(f"   + saved as images/{local_name}  ({local_file.stat().st_size:,} bytes)")
            mapping[url] = local_name
            successes += 1
        else:
            print(f"   X DOWNLOAD FAILED")
            failures.append(url)

    print("\n" + "=" * 70)
    print(f"Downloaded: {successes}/{len(urls)}")
    if failures:
        print(f"Failed:     {len(failures)} (these URLs no longer work — image is gone)")

    if mapping:
        print("\nHTML-bestanden bijwerken met lokale paden...")
        n = rewrite_files(mapping)
        print(f"\nGewijzigd: {n} bestand(en).")

    if failures:
        print("\n" + "!" * 70)
        print("LET OP: deze afbeeldingen konden NIET worden gedownload.")
        print("Ze blijven extern verwijzen (en blijven kapot).")
        print("Je moet handmatig vervangingen vinden:")
        for u in failures[:10]:
            print(f"   - {u}")
        if len(failures) > 10:
            print(f"   ... +{len(failures)-10} meer")
        print("!" * 70)

    print("\nKlaar. Run nu opnieuw 'python yoggies_pipeline.py'.")


if __name__ == "__main__":
    main()
