"""
2_translate.py — Czech to Dutch product translation
===================================================
Translates every product name and description in the catalogue, then VERIFIES
the result. The previous pipeline shipped 187 descriptions that were never
translated at all: the model returned the Czech text back with a "[CZ]" marker
and nothing checked for it, so the failure passed silently into the build.

This script cannot fail that way. Every translation is validated before it is
cached, and anything that fails validation is retried, then reported.

Validation rules
  - no [CZ] / [NL] / language markers
  - no Czech diacritics (ě š č ř ů) outside quoted brand names
  - no known Czech stop-words (jsou, které, pro psy, ...)
  - length within a sane ratio of the source (no truncation, no rambling)

Cache
  data/cache/translations.json   sha256(source) -> translation
  Seeded from the project's existing translation_cache_v2.json, so work
  already paid for is not repeated.

Requires
    pip install openai
    setx OPENAI_API_KEY "sk-..."            (then reopen the terminal)

Usage
    python pipe/2_translate.py
    python pipe/2_translate.py --only-missing    # skip anything already done
    python pipe/2_translate.py --retry-failed    # re-attempt previous failures
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import time
from pathlib import Path

try:
    from openai import OpenAI
except ImportError:
    sys.exit("Missing dependency. Run: pip install openai")

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
CATALOGUE = DATA / "yoggies_catalogue.json"
OUT = DATA / "yoggies_translated.json"
CACHE_FILE = DATA / "cache" / "translations.json"
LEGACY_CACHE = ROOT / "translation_cache_v2.json"
FAILURES = DATA / "translation_failures.json"

MODEL = "gpt-4o-mini"
MAX_TOKENS = 4000
DELAY = 0.4
RETRIES = 3
CHUNK_CHARS = 3500

SYSTEM = """Je bent een professionele vertaler Tsjechisch naar Nederlands, gespecialiseerd in diervoeding voor de Belgische markt.

REGELS
1. Vertaal ALLES naar natuurlijk, vlot Nederlands. Laat geen enkel Tsjechisch woord staan.
2. Behoud de merknaam "Yoggies" en productlijnnamen (BARF, VET Gastro Sensitive, Ibérico) exact zoals ze zijn.
3. Behoud alle getallen, gewichten, percentages en maateenheden exact.
4. Behoud de alinea-indeling en opsommingstekens (•) van het origineel.
5. Schrijf voor een Belgisch/Nederlands publiek: "brokken" niet "granule", "hondenvoer" niet "krmivo".
6. Prijzen in Tsjechische kronen (Kč) laat je weg - die gelden niet in België.
7. Geef UITSLUITEND de vertaling terug. Geen inleiding, geen uitleg, geen taalmarkering zoals [CZ] of [NL], geen markdown-kopjes.
8. De tekst is commercieel: schrijf wervend en correct, geen letterlijke woord-voor-woord vertaling."""

# Signals that a "translation" is really untranslated Czech.
CZECH_MARKERS = re.compile(
    r"\[(CZ|NL|EN|DE|FR|SK|PL)\]"
    r"|[ěščřůĚŠČŘŮ]"
    r"|\b(jsou|které|která|pro psy|pro kočky|granule|krmivo|balení|masem|"
    r"lisované|vyráběny|obsahuje|díky|nebo|jako|velmi|více|Kč)\b",
    re.IGNORECASE,
)
# Brand-ish exceptions that legitimately keep diacritics.
ALLOWED = re.compile(r"Ibérico|Yoggies|B\.A\.R\.F\.|Markus-Mühle", re.IGNORECASE)


def key_of(text: str) -> str:
    return hashlib.sha256(text.strip().encode("utf-8")).hexdigest()


def load_cache() -> dict[str, str]:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    cache: dict[str, str] = {}
    if LEGACY_CACHE.exists():
        try:
            legacy = json.loads(LEGACY_CACHE.read_text(encoding="utf-8"))
            if isinstance(legacy, dict):
                cache.update({k: v for k, v in legacy.items() if isinstance(v, str)})
        except Exception:  # noqa: BLE001 - a broken legacy cache is not fatal
            pass
    if CACHE_FILE.exists():
        cache.update(json.loads(CACHE_FILE.read_text(encoding="utf-8")))
    return cache


def is_valid(source: str, result: str) -> tuple[bool, str]:
    """Return (ok, reason). Rejecting a bad translation is the whole point."""
    if not result or not result.strip():
        return False, "empty"

    probe = ALLOWED.sub("", result)
    hit = CZECH_MARKERS.search(probe)
    if hit:
        return False, f"Czech left in output: {hit.group(0)!r}"

    ratio = len(result) / max(1, len(source))
    if ratio < 0.45:
        return False, f"too short ({ratio:.2f}x source) - likely truncated"
    if ratio > 2.5:
        return False, f"too long ({ratio:.2f}x source) - model added content"

    if re.search(r"^(hier is|vertaling|translation|here is)\b", result.strip(), re.I):
        return False, "model added a preamble"
    return True, ""


def split_text(text: str) -> list[str]:
    """Split on paragraph boundaries so no sentence is ever cut in half."""
    if len(text) <= CHUNK_CHARS:
        return [text]
    chunks, current = [], ""
    for para in text.split("\n\n"):
        if len(current) + len(para) + 2 > CHUNK_CHARS and current:
            chunks.append(current.strip())
            current = para
        else:
            current = f"{current}\n\n{para}" if current else para
    if current.strip():
        chunks.append(current.strip())
    return chunks


def translate_once(client: OpenAI, text: str, kind: str) -> str:
    hint = "Dit is een PRODUCTNAAM. Houd hem kort." if kind == "name" else ""
    msg = client.chat.completions.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": f"{hint}\n\n{text}".strip()},
        ],
    )
    return (msg.choices[0].message.content or "").strip()


def translate(client, text: str, kind: str, cache: dict, stats: dict) -> tuple[str, str]:
    """Return (translation, error). Uses cache, validates, retries."""
    text = (text or "").strip()
    if not text:
        return "", ""

    k = key_of(text)
    if k in cache:
        cached = cache[k]
        ok, _ = is_valid(text, cached)
        if ok:
            stats["cached"] += 1
            return cached, ""
        del cache[k]  # a previously cached bad translation must not persist

    pieces, problems = [], []
    for chunk in split_text(text):
        result, why = "", "no attempt"
        for attempt in range(1, RETRIES + 1):
            try:
                result = translate_once(client, chunk, kind)
                ok, why = is_valid(chunk, result)
                if ok:
                    break
            except Exception as exc:  # noqa: BLE001
                why = f"api error: {exc}"
                time.sleep(2 * attempt)
            time.sleep(DELAY)
        if why:
            problems.append(why)
        pieces.append(result or chunk)

    joined = "\n\n".join(p for p in pieces if p)
    if problems:
        stats["failed"] += 1
        return joined, "; ".join(problems[:2])

    cache[k] = joined
    stats["translated"] += 1
    return joined, ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--only-missing", action="store_true")
    ap.add_argument("--retry-failed", action="store_true")
    ap.add_argument("--limit", type=int, default=0, help="translate only the first N (for a test run)")
    args = ap.parse_args()

    if not os.environ.get("OPENAI_API_KEY"):
        sys.exit("OPENAI_API_KEY is not set.\n"
                 '  PowerShell:  setx OPENAI_API_KEY "sk-..."\n'
                 "  Then close and reopen the terminal.")
    if not CATALOGUE.exists():
        sys.exit(f"{CATALOGUE.name} not found. Run pipe/1_fetch_yoggies.py first.")

    catalogue = json.loads(CATALOGUE.read_text(encoding="utf-8"))
    if args.limit:
        catalogue = catalogue[: args.limit]

    done: dict[str, dict] = {}
    if OUT.exists() and args.only_missing:
        done = {p["handle"]: p for p in json.loads(OUT.read_text(encoding="utf-8"))}

    cache = load_cache()
    print(f"{len(catalogue)} products | {len(cache)} cached translations\n")

    client = OpenAI()
    stats = {"cached": 0, "translated": 0, "failed": 0}
    output, failures = [], []

    for n, item in enumerate(catalogue, 1):
        if item["handle"] in done and not args.retry_failed:
            output.append(done[item["handle"]])
            continue

        name_nl, err_n = translate(client, item["name_cz"], "name", cache, stats)
        desc_nl, err_d = translate(client, item["description_cz"], "body", cache, stats)

        record = dict(item)
        record["name_nl"] = name_nl
        record["description_nl"] = desc_nl
        record["translation_ok"] = not (err_n or err_d)
        output.append(record)

        if err_n or err_d:
            failures.append({"handle": item["handle"], "name": item["name_cz"],
                             "name_error": err_n, "description_error": err_d})

        flag = "" if record["translation_ok"] else "  <-- FAILED"
        print(f"  [{n}/{len(catalogue)}] {item['name_cz'][:56]}{flag}")

        if n % 25 == 0:
            CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
            OUT.write_text(json.dumps(output, ensure_ascii=False, indent=1), encoding="utf-8")

    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    OUT.write_text(json.dumps(output, ensure_ascii=False, indent=1), encoding="utf-8")
    FAILURES.write_text(json.dumps(failures, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"\nfrom cache : {stats['cached']}")
    print(f"translated : {stats['translated']}")
    print(f"failed     : {stats['failed']}")
    print(f"\nWrote {OUT.relative_to(ROOT)}")
    if failures:
        print(f"{len(failures)} products need attention -> {FAILURES.relative_to(ROOT)}")
        print("Re-run with --retry-failed to try them again.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
