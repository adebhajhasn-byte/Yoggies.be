"""
2b_import_translations.py — read completed batches back into the cache
======================================================================
Counterpart to 2a_export_batches.py. Takes the reply you pasted out of the chat
interface, in the same numbered format the batch was written in:

    ===== 001 =====
    <Nederlandse vertaling>

Every block is checked with is_valid() from 2_translate.py - the same validator
the API path uses, so a translation that still contains Czech is rejected here
exactly as it would be there. Nothing unvalidated reaches the cache.

Two kinds of entry are written to data/cache/translations.json:

  paragraph  sha256(paragraph)  -> its translation. Shared boilerplate is stored
                                   once and reused by every product repeating it.
  field      sha256(full text)  -> the paragraphs rejoined with blank lines.
                                   This is the key 2_translate.py looks up, so a
                                   field is only written once all of its
                                   paragraphs have been imported.

A field whose paragraphs are spread over several batches simply completes on the
import of the last one - re-running this script is safe and idempotent.

Usage
    python pipe/2b_import_translations.py done_001.md
    python pipe/2b_import_translations.py done_*.md
    python pipe/2b_import_translations.py done_001.md --dry-run
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CATALOGUE = ROOT / "data" / "yoggies_catalogue.json"
BATCH_DIR = ROOT / "data" / "translate_batches"
INDEX = BATCH_DIR / "_index.json"

_spec = importlib.util.spec_from_file_location("t2", Path(__file__).with_name("2_translate.py"))
t2 = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(t2)

BLOCK_RE = re.compile(r"^=====\s*(\d+)\s*=====\s*$", re.MULTILINE)
FENCE_RE = re.compile(r"^\s*```.*$", re.MULTILINE)


def parse_blocks(text: str) -> dict[str, str]:
    """{block number -> translated text}. Tolerates stray code fences."""
    text = FENCE_RE.sub("", text)
    marks = list(BLOCK_RE.finditer(text))
    out: dict[str, str] = {}
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        body = text[m.end():end].strip()
        # Normalise the blank-line runs a chat reply tends to introduce, so the
        # rejoined field matches what the API path would have produced.
        out[m.group(1)] = re.sub(r"\n{3,}", "\n\n", body)
    return out


def load_field_sources() -> dict[str, str]:
    """{field key -> original Czech text}, so is_valid sees the real source."""
    if not CATALOGUE.exists():
        return {}
    catalogue = json.loads(CATALOGUE.read_text(encoding="utf-8"))
    sources = {}
    for item in catalogue:
        for col in ("name_cz", "description_cz"):
            text = (item.get(col) or "").strip()
            if text:
                sources[t2.key_of(text)] = text
    return sources


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+", help="completed batch file(s)")
    ap.add_argument("--dry-run", action="store_true", help="validate only, write nothing")
    args = ap.parse_args()

    if not INDEX.exists():
        sys.exit(f"{INDEX.relative_to(ROOT)} not found. Run pipe/2a_export_batches.py first.")

    index = json.loads(INDEX.read_text(encoding="utf-8"))
    blocks = index["blocks"]
    cache = t2.load_cache()
    before = len(cache)
    # Snapshot first: a name is a single paragraph, so its block key and its
    # field key are the same string. Without this, importing a name would look
    # like the field was "already cached" rather than completed by this run.
    fields_before = {f["key"] for f in index["fields"] if f["key"] in cache}

    stats = Counter()
    rejects: list[tuple[str, str, str]] = []   # (number, reason, sample)
    unknown: list[str] = []
    seen: set[str] = set()

    for path in args.files:
        p = Path(path)
        if not p.exists():
            sys.exit(f"{p} not found")
        parsed = parse_blocks(p.read_text(encoding="utf-8"))
        print(f"{p.name}: {len(parsed)} blocks parsed")
        stats["files"] += 1

        for number, translation in parsed.items():
            rec = blocks.get(number)
            if rec is None:
                unknown.append(number)
                continue
            seen.add(number)
            ok, why = t2.is_valid(rec["source"], translation)
            if not ok:
                stats["rejected"] += 1
                rejects.append((number, why, translation[:70]))
                continue
            stats["accepted"] += 1
            if cache.get(rec["key"]) != translation:
                cache[rec["key"]] = translation
                stats["written"] += 1
            else:
                stats["unchanged"] += 1

    # Rejoin every field whose paragraphs are now all present.
    sources = load_field_sources()
    for field in index["fields"]:
        if field["key"] in cache:
            if field["key"] in fields_before:
                stats["fields_already"] += 1
            else:
                stats["fields_written"] += 1   # a name, completed by this import
            continue
        parts, complete = [], True
        for member in field["members"]:
            key = (blocks[member["block"]]["key"] if "block" in member
                   else member["cached_key"])
            value = cache.get(key)
            if value is None:
                complete = False
                break
            parts.append(value)
        if not complete:
            stats["fields_incomplete"] += 1
            continue

        joined = "\n\n".join(parts)
        source = sources.get(field["key"])
        if source is None:
            stats["fields_no_source"] += 1
            continue
        ok, why = t2.is_valid(source, joined)
        if not ok:
            stats["fields_rejected"] += 1
            rejects.append((f"field:{field['handle']}", why, joined[:70]))
            continue
        cache[field["key"]] = joined
        stats["fields_written"] += 1

    print()
    print(f"blocks accepted    : {stats['accepted']}"
          f"  ({stats['written']} new, {stats['unchanged']} unchanged)")
    print(f"blocks rejected    : {stats['rejected']}")
    print(f"fields completed   : {stats['fields_written']}"
          f"  ({stats['fields_already']} already cached,"
          f" {stats['fields_incomplete']} awaiting other batches)")
    if stats["fields_rejected"]:
        print(f"fields rejected    : {stats['fields_rejected']}")

    missing = [n for n in blocks if blocks[n]["batch"] in
               {blocks[s]["batch"] for s in seen} and n not in seen]
    if missing:
        print(f"\nMISSING from the batches you imported: {len(missing)} block(s)")
        print("  " + ", ".join(sorted(missing, key=int)[:20])
              + (" ..." if len(missing) > 20 else ""))
    if unknown:
        print(f"\nUNKNOWN block numbers (not in _index.json): {len(unknown)}")
        print("  " + ", ".join(sorted(set(unknown))[:20]))

    if rejects:
        print(f"\nREJECTED ({len(rejects)}) - not written to the cache:")
        for number, why, sample in rejects[:25]:
            print(f"  [{number}] {why}")
            print(f"      {sample}")
        if len(rejects) > 25:
            print(f"  ... and {len(rejects) - 25} more")
        print("\n  Re-translate these blocks and re-run this script; imports are"
              "\n  idempotent, so importing the whole batch again is safe.")

    if args.dry_run:
        print(f"\nDry run - cache untouched ({before} entries).")
        return 1 if rejects else 0

    t2.CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    t2.CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=1),
                             encoding="utf-8")
    print(f"\nCache: {before} -> {len(cache)} entries"
          f" ({t2.CACHE_FILE.relative_to(ROOT)})")
    return 1 if rejects else 0


if __name__ == "__main__":
    raise SystemExit(main())
