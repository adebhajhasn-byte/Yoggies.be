"""
2a_export_batches.py — paste-ready translation batches
======================================================
The API route (2_translate.py) is not the only way to fill the cache. This
script exports everything still needing translation as Markdown files sized for
a chat interface; 2b_import_translations.py reads the completed files back.

Deduplication is at PARAGRAPH level, not per description. Dozens of products
repeat the same boilerplate (the cold-pressed kibble paragraph, the storage
advice, the BARF handling note); translating each one once is a large saving
over translating every description whole.

Cache keys are unchanged: 2_translate.py looks up sha256(full_field_text), so
the importer reassembles a description from its translated paragraphs and
stores it under the whole-field key. Paragraph-level keys are stored too, which
is what makes the repeated boilerplate free the second time it appears.

Output (data/translate_batches/)
    batch_001.md ...   the Dutch instructions + numbered Czech blocks
    _index.json        block number -> {key, source, chars, batch, kind}

Usage
    python pipe/2a_export_batches.py
    python pipe/2a_export_batches.py --max-chars 15000
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CATALOGUE = ROOT / "data" / "yoggies_catalogue.json"
OUT_DIR = ROOT / "data" / "translate_batches"
MAX_CHARS = 15000

# Load 2_translate.py for key_of/load_cache/SYSTEM so the two stay in lockstep.
# The leading digit makes it a non-identifier, so it cannot be imported by name.
_spec = importlib.util.spec_from_file_location("t2", Path(__file__).with_name("2_translate.py"))
t2 = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(t2)

# Appended after the SYSTEM prompt so the reply comes back parseable.
FORMAT_RULES = """\

## Formaat van je antwoord

Hieronder staan genummerde blokken Tsjechische tekst. Vertaal ze allemaal.

Geef je antwoord in exact hetzelfde formaat terug: dezelfde scheidingslijn met
hetzelfde nummer, daaronder uitsluitend de Nederlandse vertaling.

    ===== 001 =====
    <Nederlandse vertaling>

    ===== 002 =====
    <Nederlandse vertaling>

Sla geen enkel nummer over en voeg geen nummers toe. Zet geen inleiding,
samenvatting of commentaar in je antwoord - alleen de genummerde blokken.
"""


def paragraphs(text: str) -> list[str]:
    """Split a description the same way split_text() finds chunk boundaries."""
    return [p.strip() for p in text.split("\n\n") if p.strip()]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--max-chars", type=int, default=MAX_CHARS,
                    help=f"max Czech characters of payload per batch (default {MAX_CHARS})")
    args = ap.parse_args()

    if not CATALOGUE.exists():
        sys.exit(f"{CATALOGUE.name} not found. Run pipe/1_fetch_yoggies.py first.")

    catalogue = json.loads(CATALOGUE.read_text(encoding="utf-8"))
    cache = t2.load_cache()

    stats = Counter()
    fields: list[dict] = []          # one per name_cz / description_cz still needed
    blocks: dict[str, dict] = {}     # unique block text -> record (dedup happens here)
    order: list[str] = []            # block texts, names first then paragraphs

    for item in catalogue:
        for kind, col in (("name", "name_cz"), ("description", "description_cz")):
            text = (item.get(col) or "").strip()
            if not text:
                continue
            stats["field_values"] += 1
            stats["chars_raw"] += len(text)

            field_key = t2.key_of(text)
            if field_key in cache:
                stats["fields_cached"] += 1
                continue
            stats["fields_todo"] += 1

            parts = [text] if kind == "name" else paragraphs(text)
            member_texts = []
            for part in parts:
                if t2.key_of(part) in cache:
                    stats["paragraphs_cached"] += 1
                    member_texts.append(part)
                    continue
                if part not in blocks:
                    blocks[part] = {"kind": kind, "chars": len(part)}
                    order.append(part)
                    stats["chars_deduped"] += len(part)
                else:
                    stats["paragraphs_duplicate"] += 1
                member_texts.append(part)

            fields.append({
                "key": field_key,
                "kind": kind,
                "handle": item.get("handle", ""),
                "chars": len(text),
                "parts": member_texts,
            })

    # Names first - they pack far more densely than descriptions.
    order.sort(key=lambda t: 0 if blocks[t]["kind"] == "name" else 1)

    # Assign global block numbers in emission order.
    for n, text in enumerate(order, 1):
        blocks[text]["number"] = f"{n:03d}"

    # Pack into batches. Names and paragraphs never share a file, and a block is
    # never split across two files - an oversized block gets a batch to itself.
    batches: list[list[str]] = []
    current: list[str] = []
    current_chars = 0
    current_kind = None
    for text in order:
        kind = blocks[text]["kind"]
        size = len(text)
        too_big = current and (current_chars + size > args.max_chars)
        if current and (kind != current_kind or too_big):
            batches.append(current)
            current, current_chars = [], 0
        current.append(text)
        current_chars += size
        current_kind = kind
        if size > args.max_chars:
            stats["oversized_blocks"] += 1
    if current:
        batches.append(current)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for stale in OUT_DIR.glob("batch_*.md"):
        stale.unlink()

    index_blocks: dict[str, dict] = {}
    for i, batch in enumerate(batches, 1):
        name = f"batch_{i:03d}.md"
        body = [t2.SYSTEM, FORMAT_RULES]
        for text in batch:
            rec = blocks[text]
            body.append(f"\n===== {rec['number']} =====\n{text}\n")
            index_blocks[rec["number"]] = {
                "key": t2.key_of(text),
                "kind": rec["kind"],
                "chars": rec["chars"],
                "batch": name,
                "source": text,
            }
        (OUT_DIR / name).write_text("\n".join(body), encoding="utf-8")

    # Map each field to the block numbers it reassembles from. A part already in
    # the cache has no block, so it is recorded by key for the importer to read.
    index_fields = []
    for f in fields:
        members = []
        for part in f["parts"]:
            rec = blocks.get(part)
            members.append({"block": rec["number"]} if rec and "number" in rec
                           else {"cached_key": t2.key_of(part)})
        index_fields.append({
            "key": f["key"], "kind": f["kind"], "handle": f["handle"],
            "chars": f["chars"], "members": members,
        })

    (OUT_DIR / "_index.json").write_text(json.dumps({
        "block_count": len(index_blocks),
        "batch_count": len(batches),
        "max_chars": args.max_chars,
        "blocks": index_blocks,
        "fields": index_fields,
    }, ensure_ascii=False, indent=1), encoding="utf-8")

    saved = stats["chars_raw"] - stats["chars_deduped"]
    print(f"catalogue          : {len(catalogue)} products")
    print(f"cache entries      : {len(cache)}")
    print()
    print(f"field values       : {stats['field_values']}"
          f"  ({stats['fields_cached']} cached, {stats['fields_todo']} to translate)")
    print(f"unique blocks       : {len(index_blocks)}"
          f"  ({stats['paragraphs_duplicate']} duplicate paragraphs collapsed)")
    print()
    print(f"chars, every field  : {stats['chars_raw']:>9,}")
    print(f"chars after dedup   : {stats['chars_deduped']:>9,}")
    print(f"saved               : {saved:>9,}"
          f"  ({saved / max(1, stats['chars_raw']):.1%})")
    print()
    print(f"batches written     : {len(batches)} -> {OUT_DIR.relative_to(ROOT)}")
    if stats["oversized_blocks"]:
        print(f"  note: {stats['oversized_blocks']} block(s) exceed --max-chars on their "
              f"own and were given a batch each (paragraphs are never split)")
    print("\nNext: paste each batch into the chat, save the reply, then")
    print("      python pipe/2b_import_translations.py <completed-file>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
