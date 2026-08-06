# Yoggies.be — Runbook

Everything you need to rebuild the site correctly, in order.

---

## Where things live

```
yoggies.be/
├── pipe/                       the data pipeline - run in numbered order
│   ├── 1_fetch_yoggies.py      pull all 526 products from yoggies.cz
│   ├── 2_translate.py          Czech -> Dutch, with verification
│   ├── 2a_export_batches.py    write untranslated text to numbered .md batches
│   ├── 2b_import_translations.py  read finished batches back into the cache
│   ├── 3_match_doglife.py      link products to DogLife purchase pages
│   ├── 4_build_master.py       assemble READY_FOR_BUILD_FINAL.xlsx
│   └── audit.py                check everything, run this constantly
│
├── data/                       all intermediate data lives here
│   ├── doglife/*.xlsx          YOUR DogLife exports - add new ones here
│   ├── cache/translations.json translation cache (do not delete, it saves money)
│   ├── translate_batches/      the manual-translation workspace (see below)
│   │   ├── _index.json         block -> product/paragraph map, written by 2a
│   │   ├── batch_001..023.md   Czech source, numbered blocks
│   │   └── done_001..023.md    your Dutch translations, same numbering
│   ├── yoggies_catalogue.json  written by step 1
│   ├── yoggies_translated.json written by step 2
│   ├── doglife_matches.json    written by step 3
│   ├── match_decisions.json    your y/n verdicts - always beat the scoring
│   ├── review_matches.xlsx     uncertain matches - you decide
│   └── doglife_unused.xlsx     DogLife products not on the site yet
│
├── READY_FOR_BUILD_FINAL.xlsx  the master file the site builder reads
├── READY_FOR_BUILD_NL.xlsx     build output of the site builder, not an input
├── yoggies_pipeline.py         the site builder (unchanged, still yours)
├── translation_cache_v2.json   the cache the SITE BUILDER reads - keep it
├── localize_images.py          downloads all images locally
├── image_map.json              remote image URL -> local file - keep it
├── missing_report.txt          written by every build; "0 missing" = good
├── atlas_nl_data.py            the 27 breed pages
├── atlas_raw/, atlas_scraped.json, scrape_atlas.py, translate_atlas.py
│                               how atlas_nl_data.py was produced; provenance only
├── templates/                  page templates (moved out of site/)
├── site/                       BUILD OUTPUT - never edit by hand
│   └── images/                 NOT VERSIONED - see below
└── legacy/                     superseded scripts, kept for reference only
```

Two translation caches exist and both are load-bearing:
`data/cache/translations.json` is what the *pipeline* (`pipe/2*`) reads and
writes, `translation_cache_v2.json` is what the *site builder* reads. The
root-level `translation_cache.json` is the superseded v1 and is only referenced
by `legacy/build_site.py`.

**One rule: never edit anything inside `site/`.** It is regenerated on every
build and your changes will be erased. Fix the generator or the data instead.

**`site/images/` is deliberately not versioned.** The 6,096 files there are
build output: `localize_images.py` rebuilds every one of them from
`image_map.json`, which *is* in git. Keeping 157.7 MB of binaries in history
would bloat every clone to buy nothing. After a fresh clone, run
`python localize_images.py` before `python yoggies_pipeline.py` — downloads
resume where they stopped, so an interrupted run is safe to repeat.

Only 1,653 of those files are referenced by a page. The other 4,443 (134.8 MB)
are originals and `-200/-600/-1200.webp` derivatives belonging to products
whose page was merged away or excluded. Harmless, and they cost nothing but
disk — but if upload size matters, they are safe to prune: every page reference
was checked against the folder and nothing is missing.

---

## First-time setup

```powershell
pip install requests pillow openpyxl openai
```

An API key is only needed if you translate through the API. The manual route
below needs no key and no budget — see "Translating without an API key".

---

## The full rebuild, in order

### 1. Get the catalogue
```powershell
python pipe/1_fetch_yoggies.py
```
Pulls all products from yoggies.cz through Shopify's own JSON endpoint —
complete Czech descriptions, every image, every size variant, product types
and tags. Expect around **526 products**.

### 2. Translate
```powershell
python pipe/2_translate.py
```
Czech to Dutch, then **verifies every result**. Anything containing Czech
diacritics, a `[CZ]` marker, or a suspicious length ratio is rejected and
retried rather than cached. Failures land in `data/translation_failures.json`.

The script insists on `OPENAI_API_KEY` even when it will not call the API.
When everything is already in the cache — which is the case today — any
non-empty value works and no request is made:

```powershell
$env:OPENAI_API_KEY="cache-only"; python pipe/2_translate.py
```

A run that reports `translated : 0` made zero API calls and cost nothing.

### 3. Match to DogLife
```powershell
python pipe/3_match_doglife.py
```
Then open `data/review_matches.xlsx`, put `y` or `n` in the **keep** column,
and lock your decisions in:
```powershell
python pipe/3_match_doglife.py --apply-review
```
Your decisions are remembered in `data/match_decisions.json` and always beat
the automatic scoring, so you never have to review the same pair twice.

To add more DogLife categories later: drop the new `.xlsx` into
`data/doglife/` and re-run. Files are merged and de-duplicated by URL.

### 4. Build the master file
```powershell
python pipe/4_build_master.py --dry-run     # look first
python pipe/4_build_master.py               # then write
```
Ingredients, dosages and your hand-built calculator flags are read out of the
existing master and carried across. The old file is backed up with a timestamp
before it is replaced — **delete those backups periodically**, git already
holds the history.

### 5. Images and site
```powershell
python localize_images.py
python yoggies_pipeline.py
```

### 6. Check
```powershell
python pipe/audit.py
```
Run this after every step. `--verbose` lists the offending items.

---

## Translating without an API key

There is no API budget, so all 2,566 text blocks were translated by hand
through a chat interface. The two scripts that make that safe are
`2a_export_batches.py` and `2b_import_translations.py`. Use the same loop for
any new product text.

**That work is finished.** All 23 batches are exported and all 23 `done_` files
are imported; `data/cache/translations.json` holds 4,323 entries and a build
today reports `missing=0`. You only re-enter this loop when new Czech text
arrives from yoggies.cz.

**1. Export what is missing.**
```powershell
python pipe/2a_export_batches.py
```
Splits every untranslated name and description into paragraphs, de-duplicates
them (shared boilerplate is exported once, not once per product) and writes
them to `data/translate_batches/batch_001.md`, `batch_002.md`, … Each file
opens with the translator brief and then lists numbered blocks:

```
===== 157 =====
Yoggies jogurt pro psy a kočky 150 g
```

`_index.json` records which block belongs to which product and paragraph. Do
not edit it.

**2. Translate a batch.** Reply in exactly the same format — same separator,
same number, Dutch underneath — and save it as `done_001.md` next to the
batch. Skip no numbers and add none.

**3. Import it back.**
```powershell
python pipe/2b_import_translations.py data/translate_batches/done_001.md
python pipe/2b_import_translations.py data/translate_batches/done_*.md
python pipe/2b_import_translations.py data/translate_batches/done_001.md --dry-run
```
Every block is checked with the same `is_valid()` the API path uses, so a
translation that still contains Czech is rejected here exactly as it would be
there. Rejected blocks are listed with the reason and are **not** written.
Fix them in the `done_` file and re-run — imports are idempotent, so
re-importing a whole batch is safe.

A field (a full description) is only written once every one of its paragraphs
has been imported, so a description split across several batches simply
completes when the last one lands.

### What the validator rejects

- `[CZ]` / `[NL]` markers
- the Czech letters `ě š č ř ů` outside `Yoggies`, `Ibérico`, `B.A.R.F.`,
  `Markus-Mühle`
- Czech stop-words: `jsou, které, pro psy, pro kočky, granule, krmivo,
  balení, masem, lisované, obsahuje, díky, nebo, jako, velmi, více, Kč`
- a translation shorter than 0.45× or longer than 2.5× its source
- a chatty preamble ("Hier is de vertaling…")

### House style

- natural commercial Dutch for a Belgian audience
- **the site addresses the reader as "u" / "uw"**, not "je" / "jouw"
  (2,711 vs 26 occurrences — match it)
- keep `Yoggies`, `BARF`, `VET Gastro Sensitive`, `Ibérico` exactly
- **brokken**, never *granulaat*; **MINIbrokken**, never *MINIgranule*
- keep every number and weight; drop the `Kč` symbol but keep the amount
- CSS blocks pass through unchanged apart from their Czech comments

### Writing a description by hand

Five products have no description on yoggies.cz at all (`body_html` is null),
so there is nothing to translate. Hand-written Dutch goes in the
`description_nl` column of `READY_FOR_BUILD_FINAL.xlsx`. `4_build_master.py`
resolves that field as **fresh translation → the previous master → Czech**, so
anything you type there is picked up as "preserved" and survives every future
rebuild.

---

## Current state — what audit.py reports today

```
21 passed   2 warnings   0 failed
```

| | |
|---|---|
| Products | 526 |
| Pages built | 647 html files — 471 real pages + 176 redirect stubs |
| | 426 product pages, 27 breed pages, 9 category pages, 4 filter pages, 5 standalone (index, calculator, hondenatlas, over-yoggies, Dierenkliniek) |
| | stubs: 152 retired size variants, 23 Czech breed URLs, 1 `home.html` |
| Images | 6,096 local files (157.7 MB), 0 missing, 0 hotlinked |
| Site folder | 193.5 MB total, 35.9 MB without `images/` |
| Translation | 2,566 blocks / 1,042 fields, 0 Czech left in names or descriptions |
| DogLife | 213 linked, 202 information-only, 111 awaiting review |

Every data and site check passes: no Czech text, no Kč, no URL collisions, no
broken links, no missing meta descriptions or canonicals, no missing alt text,
no empty descriptions, every referenced image present on disk.

`sitemap.xml` lists the 471 real pages and leaves every redirect stub out.

### Two warnings, both decisions rather than defects

**1. 111 uncertain DogLife matches** are waiting in
`data/review_matches.xlsx`. Put `y`/`n` in the **keep** column and run
`python pipe/3_match_doglife.py --apply-review`. Decisions already taken live
in `data/match_decisions.json` and are never asked again.

**2. 417 product pages show no call to action.** This is the old question of
what the 202 products DogLife does not sell should display. Since yoggies.be
is an information site and DogLife is the shop, showing no purchase block at
all is cleaner than an apology — but that is your call, and the fix lives in
the purchase block in `yoggies_pipeline.py` (around line 1550).

### Known cosmetic issues

**Four merged size groups share a page with an identically named product.**
`_merge_size_variants()` renames a group to its base name ("Yoggies B.A.R.F.
100% kip met probiotica"), and for four groups yoggies.cz *also* lists a
separate sizeless product under exactly that name. Both rows resolve to the
same slug, so 430 product rows write 426 files and the second row overwrites
the first. No link breaks — every reference goes through the same slug table —
but the search dropdown lists those four names twice pointing at one page, and
the losing row's own description never appears. The fix is in `_build_html()`
(`yoggies_pipeline.py` ~line 1360): key the slug de-duplication on the row, not
on `name_nl`. Left alone deliberately — changing slugs at hand-over would move
live URLs for no functional gain.

**Register slip on eight product pages.**
Four paragraphs were already in the translation cache before the manual pass
and were therefore never exported for translation. They use "je" where the
rest of the site uses "u", which is visible as a register change partway down
eight product pages (both Chlorella sizes, both Mineralen sizes, three RAGÚ
kalkoen/zalm variants, the rice-flake bijgerecht). Harmless, but worth
retranslating when convenient.

---

## How to deploy

**`site/` is the deliverable.** It is a plain static site — HTML, CSS, JS,
images, `sitemap.xml`, `robots.txt` — with no build step, no server code and no
database. Upload the *contents* of `site/` to the web root of any static host
(Netlify, Cloudflare Pages, GitHub Pages, S3 + CloudFront, or plain FTP to
shared hosting) and it works as-is. `index.html` is the front page; every
internal link is relative, so the site also runs from a subdirectory or
straight off disk.

```powershell
# what you upload — 193.5 MB, 6,755 files
site/
```

Two things to know before going live:

- **Absolute URLs assume `https://yoggies.be/`.** Canonical tags, `og:image`
  and `sitemap.xml` are written against that domain. Deploy somewhere else and
  those tags point at yoggies.be while the pages themselves still work. The
  domain is `BASE_URL` at `yoggies_pipeline.py:180` — change it and rebuild.
- **Redirect stubs are `<meta http-equiv="refresh">`, not HTTP 301s.** They
  work on any host with no configuration. If your host supports real redirects,
  the 176 stubs are the list to convert.

Nothing else needs deploying. `pipe/`, `data/`, the `.xlsx` files and the
generator scripts are the workshop, not the product.

---

## Gotchas

**Console output is UTF-8; a cp1252 terminal kills the build.** On Windows the
builder prints `→` and `—`, and in a console that is not UTF-8 the run dies
with `UnicodeEncodeError: 'charmap' codec can't encode character '→'`
partway through — after the product pages, before the sitemap. Set the encoding
first:

```powershell
$env:PYTHONIOENCODING="utf-8"; python yoggies_pipeline.py
```

**New DogLife matches bring new images.** A match links the DogLife product
photo, which is not yet in `image_map.json`, so the first build after a
successful `3_match_doglife.py` ends with `Image sweep: N refs still external`.
That is not an error — run `python localize_images.py` and build again. The
sweep line must read clean before you ship.

**The build will not finish over a slow network or synced folder.** Run from
a local disk. Against a mounted Downloads folder the build was killed partway
through, consistently after the Dierenkliniek step and with no error message;
on local disk the same build finishes in about 8 seconds. If you must keep the
project on a synced folder, copy it to a local scratch directory, build there,
and copy `site/` back.

**`audit.py` must see `site/images/`.** If you build in a scratch directory
without the images, the link checker reports ~1,220 broken paths that are not
real. Symlink or copy the image folder before trusting that number.

**Czech personal names are whitelisted individually** in `audit.py`, not by
allowing Czech letters generally — so the check still catches genuinely
untranslated text. If you credit another Czech name on the site, add it to
`ALLOWED` by name.

---

## If something goes wrong

Every step is safe to re-run. The master file is backed up before it is
overwritten, translations are cached, image downloads resume where they
stopped, and batch imports are idempotent. If a build looks wrong, restore the
most recent `READY_FOR_BUILD_FINAL.backup-*.xlsx` (or `git checkout` the
master) and run `pipe/audit.py` to see what changed.
