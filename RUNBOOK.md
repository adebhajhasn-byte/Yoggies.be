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
│   ├── doglife_unused.xlsx     DogLife products not on the site yet
│   └── build_manifest.json     what the last build published; audit.py reads it
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

**`site/images/` is deliberately not versioned.** The 6,104 files there are
build output: `localize_images.py` rebuilds every one of them from
`image_map.json`, which *is* in git. Keeping 157.8 MB of binaries in history
would bloat every clone to buy nothing. After a fresh clone, run
`python localize_images.py` before `python yoggies_pipeline.py` — downloads
resume where they stopped, so an interrupted run is safe to repeat.

Only 1,662 of those files are referenced by a page. The other 4,442 (128.6 MB)
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

`data/yoggies_raw.json` keeps the API response verbatim, so extraction can be
redone without touching the network:

```powershell
python pipe/1_fetch_yoggies.py --from-raw
```

Use that after any change to `html_to_text()` — it rebuilds
`data/yoggies_catalogue.json` from the saved response and downloads nothing.

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
are imported; `data/cache/translations.json` holds 4,367 entries and a build
today reports `missing=0`. You only re-enter this loop when new Czech text
arrives from yoggies.cz.

(4,323 of those came from the batches; the other 44 are the CSS-cleaned
descriptions re-keyed in place — see "The stylesheet leak" below.)

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
- **no CSS or JavaScript should reach you at all.** It used to: `html_to_text()`
  stripped `<style>` tags but kept the stylesheet between them, so 111
  descriptions arrived as `.yoggies-section { max-width: 1100px; … }` and were
  dutifully "translated" comment by comment. The extractor now removes
  `<style>`, `<script>`, `<noscript>` and HTML comments element *and* contents,
  and `audit.py` FAILs if any description or page still shows code as text. If a
  batch ever hands you a stylesheet again, that is a bug upstream — say so
  rather than translating it.

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
25 passed   1 warnings   0 failed
```

| | |
|---|---|
| Products | 526 rows → 430 product pages (56 size groups merged) |
| Pages built | 657 html files — 475 real pages + 182 redirect stubs |
| | 430 product pages, 27 breed pages, 9 category pages, 4 filter pages, 5 standalone (index, calculator, hondenatlas, over-yoggies, Dierenkliniek) |
| | stubs: 152 retired size variants, 23 Czech breed URLs, 6 disambiguated group URLs, 1 `home.html` |
| Images | 6,104 local files (157.8 MB), 0 missing, 0 hotlinked |
| Site folder | 194.0 MB total, 36.2 MB without `images/` |
| Translation | 2,566 blocks / 1,042 fields, 0 Czech left in names or descriptions |
| DogLife | 216 linked, 310 information-only, 0 awaiting review |

Every data and site check passes: no Czech text, no Kč, no URL collisions, no
broken links, no missing meta descriptions or canonicals, no missing alt text,
no empty descriptions, no CSS or JavaScript rendered as prose, every referenced
image present on disk, and every product row has a page of its own.

`sitemap.xml` lists the 475 real pages and leaves every redirect stub out.

### The stylesheet leak, and how the Dutch survived it

`html_to_text()` flattened `<style>` by stripping the tag and keeping what was
between it and `</style>`. **111 of 526 descriptions** therefore began with the
product page's stylesheet, and 64 built pages rendered it as body copy — the
first card on `voor-katten.html` opened with `.yoggies-section { max-width:
1100px; … }`. The extractor now drops `<style>`, `<script>`, `<noscript>` and
HTML comments with their contents before any tag flattening happens.

Re-extracting the Czech was free (`--from-raw`). The Dutch was the problem: the
translation cache is keyed on `sha256` of the **whole** description, so cleaning
the Czech invalidated all 111 keys and a re-translation would have meant 111 API
calls against a budget that does not exist. It was recovered instead, because
`description_cz` and `description_nl` have **identical non-blank line counts**
for every affected product (rule 4 of the translator brief — preserve the
paragraph layout). The CSS lines were located in the Czech and dropped at the
same line positions in the Dutch; each result was checked three ways (the kept
Czech lines must reproduce the cleaned Czech exactly, the Dutch must contain no
code afterwards, and the line counts must still agree) before it was written
back to both caches. All 111 passed, `2_translate.py` then reported
`translated : 0`, and no translation work was lost or repaid for.

That line-for-line correspondence is worth remembering: it is what makes any
future surgery on the Czech recoverable on the Dutch side without paying for it
twice.

### The one remaining warning

**419 product pages show no call to action.** This is the old question of
what the 310 products DogLife does not sell should display. Since yoggies.be
is an information site and DogLife is the shop, showing no purchase block at
all is cleaner than an apology — but that is your call, and the fix lives in
the purchase block in `yoggies_pipeline.py` (around line 1550).

### The DogLife review is closed

All 111 uncertain pairs were judged and stored in `data/match_decisions.json`;
`review_matches.xlsx` is empty and the audit check passes. **10 accepted,
101 rejected.** The rule applied was: accept only when pack size *and*
flavour/variant both match, checking the DogLife product page itself whenever
the two titles alone were not conclusive.

What the rejections were, in bulk: a different recipe under a similar name
(VET Gastro Sensitive offered for lamsvlees/Veggie/Insect Derma), a different
brand (LIZA, Fitmin, TRIXIE, Happy Dog), a different colour or design in the
SodaPup likmat range (the DogLife pages carry no colour selector — green is
its own product), MINI kibble pointed at the classic-size page (DogLife has
*separate* `…-mini` pages for Ibérico and Kalkoen/Gierst), and every "5+1"
multipack, because DogLife sells those cans singly.

Two traps worth remembering: **LUPO BIOTIN** looks like a certain match on
title but DogLife sells 400 g/800 g while ours is 180 g, and **Yoggies ACTIVE
vlees van wild** is a €2,70 can, not the 2–15 kg kibble whose name it echoes.

**One wrong link was found outside the review and corrected the same way.**
Four "kalkoen- en kalfsvlees met appel" rows were auto-matched — above the
0.70 threshold, so they never reached the review file — to DogLife's *kalkoen
met bataten*, a different recipe. The 400 g row now points at the real
"kalkoen en kalfvlees met appels" page; the 800 g row and both 5+1 multipacks
have no link, because DogLife does not sell those. Worth knowing that a high
score is not proof: if you spot another, put the right URL (or `null`) in
`data/match_decisions.json` and re-run the matcher.

**The four "5+1" multipacks that pointed at a single can are now unlinked.**
Two eend (400 g and 800 g, both aimed at the same sizeless eend page), one rund
mariadistel 400 g and one kangoeroe 800 g. A multipack that sends the buyer to
a single tin is wrong however right the flavour is, so all four are `null` in
`data/match_decisions.json` and are information-only, like the other 40 bundles.
No "5+1", "6+1", "voordeelverpakking", "24 x 185 g" or Proefbox product carries
a purchase link now — DogLife sells none of them.

Multipack→multipack matches were left alone and are correct: the 12x150 g
B.A.R.F. menus, the 6x250 g B.A.R.F. tubs and Fit blokjes 2x47 g all point at
the matching DogLife pack size.

**Still open, and not a pack-size problem:** rows 85 and 86 — "gedroogde
kauwsnacks – Konijnenoren met vacht 4 st." and "– Bullenpees 5 st." — are both
matched to `fit-beentje-voor-honden-118-g`, which is neither rabbit ears nor
bull pizzle. That looks like two wrong matches to a different product, but it
is a judgement call about what DogLife actually stocks, so it was left for you.

Decisions always beat the scoring, so re-running the matcher never re-opens
them. If you disagree with one, edit the `keep` column and re-apply — or edit
`data/match_decisions.json` directly.

### How merged groups are named

`_merge_size_variants()` gives a merged group the group's base name, e.g.
"Yoggies B.A.R.F. 100% kip met probiotica". yoggies.cz also sells some of those
products as one sizeless *combined listing* under exactly that name, and the
two used to slugify to the same file: the group overwrote the standalone and
four products had no page at all.

Now, when a group's base name is already taken by a row that survives merging,
the group gets a suffix naming the packs it covers — "… - 700 g en 1300 g" —
the way the yoghurt multipack carries "6 of 12 stuks". **The standalone listing
keeps the plain URL**; the group moves to the suffixed one, and its old
`<base>-1.html` URL stays alive as a redirect stub. Six groups are affected:
three B.A.R.F. menus, the kalkoen/kalfsvlees blikvoer and two gekookt menus.

`audit.py` now compares `data/build_manifest.json` (what the build says it
wrote) against the real pages on disk, so a slug that swallows a product page
fails the audit instead of passing unnoticed.

### Known cosmetic issue

**Register slip on eight product pages.**
Four paragraphs were already in the translation cache before the manual pass
and were therefore never exported for translation. They use "je" where the
rest of the site uses "u", which is visible as a register change partway down
eight product pages (both Chlorella sizes, both Mineralen sizes, three RAGÚ
kalkoen/zalm variants, the rice-flake bijgerecht). Harmless, but worth
retranslating when convenient.

### The floating contact widget

A round green button sits bottom-right on all **475 real pages** (redirect stubs
are skipped — they are gone before anyone could click). It opens a panel with
Naam, a topic dropdown, a message box and two buttons: WhatsApp and E-mail.
Both are client-side only — no backend, no database, no third party. WhatsApp
opens `wa.me/<digits>?text=…`, E-mail opens a `mailto:` with the topic as
subject. The page title and URL are appended to every message, so you can see
what the visitor was looking at.

**It does not store the phone number or the e-mail address.** The script reads
them out of the page's own `<footer>` at runtime — the `tel:` and `mailto:`
hrefs — so the footer stays the single source of truth. Change the footer and
the widget follows without touching `yoggies_pipeline.py`.

It reads the `tel:` **href**, not the visible text, on purpose: `+32 (0)475 34
48 68` strips to `320475344868`, and that trunk zero is not dialable. The href
`tel:+32475344868` gives the `32475344868` wa.me needs. `(0)` is stripped
defensively in case the pretty form ever lands in the href.

Lives in `_CONTACT_WIDGET` / `_inject_contact_widget()` next to
`_content_fixups`, injected by the post-build sweep so every page type gets it
in one place. Injection removes any existing copy first, including its trailing
newline — the static pages (index, calculator, hondenatlas, Dierenkliniek) are
edited in place rather than regenerated, so without that it would stack a
little more onto every page on every rebuild.

The button hides itself whenever the footer or a `.ph-order-btn` scrolls into
view, so it never sits on top of the footer links or the Bestel button on a
phone. Escape and an outside click close it; it is a non-modal popover, so
tabbing out closes it rather than trapping focus.

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
# what you upload — 194.0 MB, 6,773 files
site/
```

Two things to know before going live:

- **Absolute URLs assume `https://yoggies.be/`.** Canonical tags, `og:image`
  and `sitemap.xml` are written against that domain. Deploy somewhere else and
  those tags point at yoggies.be while the pages themselves still work. The
  domain is `BASE_URL` at `yoggies_pipeline.py:180` — change it and rebuild.
- **Redirect stubs are `<meta http-equiv="refresh">`, not HTTP 301s.** They
  work on any host with no configuration. If your host supports real redirects,
  the 182 stubs are the list to convert.

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
