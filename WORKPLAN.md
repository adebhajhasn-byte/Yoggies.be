# Yoggies.be — Remediation Work Order

Hand this file to Claude Code in the project root. Work top to bottom; each
phase has acceptance criteria you can verify before moving on.

## Ground rules

1. **`yoggies_pipeline.py` is the only builder.** It reads
   `READY_FOR_BUILD_FINAL.xlsx`, writes `READY_FOR_BUILD_NL.xlsx`, and generates
   everything under `site/`. `build_site.py`, `build_site_final.py` and
   `verify_and_build.py` are older forks of the same logic — do not edit them,
   archive them (Phase 0).
2. **Never patch files under `site/`.** They are build output and every run
   overwrites them. All fixes go into the generator or into the Excel.
3. **Commit before starting.** `git init` if the project is not under version
   control yet — several phases rewrite hundreds of files.
4. After each phase, run the build and re-check, so a regression is traceable to
   one phase.

Baseline numbers from the audit (use them to confirm fixes landed):

| Metric | Before |
|---|---|
| Product pages | 352 |
| Product pages with a working buy link | 114 |
| Product pages showing "niet beschikbaar bij DogLife" | 213 |
| Pages showing a buy link *and* "niet beschikbaar" | 18 |
| Pages with neither | 7 |
| External image references | 40,035 (34,981 yoggies.cz / 5,054 doglife.be) |
| Unique external image files | 640 |
| Pages with a meta description | 0 / 400 |
| `<img>` tags without `alt` | 32,763 / 35,183 |
| Product pages reachable by a crawler | 76 (rest are `onclick` only) |

---

## Phase 0 — Cleanup

- [ ] Add `.gitignore` with `.venv/`, `__pycache__/`, `.idea/`, `*.pyc`.
      `.venv` is 9,110 files and must never be committed or zipped again.
- [ ] `mkdir legacy/` and move `build_site.py`, `build_site_final.py`,
      `verify_and_build.py`, `merge_calculator.py`, `merge_extra_pages.py`,
      `fix_home_categories.py`, `replace_badge_icons.py`,
      `replace_product_calculator.py` into it. Add `legacy/README.md` saying
      these are superseded by `yoggies_pipeline.py`.
- [ ] Replace the hardcoded Windows paths in the moved scripts with
      `Path(__file__).resolve().parent.parent`. They currently read
      `Path(r'C:\Users\AdibHAJHASAN\IdeaProjects\yoggies.be\site')` and cannot
      run on any other machine.
- [ ] Move `site/product.html` and `site/hoofdpageproduct.html` (plus their
      `.css`/`.js`) into `templates/`. They are source templates sitting in the
      output folder, and their `../home.css` links are broken where they are.
      Update `SITE_DIR / "product.html"` references in `yoggies_pipeline.py`
      (around lines 904 and 991) to the new location.
- [ ] Stop writing `site/home.html`. It is a byte-identical copy of
      `site/index.html` and creates duplicate content. Keep `index.html` only;
      emit `home.html` as a redirect stub if any old link depends on it.
- [ ] Fix the diagnostics encoding: `diagnose_output.txt` is written as UTF-16
      with cp850 mojibake (`patÚ`, `monoprote´ne`, `û`). Write it with
      `encoding="utf-8"` explicitly and set `PYTHONIOENCODING=utf-8` in the run
      config instead of relying on the console codepage.

**Accept when:** `python yoggies_pipeline.py` still builds a complete site, and
`git status` shows no `.venv` or `__pycache__` entries.

---

## Phase 1 — Images (do this before Phase 2; it changes what the templates emit)

`localize_images.py` is provided. It reads every image URL out of
`READY_FOR_BUILD_FINAL.xlsx` plus `atlas_nl_data.py` and the templates,
downloads all **640** unique files into `site/images/`, and generates WebP
variants at 200 / 600 / 1200 px. It writes `image_map.json` mapping each
original URL to its local paths, and is safe to re-run.

- [ ] `pip install requests pillow openpyxl`
- [ ] `python localize_images.py` (expect roughly 130 MB downloaded, a few
      minutes). Check `image_failures.txt`; re-run to retry failures.
- [ ] Teach `yoggies_pipeline.py` to resolve every image through
      `image_map.json`. Add one helper near the top:

  ```python
  _IMG_MAP = json.loads(Path("image_map.json").read_text(encoding="utf-8")) \
      if Path("image_map.json").exists() else {}

  def img_src(url, size="md", depth=""):
      """Return a local path for an external image URL, falling back to the URL."""
      entry = _IMG_MAP.get(url) or _IMG_MAP.get(url.replace("&amp;", "&"))
      if not entry:
          return url
      return depth + entry.get(size, entry.get("orig", url))
  ```

  Then route **every** `<img src=...>` the pipeline emits through it. The call
  sites are the product page builder (~line 904+), the category card builder
  (line 1401), the nav dropdown builder (~line 1472), the breed page builder
  (~line 870) and the homepage builder (~line 1064).
- [ ] Use the right size at each call site. This is where the performance win
      is: **nav dropdown and category cards must use `size="thumb"` (200 px)**,
      not the current 1200 px originals.
- [ ] Add `srcset` on the product hero only:
      `srcset="{md} 600w, {lg} 1200w" sizes="(max-width: 700px) 100vw, 600px"`.
- [ ] Add `loading="lazy"` and `decoding="async"` to every non-hero image.
- [ ] Add a real `alt` everywhere. For product images use the product name; for
      badge icons use the icon label from `icons_text`. Never emit an empty
      `alt` on a content image.

**Accept when:** `grep -r 'src="http' site/ | grep -v fonts.googleapis` returns
nothing, and a product page requests no image wider than 200 px for the nav.

---

## Phase 2 — Generator fixes (SEO, markup, i18n)

All of these are edits inside `yoggies_pipeline.py`.

- [ ] **Line 1401 — make cards crawlable.** Replace
      `<div class="pg-card" onclick="window.location='products/{slug}.html'">`
      with a real anchor:
      `<a class="pg-card" href="{depth}products/{slug}.html">` … `</a>`.
      Add `.pg-card { display:block; text-decoration:none; color:inherit; }` to
      the category CSS. This is what makes 267 orphan product pages reachable by
      Google and by keyboard users.
- [ ] **`<head>` for every page type.** Currently 0 of 400 pages have any of
      these. Add to the product, category, breed, home and static page builders:
      - `<meta name="description" content="...">` — first 155 characters of the
        product/breed description, HTML-escaped, cut on a word boundary.
      - `<link rel="canonical" href="https://yoggies.be/{path}">`
      - `og:title`, `og:description`, `og:image` (absolute URL), `og:type`,
        `og:url`, plus `twitter:card=summary_large_image`.
      - `<link rel="icon" href="{depth}images/favicon.svg">` (add the asset).
- [ ] **Truncation on word boundaries.** Category card descriptions currently
      cut mid-word (`milieu`, `sint-janskrui`). Cut at the last space before the
      limit and append `…`.
- [ ] **Currency.** Four gift-card products are still priced in Czech crowns
      (`cadeaubon-ter-waarde-van-500-kc`). Convert the names, slugs and body
      text to euro, or drop them from the Belgian catalogue — they are not
      redeemable at DogLife anyway.
- [ ] **Untranslated Czech.** `granule` still appears in product body copy where
      it should be `brokken` (see `yoggies-brokken-kip-rund`,
      `yoggies-brokken-lam-witvis`). Add `granule → brokken`,
      `Granule → Brokken` to the translation overrides and rebuild; then re-run
      the Czech-detection check across `site/`.
- [ ] **Footer terms link.** "Algemene voorwaarden" points at
      `https://doglife.be/privacy/`, the same URL as the privacy policy. Point
      it at the real terms page.
- [ ] **Number format.** `€49.` should be `€ 49,00` — Dutch/Belgian convention
      uses a comma decimal separator.
- [ ] **Search index.** `site/search-index.json` holds 326 of 352 products. The
      26 missing are size variants (oils, CBD, minerals, goat milk, chlorella).
      Either index them, or index the parent and make the variant selector
      deep-linkable — but they must not be unreachable from search.
- [ ] **Duplicate DOM ids.** The inline logo SVG is emitted twice per page
      (header + footer), duplicating `yoggies_logo_on_light_bg_RGB`, `Path_7`…
      `Path_15`. Emit the SVG once as `site/images/logo.svg` and reference it
      with `<img>`, or namespace the ids per instance.
- [ ] **Generate `sitemap.xml` and `robots.txt`** at the end of the build, from
      the same page list the builder already has.

**Accept when:** every page in `site/` has a non-empty meta description and a
canonical URL, and no page contains `onclick="window.location`.

---

## Phase 3 — Breed URL slugs

The 27 breed pages have Dutch content under Czech filenames:
`ras-nemecky-ovcak.html` contains `<h1>Duitse herdershond</h1>`.

- [ ] Add an `nl_slug` field to each entry in `_BREEDS_NL` (`atlas_nl_data.py`)
      and generate `ras-duitse-herdershond.html` etc.
- [ ] Emit a redirect stub at each old Czech filename, using the same pattern
      already used for `site/yoghurt.html` (meta refresh + `location.replace`).
- [ ] Update the atlas index cards (line ~773) and any cross-links.

Suggested mapping (verify against the Dutch names already in `atlas_nl_data.py`):
`nemecky-ovcak → duitse-herdershond`, `zlaty-retrivr → golden-retriever`,
`labradorsky-retrivr → labrador-retriever`, `jazevcik → teckel`,
`civava → chihuahua`, `mops → mopshond`, `rotvajler → rottweiler`,
`dobrman → dobermann`, `nemecky-boxer → duitse-bokser`,
`francouzsky-buldocek → franse-bulldog`, `belgicky-ovcak → belgische-herder`,
`australsky-ovcak → australische-herder`, `bernsky-salasnicky-pes →
berner-sennenhond`, `border-kolie → border-collie`, `chodsky-pes →
boheemse-herder`, `havansky-psik → havanezer`, `kavkazsky-pastevecky-pes →
kaukasische-herder`, `stafordsirsky-bulterier → staffordshire-bull-terrier`,
`tibetska-doga → tibetaanse-mastiff`, `yorkshirsky-terier → yorkshire-terrier`,
`samojed → samojeed`, `pitbul → pitbull`.

**Accept when:** all 27 breed pages resolve at Dutch slugs and the 27 Czech
filenames redirect.

---

## Phase 4 — The purchase path (needs the business decision)

From `READY_FOR_BUILD_NL.xlsx` → `match_status`:

| Status | Count | Meaning |
|---|---|---|
| `MATCH` | 122 | real DogLife product URL |
| `NOT_ON_DOGLIFE` | 212 | fallback `doglife.be/?s=...` search URL |
| `SEARCH` | 5 | uncertain match, search URL |
| `manual` | 1 | hand-mapped |

161 of the 212 are in the `speelgoed` category, so this is not a matcher bug —
DogLife genuinely does not carry the toys and accessories.

Regardless of which option is chosen, fix these two first:

- [ ] **The 18 contradictory pages** show a "Bestel bij DogLife.be" button and
      the "niet beschikbaar" note at the same time. The template at
      `yoggies_pipeline.py:1552` must branch on `match_status`, not render both.
- [ ] **The 7 pages with neither** must always render one clear call to action.

Then apply the chosen policy for the 212 in one place, keyed on `match_status`.

---

## Phase 5 — Adding new products

The catalogue is data, not code. To add a product:

1. Append a row to `READY_FOR_BUILD_FINAL.xlsx` (sheet `READY_FOR_BUILD`) with
   at minimum: `row_num`, `category` (one of `brokken`, `barf`, `blikvoer`,
   `snacks`, `gekookt`, `supplementen`, `speelgoed`, `yoghurt`), `name_nl`,
   `description_nl`, `image_main`, `doglife_url`, `match_status`.
   `ingredients_cz` / `dosage_cz` may be blank — 159 and 170 existing rows
   already are.
2. Set `has_calculator=TRUE` only if `dosage_cz` contains a feeding percentage
   range; the calculator is generated from it.
3. Re-run `python localize_images.py` to pull the new images.
4. Re-run `python yoggies_pipeline.py`.

The slug is derived from `name_nl`, so pick the name carefully — changing it
later changes the URL and needs a redirect stub.

---

## Verification script

Run this after the full build; it should print all-clear.

```bash
python - <<'EOF'
import glob, re, collections
pages = glob.glob('site/**/*.html', recursive=True)
no_desc  = [f for f in pages if 'name="description"' not in open(f, encoding='utf-8', errors='replace').read()]
no_canon = [f for f in pages if 'rel="canonical"'    not in open(f, encoding='utf-8', errors='replace').read()]
onclick  = [f for f in pages if 'onclick="window.location' in open(f, encoding='utf-8', errors='replace').read()]
hotlink  = collections.Counter()
noalt = total = 0
for f in pages:
    s = open(f, encoding='utf-8', errors='replace').read()
    for u in re.findall(r'<img[^>]+src="(http[^"]+)"', s):
        hotlink[re.sub(r'https?://([^/]+).*', r'\1', u)] += 1
    for tag in re.findall(r'<img[^>]*>', s):
        total += 1
        if 'alt=' not in tag:
            noalt += 1
print(f"pages                  : {len(pages)}")
print(f"missing description    : {len(no_desc)}")
print(f"missing canonical      : {len(no_canon)}")
print(f"onclick cards left     : {len(onclick)}")
print(f"img without alt        : {noalt}/{total}")
print(f"remaining hotlinks     : {dict(hotlink)}")
EOF
```

Target: zeros everywhere, and `remaining hotlinks` empty.
