# legacy/

Superseded scripts and data files. Kept for reference only — do not run them.

## Replaced builders
`build_site.py`, `build_site_final.py`, `verify_and_build.py` were three
overlapping forks of the same site builder. `yoggies_pipeline.py` in the
project root is the only builder now.

## One-off patch scripts
`fix_home_categories.py`, `merge_calculator.py`, `merge_extra_pages.py`,
`replace_badge_icons.py`, `replace_product_calculator.py` patched files
inside `site/`. That work is now done by the generator, so patching build
output is no longer needed — and would be erased on the next build anyway.

## Superseded helpers
`download_images.py` handled a handful of template images and rewrote built
HTML. Replaced by `localize_images.py`, which is data-driven and covers all
640+ images.
`manual_variants.py`, `diagnose_variants.py`, `check_site.py`, `test_models.py`
were diagnostic one-offs. `pipe/audit.py` replaces them.

## Old spreadsheets
`READY_FOR_BUILD_NL.xlsx` and `YOGGIES_FINAL_NL.xlsx` were intermediate copies
of the same catalogue. `READY_FOR_BUILD_FINAL.xlsx` in the root is the single
source of truth.
