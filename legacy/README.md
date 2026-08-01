# legacy/ — superseded scripts

Everything in this folder is **superseded by `yoggies_pipeline.py`**, which is
the only supported builder. These files are kept for reference and for the
occasional archaeology question ("why is this markup like that?"). They are not
part of the build and are not maintained.

Do not run them against a live `site/` — they were written against older data
shapes and will happily overwrite generated pages with stale markup.

| File | What it was |
|---|---|
| `build_site.py` | First site generator, read `YOGGIES_FINAL_NL.xlsx`. |
| `build_site_final.py` | Second generator, read `READY_FOR_BUILD_FINAL.xlsx`. Closest ancestor of the current pipeline. |
| `verify_and_build.py` | Translation-cache verifier + Excel writer, run in a Linux sandbox. |
| `merge_calculator.py` | One-off: patched the dosage calculator into `index.html`. |
| `merge_extra_pages.py` | One-off: merged extra static pages into the site. |
| `fix_home_categories.py` | One-off: rewrote the homepage category buttons. |
| `replace_badge_icons.py` | One-off: swapped the product badge icon set. |
| `replace_product_calculator.py` | One-off: replaced the calculator block on product pages. |

## Changes made when archiving

All of these hardcoded a path from one developer's machine
(`C:\Users\AdibHAJHASAN\IdeaProjects\yoggies.be\site`) or a Linux sandbox
(`/home/claude/...`) and could not run anywhere else. Each now derives its
paths from its own location:

```python
ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / 'site'
```

Template references were also repointed from `site/` to `templates/`, since the
source templates no longer live in the output folder.
