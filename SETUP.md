# Setting up on a fresh laptop

From nothing to a built site. Windows/PowerShell is assumed — that is what this
project has always been built on. `RUNBOOK.md` covers the pipeline itself; this
file only covers getting a new machine to the point where the runbook applies.

---

## 1. Install Python 3.13

Get it from <https://www.python.org/downloads/> — the project is built and run
on **3.13.14**. Tick **"Add python.exe to PATH"** in the installer.

```powershell
python --version      # expect Python 3.13.x
```

Nothing else needs installing globally. `git` you already have if you cloned.

## 2. Clone

```powershell
git clone <repo-url> yoggies.be
cd yoggies.be
```

## 3. Create the virtual environment

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

If PowerShell refuses to run the activate script, allow signed local scripts
for your user once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

You can also skip activation entirely and call the interpreter directly —
`.venv\Scripts\python.exe pipe\audit.py`. Every command below works either way.

## 4. Install the dependencies

```powershell
pip install -r requirements.txt
```

Nine pinned packages: requests, openai, anthropic, openpyxl, pillow, pandas,
cloudscraper, beautifulsoup4, httpx.

## 5. Fetch the images — **before** you build

`site/images/` is **not in the repo** (173 MB of binaries that regenerate).
A fresh clone has no images at all, so this has to run first:

```powershell
python localize_images.py
```

It reads `image_map.json` (which **is** in the repo) and re-downloads roughly
1,019 originals from yoggies.cz, then writes the derived sizes. Expect several
minutes and about 173 MB on disk. It needs a working internet connection and
yoggies.cz to be up.

Then build:

```powershell
python yoggies_pipeline.py
python pipe\audit.py
```

A clean run ends at **25 passed, 1 warning, 0 failed**. If you get failures,
the order above is the first thing to check.

---

## Gotchas

### Set `PYTHONIOENCODING=utf-8` on Windows

The catalogue is full of Czech (`ř`, `ě`, `ú`) and the Dutch carries `€` and
`–`. The Windows console defaults to cp1252 and any script that prints a
product name will die with `UnicodeEncodeError: 'charmap' codec can't encode
character`. Set it for the session:

```powershell
$env:PYTHONIOENCODING = "utf-8"
```

Or permanently, once:

```powershell
[Environment]::SetEnvironmentVariable("PYTHONIOENCODING","utf-8","User")
```

This bites hardest on ad-hoc one-liners that print names. The pipeline scripts
mostly write files rather than printing, so a build can succeed while a
five-line inspection script crashes on the same data.

### audit.py reports phantom broken links when images are missing

If you run `pipe\audit.py` on a clone where step 5 has not run yet, the "no
broken internal links" check reports roughly **1,220 broken paths**. They are
not broken — every one is an `<img src>` pointing into `site/images/`, and the
audit is correctly telling you the file is not on disk. Run
`localize_images.py` and the count returns to 0. Do not go hunting for a bug in
the link checker or start editing paths.

### The two translation caches are load-bearing — never delete them

```
translation_cache_v2.json      1.9 MB
data/cache/translations.json   3.3 MB
```

Both are in the repo, and both must stay there. They hold the Dutch for the
whole catalogue: 2,566 blocks across 1,042 fields. They are keyed by
**sha256 of the full source text**, so any change to the Czech side — even
whitespace — invalidates that entry and the text has to be paid for again
through a translation API.

Deleting either one does not "just regenerate". It means re-translating 526
products against an API budget, and the Dutch has already had a manual pass
over it that no re-run will reproduce. If a build reports
`translated : 0`, that is the caches doing their job.

The same applies to `data/yoggies_raw.json` — it is the verbatim API response
and the reason `pipe\1_fetch_yoggies.py --from-raw` can redo extraction
without hitting yoggies.cz again.

---

## What is NOT in the repo

Everything ignored, and whether a fresh clone gets it back on its own:

| Path | Size | Comes back how |
|---|---|---|
| `site/images/` | 173 MB | **Automatic** — `python localize_images.py` rebuilds all of it from `image_map.json`. Needs internet and yoggies.cz reachable. |
| `.venv/` | ~200 MB | **Automatic** — steps 3 and 4. |
| `__pycache__/`, `pipe/__pycache__/` | small | **Automatic** — Python writes them. |
| `.idea/` | small | **Not needed.** PyCharm settings, personal to the machine. |
| `READY_FOR_BUILD_FINAL.backup-*.xlsx` | ~670 KB each | **Not needed.** `4_build_master.py` timestamps a backup on every run; git already holds that history. Delete them periodically. |

**Nothing irreplaceable lives outside the repo.** The files that could not be
recreated — both translation caches, `image_map.json`, `data/yoggies_raw.json`,
`data/match_decisions.json` and the master workbook — are all tracked and come
down with the clone.

The one real dependency on the outside world is `localize_images.py` needing
yoggies.cz to still serve those 1,019 URLs. If that ever stops being true, the
images become genuinely unrecoverable from this repo alone, so a copy of
`site/images/` on a disk somewhere is cheap insurance.
