# sort_photos

Group edited photos together with the camera originals they were made from,
so all versions of the same shot sit in one place. Built for photo-library /
photoshoot workflows (large RAW/JPEG libraries, Photoshop-edited derivatives).

- **Runs 100% locally** — no cloud services, no per-image fees. Cost is effectively zero.
- **Non-destructive** — originals are never moved, renamed, or modified. Files are only *copied* into a new output folder.
- **Reusable** — works per shoot or batch across many shoots, with configurable folder names.

---

## Why this works (and why it's cheap)

When a photo is edited in Photoshop, the editor *preserves the original capture
timestamp* (`Exif.DateTimeOriginal`) but overwrites `DateTime` and adds
`Software = Adobe Photoshop`. That preserved timestamp is a deterministic link
back to the original frame — no image analysis required.

On the sample shoot (`_sort images`, 1,799 originals + 159 edits):

| Method | Count | Meaning |
|---|---|---|
| `exif` | 63 | Unique capture-second match — certain |
| `exif+burst` | 95 | Matched to the correct capture *second*, disambiguated within a burst |
| `unmatched` | 1 | No timestamp link; flagged for manual review |

**158 / 159 edits matched (99.4%)**, with only 1 needing a human look.

---

## Matching strategy (three tiers, cheapest first)

1. **Tier 1 — EXIF `DateTimeOriginal`** (free, instant, deterministic)
   Each edit is matched to the original with the identical capture timestamp.

2. **Tier 2 — Burst disambiguation** (only runs when needed)
   When several originals share the same capture *second* (burst/continuous
   shots), the edit is compared only against those few candidates using a
   downscaled, contrast-normalised grayscale MSE, which keeps far more spatial
   detail than a perceptual hash. A `burst_margin` records how confidently the
   exact frame was chosen.

3. **Tier 3 — pHash fallback** (only for edits with no timestamp link)
   The edit is perceptually hashed (pHash, orientation-normalised so 90°-rotated
   edits still match) against every original, accepting a match only within a
   distance threshold. Otherwise it is listed as `unmatched`.

> **Note on confidence:** heavy edits (e.g. background replacement) change an
> image so much that *any* pixel/visual comparison becomes unreliable — this was
> measured and confirmed. That is exactly why metadata does the primary work and
> visual matching is only a tie-breaker. Every row is labelled with its method
> and confidence so you always know what is certain vs. best-guess.

---

## Requirements

- Python 3
- Pillow, numpy, imagehash, openpyxl, requests

Install everything at once (`python` on Windows, `python3` on macOS):

```bash
python3 -m pip install -r requirements.txt
```

---

## Running on Windows

`sort_photos.py` and `categorize.py` are pure Python and run unchanged on
Windows. Only the way you launch them differs.

1. **Install Python** (https://www.python.org/downloads/) — tick "Add Python to PATH".
2. **Open a terminal in the project folder** (in File Explorer type `cmd` in the
   address bar, or in VS Code use Terminal → New Terminal).
3. **Install dependencies once:**  `python -m pip install -r requirements.txt`
4. **Create your key file** — the `.env` holding your API key is NOT cloned
   (it is git-ignored). Recreate it:
   ```bat
   copy .env.example .env
   notepad .env
   ```
   then paste your Anthropic key after `ANTHROPIC_API_KEY=` and save.
5. **Run it** — any of:
   ```bat
   run.bat
   ```
   ```bat
   python run.py
   ```
   ```bat
   python sort_photos.py --shoot "_sort images" --out "sorted_output"
   python categorize.py --edit-dir "_sort images/Edit" --report sorted_output/report.csv --out sorted_output/photos.xlsx
   ```
6. **Open the result:** `start sorted_output\photos.xlsx`

> Use forward slashes `/` in the command arguments (as above) — they work on
> Windows too and avoid backslash-escaping issues.

### Windows vs macOS — what actually changes
| Thing | Windows |
|---|---|
| `sort_photos.py`, `categorize.py` | No change — already cross-platform |
| `python3` | use `python` |
| `./run.sh` | use `run.bat` or `python run.py` |
| `.env` (API key) | recreate from `.env.example` |

---

## Usage

### Single shoot

```bash
python3 sort_photos.py --shoot "_sort images" --out "sorted_output"
```

`_sort images` must contain two sub-folders (originals + edits). By default the
tool expects `Camera Original ID` and `Edit`, and otherwise auto-detects via
folder-name hints (`original`, `camera`, `edit`, `final`, ...).

### Batch — many shoots at once

```bash
python3 sort_photos.py --root "/path/to/AllShoots" --out "sorted_output"
```

Every sub-folder of `AllShoots` is treated as its own shoot and given its own
output folder.

### Custom sub-folder names

```bash
python3 sort_photos.py --shoot "MyShoot" --out "out" \
  --originals-name "RAW" --edits-name "Retouched"
```

### Options

| Flag | Default | Purpose |
|---|---|---|
| `--shoot` | — | One shoot folder (mutually exclusive with `--root`) |
| `--root` | — | Parent folder; each sub-folder is a shoot |
| `--out` | *(required)* | Output folder (created if missing) |
| `--originals-name` | `Camera Original ID` | Name of the originals sub-folder |
| `--edits-name` | `Edit` | Name of the edits sub-folder |
| `--hash-size` | `16` | pHash size (bits per side) |
| `--phash-threshold` | `48` | Max pHash distance for a Tier-3 match |

---

## Output

```
sorted_output/
  report.csv
  groups/
    0001_4C8A2376/          ← one folder per edited photo
      4C8A2376.JPG          ← the original
      PRHM5402DF.JPG        ← its edit
    0002_...
  unmatched_edits/          ← edits with no confident match (manual review)
    PRHM5472DF.JPG
```

**`report.csv` columns**

| Column | Meaning |
|---|---|
| `group_id` | Group/folder number (`0001`, `0002`, ...) |
| `edit_file` | The edited file name |
| `matched_original` | The original it was matched to |
| `method` | `exif`, `exif+burst`, `phash`, or `unmatched` |
| `confidence` | `1.0` certain; `0.9` high; `0.6` medium; `0.3` frame uncertain within a burst |
| `distance` | Comparison distance (0 for exact `exif` matches) |
| `burst_margin` | How clearly the exact burst frame won (blank when not a burst) |

**No files are renamed.** Editing is grouped, not renamed, by design.

---

## Scale & cost

- This shoot (~1,958 files): ~10–15 seconds.
- Larger shoots stay cheap: reading EXIF is milliseconds per file, and the
  visual (MSE/pHash) stages run only on the small uncertain subset — never on
  every image. No per-image or cloud charges at any scale.

---

## Step 2 — categorize + captions (`categorize.py`)

Adds a **category** and a **one-line caption** to every edit, then writes an Excel
sheet. Uses a vision LLM (Anthropic Claude, default `claude-haiku-4-5` — cheap and
fast). It joins to `report.csv` so each row also carries the matched original.

```bash
python3 categorize.py --edit-dir "_sort images/Edit" \
                      --report sorted_output/report.csv \
                      --out sorted_output/photos.xlsx
```

**Categories (fixed):** `Portrait`, `Traditional/Festival`, `White Background`,
`Others` — assigned in this priority order (first match wins), as requested:

1. Portrait — a person is the main subject, plain/neutral background
2. Traditional/Festival — traditional/ethnic attire or cultural/festive scene
3. White Background — plain white studio backdrop
4. Others — everything else (product, warehouse, scenery, ...)

**The API key** is read from (in order): `--api-key`, `$ANTHROPIC_API_KEY`, or the
`.env` file next to this script (git-ignored). Nothing is hard-coded.

**Cost / scale:** each image is downscaled (longest side 640 px) and sent once.
Roughly ~$0.001–0.002 per image with Haiku → pennies for a normal shoot
(159 images ≈ a few cents). Results are cached in `labels_cache.json`, so re-runs
are free; use `--no-cache` after changing the prompt.

**Excel columns** (`sorted_output/photos.xlsx`): `group_id, category, caption,
edit_file, matched_original` — sorted by category, colour-coded, header frozen.

The match diagnostics (`method, confidence, distance, burst_margin, caption`) are
written to `sorted_output/categorize_debug.log` instead, for debugging.

### categorize.py options

| Flag | Default | Purpose |
|---|---|---|
| `--edit-dir` | *(required)* | Folder holding the edits |
| `--report` | `sorted_output/report.csv` | Edit→original map from step 1 |
| `--out` | `sorted_output/photos.xlsx` | Output Excel file |
| `--model` | `claude-haiku-4-5-20251001` | Anthropic vision model |
| `--api-key` | env/.env | Anthropic API key |
| `--workers` | `6` | Parallel requests |
| `--limit` | `0` | Only first N images (testing) |
| `--no-cache` | off | Ignore cached labels |

---

## Reviewed / tuning

- **Confirm low-confidence rows:** filter `report.csv` for `confidence=0.3`
  (same burst, exact frame uncertain) or `method=unmatched` and eyeball those
  folders first.
- **Loosen or tighten visuals:** raise `--phash-threshold` to accept weaker
  Tier-3 matches, or lower it to be stricter.