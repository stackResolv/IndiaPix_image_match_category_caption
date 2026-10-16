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
- **rawpy** — *optional*, only needed to read Canon RAW (`.cr3`, `.cr2`) and
  other RAW formats. Without it the pipeline still runs on JPEG/PNG/TIFF/WEBP/
  BMP and any RAW files are simply skipped (with a one-time warning).

Install everything at once (`python` on Windows, `python3` on macOS):

```bash
python3 -m pip install -r requirements.txt
```

### Supported image formats

| Kind | Extensions | Decoder |
|---|---|---|
| Raster | `.jpg .jpeg .png .tif .tiff .webp .bmp` | Pillow |
| Canon RAW | `.cr3 .cr2 .crw` | rawpy (libraw) |
| Other RAW | `.nef .arw .raf .orf .rw2 .dng .pef .srw .raw …` | rawpy (libraw) |

> **CR3 support:** Canon CR3 is a BMFF/HEIF-based RAW format that Pillow cannot
> decode. The tool now reads CR3 (and CR2 and other RAW formats) via **rawpy**,
> for both matching (`sort_photos.py`) and captioning (`categorize.py`). RAW
> files are decoded at half resolution — full detail is not needed for hashing
> or the 640 px vision payload. If `rawpy` is not installed, RAW files are
> skipped with a warning rather than crashing the run.
>
> RAW/CR3 files can appear in **either** the originals folder or (less commonly)
> the edits folder — both are handled.

---

## Running on Windows

Everything is pure Python and runs unchanged on Windows — only the launch
command differs. `run.bat` is the Windows counterpart of `run.sh` and both
delegate to the same `run.py`, so behaviour is identical across platforms.

1. **Install Python** (https://www.python.org/downloads/) — tick
   **"Add Python to PATH"** during setup.
2. **Open a terminal in the project folder** (in File Explorer type `cmd` in the
   address bar, or in VS Code use Terminal → New Terminal).
3. **Install dependencies once:**
   ```bat
   python -m pip install -r requirements.txt
   ```
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
   run.bat "input_folder"
   ```
   ```bat
   python run.py "input_folder"
   ```
   ```bat
   python sort_photos.py --shoot "_sort images" --out "sorted_output"
   python categorize.py --edit-dir "_sort images/Edit" --report sorted_output\report.csv --out sorted_output\photos.xlsx
   ```
6. **Open the result:** `start input_folder\_sort images\_out\photos.xlsx`

> The tool works with **forward slashes `/` or backslashes `\`** in arguments —
> either is fine on Windows.

`run.bat` handles the common Windows papercuts for you:
- **Finds Python automatically** — uses the `py -3` launcher if present, else
  `python` on PATH; prints a clear message if neither is found.
- **UTF-8 console** (`chcp 65001`) + `PYTHONUTF8=1` so non-ASCII filenames never
  crash with a `UnicodeEncodeError`.
- **Keeps the window open on error** (`pause`) so a failure is readable even when
  the file is double-clicked.

### Windows vs macOS — what actually changes
| Thing | Windows | macOS / Linux |
|---|---|---|
| Launcher | `run.bat` (or `python run.py`) | `./run.sh` (or `python3 run.py`) |
| Interpreter | `python` / `py -3` | `python3` |
| Console UTF-8 | handled by `run.bat` (`chcp 65001`) | handled by `run.sh` |
| `.env` (API key) | recreate from `.env.example` | same |

> You can also **double-click `run.bat`** in File Explorer — it runs with the
> default shoot and pauses at the end so you can read the output.

### Windows PowerShell

Everything works the same in **PowerShell** (the default terminal in VS Code on
Windows). Two small differences vs `cmd.exe`:

- Prefix a script in the current folder with `.\` — PowerShell does not run
  local executables from the current directory unless you do.
- Quote paths that contain spaces (e.g. `"_sort images"`), same as `cmd`.

```powershell
# from the project folder
.\run.bat                                   # default shoot "_sort images"
.\run.bat "input_folder"                    # every shoot inside it
.\run.bat "input_folder" "my_output"        # batch + custom output base
.\run.bat "input_folder" --provider openai  # use OpenAI for captions

# or call Python directly (same behaviour as run.bat)
python run.py "input_folder"
python run.py "input_folder" --provider openai --model gpt-4o

# the individual steps, if you prefer
python sort_photos.py --shoot "input_folder" --out "sorted_output"
python categorize.py --edit-dir "input_folder/Edit" --report "sorted_output/report.csv" --out "sorted_output/photos.xlsx"
```

If PowerShell blocks `.\run.bat` with a script-execution message, that is the
**Execution Policy**, not the tool. Either run it through `cmd` instead:

```powershell
cmd /c run.bat "input_folder"
```

or unblock for the current session only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

> Installing dependencies and opening results are identical in PowerShell:
> `python -m pip install -r requirements.txt` and
> `start "input_folder\_out\photos.xlsx"`.

| Task | cmd.exe | PowerShell |
|---|---|---|
| Run the pipeline | `run.bat "input_folder"` | `.\run.bat "input_folder"` |
| Call Python directly | `python run.py "input_folder"` | `python run.py "input_folder"` |
| Install deps | `python -m pip install -r requirements.txt` | `python -m pip install -r requirements.txt` |
| Open the Excel | `start ...\photos.xlsx` | `start ...\photos.xlsx` |

---

## Usage

### Single shoot

```bash
python3 sort_photos.py --shoot "_sort images" --out "sorted_output"
```

`_sort images` must contain two sub-folders (originals + edits). By default the
tool expects `JPEG` and `Edit`, and otherwise auto-detects via folder-name hints
(`original`, `camera`, `raw`, `edit`, `final`, ...).

### Batch — many shoots at once

Point any entry point at the **parent folder** and every shoot inside it is
processed. A folder is treated as a shoot if it contains an originals
(`JPEG`) and/or edits (`Edit`) sub-folder.

```bash
# run the whole pipeline (match + captions) for every shoot under input_folder
python3 run.py "input_folder"

# or just the matching step
python3 sort_photos.py --root "input_folder" --out "sorted_output"
```

Each shoot gets its own flat `match/` folder and its own `_out/` (report + Excel).
With `run.py`, omitting the output argument writes results to `<shoot>/_out`.

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
| `--originals-name` | `JPEG` | Name of the originals sub-folder |
| `--edits-name` | `Edit` | Name of the edits sub-folder |
| `--hash-size` | `16` | pHash size (bits per side) |
| `--phash-threshold` | `48` | Max pHash distance for a Tier-3 match |

---

## Output

The grouped photos go into a single flat `match/` folder **inside each shoot
folder** (next to the originals and edits folders). Every matched original is
copied in beside its edit(s) — there are **no sub-folders**, so all pairs sit in
`match/` together. Edits with no matching original are **ignored** (not copied
anywhere).

```
_sort images/
  JPEG/                     ← source (never modified)
  Edit/                     ← source (never modified)
  match/                    ← created by the tool, flat, one level only
    4C8A2376.JPG            ← an original
    PRHM5402DF.JPG          ← its matched edit
    4C8A2377.JPG            ← another original
    PRHM5403DF.JPG
    ...

sorted_output/              ← report + Excel only
  report.csv
```

Every entry in `match/` belongs to a pair: an original is only copied when it has
at least one matched edit, and an edit is only copied when it has a matched
original. So `match/` never contains a lone edit (or a lone original).

> If two different source files happen to share the same file name, the later
> one is copied with a numeric suffix (`name_1.JPG`) so nothing is overwritten.

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

**No files are renamed.** Editing is grouped, not renamed, by design. Matched
files are only *copied* into `match/`; the originals and edits folders are left
untouched.

---

## Scale & cost

- This shoot (~1,958 files): ~10–15 seconds.
- Larger shoots stay cheap: reading EXIF is milliseconds per file, and the
  visual (MSE/pHash) stages run only on the small uncertain subset — never on
  every image. No per-image or cloud charges at any scale.

### Progress logging

Both steps print **timestamped, line-flushed** progress so a long run never
looks stuck (useful for big shoots or slow API responses):

```
  [   0.0s] reading capture time from 20 original(s) ...
  [   1.2s]   originals read: 20/20
  [   1.2s] indexed 20 distinct capture time(s)
  [   1.2s] matching 7 edit(s) to originals ...
  [   1.3s] copying 7 matched group(s) into 'input_folder/match' ...
```

`categorize.py` prints a per-image line with elapsed seconds, and — while API
requests are in flight — a **heartbeat every 15 s** (`... still working: N/M
done ...`). A fully-cached run states `nothing to do — all labels cached`
instead of silently finishing.

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

**Provider / model** — two vision providers are supported, chosen with
`--provider` (default `auto`, which picks whichever key is present):

| `--provider` | Default model | Key used |
|---|---|---|
| `anthropic` (default) | `claude-haiku-4-5-20251001` | `ANTHROPIC_API_KEY` |
| `openai` | `gpt-4o-mini` | `OPENAI_API_KEY` |
| `auto` | whichever key is present | — |

```bash
# Step 2 only — use OpenAI instead of Anthropic
python3 categorize.py --edit-dir "_sort images/Edit" \
                      --report sorted_output/report.csv \
                      --out sorted_output/photos.xlsx --provider openai

# override the model too
python categorize.py ... --provider openai --model gpt-4o
```

**Run the whole pipeline (match + captions) with OpenAI in one command** — the
provider flag is passed straight through `run.py` (and therefore `run.bat` /
`run.sh`):

```bash
# macOS / Linux
./run.sh "input_folder" --provider openai

# Windows
run.bat "input_folder" --provider openai

# or directly, with a specific model
python3 run.py "input_folder" --provider openai --model gpt-4o
```

> **Note:** the OpenAI account must have credits (a key with no balance returns
> HTTP 429 "no credits remaining"). Anthropic worked in testing because that key
> has credit. Both providers cost roughly the same (~a few cents per 159 images).

**The API key** is read from (in order): `--api-key`, the provider's env var
(`ANTHROPIC_API_KEY` or `OPENAI_API_KEY`), or the `.env` file next to this script
(git-ignored). Nothing is hard-coded.

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
| `--provider` | `auto` | `auto`, `anthropic`, or `openai` |
| `--model` | provider default | e.g. `claude-haiku-4-5-20251001` / `gpt-4o-mini` |
| `--api-key` | env/.env | API key for the chosen provider |
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