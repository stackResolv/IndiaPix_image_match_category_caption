# sort_photos

Group edited photos together with the camera originals they were made from, then
write the IndiaPix metadata sheet for every pair. Built for photo-library /
photoshoot workflows (large RAW/JPEG libraries, Photoshop-edited derivatives).

Two input layouts are supported, and `run.py` **detects which one you have**:

| Your input | What the tool does |
|---|---|
| A **shoot** folder with `JPEG` + `Edit` sub-folders | matches each edit to its original by capture timestamp |
| One **folder per set** (`IMG_0001/` holding a pair), 30,000-40,000 of them | reads each photo's role from its own metadata — see *Batch — one folder per set* |

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
   edits still match) and compared against a **pre-built index of every
   original**, accepting a match only within a distance threshold. Otherwise it
   is listed as `unmatched`. See *Tier-3 performance* below — the index is built
   once per run, and comparison is ~1 ms per edit rather than minutes.

   > **Accuracy caveat, measured.** On the sample shoot, given 104 pairs whose
   > original was already known from the timestamp, the visual fallback put the
   > true source at rank 1 for only **66%** of them (`--hash-size 16`). With 8 it
   > fell to **48%**, so 8 is *not* used. Perceptual hashing cannot reliably
   > distinguish visually similar frames — that is exactly why metadata does the
   > primary work. Treat `method=phash` rows as "best guess" and prefer
   > `--no-visual-fallback` when the images are look-alikes of each other.

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
command differs. Three launchers ship with the project and all three delegate
to the same `run.py`, so behaviour is identical across platforms:

| Launcher | Shell | Notes |
|---|---|---|
| `run.bat` | `cmd.exe` | Double-clickable; pauses on error |
| `run.ps1` | **PowerShell** | Sets UTF-8 + reports the window width |
| `run.sh` | macOS / Linux | Native shell counterpart |

All of them set `PYTHONUTF8=1`, `PYTHONIOENCODING=utf-8` and
`PIPELINE_WIDTH`, so the narration you see on Windows is **the same as on
macOS** — wrapped to your window, with no `?` where an em dash should be.

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
   python categorize.py --edit-dir "_sort images/Edit" --originals-dir "_sort images/JPEG" --report sorted_output\report.csv --out sorted_output\photos.xlsx
   ```
6. **Open the result:** `start input_folder\_sort images\_out\photos.xlsx`

> The tool works with **forward slashes `/` or backslashes `\`** in arguments —
> either is fine on Windows.

`run.bat` handles the common Windows papercuts for you:
- **Runs from its own folder** — `cd /d "%~dp0"`, so it works when called from
  another directory, by full path, or double-clicked in Explorer. Relative paths
  like `input_folder` always resolve against the project folder.
- **Finds Python automatically** — uses the `py -3` launcher if present, else
  `python` on PATH; prints a clear message if neither is found.
- **UTF-8 console** (`chcp 65001`) + `PYTHONUTF8=1` so non-ASCII filenames never
  crash with a `UnicodeEncodeError`.
- **Forwards arguments untouched** (`%*`), so quoted paths with spaces work from
  `cmd.exe`, PowerShell, or Explorer.
- **Keeps the window open on error** (`pause`) so a failure is readable even when
  the file is double-clicked.
- **Clear error if Excel has the workbook open** — Windows locks the `.xlsx` while
  it is open, so instead of a Python traceback you get:
  `error: cannot write '...photos.xlsx' - it is open in Excel. Close the file ...`

### Windows vs macOS — what actually changes
| Thing | cmd.exe | PowerShell | macOS / Linux |
|---|---|---|---|
| Launcher | `run.bat` | `.\run.ps1` (or `.\run.bat`) | `./run.sh` |
| Interpreter | `python` / `py -3` | `python` / `py -3` | `python3` |
| Console UTF-8 | `chcp 65001` in `run.bat` | set by `run.ps1` | set by `run.sh` |
| Wrap width | `PIPELINE_WIDTH` in `run.bat` | detected via `RawUI` | detected via `tput` |
| `.env` (API key) | recreate from `.env.example` | same | same |

> You can also **double-click `run.bat`** in File Explorer — it runs with the
> default shoot and pauses at the end so you can read the output.

### Windows PowerShell

PowerShell is the default terminal in VS Code on Windows, so `run.ps1` exists as
a first-class launcher alongside `run.bat`. Two small differences vs `cmd.exe`:

- Prefix a script in the current folder with `.\` — PowerShell does not run
  local executables from the current directory unless you do.
- Quote paths that contain spaces (e.g. `"_sort images"`), same as `cmd`.

```powershell
# from the project folder — PowerShell launcher (recommended)
.\run.ps1                                   # default shoot "_sort images"
.\run.ps1 input_folder                      # every shoot inside it
.\run.ps1 input_folder --provider openai    # use OpenAI for captions
.\run.ps1 "input_folder" "my_output"        # batch + custom output base

# the .bat launcher also works from PowerShell
.\run.bat "input_folder"

# or call Python directly (identical behaviour)
python run.py "input_folder"
python run.py "input_folder" --provider openai --model gpt-4o

# the individual steps, if you prefer
python sort_photos.py --shoot "input_folder" --out "sorted_output"
python categorize.py --edit-dir "input_folder/Edit" --report "sorted_output/report.csv" --out "sorted_output/photos.xlsx"
```

**What `run.ps1` handles for you** (all of it silently):

- **Execution policy.** If PowerShell refuses to run the file, that is the
  Execution Policy, not the tool. Run it once with:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\run.ps1 "input_folder"
  ```
  or unblock just this session:
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```
- **UTF-8 console.** Sets `[Console]::OutputEncoding`, `$OutputEncoding` and
  `chcp 65001`, so non-ASCII filenames and captions print correctly instead of
  as `?`.
- **Console width.** Reads `$Host.UI.RawUI.WindowSize.Width` and passes it as
  `PIPELINE_WIDTH`, so the step banners and explanations wrap to *your* window
  instead of assuming a wide terminal.
- **Python discovery.** Prefers the `py -3` launcher, then `python`, then
  `python3`, and prints a clear message if none is found.
- **Right working directory.** `Push-Location $PSScriptRoot`, so it works no
  matter where you invoke it from.

Fallbacks if `.ps1` files are blocked entirely:

```powershell
cmd /c run.bat "input_folder"
python run.py "input_folder"
```

> Installing dependencies and opening results are identical in PowerShell:
> `python -m pip install -r requirements.txt` and
> `start "input_folder\_out\photos.xlsx"`.

#### Console output on Windows

The narration is written for an **80-column** window by default (a fresh
`cmd.exe`/PowerShell console is exactly that). Nothing overflows, and nothing is
silently truncated — long explanations wrap with a hanging indent:

```
  ------------------------------------------------------------------------------
  Step 2/5: reading capture timestamps from 1799 original(s)
  ------------------------------------------------------------------------------
    what: the date+time each frame was shot (EXIF DateTimeOriginal).
    why : Photoshop keeps that stamp, so it is the join key back to the
      original frame - a timestamp match is exact, no guessing.
    how : reads only the metadata block (no image decoding), 32 threads.
  [   0.6s]   originals read: 600/1799 (33%) [0s, ~0s left]
```

Width is detected automatically (the Windows console API is used when
`shutil` reports nothing useful) and clamped to 60–110 columns. To force a
width — useful when piping into a file or a narrow panel:

```powershell
$env:PIPELINE_WIDTH = 100 ; .\run.ps1 input_folder
```

```bat
set PIPELINE_WIDTH=100 && run.bat "input_folder"
```

Characters a legacy code page cannot represent degrade to ASCII look-alikes
(`—` becomes `-`, `→` becomes `->`) rather than raising
`UnicodeEncodeError`, so a run never dies because of a caption.

| Task | cmd.exe | PowerShell |
|---|---|---|
| Run the pipeline | `run.bat "input_folder"` | `.\run.bat "input_folder"` |
| Call Python directly | `python run.py "input_folder"` | `python run.py "input_folder"` |
| Install deps | `python -m pip install -r requirements.txt` | `python -m pip install -r requirements.txt` |
| Open the Excel | `start ...\photos.xlsx` | `start ...\photos.xlsx` |

> **No changes are needed to run under PowerShell.** `run.bat` works there
> unchanged (prefix it with `.\`), and `run.ps1` exists if you prefer a native
> launcher. Both, plus `run.sh`, change into the project folder first, so
> `run.bat "input_folder"` finds the same shoot no matter which directory your
> prompt is in. The only two things to remember in PowerShell are `.\run.bat`
> (PowerShell will not run a script from the current folder without `.\`) and
> execution policy for `run.ps1` (see below).

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

### Batch — one folder per set (30,000–40,000 pairs)

Some clients send the collection already grouped, with **no `JPEG` / `Edit`
sub-folders** — just one folder per set, holding the camera original and its
edit side by side:

```
input_folder/
  IMG_0001/
    _77A0001.CR3      ← camera original
    _77A0001.JPG      ← its edit
  IMG_0002/
    _77A0002.JPG
    _77A0002_final.JPG
  ...
  IMG_40000/
```

`run.py` **detects this layout automatically** — no extra flag — and hands off
to `sets_pipeline.py`:

```bash
# one command, works for either layout
python3 run.py "input_folder"

# or drive the set-folder pipeline directly
python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output"
```

Because there are no folder names to say which photo is which, **each photo's
role is decided from its own metadata**:

| Signal | Weight | Why it is trusted |
|---|---|---|
| EXIF `Software` names an editor | `+4` | Photoshop/Lightroom wrote it — near-conclusive |
| This file is a RAW, the other is not | `-4` | a camera produces the RAW, an export produces the JPEG |
| Edit date later than capture date | `+2` | `DateTimeOriginal` is preserved, `DateTime` is rewritten |
| Filename hint (`_final`, `_edit`, `_retouch`) | `+2` | corroboration only, never decisive alone |
| Smaller pixel area / byte size, later mtime | `+1` | weak — an aggressive crop can go either way |

The result is written to **`report.csv` before any caption is paid for**, so the
decisions can be reviewed with zero API cost:

```bash
# 1) pair the sets and stop — no API calls, no cost
python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output" --dry-run

# 2) a cheap pilot on the first 200 sets
python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output" --limit 200

# 3) the full run (resumable — a re-run never pays twice for the same set)
python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output"
```

#### Sets grouped by operation (`metadata/<operation>/<SET_…>`)

Some deliveries group the sets by **what was done** rather than by date — the
folder name is the operation, and the sets sit inside it:

```
metadata/
  color_tonal/
    SET_000001/
      SET_000001_00_original.jpg
      SET_000001_01_final.jpg
    SET_000002/
  sharpen_denoise/
    SET_000001/
    SET_000002/
  color_tonal;white_balance/          ← the folder may name several
    SET_000001/
```

Point the pipeline at the parent folder:

```bash
python3 sets_pipeline.py --sets-root "metadata" --out "metadata_out" --dry-run
python3 sets_pipeline.py --sets-root "metadata" --out "metadata_out"
```

#### `edit_operation` — the folder name, verbatim, then the model's operations

The operation folder's name is the client's own label, so it is written **first
and exactly as typed** — it is never canonicalised, split or re-cased. A name may
contain `;` or `,` and still be one folder name:

| Folder name | `edit_operation` written |
|---|---|
| `color_tonal/` | `color_tonal;color_tonal_model_ops…` |
| `white balance, crop perspective/` | `white balance, crop perspective;color_tonal` |
| `color_tonal;white_balance/` | `color_tonal;white_balance;crop_perspective` |

The model's own operations follow, in the order it reported them (most
important first). An operation the folder name already mentions is not repeated
behind it. A crop the file metadata proves is appended last. If the input is not
grouped by operation (a date folder, a bare `IMG_0001/`), the cell is just the
model's list.

**The folder name always leads — and it is never lost.** The one rule a client
checks first is that a set filed under `sharpen_denoise/` must have
`sharpen_denoise` as the **first** term of its cell. That holds even when the
tool is pointed straight *at* the operation folder:

```bash
python3 sets_pipeline.py --sets-root "metadata/sharpen_denoise" --out "out"
# -> every cell starts with "sharpen_denoise;…"
```

Every run ends by asserting it, printing either

```
[ok] every set that sits in an operation folder leads its cell with that
     folder's name, verbatim
```

or a `[FAIL]` list naming each offending `set_id`, the folder it should have
led with, and the cell it actually wrote. A violation is never silent.

**No final row is left blank — but no operation is ever invented either.** The
order of evidence is:

1. the client's operation folder name (**verbatim, first, always**),
2. the model's own `edit_operation` answer,
3. anything the file metadata proves (a genuinely changed aspect ratio),
4. **only if 1–3 produced nothing**, the label's own `description` / `title`
   (`Same scene with the crowd removed and a warmer colour grade` →
   `object_removal;color_tonal`),
5. only if *that* is empty too, `none`.

Step 4 fills a blank cell; it never extends a real answer. If the model said
`object_removal` and the description happens to say "scissors added", the cell
stays `object_removal` — writing `compositing` on top would record an operation
that neither the model reported nor a human confirmed. Blank is a defect;
invented is worse.

**Every cell is traceable.** `categorize_debug.log` carries four extra columns
per final row, so a reviewer can see the decision instead of re-running the
model:

| Column | Shows |
|---|---|
| `folder_operation` | the client's operation folder for that set (`''` when not grouped) |
| `op_sources` | where each term came from — `sharpen_denoise<-folder object_removal<-model` |
| `model_edit_operation` | the model's raw answer, before any merging |
| `description_used_for_prose` | the sentence that filled the cell (only when it was blank) |

The run log summarises the same thing: how many terms came from the folder, the
model, the metadata, the prose or the fallback, how many final rows had an empty
model answer, and how many were therefore filled from the description.

#### `set_id` — the set folder's own name

`set_id` is taken **directly from the set folder name** and is never renumbered
or reformatted:

```
metadata/white_balance/SET_037501/   ->  set_id = SET_037501
metadata/color_tonal/SET_000002/     ->  set_id = SET_000002
```

The client's naming convention is preserved exactly, so a sheet can always be
traced back to the folder it came from.

#### The Excel is identical to the shoot workflow

Same 17 template columns in the same order, same trailing `category` column, same `SET_000001` ids, same colour coding — one
`00 original` row followed by one `final` row per edit. Only the way the pairs
are found differs.

#### Scale: 40,000 sets (~80,000 rows)

Measured on a generated 40,000-set / 80,000-photo library:

| Stage | Time | Notes |
|---|---|---|
| Discover 40,000 folders | 0.7 s | one `scandir` per folder, no file opens |
| Read 80,000 photos' metadata | 6 s | EXIF/headers only — **no image is decoded** |
| Pair + write `report.csv` | 0.5 s | 80,000 rows |
| Write `photos.xlsx` (80,000 rows) | 10.5 s | streamed, **peak memory 305 MB** |

The sheet is written with openpyxl's **write-only mode** (`excelutil.py`), which
streams rows straight into the worksheet XML. A normal workbook would hold every
cell in memory — roughly a gigabyte at 80,000 rows — and the run would die just
before saving. With streaming, memory is flat: 200 sets and 40,000 sets cost the
same.

Interrupting a run is safe: labels are checkpointed every 100 sets (written to a
temp file, then atomically renamed, so a crash cannot corrupt the cache), and a
re-run continues instead of paying twice.

Set-folder options:

| Flag | Default | Purpose |
|---|---|---|
| `--sets-root` | *(required)* | Folder whose sub-folders are the sets |
| `--out` | `sorted_output` | Output folder (`photos.xlsx`, `report.csv`) |
| `--provider` | `auto` | `auto`, `anthropic`, or `openai` |
| `--model` | provider default | e.g. `claude-haiku-4-5-20251001` / `gpt-4o-mini` |
| `--workers` | `8` | Parallel API requests |
| `--limit` | `0` | Only the first N sets (a cheap pilot) |
| `--dry-run` | off | Pair + write `report.csv` only — stops before any API call |
| `--no-cache` | off | Ignore cached labels and re-classify |

> Sets with **one photo** get a single `final` row (classified on its own);
> sets with **three or more** get one `original` row plus one `final` row per
> extra photo. Both are reported in the run log, never silently dropped.
> Nested layouts (`input_folder/2025-03-14/IMG_0001/`) are discovered too.

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
| `--hash-size` | `16` | pHash size (bits per side). 16 identifies the true source far more often than 8 |
| `--phash-threshold` | `48` | Max pHash distance for a Tier-3 match. Raise to accept weaker matches |
| `--no-visual-fallback` | off | Skip the catalogue comparison; no-timestamp edits become `unmatched` instantly |

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

### Tier-3 performance

Edits with no capture timestamp used to crawl: the old code ran a Python loop
over every original and printed `compared 250/4074` once or twice a minute.
Two things changed, both measured on the sample shoot:

| Stage | Before | After |
|---|---|---|
| Index 1,799 originals | ~6.0 s sequential, inside the crawl | **2.3 s** parallel, as a labelled step |
| Compare one edit vs 1,799 | 5.2 ms | **3.6 ms** vectorised |
| First orphan edit (4,074 originals) | ~16 min of crawling | **~0.1 s** |
| All 4,312 edits (4,074 originals) | ~20 min | **~16 s** |

How it works: the catalogue is hashed **once**, in parallel, into a `uint8`
matrix (one row per original). Comparing an edit is a single NumPy XOR plus a
byte-popcount lookup table — no Python loop over files, no repeated decoding.

```
  building visual index (reads every original once) ...
  hashed 1200/4074 (29%) [1m12s, ~2m30s left]
  index ready in 94.3s: 4074/4074 usable hash(es)
```

The index is built **only if some edit actually needs it**, so a fully
timestamped shoot never pays for it. Results are identical to the old loop —
verified against it on 30 edits (byte-for-byte the same source and distance).

Two approaches were tested and **rejected**:

- **`--hash-size 8`**: halves the work but identifies the true source only 48%
  of the time instead of 66%, and produced a demonstrable false positive (the
  claimed original ranked 883rd of 1,799 by a higher-detail metric).
- **Draft-mode JPEG decoding** (`Image.draft`): 1.7× faster on a 6000×4000 file,
  but shifts hashes by ~36/256 bits — enough to break the threshold.

New flag:

| Flag | Purpose |
|---|---|
| `--no-visual-fallback` | Skip the catalogue comparison entirely. Edits with no timestamp are marked `unmatched` immediately (instant). Use this when "closest visual match" would just pick a look-alike. |

### Progress logging

Every step is **announced and explained** — the log tells you what is happening,
not just that something is happening. Each step prints a heading, a one-line
*what / why / how*, then progress you can act on:

```
  ------------------------------------------------------------------------------
  Step 1/5: locating the originals and edits folders
  ------------------------------------------------------------------------------
    originals : JPEG  (1799 image(s))
    edits     : Edit  (159 image(s))
    output    : input_folder/_out  (report written here)
    match dir : input_folder/match

  ------------------------------------------------------------------------------
  Step 2/5: reading capture timestamps from 1799 original(s)
  ------------------------------------------------------------------------------
    what: the date+time each frame was shot (EXIF DateTimeOriginal).
    why : Photoshop keeps that stamp, so it is the join key back to the
      original frame - a timestamp match is exact, no guessing.
    how : reads only the metadata block (no image decoding), 32 threads.
  [   0.6s]   done in 0.6s: 1220 distinct timestamp(s) across 1799 file(s)
  [   0.6s]   312 of those timestamp(s) are shared by several frames (bursts) -
              those need visual disambiguation in step 3

  ------------------------------------------------------------------------------
  Step 3/5: matching 159 edit(s) to their original frames
  ------------------------------------------------------------------------------
    how : 1) same capture second? -> certain (free).
          2) several frames share that second (a burst)? -> compare pixels.
          3) no timestamp at all? -> compare perceptual hashes.
  [   7.8s]   how the matches were made:
  [   7.8s]      63  certain - unique capture timestamp
  [   7.8s]      95  timestamp + pixel check (burst frame chosen)
  [   7.8s]       1  no match found - needs manual review

  ------------------------------------------------------------------------------
  Step 4/5: copying 157 matched pair(s) into the match folder
  ------------------------------------------------------------------------------
    note: files are only COPIED - your sources are never moved or renamed.
  [   8.0s]   315 file(s) copied in total

  ------------------------------------------------------------------------------
  RESULT
    edits matched   : 158 of 159
    edits unmatched : 1
    original groups : 157
  ------------------------------------------------------------------------------
```

**On small shoots every match is explained individually**, so you can see exactly
which frame was chosen and why:

```
    [   0.1s] [1/7] _77A0001.JPG
    [   0.1s]     -> EXIF timestamp match -> _77A0001.CR3 (certain)
    [   0.1s] [2/7] _77A0002.JPG
    [   0.1s]     -> burst tie broken by pixel comparison -> _77A0002.CR3
                        (confidence 0.9, margin 2.4)
```

On larger shoots it prints a method breakdown at the end of step 3 instead, and
**lists the low-confidence matches by name** so you know exactly which rows to
check first.

`categorize.py` narrates its own five steps (load report → choose model → check
cache → classify → write Excel). Each image prints the label and caption it
received:

```
  [   4/159] PRHM5402DF.JPG
        -> Portrait | Woman in a mustard saree against a white backdrop
           (3% done, 12s, ~380s left)
```

All output is **timestamped and line-flushed**, so a long run never looks stuck
(useful for big shoots or slow API responses).

**Reading the `originals read` counter:** metadata is read on a thread pool, so
a 4,000-file folder indexes in **seconds**. Progress is printed every 100 files
(every file on shoots of 100 or fewer) with percent, elapsed time and an ETA.

**Reading the `edits processed` counter:** the number is how many edits are
**finished** (not the one in progress), and it always ends at `N/N (100%)`.
Small shoots (≤ 20 edits) print every edit; larger shoots print every 5th plus
the last one. Each line carries elapsed time and a rough ETA.

If a counter ever does sit still, the log now names the culprit instead of going
silent. A file that takes more than 15 s during indexing is reported by name:

```
  [  15.2s]   ... still reading 'IMG_4471.CR3' (16s on this file)
```

and the visual index announces itself before it starts, with progress and ETA:

```
  [   0.1s]   building visual index (reads every original once) ...
  [   0.4s]   index ready in 0.3s: 300/300 usable hash(es)
```

> **Why indexing no longer stalls.** Two things made `originals read` appear
> frozen at an arbitrary number (e.g. `1700/4074`):
> 1. `capture_time()` used to call **rawpy first**, which opens and parses the
>    whole RAW container — seconds per file, and it could hang on a damaged one.
>    It now reads the embedded TIFF/Exif block straight from the file (mmap +
>    struct unpack, microseconds) and only falls back to `rawpy` when that
>    yields nothing. For Canon CR3 the `CMT1` box is located directly.
> 2. Indexing ran on a **single thread** and printed only every 20 files, so one
>    slow file looked identical to a hang. It is now parallel and reports every
>    100 files with an ETA.
>
> During matching, metadata reads also use the bounded (`fast=True`) path, so a
> single pathological file can no longer freeze the run. A file whose timestamp
> cannot be read simply falls through to the pHash stage — matching quality is
> unchanged.

`categorize.py` prints a per-image line with elapsed seconds, and — while API
requests are in flight — a **heartbeat every 15 s** (`... still working: N/M
done ...`). A fully-cached run states `nothing to do — all labels cached`
instead of silently finishing.

---

## Step 2 — categorize + metadata (`categorize.py`)

Sends the **original and its edit to the model together**, so the sheet can name what
actually changed, then writes an Excel sheet in the **IndiaPix metadata template
layout** — the same 17 columns in the same order (plus our `category` appended),
one `00 original` row plus a `final` row per edit. Uses a vision LLM (Anthropic
Claude, default
`claude-haiku-4-5` — cheap and fast). It joins to `report.csv` so every row in a
set carries the same `set_id`.

```bash
python3 categorize.py --edit-dir "_sort images/Edit" \
                      --originals-dir "_sort images/JPEG" \
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
                      --originals-dir "_sort images/JPEG" \
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

# Windows (cmd.exe)
run.bat "input_folder" --provider openai
```

```powershell
# Windows (PowerShell)
.\run.ps1 input_folder --provider openai

# or directly, with a specific model
python run.py "input_folder" --provider openai --model gpt-4o
```

> **Note:** the OpenAI account must have credits (a key with no balance returns
> HTTP 429 "no credits remaining"). Anthropic worked in testing because that key
> has credit. Both providers cost roughly the same (~a few cents per 159 images).

**The API key** is read from (in order): `--api-key`, the provider's env var
(`ANTHROPIC_API_KEY` or `OPENAI_API_KEY`), or the `.env` file next to this script
(git-ignored). Nothing is hard-coded.

**Cost / scale:** each pair is downscaled (longest side 640 px per image) and sent
once — two images per call, but still one request per edit. Roughly
~$0.001–0.002 per edit with Haiku → pennies for a normal shoot (159 edits ≈ a few
cents). Results are cached in `labels_cache_v8.json` (set-folder runs use
`labels_cache_sets_v6.json`), so re-runs are free; use `--no-cache` after
changing the prompt.

> The cache filename tracks the prompt shape. `v8` (and `_sets_v6`) hold the
> client's eight-operation vocabulary **and** the current copy spec — a
> 9–10 meaningful-word title, one ~20-word description and a 13+ keyword list.
> Caches from earlier builds (`v7`/`_sets_v5` hold multi-sentence 35–70 word
> descriptions; `v6` and older also hold 7–8 word titles and 8–15 keywords) are
> **ignored** rather than silently mixed in. Expect one full pass after
> upgrading — the alternative is a sheet with a mixture of old and new formats,
> which would break filtering.

### One original row + one row per edit

A **set** is a matched original plus every edit made from it. Each set produces
**one `00 original` row followed by one `final` row per edit** (`01`, `02`, …),
mirroring `IndiaPix_metadata_template.csv`:

| Row | `edit_state` | `role` | Content |
|---|---|---|---|
| 1 | `00` | `original` | The camera file. `edit_operation` blank |
| 2..n | `01`, `02`, … | `final` | Each delivered edit. `edit_operation` describes its differences |

If an original has two edits, you get three rows — the original is written
**once**, not once per edit, so no `edit_state` repeats inside a set.

### The eight operations in the sheet

They appear in one column, `edit_operation`, on the `final` rows only:

| Column | Holds | Use it for |
|---|---|---|
| `edit_operation` | **all** operations applied, semicolon-separated (`color_tonal;skin_portrait`) | knowing the full edit history of an image |

The `00 original` rows leave it blank — an unedited camera file has no edit type.
Because the value is a semicolon-separated list, split that column on `;` to
count each operation (Excel: Data → Text to Columns; or `=LEN(...)-LEN(SUBSTITUTE(...))+1`).

`edit_operation` uses **the client's eight operations**, semicolon-separated:

| Operation code | Meaning | Looks for |
|---|---|---|
| `background` | Background removal or replacement | the backdrop replaced, removed to white, blurred or swapped; subject keeps its shape |
| `color_tonal` | Colour grading and tonal or exposure correction | brightness, exposure, contrast, saturation, vibrance, curves, LUT, vignette — same framing, different light/colour mood |
| `crop_perspective` | Crop, reframing, and perspective or lens correction | tighter/looser crop, different aspect ratio, straightened horizon, rotation, corrected keystone |
| `white_balance` | White balance correction | colour-temperature/tint correction only — a cast removed or added, greys reading neutral |
| `sharpen_denoise` | Sharpening and noise reduction | detail only — sharpening/clarity/texture, or noise/grain reduction |
| `compositing` | Compositing or element addition | elements combined — something **added**, sky swapped, montage, text/watermark overlay |
| `skin_portrait` | Skin and portrait retouching | portrait retouching — skin smoothing, blemish healing, teeth/eye work, slimming |
| `object_removal` | Object, blemish, or distraction removal | something visible in the original is **gone**, area rebuilt |

Combinations are written the same way, most important first, e.g.
`color_tonal;white_balance` = "colour grading and tonal correction and white
balance correction", or `color_tonal;white_balance;crop_perspective` = "…
and crop, reframing and perspective correction".

The model reports every operation it can actually see, so most final rows carry
two or three. `crop_perspective` is also added automatically when the two files'
aspect ratios genuinely differ (orientation alone never counts) — that is a
*measured* signal, independent of what the model saw.

**Values are canonicalised before they reach a cell.** Whatever the model (or an
older cache) says — `crop`, `Colour Grade`, `WB`, `sky replacement`,
`teeth whitening`, `colour grading and retouch` — is mapped onto the eight, so
a column filter always matches. Unknown terms are dropped. Output order is fixed
(`color_tonal` first, then `crop_perspective`, …), so re-running a pair writes a
byte-identical cell.

`title` and `keywords` are shared by the set; `description` differs — the
original row describes the untouched scene, the final row usually starts with
`"Same scene with …"` and names only the change.

Example:

| file_id | set_id | edit_state | role | edit_operation |
|---|---|---|---|---|
| `_77A0001.CR3` | SET_000001 | 00 | original | |
| `_77A0001.JPG` | SET_000001 | 01 | final | `color_tonal;skin_portrait` |

An edit with no matched original still gets one `final` row (classified from its
own image alone); the run log reports `N set(s) + M single edit(s)`.

**Excel columns** (`sorted_output/photos.xlsx`) — the 17
`IndiaPix_metadata_template.csv` columns first, in exactly the template order,
then one derived column: `category` (18th):

| Column | Filled with |
|---|---|
| `file_id` | The **file name, with extension** — the original's for row `00`, the edit's for the `01`/`02` rows (e.g. `_77A0001.CR3`, `_77A0001.JPG`) |
| `set_id` | The match group from `report.csv`, template style: `SET_` + 6 digits (`SET_000001`, `SET_000002`, ...) |
| `edit_state` | `00` for the original row, then `01`, `02`, ... for each edit in that set |
| `role` | `original` or `final` |
| `edit_operation` | Final rows only: semicolon-separated operations, e.g. `object_removal;color_tonal` |
| `edit_operation_note` | *blank* |
| `title` | LLM editorial title — **9–10 meaningful words** (articles *a / an / the* do **not** count towards the 9–10, so the cell may run to 11–13 words in total), sentence case (first word + proper nouns capitalised), no trailing period: `Woman buying fresh spices at a busy Jaipur market stall`. Long answers are trimmed from the end without leaving a dangling article or conjunction, a forbidden `image of` / `photograph of` opener is stripped, and Title Case is folded down. Shared by the set |
| `description` | LLM description, no trailing period, **one compact sentence of about 20 words** (18–22, hard-capped at 25) — original row: the untouched scene (`Woman selecting spices at a busy Jaipur market stall, warm daylight, crowd behind her`); final row: `Same scene with …`, naming only what changed (`Same scene with the crowd removed and a warmer colour grade`). A longer answer is cut at the first sentence end or at a comma boundary |
| `keywords` | LLM keywords — **13 to 18**, comma-separated with **no space after the commas**, proper names capitalised (`spice market,turmeric,vendor,woman,saree,India,Jaipur,street,crowd,daylight,warm tones,travel,culture,retail`). A longer list is trimmed to the 18 most specific; a shorter one is topped up from the row's own title/description so the **13-keyword minimum** is always met |
| `file_format` | **The filename extension**, lowercase — `photo.jpg` → `jpg`, `IMG_0001.CR3` → `cr3`, `scan.tiff` → `tiff`. Taken from the name, never from Pillow (which would say `JPEG`/`MPO` and cannot even open a RAW) |
| `resolution` | Actual pixel size, e.g. `6000x4000` |
| `color_space` | `sRGB` / `Adobe RGB` / `ProPhoto RGB` / `Grayscale` / `CMYK` (ICC profile first, then Exif ColorSpace, then PIL mode) |
| `file_size_mb` | File size in MiB, one decimal |
| `date_created` | Original row: EXIF `DateTimeOriginal` (`2025-03-14`). Final row: the *edit* date from EXIF `DateTime`, else file mtime (`2026-09-02`) — the same split the template shows |
| `date_contributed` | The run date (`YYYY-MM-DD`) |
| `asset_category` | *blank* |
| `ai_used` | *blank* |
| `category` | **extra 18th column**: `Portrait`, `Traditional/Festival`, `White Background`, `Others` or `ERROR` |

Rows are grouped by category (a set's rows stay adjacent), the header is frozen
at `B2`, and the table has an auto-filter. The category cell is colour-coded per
category, the same tint is applied to the `file_id` cell so a set is scannable
from the first column, and the `original` rows are shaded grey.

The match diagnostics (`method, confidence, distance, burst_margin`), whether an
original was found, the classifier category, and the full `edit_operation`,
`title`, `original_description`, `final_description`, `keywords` text are written
to `sorted_output/categorize_debug.log` for debugging.

### categorize.py options

| Flag | Default | Purpose |
|---|---|---|
| `--edit-dir` | *(required)* | Folder holding the edits |
| `--originals-dir` | auto | Folder holding the originals (else `<shoot>/match`, else a sibling folder) |
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
- **Set-folder pairs:** run `sets_pipeline.py --dry-run` first and filter
  `report.csv` on `confidence`. Rows at `0.35` had no metadata signal at all and
  were paired by "the later-written file is the edit". If a whole batch scores
  `0.35`, the client probably stripped EXIF — spot-check a few sets before paying
  for 40,000 captions.
- **Run the offline tests** after changing the pairing weights:
  ```bash
  python3 tests/test_sets_pipeline.py
  ```
  They build synthetic RAW/JPEG/dated/blind pairs and assert the role decision,
  the 18 columns, the colour coding and the frozen header — no API key
  and no network needed.
