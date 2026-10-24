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
cents). Results are cached in `labels_cache_v3.json`, so re-runs are free; use
`--no-cache` after changing the prompt.

> The cache filename tracks the prompt shape: `labels_cache_v3.json` holds the
> pair-comparison fields. Older caches (`labels_cache.json`,
> `labels_cache_v2.json`) cannot supply `edit_operation` /
> `original_description`, so they are not reused — expect one full pass after
> upgrading.

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

`edit_operation` uses the template's vocabulary, semicolon-separated, most
important first: `crop`, `resize`, `straightening`, `color_tonal`, `retouch`,
`object_removal`, `object_addition`, `background_replacement`, `sharpening`,
`vignette`, `text_overlay`. The model reports what it can see; a `crop` is also
added automatically when the two files' aspect ratios genuinely differ
(orientation alone never counts). `title` and `keywords` are shared by the set;
`description` differs — the original row describes the untouched scene, the final
row usually starts with `"Same scene with …"` and names only the change.

Example:

| file_id | set_id | edit_state | role | edit_operation |
|---|---|---|---|---|
| `_77A0001.CR3` | SET_000001 | 00 | original | |
| `_77A0001.JPG` | SET_000001 | 01 | final | `color_tonal;retouch` |

An edit with no matched original still gets one `final` row (classified from its
own image alone); the run log reports `N set(s) + M single edit(s)`.

**Excel columns** (`sorted_output/photos.xlsx`) — the 17
`IndiaPix_metadata_template.csv` columns first, in exactly the template order,
plus the classifier's own `category` appended as an **18th trailing column**:

| Column | Filled with |
|---|---|
| `file_id` | The **file name, with extension** — the original's for row `00`, the edit's for the `01`/`02` rows (e.g. `_77A0001.CR3`, `_77A0001.JPG`) |
| `set_id` | The match group from `report.csv`, template style: `SET_` + 6 digits (`SET_000001`, `SET_000002`, ...) |
| `edit_state` | `00` for the original row, then `01`, `02`, ... for each edit in that set |
| `role` | `original` or `final` |
| `edit_operation` | Final rows only: semicolon-separated operations, e.g. `object_removal;color_tonal` |
| `edit_operation_note` | *blank* |
| `title` | LLM editorial title (4–10 words, Title Case, no trailing period), shared by the set |
| `description` | LLM description, no trailing period — original row: the untouched scene (`Woman selecting spices at a market stall, crowd in background`); final row: what changed (`Same scene with background crowd removed and warm color grade`) |
| `keywords` | LLM keywords, comma-separated, **no space after the commas**, proper names capitalised (`market,spices,woman,India,street`) |
| `file_format` | `JPEG`, `TIFF`, `NEF`, `CR3`, … (real format; RAW from the extension) |
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