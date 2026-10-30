#!/usr/bin/env python3
"""
run.py — cross-platform one-command runner for the photo pipeline.
Works identically on Windows (cmd / PowerShell), macOS and Linux.

Usage:
    python run.py                                  # default: shoot "_sort images"
    python run.py "input_folder"                    # batch: every shoot inside it
    python run.py "MyShoot" "my_output"             # single shoot + custom output
    python run.py "input_folder" "my_output"        # batch + custom output base
    python run.py "input_folder" --provider openai  # use OpenAI for captions
    python run.py "input_folder" --provider openai --model gpt-4o

A shoot is any folder that contains an originals sub-folder (default "JPEG")
and/or an edits sub-folder (default "Edit").

  * If the path you pass IS a shoot, only that shoot is processed.
  * Otherwise the path is treated as a parent folder and every sub-folder
    that looks like a shoot is processed.

Runs both steps per shoot:
    1. sort_photos.py  -> matches each edit with its original, builds <shoot>/match
    2. categorize.py   -> classifies + captions each edit, writes photos.xlsx

Output per shoot (unless a custom output is given):
    <shoot>/match/            matched pairs (flat)
    <shoot>/_out/             report.csv, photos.xlsx, categorize_debug.log
"""

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

from consoleutil import hr, safe, setup_console, term_width, wrap

# Console prep: UTF-8 where possible, and the real width so the narration fits
# a narrow Windows PowerShell / cmd window instead of wrapping into nonsense.
setup_console()

HERE = os.path.dirname(os.path.abspath(__file__))
_MAIN_T0 = time.time()


def say(text, indent=0, hanging=None):
    """Print wrapped, encoding-safe text, indented under the current section.

    ``hanging`` controls how far continuation lines are indented relative to
    the first; two spaces keeps a bullet's continuation visually attached.
    """
    for line in wrap(safe(text), indent=indent,
                    subsequent_indent=hanging if hanging is not None else "  "):
        print(" " * indent + line, flush=True)


def rule(char="-"):
    print(hr(char), flush=True)


def run(cmd):
    """Run a child step, echoing the command and how long it took."""
    say("$ " + " ".join(cmd))
    env = dict(os.environ)
    # Force child scripts to use UTF-8 for files + console (Python 3.7+).
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env.setdefault("PIPELINE_WIDTH", str(term_width()))
    t0 = time.time()
    subprocess.check_call(cmd, cwd=HERE, env=env)
    say(f"(this step took {time.time() - t0:.1f}s)")


def looks_like_shoot(path, originals_name, edits_name):
    """True if `path` contains the originals and/or edits sub-folder."""
    p = Path(path)
    if not p.is_dir():
        return False
    names = {d.name for d in p.iterdir() if d.is_dir()}
    return (originals_name in names) or (edits_name in names)


# Extensions that count as a photo when deciding the input layout. Kept local so
# run.py never has to import Pillow just to look at a folder.
_PHOTO_EXTS = {
    ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp", ".bmp",
    ".cr3", ".cr2", ".crw", ".nef", ".nrw", ".arw", ".srf", ".sr2", ".raf",
    ".orf", ".rw2", ".pef", ".dng", ".raw", ".rwl", ".srw", ".3fr", ".fff",
    ".iiq", ".x3f", ".mrw", ".erf", ".kdc", ".dcr",
}


def _holds_photos(path):
    """True if `path` directly contains at least one photo file."""
    try:
        with os.scandir(path) as it:
            for e in it:
                if e.is_file() and \
                        os.path.splitext(e.name)[1].lower() in _PHOTO_EXTS:
                    return True
    except OSError:
        return False
    return False


def _walk_has_photos(path, max_depth=12):
    """True if any photo file lives under `path`, at any sensible depth.

    Deep on purpose: a client may nest their delivery several levels
    (``.../lense correction/INDIAPIX_BATCH_001/SET_000001/*.jpg``) and the layout
    must still be recognised. Our own output folders (``_out``) are skipped.
    """
    base = str(path).rstrip(os.sep)
    base_depth = base.count(os.sep)
    for root, dirs, files in os.walk(base):
        depth = root.rstrip(os.sep).count(os.sep) - base_depth
        dirs[:] = sorted(d for d in dirs if not d.startswith("_") and
                         d.lower() not in ("match", "__macosx", ".git"))
        if depth >= max_depth:
            dirs[:] = []
        for f in files:
            if os.path.splitext(f)[1].lower() in _PHOTO_EXTS:
                return True
    return False


def _find_shoot_dirs(path, originals_name, edits_name, max_depth=8):
    """Every folder at any depth that holds a JPEG/Edit sub-folder.

    Checking only the immediate children is not enough: a shoot can sit inside a
    dated or per-client folder (``.../2024_02_05/INDIAPIX_BATCH_001/JPEG``).
    """
    found = []
    base = str(path).rstrip(os.sep)
    base_depth = base.count(os.sep)
    for root, dirs, files in os.walk(base):
        depth = root.rstrip(os.sep).count(os.sep) - base_depth
        dirs[:] = sorted(d for d in dirs if not d.startswith("_") and
                         d.lower() not in ("__macosx", ".git"))
        if depth >= max_depth:
            dirs[:] = []
        if originals_name in dirs or edits_name in dirs:
            found.append(root)
            dirs[:] = []          # a shoot's own children are not shoots
    return found


def detect_layout(path, originals_name, edits_name):
    """Work out how the input is organised: ``"shoot"`` or ``"sets"``.

      shoot — somewhere under this path is an originals folder (`JPEG`) and/or
              an edits folder (`Edit`). The original workflow: match every edit
              to its original by capture timestamp, then categorise.
      sets  — there are photos but **no** Original/Edit folders anywhere, so the
              delivery is one folder per set (original + edit side by side). The
              role of each photo is then read from its own metadata.

    The set-folder layout is the fallback whenever photos exist, which is what
    makes a nested client tree (``.../<batch>/<SET_xxx>/``) work without any
    flags and without demanding ``JPEG``/``Edit`` folders that are not there.
    """
    p = Path(path)
    if not p.is_dir():
        return "shoot"

    # 1. This folder is itself a shoot.
    if looks_like_shoot(p, originals_name, edits_name):
        return "shoot"
    # 2. A shoot exists somewhere below (any depth).
    if _find_shoot_dirs(path, originals_name, edits_name):
        return "shoot"
    # 3. Photos exist below but no shoot folders: one folder per set.
    if _walk_has_photos(path):
        return "sets"
    return "shoot"


def detect_shoots(path, originals_name, edits_name):
    """Every shoot folder under `path` (the path itself, or any descendant).

    Returns ``[]`` when there is no shoot. The caller then falls back to the
    set-folder pipeline instead of exiting: a plain tree of photo folders is a
    valid input and must never produce "no shoot found".
    """
    p = Path(path)
    if not p.is_dir():
        sys.exit(f"error: '{path}' is not a directory")
    if looks_like_shoot(p, originals_name, edits_name):
        return [str(p)]
    return _find_shoot_dirs(path, originals_name, edits_name)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", nargs="?", default="_sort images",
                    help="a shoot folder, or a parent folder of shoots "
                         "(default: '_sort images')")
    ap.add_argument("out", nargs="?", default=None,
                    help="output folder. Single shoot -> used as-is; batch -> a base "
                         "folder. If omitted, each shoot writes to <shoot>/_out")
    ap.add_argument("--originals-name", default="JPEG",
                    help="name of the originals sub-folder")
    ap.add_argument("--edits-name", default="Edit",
                    help="name of the edits sub-folder")
    ap.add_argument("--provider", default="auto", choices=["auto", "anthropic", "openai"],
                    help="vision LLM provider for captions (default: auto-detect from keys)")
    ap.add_argument("--model", default=None,
                    help="model name (default: claude-haiku-4-5 for anthropic, "
                         "gpt-4o-mini for openai)")
    ap.add_argument("--workers", type=int, default=None,
                    help="parallel caption requests (default: 6 for a shoot, "
                         "8 for a set-folder run)")
    ap.add_argument("--limit", type=int, default=0,
                    help="set-folder layout only: process just the first N sets")
    ap.add_argument("--dry-run", action="store_true",
                    help="set-folder layout only: find and pair the sets, write "
                         "report.csv, and stop before any API call")
    args = ap.parse_args()

    # A path that does not exist is the one thing we cannot recover from, so
    # say exactly what was tried rather than failing deeper down.
    if not os.path.isdir(args.input):
        sys.exit(f"error: '{args.input}' is not a folder.\n"
                 f"       Pass the folder holding your photos (or the batch "
                 f"folder that contains it).\n"
                 f"       Tip: you can drag the folder onto run.bat instead of "
                 f"typing the path.")

    # The client now sends two possible layouts. Detect which one this is, so
    # the same command works for both and no flag has to be remembered.
    layout = detect_layout(args.input, args.originals_name, args.edits_name)
    if layout == "sets":
        run_sets(args)
        return

    shoots = detect_shoots(args.input, args.originals_name, args.edits_name)
    if not shoots:
        # No JPEG/Edit folders anywhere, but there are photos: this is the
        # folder-per-set layout. Run that instead of erroring out.
        print(flush=True)
        say(f"note          : no '{args.originals_name}' / '{args.edits_name}' "
            f"folders under this path - treating it as one folder per set.")
        run_sets(args)
        return
    batch = len(shoots) > 1

    print(flush=True)
    rule("=")
    say("PHOTO PIPELINE")
    rule("=")
    say(f"input         : {args.input}")
    say(f"shoot(s)      : {len(shoots)}"
        + ("  (batch - each shoot is processed in turn)" if batch else ""))
    say(f"originals dir : {args.originals_name}    edits dir : {args.edits_name}")
    say(f"provider      : {args.provider}")
    if os.name == "nt":
        say(f"platform      : Windows  (console width {term_width()} columns)")
    else:
        say(f"platform      : {sys.platform}  (console width {term_width()} columns)")
    print(flush=True)
    say("What will happen, per shoot:")
    say("1. sort_photos.py - read capture timestamps, match every edit to the "
        "original it came from, copy each pair into <shoot>/match/, and write "
        "report.csv describing how each match was made.", indent=2)
    say("2. categorize.py  - send each edit to the vision model, get a category "
        "+ caption, and write photos.xlsx.", indent=2)
    say("Nothing in your source folders is moved, renamed or deleted.", indent=2)

    for shoot_idx, shoot in enumerate(shoots, 1):
        if args.out:
            out = os.path.join(args.out, os.path.basename(shoot)) if batch else args.out
        else:
            out = os.path.join(shoot, "_out")
        edit_dir = os.path.join(shoot, args.edits_name)

        print(flush=True)
        rule("#")
        say(f"SHOOT {shoot_idx}/{len(shoots)}: {os.path.basename(shoot)}")
        say(f"output -> {out}")
        rule("#")
        print(flush=True)

        say("Step 1/2: match each edit to its original")
        say("reads timestamps (free), falls back to pixel/hash comparison only "
            "where needed; copies pairs into the shoot's match/ folder.", indent=2)
        run([sys.executable, "sort_photos.py",
             "--shoot", shoot, "--out", out,
             "--originals-name", args.originals_name,
             "--edits-name", args.edits_name])

        print(flush=True)
        say("Step 2/2: categorise + caption each edit")
        say(f"sends each edit to the vision model (provider: {args.provider}); "
            "results are cached, so a re-run never pays for the same image twice.",
            indent=2)
        cat_cmd = [sys.executable, "categorize.py",
                   "--edit-dir", edit_dir,
                   "--originals-dir", os.path.join(shoot, args.originals_name),
                   "--report", os.path.join(out, "report.csv"),
                   "--out", os.path.join(out, "photos.xlsx"),
                   "--provider", args.provider]
        if args.model:
            cat_cmd += ["--model", args.model]
        if args.workers:
            cat_cmd += ["--workers", str(args.workers)]
        if args.limit:
            cat_cmd += ["--limit", str(args.limit)]
        run(cat_cmd)

    print(flush=True)
    rule("=")
    say(f"DONE - {len(shoots)} shoot(s) processed in {time.time() - _MAIN_T0:.1f}s")
    rule("=")
    say("What to open:")
    for shoot in shoots:
        out = (os.path.join(args.out, os.path.basename(shoot)) if batch else args.out) \
            if args.out else os.path.join(shoot, "_out")
        say(f"{os.path.basename(shoot)}:", indent=2)
        say("photos.xlsx  - categories + captions (start here)", indent=4)
        say(os.path.join(out, "photos.xlsx"), indent=6)
        say("report.csv   - how each edit was matched to its original", indent=4)
        say(os.path.join(out, "report.csv"), indent=6)
        say("match/       - the matched pairs, side by side", indent=4)
        say(os.path.join(shoot, "match"), indent=6)
    print(flush=True)
    say("If any rows look wrong, check 'confidence' in report.csv - values of "
        "0.3, or method=unmatched, are the ones worth eyeballing first.")


def run_sets(args):
    """Run the set-folder workflow (one folder per set, 2 photos each).

    Delegates to sets_pipeline.py, which reads each photo's role from its own
    metadata instead of from Original/Edit folder names, and streams the Excel
    sheet so 40,000 sets (~80,000 rows) fits in memory.
    """
    print(flush=True)
    rule("=")
    say("PHOTO PIPELINE - SET-FOLDER LAYOUT")
    rule("=")
    say(f"input    : {args.input}")
    say(f"detected : one folder per set (original + edit side by side)")
    say(f"provider : {args.provider}")
    if args.dry_run:
        say("mode     : DRY RUN - pairs the sets and writes report.csv only")
    print(flush=True)
    say("Why this differs from a shoot: there are no 'JPEG' / 'Edit' folders, "
        "so each photo's role is read from its own metadata (EXIF Software, "
        "edit date, pixel size, byte size) before anything is captioned.")
    say("Nothing in your source folders is moved, renamed or deleted.")

    out = args.out or os.path.join(args.input, "_out")
    cmd = [sys.executable, "sets_pipeline.py",
           "--sets-root", args.input,
           "--out", out,
           "--provider", args.provider]
    if args.model:
        cmd += ["--model", args.model]
    if args.workers:
        cmd += ["--workers", str(args.workers)]
    if args.limit:
        cmd += ["--limit", str(args.limit)]
    if args.dry_run:
        cmd += ["--dry-run"]
    print(flush=True)
    run(cmd)

    print(flush=True)
    rule("=")
    say(f"DONE in {time.time() - _MAIN_T0:.1f}s")
    rule("=")
    say("What to open:")
    say("photos.xlsx  - categories + captions, same 17-column template "
        "(start here)", indent=2)
    say(os.path.join(out, "photos.xlsx"), indent=4)
    say("report.csv   - which photo of each set was treated as the original, "
        "and why", indent=2)
    say(os.path.join(out, "report.csv"), indent=4)
    print(flush=True)
    say("Rows with a low 'confidence' in report.csv are the ones to eyeball "
        "first: they were paired without a strong metadata signal.")


if __name__ == "__main__":
    main()
