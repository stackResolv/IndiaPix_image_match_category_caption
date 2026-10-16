#!/usr/bin/env python3
"""
run.py — cross-platform one-command runner for the photo pipeline.
Works on Windows, macOS and Linux (uses only Python, no shell scripting).

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
    2. categorize.py   -> adds category + caption, writes Excel

Output per shoot (unless a custom output is given):
    <shoot>/match/            matched pairs (flat)
    <shoot>/_out/             report.csv, photos.xlsx, categorize_debug.log
"""

import argparse
import subprocess
import sys
import os
from pathlib import Path

# Windows-safe console: force UTF-8 with a lossy fallback so non-ASCII paths or
# filenames never crash printing on a cp1252 console.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))


def run(cmd):
    print("$ " + " ".join(cmd), flush=True)
    env = dict(os.environ)
    # Force child scripts to use UTF-8 for files + console (Python 3.7+).
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    subprocess.check_call(cmd, cwd=HERE, env=env)


def looks_like_shoot(path, originals_name, edits_name):
    """True if `path` contains the originals and/or edits sub-folder."""
    p = Path(path)
    if not p.is_dir():
        return False
    names = {d.name for d in p.iterdir() if d.is_dir()}
    return (originals_name in names) or (edits_name in names)


def detect_shoots(path, originals_name, edits_name):
    """Return a list of shoot folder paths (single shoot, or all sub-shoots)."""
    p = Path(path)
    if not p.is_dir():
        sys.exit(f"error: '{path}' is not a directory")

    # The path itself is a shoot.
    if looks_like_shoot(p, originals_name, edits_name):
        return [str(p)]

    # Otherwise treat it as a parent: each sub-folder that looks like a shoot.
    shoots = [str(d) for d in sorted(p.iterdir())
              if d.is_dir() and looks_like_shoot(d, originals_name, edits_name)]
    if not shoots:
        sys.exit(f"error: no shoot found in '{path}' "
                 f"(expected sub-folders containing '{edits_name}' or '{originals_name}')")
    return shoots


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
    args = ap.parse_args()

    shoots = detect_shoots(args.input, args.originals_name, args.edits_name)
    batch = len(shoots) > 1

    print(f"Processing {len(shoots)} shoot(s) from '{args.input}'")

    for shoot in shoots:
        if args.out:
            out = os.path.join(args.out, os.path.basename(shoot)) if batch else args.out
        else:
            out = os.path.join(shoot, "_out")
        edit_dir = os.path.join(shoot, args.edits_name)

        print(flush=True)
        print("=" * 66, flush=True)
        print(f" {os.path.basename(shoot)} -> {out}", flush=True)
        print("=" * 66, flush=True)
        print(" Step 1/2: matching edits to originals", flush=True)
        run([sys.executable, "sort_photos.py",
             "--shoot", shoot, "--out", out,
             "--originals-name", args.originals_name,
             "--edits-name", args.edits_name])

        print(flush=True)
        print(f" Step 2/2: categorizing + captioning edits  (provider: {args.provider})", flush=True)
        cat_cmd = [sys.executable, "categorize.py",
                   "--edit-dir", edit_dir,
                   "--report", os.path.join(out, "report.csv"),
                   "--out", os.path.join(out, "photos.xlsx"),
                   "--provider", args.provider]
        if args.model:
            cat_cmd += ["--model", args.model]
        run(cat_cmd)

    print(flush=True)
    print("Done. Results:", flush=True)
    for shoot in shoots:
        print(f"  {os.path.basename(shoot)}:", flush=True)
        print(f"    Match : {os.path.join(shoot, 'match')}  (all matched pairs, flat)", flush=True)
        out = (os.path.join(args.out, os.path.basename(shoot)) if batch else args.out) \
            if args.out else os.path.join(shoot, "_out")
        print(f"    Excel : {os.path.join(out, 'photos.xlsx')}", flush=True)
        print(f"    Report: {os.path.join(out, 'report.csv')}", flush=True)


if __name__ == "__main__":
    main()