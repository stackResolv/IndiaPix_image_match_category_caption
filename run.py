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
                   "--report", os.path.join(out, "report.csv"),
                   "--out", os.path.join(out, "photos.xlsx"),
                   "--provider", args.provider]
        if args.model:
            cat_cmd += ["--model", args.model]
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


if __name__ == "__main__":
    main()