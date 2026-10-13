#!/usr/bin/env python3
"""
run.py — cross-platform one-command runner for the photo pipeline.
Works on Windows, macOS and Linux (uses only Python, no shell scripting).

Usage:
    python run.py                              # defaults: shoot "_sort images" -> "sorted_output"
    python run.py "MyShoot" "my_output"        # custom shoot folder + output folder
    python run.py --edit-dir "MyShoot/Edits" --originals-dir "MyShoot/RAW"

Runs both steps:
    1. sort_photos.py  -> groups each edit with its matching original
    2. categorize.py   -> adds category + caption, writes Excel
"""

import argparse
import subprocess
import sys
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def run(cmd):
    print("$ " + " ".join(cmd))
    subprocess.check_call(cmd, cwd=HERE)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("shoot", nargs="?", default="_sort images",
                    help="shoot folder name (default: '_sort images')")
    ap.add_argument("out", nargs="?", default="sorted_output",
                    help="output folder (default: 'sorted_output')")
    ap.add_argument("--originals-name", default="Camera Original ID",
                    help="name of the originals sub-folder")
    ap.add_argument("--edits-name", default="Edit",
                    help="name of the edits sub-folder")
    args = ap.parse_args()

    shoot = args.shoot
    out = args.out
    edit_dir = os.path.join(shoot, args.edits_name)

    print("=" * 66)
    print(" Step 1/2: matching edits to originals")
    print(f"   shoot : {shoot}")
    print(f"   out   : {out}")
    print("=" * 66)
    run([sys.executable, "sort_photos.py",
         "--shoot", shoot, "--out", out,
         "--originals-name", args.originals_name,
         "--edits-name", args.edits_name])

    print()
    print("=" * 66)
    print(" Step 2/2: categorizing + captioning edits")
    print("=" * 66)
    run([sys.executable, "categorize.py",
         "--edit-dir", edit_dir,
         "--report", os.path.join(out, "report.csv"),
         "--out", os.path.join(out, "photos.xlsx")])

    print()
    print("Done. Results:")
    print(f"  Excel   : {os.path.join(out, 'photos.xlsx')}")
    print(f"  Report  : {os.path.join(out, 'report.csv')}")
    print(f"  Groups  : {os.path.join(out, 'groups')}")
    print(f"  Debug   : {os.path.join(out, 'categorize_debug.log')}")


if __name__ == "__main__":
    main()