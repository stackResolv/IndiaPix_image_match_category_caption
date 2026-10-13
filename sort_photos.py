#!/usr/bin/env python3
"""
sort_photos.py — group edited photos with their camera originals.

Goal: given a shoot that contains a folder of camera originals and a folder of
edited derivatives, produce (1) a CSV report mapping every edit to its original
and (2) a grouped output folder where each original sits next to its edit(s).

Matching strategy, cheapest / most reliable first:

  Tier 1  EXIF DateTimeOriginal match ......... deterministic, free, ~99% on
                                                 Photoshop-edited sets because
                                                 Photoshop preserves the capture
                                                 timestamp.
  Tier 2  pHash disambiguation ................ when several originals share the
                                                 same capture second (bursts), the
                                                 perceptual hash of the edit is
                                                 compared only against those few
                                                 candidates.
  Tier 3  pHash fallback ...................... edits with no timestamp match are
                                                 hashed against every original; a
                                                 match is accepted only within a
                                                 distance threshold.

Everything runs locally (Pillow + imagehash). No cloud services, no cost, and the
tool never modifies, moves or renames the source files — it only copies.
"""

import argparse
import csv
import os
import shutil
import sys
from pathlib import Path

from PIL import Image, ExifTags
import imagehash
import numpy as np


IMG_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp", ".bmp"}

# EXIF tag ids we care about
TAG_DATETIME = 306            # DateTime
TAG_DATETIME_ORIGINAL = 36867  # DateTimeOriginal (capture time)
TAG_SUBSEC_TIME_ORIGINAL = 37520  # SubsecTimeOriginal

# Heuristic folder-name hints for auto-detection
ORIGINAL_HINTS = ("original", "orig", "camera", "raw", "source", "unedit")
EDIT_HINTS = ("edit", "retouch", "final", "export", "derived", "retouched")


# --------------------------------------------------------------------------- #
# I/O helpers
# --------------------------------------------------------------------------- #
def list_images(folder: str):
    """Return a sorted list of image file paths directly inside `folder`."""
    if not os.path.isdir(folder):
        return []
    out = []
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and os.path.splitext(name)[1].lower() in IMG_EXTS:
            out.append(path)
    return out


def read_capture_time(path: str):
    """Return the capture timestamp string (DateTimeOriginal) or None."""
    try:
        with Image.open(path) as im:
            exif = im.getexif()
            try:
                ifd = exif.get_ifd(ExifTags.IFD.Exif)
            except Exception:
                ifd = {}
            dto = ifd.get(TAG_DATETIME_ORIGINAL)
            if dto:
                return str(dto)
            dt = exif.get(TAG_DATETIME)
            if dt:
                return str(dt)
    except Exception:
        return None
    return None


def phash_of(path: str, hash_size: int = 16, normalize_orientation: bool = True):
    """Perceptual hash, normalised to landscape so 90-degree rotations still match."""
    try:
        with Image.open(path) as im:
            im = im.convert("L")
            if normalize_orientation and im.height > im.width:
                im = im.rotate(-90, expand=True)
            return imagehash.phash(im, hash_size=hash_size)
    except Exception:
        return None


def gray_vec(path: str, n: int = 64, normalize_orientation: bool = True):
    """Downscaled, contrast-normalised grayscale vector for MSE comparison.

    This keeps far more spatial detail than pHash and is used to pick the right
    frame inside a same-second burst, and to score how confident that pick is.
    """
    try:
        with Image.open(path) as im:
            im = im.convert("L")
            if normalize_orientation and im.height > im.width:
                im = im.rotate(-90, expand=True)
            im = im.resize((n, n))
            a = np.asarray(im, dtype=np.float32)
        return (a - a.mean()) / (a.std() + 1e-6)
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# Shoot / folder resolution
# --------------------------------------------------------------------------- #
def detect_subfolders(shoot_dir: str, originals_name, edits_name):
    """Work out which sub-folder holds the originals and which holds the edits."""
    subdirs = [d for d in sorted(os.listdir(shoot_dir))
               if os.path.isdir(os.path.join(shoot_dir, d))]
    if not subdirs:
        return None, None

    odir = edir = None

    # explicit names
    if originals_name and originals_name in subdirs:
        odir = originals_name
    if edits_name and edits_name in subdirs:
        edir = edits_name

    # name heuristics
    if odir is None:
        for d in subdirs:
            if any(h in d.lower() for h in ORIGINAL_HINTS):
                odir = d
                break
    if edir is None:
        for d in subdirs:
            if any(h in d.lower() for h in EDIT_HINTS):
                edir = d
                break

    # if we found one side and there are exactly two sub-folders, the other is the pair
    if odir and not edir:
        others = [d for d in subdirs if d != odir]
        if len(others) == 1:
            edir = others[0]
    if edir and not odir:
        others = [d for d in subdirs if d != edir]
        if len(others) == 1:
            odir = others[0]

    return odir, edir


def resolve_shoots(args):
    """Return a list of (shoot_path, shoot_name) tuples to process."""
    if args.shoot:
        p = Path(args.shoot)
        if not p.is_dir():
            sys.exit(f"error: --shoot '{args.shoot}' is not a directory")
        return [(str(p), p.name)]

    if args.root:
        root = Path(args.root)
        if not root.is_dir():
            sys.exit(f"error: --root '{args.root}' is not a directory")
        shoots = [(str(d), d.name) for d in sorted(root.iterdir()) if d.is_dir()]
        if not shoots:
            sys.exit(f"error: no sub-folders found in '{args.root}'")
        return shoots

    sys.exit("error: provide either --shoot <folder> or --root <parent-folder>")


# --------------------------------------------------------------------------- #
# Core processing
# --------------------------------------------------------------------------- #
def process_shoot(shoot_dir: str, out_dir: str, args):
    odir_name, edir_name = detect_subfolders(shoot_dir, args.originals_name, args.edits_name)
    if not odir_name or not edir_name:
        print(f"  [skip] could not find original/edit sub-folders in '{shoot_dir}'")
        return None

    originals = list_images(os.path.join(shoot_dir, odir_name))
    edits = list_images(os.path.join(shoot_dir, edir_name))

    print(f"  originals folder : {odir_name}  ({len(originals)} images)")
    print(f"  edits folder     : {edir_name}  ({len(edits)} images)")

    if not edits:
        print("  [skip] no edits found")
        return None

    # ---- index originals by capture time -------------------------------- #
    by_time = {}
    for o in originals:
        t = read_capture_time(o)
        if t:
            by_time.setdefault(t, []).append(o)

    # lazy caches
    orig_hash_cache = {}

    def orig_hash(o):
        if o not in orig_hash_cache:
            orig_hash_cache[o] = phash_of(o, hash_size=args.hash_size)
        return orig_hash_cache[o]

    # ---- match every edit ----------------------------------------------- #
    match_map = {}  # edit_path -> (original_path|None, method, confidence, distance)
    for e in edits:
        t = read_capture_time(e)
        candidates = by_time.get(t, []) if t else []

        # Tier 1: unique capture-time match
        if len(candidates) == 1:
            match_map[e] = (candidates[0], "exif", 1.0, 0, None)
            continue

        # Tier 2: burst — several originals share the capture second.
        # Disambiguate with downscaled-grayscale MSE (keeps more detail than
        # pHash). The edit is always attached to the correct *burst*; the margin
        # between the best and second-best candidate tells us how sure we are of
        # the exact frame within it.
        if len(candidates) > 1:
            ev = gray_vec(e)
            scored = []
            if ev is not None:
                for c in candidates:
                    cv = gray_vec(c)
                    if cv is None:
                        continue
                    scored.append((float(np.mean((ev - cv) ** 2)), c))
            if scored:
                scored.sort(key=lambda t: t[0])
                best, best_d = scored[0][1], scored[0][0]
                if len(scored) > 1 and best_d > 0:
                    margin = scored[1][0] / best_d
                else:
                    margin = float("inf")
            else:
                best, best_d, margin = candidates[0], "", 0.0
            if margin == float("inf"):
                conf = 1.0
            elif margin >= 1.5:
                conf = 0.9
            elif margin >= 1.2:
                conf = 0.6
            else:
                conf = 0.3  # near-duplicate frames: same burst, frame uncertain
            match_map[e] = (best, "exif+burst", conf, round(best_d, 2) if isinstance(best_d, float) else best_d, round(margin, 2) if margin != float("inf") else "inf")
            continue

        # Tier 3: no timestamp match — pHash against all originals
        eh = phash_of(e, hash_size=args.hash_size)
        best, best_d = None, None
        if eh is not None:
            for o in originals:
                oh = orig_hash(o)
                if oh is None:
                    continue
                d = eh - oh
                if best_d is None or d < best_d:
                    best, best_d = o, d

        if best is not None and isinstance(best_d, int) and best_d <= args.phash_threshold:
            conf = round(max(0.0, 1.0 - best_d / (args.phash_threshold * 2)), 2)
            match_map[e] = (best, "phash", conf, best_d, None)
        else:
            match_map[e] = (None, "unmatched", 0.0, best_d if best_d is not None else "", None)

    # ---- group by matched original -------------------------------------- #
    groups = {}
    unmatched = []
    for e, (o, method, conf, dist, margin) in match_map.items():
        if o:
            groups.setdefault(o, []).append(e)
        else:
            unmatched.append(e)

    # ---- write output ---------------------------------------------------- #
    groups_dir = os.path.join(out_dir, "groups")
    unmatched_dir = os.path.join(out_dir, "unmatched_edits")
    os.makedirs(groups_dir, exist_ok=True)
    if unmatched:
        os.makedirs(unmatched_dir, exist_ok=True)

    group_ids = {}
    for i, o in enumerate(sorted(groups), 1):
        gid = f"{i:04d}"
        group_ids[o] = gid
        gdir = os.path.join(groups_dir, f"{gid}_{Path(o).stem}")
        os.makedirs(gdir, exist_ok=True)
        shutil.copy2(o, os.path.join(gdir, os.path.basename(o)))
        for e in groups[o]:
            shutil.copy2(e, os.path.join(gdir, os.path.basename(e)))

    for e in unmatched:
        shutil.copy2(e, os.path.join(unmatched_dir, os.path.basename(e)))

    # ---- write CSV report ------------------------------------------------ #
    report_path = os.path.join(out_dir, "report.csv")
    with open(report_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["group_id", "edit_file", "matched_original", "method", "confidence", "distance", "burst_margin"])
        for e in sorted(match_map):
            o, method, conf, dist, margin = match_map[e]
            w.writerow([
                group_ids.get(o, ""),
                os.path.basename(e),
                os.path.basename(o) if o else "",
                method,
                conf,
                dist,
                margin,
            ])

    # ---- summary --------------------------------------------------------- #
    by_method = {}
    for _o, method, _c, _d, _m in match_map.values():
        by_method[method] = by_method.get(method, 0) + 1

    print(f"  matched groups   : {len(groups)}")
    print(f"  edits matched    : {len(match_map) - len(unmatched)}/{len(edits)}")
    for m in sorted(by_method):
        print(f"    - {m:<10}: {by_method[m]}")
    if unmatched:
        print(f"  unmatched edits  : {len(unmatched)} (see unmatched_edits/)")
    print(f"  report           : {report_path}")

    return {
        "shoot": shoot_dir,
        "out": out_dir,
        "groups": len(groups),
        "edits": len(edits),
        "matched": len(match_map) - len(unmatched),
        "unmatched": len(unmatched),
        "report": report_path,
    }


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser():
    p = argparse.ArgumentParser(
        description="Group edited photos with their camera originals (local, free).",
    )
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--shoot", help="path to a single shoot folder (contains originals + edits sub-folders)")
    src.add_argument("--root", help="path to a parent folder; every sub-folder is treated as a shoot")

    p.add_argument("--out", required=True, help="output folder (created if missing; sources are never modified)")
    p.add_argument("--originals-name", default="Camera Original ID",
                   help="name of the originals sub-folder (default: 'Camera Original ID')")
    p.add_argument("--edits-name", default="Edit",
                   help="name of the edits sub-folder (default: 'Edit')")
    p.add_argument("--hash-size", type=int, default=16,
                   help="pHash size (bits per side; total bits = size^2). default 16")
    p.add_argument("--phash-threshold", type=int, default=48,
                   help="max pHash hamming distance to accept a Tier-3 match. default 48")
    return p


def main():
    args = build_parser().parse_args()
    shoots = resolve_shoots(args)
    batch = len(shoots) > 1 or args.root is not None

    print(f"Processing {len(shoots)} shoot(s)")
    summaries = []
    for shoot_path, shoot_name in shoots:
        out_dir = os.path.join(args.out, shoot_name) if batch else args.out
        print(f"\n[{shoot_name}] -> {out_dir}")
        summary = process_shoot(shoot_path, out_dir, args)
        if summary:
            summaries.append(summary)

    print("\n=== Done ===")
    tot_edits = sum(s["edits"] for s in summaries)
    tot_matched = sum(s["matched"] for s in summaries)
    if tot_edits:
        print(f"{tot_matched}/{tot_edits} edits matched "
              f"({tot_matched / tot_edits * 100:.1f}%) across {len(summaries)} shoot(s).")


if __name__ == "__main__":
    main()