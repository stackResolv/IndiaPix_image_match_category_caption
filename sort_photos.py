#!/usr/bin/env python3
"""
sort_photos.py — group edited photos with their camera originals.

Goal: given a shoot that contains a folder of camera originals and a folder of
edited derivatives, produce (1) a CSV report mapping every edit to its original
and (2) a single flat "match" folder INSIDE the shoot that holds every matched
original next to its edit(s). Edits with no matching original are ignored.

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
import concurrent.futures as cf
import csv
import os
import shutil
import sys
import threading
import time
from pathlib import Path

import imagehash
import numpy as np

from imageutil import IMG_EXTS, capture_time, load_image


# Windows-safe console: force UTF-8 with a lossy fallback so filenames that
# contain non-ASCII characters never crash printing on a cp1252 console.
from consoleutil import hr, safe, setup_console, term_width, wrap

# Console prep: UTF-8 where possible, and the real width so the narration fits
# a narrow Windows PowerShell window instead of wrapping into nonsense.
setup_console()

# Wall-clock start, used by log() so long runs never look "stuck".
_T0 = time.time()


def log(msg, indent=2):
    """Print a timestamped progress line, wrapped to the terminal width.

    The timestamp is printed once; continuation lines are aligned under the
    message so a wrapped explanation still reads as one block.
    """
    stamp = f"[{time.time() - _T0:6.1f}s] "
    text = safe(str(msg))
    lead = len(text) - len(text.lstrip())      # honour the caller's "  what:" indent
    body = text.strip()
    col = indent + len(stamp) + lead
    lines = wrap(body, indent=col, subsequent_indent=" " * 2)
    print(" " * indent + stamp + " " * lead + lines[0], flush=True)
    for extra in lines[1:]:
        print(" " * col + extra, flush=True)


def step(n, text, total=5):
    """Announce a pipeline step so the log reads as a story, not just numbers."""
    print(flush=True)
    print("  " + hr("-", indent=2), flush=True)
    for line in wrap(safe(f"Step {n}/{total}: {text}"),
                     indent=4, subsequent_indent=" " * 4):
        print("  " + line, flush=True)
    print("  " + hr("-", indent=2), flush=True)


def note(text):
    """A plain explanatory line, wrapped and indented under the current step."""
    for line in wrap(safe(text), indent=6, subsequent_indent=" " * 2):
        print(" " * 4 + line, flush=True)


def describe_match(res):
    """Plain-English sentence explaining one match result."""
    o, method, conf, dist, margin = res
    if method == "exif":
        return f"EXIF timestamp match -> {os.path.basename(o)} (certain)"
    if method == "exif+burst":
        return (f"burst tie broken by pixel comparison -> {os.path.basename(o)} "
                f"(confidence {conf}, margin {margin})")
    if method == "phash":
        return (f"no timestamp; closest image by perceptual hash -> "
                f"{os.path.basename(o)} (distance {dist}, confidence {conf})")
    return f"no confident match (closest distance {dist}) - flagged for review"


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


def read_capture_time(path: str, fast: bool = False):
    """Return the capture timestamp string (DateTimeOriginal) or None.

    Delegates to imageutil.capture_time so RAW files (e.g. Canon .cr3) are
    handled as well as ordinary raster files. ``fast`` limits RAW handling to
    the bounded TIFF/Exif scan (see imageutil.capture_time).
    """
    return capture_time(path, fast=fast)


def phash_of(path: str, hash_size: int = 16, normalize_orientation: bool = True):
    """Perceptual hash, normalised to landscape so 90-degree rotations still match."""
    try:
        im = load_image(path)
        if im is None:
            return None
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
        im = load_image(path)
        if im is None:
            return None
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

    step(1, "locating the originals and edits folders")
    note(f"originals : {odir_name}  ({len(originals)} image(s))")
    note(f"edits     : {edir_name}  ({len(edits)} image(s))")
    note(f"output    : {out_dir}  (report written here)")
    note(f"match dir : {os.path.join(shoot_dir, 'match')}")

    if not edits:
        note("[skip] no edits found in the edits folder")
        return None
    if not originals:
        note("[warn] no originals found - every edit will be unmatched")

    # ---- index originals by capture time -------------------------------- #
    # Metadata-only reads are IO-bound, so they run on a thread pool — a
    # 4,000-file folder indexes in seconds instead of stalling. capture_time is
    # called with fast=True, which skips rawpy's whole-container parser: that
    # parser can sit for many seconds on one odd RAW file and is what made the
    # counter appear frozen. Unparseable RAW files just get no timestamp (they
    # can still match via perceptual hash), so nothing is lost.
    idx_workers = max(4, min(32, (os.cpu_count() or 4) * 4))
    step(2, f"reading capture timestamps from {len(originals)} original(s)")
    note("what: the date+time each frame was shot (EXIF DateTimeOriginal).")
    note("why : Photoshop keeps that stamp, so it is the join key back to the "
         "original frame - a timestamp match is exact, no guessing.")
    note(f"how : reads only the metadata block (no image decoding), "
         f"{idx_workers} threads.")

    in_flight = {}          # thread name -> (index, basename, started_at)
    flight_lock = threading.Lock()
    stop_hb = threading.Event()

    def _index_heartbeat():
        while not stop_hb.wait(15):
            now = time.time()
            with flight_lock:
                stuck = [(n, f, now - t) for n, (_, f, t) in in_flight.items()
                         if now - t >= 15]
            for _, fname, waited in stuck:
                log(f"  ... still reading '{fname}' ({waited:.0f}s on this file)")

    hb = threading.Thread(target=_index_heartbeat, daemon=True)
    hb.start()

    def _one(item):
        i, o = item
        tn = threading.current_thread().name
        name = os.path.basename(o)
        with flight_lock:
            in_flight[tn] = (i, name, time.time())
        try:
            return o, read_capture_time(o, fast=True), time.time()
        finally:
            with flight_lock:
                in_flight.pop(tn, None)

    by_time = {}
    t_index = time.time()
    try:
        with cf.ThreadPoolExecutor(max_workers=idx_workers) as ex:
            futures = [ex.submit(_one, item) for item in enumerate(originals, 1)]
            for done, fut in enumerate(cf.as_completed(futures), 1):
                o, t, started = fut.result()
                name = os.path.basename(o)
                if t:
                    by_time.setdefault(t, []).append(o)
                if len(originals) <= 100 or done % 100 == 0 or done == len(originals):
                    elapsed = time.time() - t_index
                    rate = done / elapsed if elapsed > 0 else 0.0
                    eta = (len(originals) - done) / rate if rate > 0 else 0.0
                    tail = f", ~{eta:.0f}s left" if done < len(originals) else ""
                    log(f"  originals read: {done}/{len(originals)} "
                        f"({100.0 * done / len(originals):.0f}%) [{elapsed:.0f}s{tail}]")
    finally:
        stop_hb.set()
    n_nostamp = len(originals) - sum(len(v) for v in by_time.values())
    n_burst = sum(1 for v in by_time.values() if len(v) > 1)
    log(f"  done in {time.time() - t_index:.1f}s: {len(by_time)} distinct "
        f"timestamp(s) across {len(originals)} file(s)")
    log(f"  {n_burst} of those timestamp(s) are shared by several frames "
        f"(bursts) - those need visual disambiguation in step 3")
    if n_nostamp:
        log(f"  {n_nostamp} file(s) have no readable timestamp; they can still "
            f"be reached by perceptual hash")

    # lazy caches
    orig_hash_cache = {}

    def orig_hash(o):
        if o not in orig_hash_cache:
            orig_hash_cache[o] = phash_of(o, hash_size=args.hash_size)
        return orig_hash_cache[o]

    def match_one(e):
        """Match a single edit. Returns (original|None, method, conf, dist, margin)."""
        # fast=True: bounded metadata read, so one odd file cannot freeze the
        # loop. A miss here only costs a Tier-3 visual match, never a wrong row.
        t = read_capture_time(e, fast=True)
        candidates = by_time.get(t, []) if t else []

        # Tier 1: unique capture-time match
        if len(candidates) == 1:
            return (candidates[0], "exif", 1.0, 0, None)

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
            return (best, "exif+burst", conf,
                    round(best_d, 2) if isinstance(best_d, float) else best_d,
                    round(margin, 2) if margin != float("inf") else "inf")

        # Tier 3: no timestamp match — pHash against all originals. This is the
        # slow path (it decodes every original), so say so before it starts.
        log(f"  no capture-time match for {os.path.basename(e)} - comparing "
            f"against all {len(originals)} original(s) (slow path) ...")
        eh = phash_of(e, hash_size=args.hash_size)
        best, best_d = None, None
        if eh is not None:
            for o_idx, o in enumerate(originals, 1):
                oh = orig_hash(o)
                if oh is not None:
                    d = eh - oh
                    if best_d is None or d < best_d:
                        best, best_d = o, d
                if o_idx % 250 == 0:
                    log(f"    ... compared {o_idx}/{len(originals)} original(s)")

        if best is not None and isinstance(best_d, int) and best_d <= args.phash_threshold:
            conf = round(max(0.0, 1.0 - best_d / (args.phash_threshold * 2)), 2)
            return (best, "phash", conf, best_d, None)
        return (None, "unmatched", 0.0, best_d if best_d is not None else "", None)

    # ---- match every edit ----------------------------------------------- #
    step(3, f"matching {len(edits)} edit(s) to their original frames")
    note("what: for each edit, find the frame it was made from.")
    note("how : 1) same capture second? -> certain (free).")
    note("      2) several frames share that second (a burst)? -> compare pixels.")
    note("      3) no timestamp at all? -> compare perceptual hashes.")
    note("Only 2 and 3 look at pixels, so the run stays fast.")
    verbose = len(edits) <= 40
    if verbose:
        note("This shoot is small, so every match is explained line by line below.")
    match_map = {}  # edit_path -> (original_path|None, method, confidence, distance, margin)
    t_match = time.time()
    for e_idx, e in enumerate(edits, 1):
        res = match_one(e)
        match_map[e] = res
        if verbose:
            log(f"[{e_idx}/{len(edits)}] {os.path.basename(e)}", indent=4)
            log(f"    -> {describe_match(res)}", indent=4)
        # Report *completed* edits (not the one about to start). Every edit is
        # printed for small shoots; otherwise every 5th, plus the final one.
        if len(edits) <= 20 or e_idx == 1 or e_idx % 5 == 0 or e_idx == len(edits):
            elapsed = time.time() - t_match
            rate = e_idx / elapsed if elapsed > 0 else 0.0
            eta = (len(edits) - e_idx) / rate if rate > 0 else 0.0
            pct = 100.0 * e_idx / len(edits)
            tail = f", ~{eta:.0f}s left" if e_idx < len(edits) else ""
            log(f"  progress: {e_idx}/{len(edits)} edit(s) matched ({pct:.0f}%)"
                f" [{elapsed:.0f}s{tail}]")
    log(f"  done in {time.time() - t_match:.1f}s")

    # How each edit was actually resolved, so the log shows the quality you got.
    from collections import Counter
    method_counts = Counter(r[1] for r in match_map.values())
    labels = {
        "exif": "certain - unique capture timestamp",
        "exif+burst": "timestamp + pixel check (burst frame chosen)",
        "phash": "visual only, no timestamp - worth a glance",
        "unmatched": "no match found - needs manual review",
    }
    log("  how the matches were made:")
    for key in ("exif", "exif+burst", "phash", "unmatched"):
        if method_counts.get(key):
            log(f"{method_counts[key]:5d}  {labels[key]}", indent=4)
    low_conf = [os.path.basename(e) for e, r in match_map.items()
                if r[1] == "exif+burst" and r[2] <= 0.3]
    if low_conf:
        log(f"{len(low_conf)} burst match(es) could not pick the exact frame "
            f"confidently - check these first:", indent=2)
        for nm in low_conf[:10]:
            log(f"- {nm}", indent=4)
        if len(low_conf) > 10:
            log(f"... and {len(low_conf) - 10} more (filter confidence=0.3 "
                f"in report.csv)", indent=4)

    # ---- group by matched original -------------------------------------- #
    groups = {}
    unmatched = []
    for e, (o, method, conf, dist, margin) in match_map.items():
        if o:
            groups.setdefault(o, []).append(e)
        else:
            unmatched.append(e)

    # ---- write output: one flat 'match' folder inside the shoot ---------- #
    # Every matched original is copied next to its edit(s) directly inside
    # <shoot>/match/. There are NO sub-folders: all pairs live flat in 'match'.
    # Edits with no matching original are simply ignored (not copied anywhere).
    os.makedirs(out_dir, exist_ok=True)
    match_dir = os.path.join(shoot_dir, "match")
    os.makedirs(match_dir, exist_ok=True)

    seen = {}  # basename -> source already copied under that name

    def copy_flat(src):
        base = os.path.basename(src)
        first = seen.get(base)
        if first is None or first == src:
            seen[base] = src
            dest = os.path.join(match_dir, base)
        else:
            # Two different source files share a name — give the later one a
            # unique suffix so nothing is overwritten.
            stem, ext = os.path.splitext(base)
            n = 1
            dest = os.path.join(match_dir, f"{stem}_{n}{ext}")
            while os.path.exists(dest):
                n += 1
                dest = os.path.join(match_dir, f"{stem}_{n}{ext}")
        shutil.copy2(src, dest)

    ordered_groups = sorted(groups)
    step(4, f"copying {len(ordered_groups)} matched pair(s) into the match folder")
    note("what: each matched original is copied next to its edit(s).")
    note("note: files are only COPIED - your sources are never moved or renamed.")
    note(f"{len(unmatched)} unmatched edit(s) are skipped entirely.")
    multi = sum(1 for v in groups.values() if len(v) > 1)
    if multi:
        note(f"{multi} original(s) have more than one edit (versions).")
    group_ids = {}
    n_files_copied = 0
    for i, o in enumerate(ordered_groups, 1):
        group_ids[o] = f"{i:04d}"
        copy_flat(o)
        n_files_copied += 1
        for e in groups[o]:
            copy_flat(e)
            n_files_copied += 1
        if i % 20 == 0 or i == len(ordered_groups):
            log(f"  groups copied: {i}/{len(ordered_groups)} "
                f"({100.0 * i / len(ordered_groups):.0f}%)")
    log(f"  {n_files_copied} file(s) copied in total")

    # ---- write CSV report ------------------------------------------------ #
    step(5, "writing the report that links each edit to its original")
    note(f"file: {os.path.join(out_dir, 'report.csv')}")
    note("one row per edit: group id, edit name, matched original, match "
         "method, confidence and distance.")
    report_path = os.path.join(out_dir, "report.csv")
    with open(report_path, "w", newline="", encoding="utf-8") as f:
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
    print(flush=True)
    print("  " + hr("-", indent=2), flush=True)
    print(safe("  RESULT"), flush=True)
    note(f"edits matched   : {len(match_map) - len(unmatched)} of {len(edits)}")
    note(f"edits unmatched : {len(unmatched)}")
    note(f"original groups : {len(groups)}")
    note(f"match folder    : {match_dir}")
    note(f"report          : {report_path}")
    print("  " + hr("-", indent=2), flush=True)

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
    p.add_argument("--originals-name", default="JPEG",
                   help="name of the originals sub-folder (default: 'JPEG')")
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

    print(safe(f"Processing {len(shoots)} shoot(s)"))
    summaries = []
    for shoot_path, shoot_name in shoots:
        out_dir = os.path.join(args.out, shoot_name) if batch else args.out
        print(flush=True)
        for line in wrap(safe(f"[{shoot_name}] -> {out_dir}"),
                         indent=0, subsequent_indent=" " * 4):
            print(line, flush=True)
        summary = process_shoot(shoot_path, out_dir, args)
        if summary:
            summaries.append(summary)

    print(flush=True)
    print(hr("="))
    print(safe("Done"))
    print(hr("=", indent=0))
    tot_edits = sum(s["edits"] for s in summaries)
    tot_matched = sum(s["matched"] for s in summaries)
    if tot_edits:
        print(safe(f"{tot_matched}/{tot_edits} edits matched "
                   f"({tot_matched / tot_edits * 100:.1f}%) across "
                   f"{len(summaries)} shoot(s)."))


if __name__ == "__main__":
    main()