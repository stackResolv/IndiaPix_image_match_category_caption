#!/usr/bin/env python3
"""
sets_pipeline.py — metadata + captions when the input is organised as
"one folder per set".

Client layout:

    input_folder/
        IMG_0001/            <- set 1
            _77A0001.CR3     <- camera original
            _77A0001.JPG     <- its edit
        IMG_0002/            <- set 2
        ...
        IMG_40000/           <- set 40000

30,000-40,000 set folders, two photos each. There are no `JPEG` / `Edit`
sub-folders, so the *role* of each photo (camera original vs edited derivative)
has to be worked out from the files themselves, not from a folder name.

Five steps:

  1. discover  — walk the root and list every set folder (extra nesting is fine).
  2. inspect   — read each photo's metadata ONCE (format, pixels, ICC, both EXIF
                 dates, EXIF Software, byte size). No image is decoded.
  3. pair      — decide which photo is the original from that metadata, and
                 write report.csv so the decision is auditable *before* any
                 money is spent on captions.
  4. classify  — send the original and its edit to the vision model together,
                 one request per set, exactly like categorize.py.
  5. write     — stream the IndiaPix template sheet (17 columns + `category`)
                 to Excel, plus a debug log.

The Excel is identical to the shoot workflow: one `00 original` row followed by
one `final` row per edit, the same 17 template columns in the same order, the
same `SET_000001` ids, the same trailing `category` column and the same colour
coding. Only the way the pairs are found differs.

Nothing is copied, moved or renamed — each set folder already holds its pair.

Usage:

    # 1) check the pairing decisions first: no API calls, no cost
    python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output" --dry-run

    # 2) the real run (resumable: a re-run never pays twice for the same set)
    python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output"

    # a cheap pilot on the first 200 sets before committing to 40,000
    python3 sets_pipeline.py --sets-root "input_folder" --out "sorted_output" --limit 200
"""

import argparse
import concurrent.futures as cf
import csv
import json
import os
import sys
import threading
import time

import imageutil
import categorize as cat
from consoleutil import hr, safe, setup_console, wrap
from excelutil import ALL_COLUMNS, StreamedSheet

setup_console()

_T0 = time.time()

# Label cache, keyed by "<set>/<edit>|<original>". Versioned because the prompt
# decides which fields exist — a new cache name avoids mixing shapes. v6 keeps
# the 9-10 meaningful-word title (articles excluded) and the 13+ keyword list
# from v5, but descriptions are back to ONE sentence of about 20 words instead
# of the multi-sentence 35-70 word text v5 held. v3 added the `edit_category`
# column, v2 was the client's eight-operation vocabulary, v1 held old op names.
CACHE_NAME = "labels_cache_sets_v6.json"

# Progress is checkpointed this often, so an interrupted 40,000-set run loses
# minutes of API spend, not hours.
CHECKPOINT_EVERY = 100

# A set is normally exactly two photos.
EXPECTED_PHOTOS = 2

# Name hints, matched against the file stem split on non-alphanumeric chars.
# Only ever corroboration: "final", "raw" and friends appear in real camera
# filenames too, so no single hint is trusted on its own.
_EDIT_NAME_HINTS = {
    "edit", "edited", "final", "fin", "retouch", "retouched", "copy", "mod",
    "modified", "export", "exported", "corrected", "done", "fix", "fixed",
    "print", "web", "social", "select", "selected",
}
_ORIG_NAME_HINTS = {
    "orig", "original", "raw", "camera", "unedit", "unedited", "source",
    "master", "neg", "negative",
}


def log(msg, indent=2):
    """Timestamped, width-wrapped progress line."""
    stamp = f"[{time.time() - _T0:7.1f}s] "
    text = safe(str(msg))
    lead = len(text) - len(text.lstrip())
    body = text.strip()
    col = indent + len(stamp) + lead
    lines = wrap(body, indent=col, subsequent_indent=" " * 2)
    print(" " * indent + stamp + " " * lead + lines[0], flush=True)
    for extra in lines[1:]:
        print(" " * col + extra, flush=True)


def step(n, text, total=5):
    print(flush=True)
    print("  " + hr("-", indent=2), flush=True)
    for line in wrap(safe(f"Step {n}/{total}: {text}"), indent=4,
                     subsequent_indent=" " * 4):
        print("  " + line, flush=True)
    print("  " + hr("-", indent=2), flush=True)


def note(text):
    for line in wrap(safe(text), indent=6, subsequent_indent=" " * 2):
        print(" " * 4 + line, flush=True)


def fmt_eta(seconds):
    """'2h 14m' / '8m 30s' / '45s' — a 40,000-set run is measured in hours."""
    seconds = max(0, int(seconds))
    if seconds >= 3600:
        return f"{seconds // 3600}h {(seconds % 3600) // 60}m"
    if seconds >= 60:
        return f"{seconds // 60}m {seconds % 60}s"
    return f"{seconds}s"


def relpath(path, root):
    """Path relative to the input root, with forward slashes (readable in CSV)."""
    try:
        return os.path.relpath(path, root).replace(os.sep, "/")
    except ValueError:
        return path


def _operation_segments(set_name, root_name=""):
    """The folder names around this set, nearest-first, that could be the label.

    ``set_name`` is the set's path relative to the run root, so ``parts[:-1]``
    are its ancestors *inside* that root. The root itself is an ancestor too —
    and it is the one that carries the label when the user points the tool
    straight at the operation folder::

        --sets-root metadata           + sharpen_denoise/SET_1  -> parents
                                          ['sharpen_denoise']   (+root 'metadata')
        --sets-root metadata/sharpen_denoise + SET_1          -> parents []
                                          (+root 'sharpen_denoise')  <-- the label

    Without the root the second case loses the client's label entirely and the
    model's guess ends up leading the cell. The root name is therefore always
    considered last, i.e. it only wins when no folder inside the root qualifies.
    """
    parts = [p for p in str(set_name or "").replace("\\", "/").split("/") if p]
    segments = list(reversed(parts[:-1]))
    root_seg = os.path.basename(os.path.normpath(str(root_name or "").strip()))
    if root_seg and root_seg not in (".", "..") and root_seg not in segments:
        segments.append(root_seg)
    return segments


def operation_folder(set_name, root_name=""):
    """The edit-operation folder that contains this set, as an ordered list.

    Some deliveries are grouped by *what was done*, not by date:

        metadata/color_tonal/SET_000001/...
        metadata/sharpen_denoise/SET_000001/...
        metadata/color_tonal;white_balance/SET_000001/...

    The folder name is the client's own label for the batch, so it is
    authoritative and written FIRST. A folder can name more than one operation,
    in any separator a human might type — ``;``, ``,``, ``+``, ``&``, ``and`` —
    and the order inside the name is kept, because it states the primary
    operation first: ``color_tonal;white_balance`` means "graded, and also
    white-balanced".

    Ancestors are checked nearest-first (including the run root itself, so
    `--sets-root metadata/sharpen_denoise` still finds `sharpen_denoise`), and
    only a segment that yields at least one of the client's eight operations is
    used. A layout grouped by date (`input/2025-03-14/IMG_0001`) therefore
    yields ``[]`` and nothing changes.
    """
    for seg in _operation_segments(set_name, root_name):
        ops = cat.ops_ordered(seg)
        if ops:
            return ops
    return []


def operation_folder_label(set_name, root_name=""):
    """The raw folder name that supplied the operations, or ``''``.

    Kept for report.csv so a reviewer sees exactly what the client typed, even
    when it names several operations (``color_tonal;white_balance``).
    """
    for seg in _operation_segments(set_name, root_name):
        if cat.ops_ordered(seg):
            return seg
    return ""


def merge_ops(folder_op, model_ops_value, extra_ops=None, hint_text="",
              fallback=""):
    """Folder operations first, then the model's, then measured ones, then prose.

    Each source keeps its own order — the folder states the batch's primary
    operation, and the model answers most-important-first — so the final cell
    reads: what the client filed it as, then what else the model could see (by
    its own confidence), then anything the file metadata proves. Duplicates are
    dropped, keeping the earliest (highest-priority) position.

    The label's own prose (`hint_text`) is used **only when the cell would
    otherwise be empty** — never to extend a real answer. A description saying
    "scissors added" beside a model answer of `object_removal` is a model error;
    appending `compositing` would record an operation no model reported and no
    human confirmed. See `categorize.complete_ops`.

    `fallback` is written only when nothing at all was found.
    """
    cell, _, _ = cat.explain_ops(folder_op, model_ops_value, extra_ops,
                                 hint_text=hint_text, fallback=fallback)
    return cell


# --------------------------------------------------------------------------- #
# Step 1 — discover the set folders
# --------------------------------------------------------------------------- #
def _scan_dir(path):
    """(image files, sub-folders) directly inside `path`."""
    files, subdirs = [], []
    try:
        with os.scandir(path) as it:
            for e in it:
                try:
                    if e.is_dir():
                        subdirs.append(e.path)
                    elif e.is_file() and \
                            os.path.splitext(e.name)[1].lower() in imageutil.IMG_EXTS:
                        files.append(e.path)
                except OSError:
                    continue
    except OSError:
        return [], []
    return sorted(files), sorted(subdirs)


def discover_sets(root):
    """Return ``(sets, ambiguous)`` for the tree under `root`.

    ``sets`` is ``[(set_dir, [image paths])]``, one entry per folder that holds
    photos *and* has no sub-folder holding photos. That single rule covers both
    layouts a client might send:

        input_folder/IMG_0001/...          flat: 40,000 children
        input_folder/2025-03-14/IMG_0001/  nested: sets grouped by day

    ``ambiguous`` lists folders that hold loose photos *and* photo sub-folders;
    the sub-folders win and the loose files are reported.
    """
    found, ambiguous = [], []

    def walk(path):
        files, subdirs = _scan_dir(path)
        has_child_sets = False
        for d in subdirs:
            before = len(found)
            walk(d)
            if len(found) > before:
                has_child_sets = True
        if not files:
            return
        if has_child_sets:
            ambiguous.append((path, len(files)))
            return
        found.append((path, files))

    walk(os.path.abspath(root))
    found.sort(key=lambda t: t[0])
    return found, ambiguous


# --------------------------------------------------------------------------- #
# Step 3 — decide which photo of the pair is the original
# --------------------------------------------------------------------------- #
def name_hints(path):
    """``(looks_like_edit, looks_like_original)`` read from the file stem."""
    stem = os.path.splitext(os.path.basename(path))[0].lower()
    tokens = "".join(c if c.isalnum() else " " for c in stem).split()
    return (any(t in _EDIT_NAME_HINTS for t in tokens),
            any(t in _ORIG_NAME_HINTS for t in tokens))


def edit_score(info, other):
    """How strongly THIS photo looks like the edited derivative.

    Returns ``(score, fired)`` where ``fired`` is a list of ``(weight, label)``
    so report.csv can state why the decision was made. The weights encode how
    far each signal can be trusted:

      +-4  EXIF ``Software`` naming an editor — the edit was written by
           Photoshop/Lightroom/Capture One. Near-conclusive.
      -4  this file is a RAW and the other is not — a camera produces the RAW;
           an export pipeline produces the JPEG.  Near-conclusive.
      +-2  edit date (EXIF ``DateTime``) later than the capture date, or a
           filename hint ("final", "retouch").  Strong corroboration.
      +1   smaller pixel area, noticeably smaller byte size, later mtime.
           Weak: a resize or an aggressive crop can go either way.

    Positive = this file is the edit, negative = this file is the original.
    """
    fired = []

    if imageutil.editor_software(info.get("software")):
        fired.append((4, "exif-software"))

    if info.get("is_raw") and not other.get("is_raw"):
        fired.append((-4, "raw-file"))
    elif other.get("is_raw") and not info.get("is_raw"):
        fired.append((2, "other-is-raw"))

    cap = imageutil._dt_epoch(info.get("capture"))
    mod = imageutil._dt_epoch(info.get("modified"))
    if cap and mod and mod - cap > 86400:
        fired.append((2, "edit-date"))

    a, b = info.get("area") or 0, other.get("area") or 0
    if a and b and a != b:
        if a < b:
            fired.append((1, "smaller-area"))
        elif a > b * 1.25:
            fired.append((-1, "larger-area"))

    na, nb = info.get("nbytes") or 0, other.get("nbytes") or 0
    if na and nb and na < nb * 0.85:
        fired.append((1, "smaller-bytes"))

    is_edit_name, is_orig_name = name_hints(info.get("path", ""))
    if is_edit_name and not is_orig_name:
        fired.append((2, "name"))
    elif is_orig_name and not is_edit_name:
        fired.append((-2, "name-original"))

    ma, mb = info.get("modified_epoch") or 0, other.get("modified_epoch") or 0
    if ma and mb and ma - mb > 60:
        fired.append((1, "later-mtime"))

    return sum(w for w, _ in fired), fired


# Strongest signal first — this order names the decision in report.csv, so the
# `method` column reports the most trustworthy reason that fired (never the
# weakest), which is what a reviewer actually wants to see.
_SIGNAL_ORDER = ["exif-software", "raw-file", "other-is-raw", "edit-date",
                 "name", "name-original", "smaller-area", "larger-area",
                 "smaller-bytes", "later-mtime"]


def decide_roles(infos):
    """Choose the camera original inside one set.

    Returns ``(orig_index, edit_indices, method, confidence, note_text)``.
    ``method`` is the strongest signal that decided it (or ``"order"`` when all
    metadata was equal), and ``confidence`` is a review hint: 1.0 certain,
    0.9 high, 0.6 medium, 0.35 "no signal, picked deterministically".
    """
    n = len(infos)
    if n == 1:
        return 0, [], "single", 0.3, "set holds one photo - written as one final row"

    scores, fired = [], {}
    for i, info in enumerate(infos):
        if n == 2:
            s, f = edit_score(info, infos[1 - i])
        else:
            # 3+ photos: score against the most camera-like companion so a
            # single stray extra file cannot flip every decision.
            others = [o for j, o in enumerate(infos) if j != i]
            base = min(others, key=lambda o: (
                sum(edit_score(o, x)[0] for x in others if x is not o)))
            s, f = edit_score(info, base)
        scores.append(s)
        fired[i] = f

    if n == 2:
        if scores[0] != scores[1]:
            edit_i = 0 if scores[0] > scores[1] else 1
        else:
            # Perfect tie: the later-written file is the edit. Deterministic and
            # true far more often than not, but flagged at confidence 0.35.
            e0 = infos[0].get("modified_epoch") or 0
            e1 = infos[1].get("modified_epoch") or 0
            edit_i = 0 if e0 > e1 else 1
        gap = abs(scores[0] - scores[1])
        confidence = 1.0 if gap >= 4 else 0.9 if gap >= 2 else 0.6 if gap == 1 else 0.35

        positive = {label for w, label in fired[edit_i] if w > 0}
        method = next((s for s in _SIGNAL_ORDER if s in positive), "order")
        if method == "order":
            why = ("no distinguishing metadata - the later-written file was "
                   "treated as the edit")
        else:
            why = "decided by: " + ", ".join(
                s for s in _SIGNAL_ORDER if s in positive)
        return 1 - edit_i, [edit_i], method, confidence, why

    # 3+ photos: lowest score = most camera-like; everything else is an edit.
    # `method` says "multi-photo" rather than naming an edit signal: in a set of
    # three or more, the fact a reviewer needs is the shape of the set, not
    # which single signal happened to fire.
    orig_i = min(range(n), key=lambda i: scores[i])
    edits = [i for i in range(n) if i != orig_i]
    negative = {label for w, label in fired[orig_i] if w < 0}
    strongest = next((s for s in _SIGNAL_ORDER if s in negative), "")
    why = (f"set holds {n} photos - wrote 1 original row + {len(edits)} final "
           f"rows (original chosen by: {strongest or 'file order'})")
    return orig_i, edits, "multi-photo", 0.6, why


def task_key(set_name, edit_path, orig_path=""):
    """Stable cache key: '<set>/<edit>[|<original>]'."""
    key = f"{set_name}/{os.path.basename(edit_path)}"
    if orig_path:
        key += f"|{os.path.basename(orig_path)}"
    return key


# --------------------------------------------------------------------------- #
# Step 2 — metadata, read once per photo
# --------------------------------------------------------------------------- #
def inspect_all(paths, workers):
    """``{path: info}`` for every photo, read on a thread pool.

    One Open/mmap per file: no image is decoded, so 80,000 photos finishes in
    minutes. Progress is printed every 500 files with an ETA, because a silent
    counter on a multi-hour run is indistinguishable from a hang.
    """
    total = len(paths)
    out = {}
    lock = threading.Lock()
    state = {"done": 0}
    t0 = time.time()

    def one(path):
        info = imageutil.inspect_image(path)
        with lock:
            state["done"] += 1
            if state["done"] % 500 == 0 or state["done"] == total:
                elapsed = time.time() - t0
                rate = state["done"] / elapsed if elapsed > 0 else 0.0
                eta = (total - state["done"]) / rate if rate > 0 else 0.0
                log(f"  {state['done']:,}/{total:,} read "
                    f"({100.0 * state['done'] / total:.0f}%) "
                    f"[{fmt_eta(elapsed)} elapsed, ~{fmt_eta(eta)} left]")
        return path, info

    with cf.ThreadPoolExecutor(max_workers=workers) as ex:
        for path, info in ex.map(one, paths):
            out[path] = info
    log(f"  metadata read in {fmt_eta(time.time() - t0)}")
    return out


# --------------------------------------------------------------------------- #
# Step 3 — pair every set and write report.csv
# --------------------------------------------------------------------------- #
def pair_sets(sets, inspected, root):
    """Turn ``[(dir, files)]`` into one record per set, with roles decided."""
    records = []
    for set_dir, files in sets:
        infos = [inspected[p] for p in files if p in inspected]
        if not infos:
            continue
        set_name = relpath(set_dir, root)
        orig_i, edit_is, method, conf, why = decide_roles(infos)
        records.append({
            "set": set_name,
            "dir": set_dir,
            "files": files,
            "infos": infos,
            "orig_i": orig_i,
            "edit_is": edit_is,
            "method": method,
            "confidence": conf,
            "why": why,
            # The operation folder name that holds this set, VERBATIM
            # ("color_tonal", "white balance, crop perspective",
            # "color_tonal;white_balance"). The client's name is authoritative
            # and must reach the sheet unchanged, so it is never canonicalised
            # or split here. "" when the layout is not grouped by operation.
            # The root's own name is checked too, so pointing --sets-root
            # straight at an operation folder (metadata/sharpen_denoise) still
            # keeps that name leading the cell.
            "op": operation_folder_label(set_name, root),
            "op_from": _operation_segments(set_name, root),
            # The set folder's own name, used verbatim as set_id.
            "set_id": os.path.basename(os.path.normpath(set_dir)),
        })
    return records


def write_report(path, records):
    """The same columns as the shoot workflow's report.csv, plus a note column.

    Writing this before any caption is paid for means the pairing can be
    reviewed (filter ``confidence``) with no API cost at all.
    """
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["group_id", "set_folder", "edit_file", "matched_original",
                    "method", "confidence", "distance", "burst_margin",
                    "folder_operation", "note"])
        for idx, r in enumerate(records, 1):
            gid = f"{idx:04d}"
            op = r.get("op") or ""
            if not r["edit_is"]:
                w.writerow([gid, r["set"],
                            os.path.basename(r["files"][r["orig_i"]]), "",
                            "single", r["confidence"], "", "", op, r["why"]])
                continue
            orig_name = os.path.basename(r["files"][r["orig_i"]])
            for ei in r["edit_is"]:
                w.writerow([gid, r["set"], os.path.basename(r["files"][ei]),
                            orig_name, r["method"], r["confidence"], "", "",
                            op, r["why"]])
    return path


# --------------------------------------------------------------------------- #
# Building the work list and the sheet rows
# --------------------------------------------------------------------------- #
def build_tasks(records):
    """One task per row that needs a vision call (i.e. per edit / lone photo)."""
    tasks = []
    for r in records:
        orig_path = r["files"][r["orig_i"]]
        if not r["edit_is"]:
            # A lone photo still gets its one final row, classified on its own.
            tasks.append({
                "key": task_key(r["set"], orig_path),
                "rec": r,
                "edit_i": r["orig_i"],
                "edit_path": orig_path,
                "orig_path": "",
            })
            continue
        for ei in r["edit_is"]:
            tasks.append({
                "key": task_key(r["set"], r["files"][ei], orig_path),
                "rec": r,
                "edit_i": ei,
                "edit_path": r["files"][ei],
                "orig_path": orig_path,
            })
    return tasks


def row_dict(file_id, set_id, edit_state, role, title, description, keywords,
             info, category, date, edit_operation="", edit_operation_note=""):
    """One Excel row, keyed by column name (blank cells are empty strings)."""
    return {
        "file_id": file_id,
        "set_id": set_id,
        "edit_state": edit_state,
        "role": role,
        "edit_operation": edit_operation,
        # Where each removed object used to be. The description is one short
        # sentence and cannot carry several locations, so the detail lives here
        # (the template has this column for exactly that and left it blank).
        "edit_operation_note": edit_operation_note,
        "title": title,
        "description": description,
        "keywords": keywords,
        "file_format": info.get("file_format", ""),
        "resolution": info.get("resolution", ""),
        "color_space": info.get("color_space", ""),
        "file_size_mb": info.get("file_size_mb", ""),
        "date_created": date,
        "date_contributed": "",
        "asset_category": "",
        "ai_used": "",
        "category": category,
    }


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def build_parser():
    p = argparse.ArgumentParser(
        description="IndiaPix metadata sheet for input organised as one folder "
                    "per set (each holding a camera original + its edit).")
    p.add_argument("--sets-root", required=True,
                   help="folder whose sub-folders are the sets")
    p.add_argument("--out", default="sorted_output",
                   help="output folder for photos.xlsx / report.csv "
                        "(default: sorted_output)")
    p.add_argument("--provider", default="auto",
                   choices=["auto", "anthropic", "openai"],
                   help="vision LLM provider (default: auto-detect from keys)")
    p.add_argument("--model", default=None,
                   help="model name (default: claude-haiku-4-5 for anthropic, "
                        "gpt-4o-mini for openai)")
    p.add_argument("--api-key", default=None, help="API key (else env/.env)")
    p.add_argument("--workers", type=int, default=8,
                   help="parallel API requests (default 8)")
    p.add_argument("--limit", type=int, default=0,
                   help="only the first N sets (a cheap pilot run)")
    p.add_argument("--dry-run", action="store_true",
                   help="discover + pair only, then stop: writes report.csv so "
                        "the pairing can be checked before paying for captions")
    p.add_argument("--no-cache", action="store_true",
                   help="ignore cached labels and re-classify everything")
    return p


def stage_discover_and_pair(args, root, out_dir):
    """Steps 1-3: find the sets, read their metadata, decide the roles.

    Returns ``(records, tasks, n_uncertain)``.
    """
    # ---- 1. discover ---------------------------------------------------- #
    step(1, "finding the set folders")
    sets, ambiguous = discover_sets(root)
    if not sets:
        sys.exit(f"error: no set folders containing photos were found under "
                 f"'{root}'\n       expected: {root}/<set name>/<two photos>")
    if args.limit:
        sets = sets[: args.limit]
    n_sets = len(sets)
    n_photos = sum(len(f) for _, f in sets)
    log(f"  {n_sets:,} set folder(s) holding {n_photos:,} photo(s) "
        f"({n_photos / n_sets:.2f} per set)")

    if ambiguous:
        note(f"{len(ambiguous)} folder(s) held loose photos AND photo "
             f"sub-folders; the sub-folders were used as the sets:")
        for path, count in ambiguous[:5]:
            note(f"  skipped {count} photo(s) directly in {relpath(path, root)}")

    odd = [(p, len(f)) for p, f in sets if len(f) != EXPECTED_PHOTOS]
    if odd:
        note(f"{len(odd)} set(s) do not hold exactly {EXPECTED_PHOTOS} photos "
             f"- they are still processed:")
        for path, count in odd[:5]:
            note(f"  {relpath(path, root)}: {count} photo(s)")
        if len(odd) > 5:
            note(f"  ... and {len(odd) - 5} more")

    # ---- 2. metadata ---------------------------------------------------- #
    step(2, f"reading metadata from {n_photos:,} photo(s)")
    note("what: format, pixel size, colour space, EXIF Software, capture date, "
         "edit date and byte size.")
    note("why : these decide which photo of each pair is the camera original, "
         "and they fill the technical columns of the sheet.")
    note("how : file headers and EXIF only - no image is decoded here, so this "
         "takes minutes even at 80,000 photos.")
    workers = max(4, min(32, (os.cpu_count() or 4) * 4))
    inspected = inspect_all([p for _, files in sets for p in files], workers)

    # ---- 3. pair -------------------------------------------------------- #
    step(3, "pairing each set: which photo is the original, which is the edit")
    note("what: the edit is normally written by Photoshop/Lightroom, so its "
         "EXIF Software tag, edit date, pixel size and byte size all point to it.")
    note("why : the sheet needs one `00 original` row plus a `final` row per "
         "edit, and there are no Original/Edit folders to say which is which.")
    records = pair_sets(sets, inspected, root)
    n_uncertain = sum(1 for r in records
                      if r["confidence"] <= 0.35 and len(r["infos"]) > 1)
    tasks = build_tasks(records)
    n_rows = len(records) + len(tasks)
    n_lone = sum(1 for r in records if not r["edit_is"])

    by_op = {}
    for r in records:
        if r.get("op"):
            by_op[r["op"]] = by_op.get(r["op"], 0) + 1
    if by_op:
        log(f"  {sum(by_op.values()):,} set(s) sit in an edit-operation folder - "
            f"that folder name is written, verbatim, as the first "
            f"edit_operation:")
        for name in sorted(by_op):
            log(f"    {name:42s} {by_op[name]:,} set(s)", indent=4)

    log(f"  {len(records):,} set(s) -> {n_rows:,} sheet row(s) "
        f"({len(tasks):,} caption request(s))")
    log(f"  {len(records) - n_uncertain:,} pair(s) decided by a strong signal, "
        f"{n_uncertain:,} with none (confidence 0.35)")
    if n_lone:
        log(f"  {n_lone:,} set(s) hold a single photo - each gets one final row")

    report_path = write_report(os.path.join(out_dir, "report.csv"), records)
    low = sum(1 for r in records if r["confidence"] < 0.6)
    log(f"  report -> {report_path}  ({n_rows:,} row(s))")
    if low:
        note(f"{low:,} pair(s) scored below 0.6 - filter report.csv on the "
             f"confidence column to review them before the caption run.")
    return records, tasks, n_uncertain


# --------------------------------------------------------------------------- #
# Step 4 — classify (one vision call per edit)
# --------------------------------------------------------------------------- #
def load_cache(cache_path, no_cache):
    """Read the label cache, keeping only good labels (errors are retried)."""
    if no_cache or not os.path.isfile(cache_path):
        return {}
    try:
        with open(cache_path, encoding="utf-8") as fh:
            loaded = json.load(fh)
    except Exception:
        return {}
    if not isinstance(loaded, dict):
        return {}
    return {k: cat._sanitise_label(v) for k, v in loaded.items()
            if cat._is_good_label(v)}


def stage_classify(args, tasks, out_dir):
    """Step 4: label every edit, caching aggressively so a re-run is free.

    Returns ``(results, cache_path)``.
    """
    provider = cat.detect_provider(args.provider)
    api_key = cat.resolve_key(args.api_key, provider)
    if not api_key:
        keyname = "OPENAI_API_KEY" if provider == "openai" else "ANTHROPIC_API_KEY"
        sys.exit(f"error: no {provider} API key (use --api-key, ${keyname}, or "
                 f"the .env file next to this script)")
    model = args.model or (cat.DEFAULT_OPENAI_MODEL if provider == "openai"
                           else cat.DEFAULT_ANTHROPIC_MODEL)

    cache_path = os.path.join(out_dir, CACHE_NAME)
    cache = load_cache(cache_path, args.no_cache)
    todo = [t for t in tasks if t["key"] not in cache]

    step(4, f"classifying {len(todo):,} edit(s) with the vision model")
    note(f"provider : {provider} / {model}")
    note(f"cache    : {cache_path}")
    note(f"result   : {len(tasks) - len(todo):,} already labelled, "
         f"{len(todo):,} to classify")
    if not todo:
        note("nothing to do - every set is already labelled; writing the sheet")
        return {t["key"]: dict(cache[t["key"]]) for t in tasks}, cache_path

    note(f"workers  : {args.workers} set(s) in flight at once")
    note("each set sends BOTH photos (original + edit) in one request, so the "
         "model can name the real difference between them.")
    note(f"progress is saved every {CHECKPOINT_EVERY} labels, so an interrupted "
         f"run resumes instead of paying twice.")

    results = {}
    lock = threading.Lock()
    state = {"since": 0}
    done_count = [0]
    n_asked_again = [0]
    t0 = time.time()

    def save_cache(force=False):
        with lock:
            if not force and state["since"] < CHECKPOINT_EVERY:
                return
            state["since"] = 0
            tmp = cache_path + ".tmp"
            try:
                with open(tmp, "w", encoding="utf-8") as fh:
                    json.dump(cache, fh, indent=0)
                os.replace(tmp, cache_path)   # atomic: a crash cannot truncate it
            except Exception as e:
                log(f"  [warn] could not write cache: {e}")

    def work(task):
        try:
            edit_b64 = cat.image_to_data_url(task["edit_path"])
            orig_b64 = (cat.image_to_data_url(task["orig_path"])
                        if task["orig_path"] else None)
            # classify_pair re-asks once when the model used object_removal but
            # never said WHAT was removed, so the cell is actionable.
            label, asked_again = cat.classify_pair(orig_b64, edit_b64, api_key,
                                                   model, provider)
            if asked_again:
                with lock:
                    n_asked_again[0] += 1
        except Exception as e:
            label = cat._error_label(e)
        with lock:
            done_count[0] += 1
            if cat._is_good_label(label):
                cache[task["key"]] = label
                state["since"] += 1
        save_cache()
        return task["key"], label

    # 40,000 lines of per-image output would bury the progress; print every 20th
    # (and always the last), with a percentage and an ETA measured in hours.
    every = 20 if len(todo) > 200 else 1
    with cf.ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futures = [ex.submit(work, t) for t in todo]
        for fut in cf.as_completed(futures):
            key, label = fut.result()
            results[key] = label
            n = done_count[0]
            elapsed = time.time() - t0
            rate = n / elapsed if elapsed > 0 else 0.0
            eta = (len(todo) - n) / rate if rate > 0 else 0.0
            if n % every == 0 or n == len(todo):
                log(f"  [{n:,}/{len(todo):,}] {100.0 * n / len(todo):.0f}%  "
                    f"{label.get('category', 'ERROR')}  "
                    f"[{fmt_eta(elapsed)} elapsed, ~{fmt_eta(eta)} left]")
            elif n % 200 == 0:
                log(f"  ... {n:,}/{len(todo):,} done, ~{fmt_eta(eta)} left")

    save_cache(force=True)
    log(f"  classified {len(todo):,} edit(s) in {fmt_eta(time.time() - t0)}; "
        f"{len(cache):,} label(s) cached")

    # Fill in labels that came from the cache rather than from this run.
    for t in tasks:
        results.setdefault(t["key"], dict(cache.get(t["key"])
                                          or cat._error_label("not classified")))
    return results, cache_path


# --------------------------------------------------------------------------- #
# Step 5 — stream the sheet
# --------------------------------------------------------------------------- #
def stage_write(records, tasks, results, out_dir):
    """Write photos.xlsx (streamed) + categorize_debug.log.

    Rows are grouped by category so similar photos sit together, and a set's
    rows stay adjacent — exactly like the shoot workflow. Only the *order* is
    computed as a list of indices, so memory stays flat at 80,000 rows.

    Every final row's `edit_operation` is logged with its evidence
    (`op_sources`: which term came from the folder, the model, the metadata or
    the prose), and the run ends with a hard invariant: when a set sits in an
    operation folder, that folder's name must be the FIRST term of the cell.
    A violation is counted, printed and written to the debug log rather than
    passing silently.
    """
    step(5, "writing the Excel workbook and the debug log")
    note("rows    : one `00 original` row, then `01`, `02`, ... per edit")
    note("grouping: by category, then by set name - a set's rows stay adjacent")
    note("columns : 17 template + `category` (subject type) = 18")
    note("memory  : the sheet is streamed, so 80,000 rows costs the same memory "
         "as 200")
    note("evidence: every final cell records where each operation came from "
         "(folder / model / measured / prose / fallback)")
    note("invariant: the operation folder's name ALWAYS leads the cell when the "
         "set sits in one - violations are reported, never silent")
    violations = []
    reviews = []
    op_sources_seen = {}
    n_filled_by_prose = 0
    n_empty_from_model = 0

    # The first task of a set supplies the shared title / keywords / category.
    first_task = {}
    for t in tasks:
        first_task.setdefault(id(t["rec"]), t)

    order = {c: i for i, c in enumerate(cat.CATEGORIES)}

    def cat_of(rec):
        t = first_task.get(id(rec))
        return results.get(t["key"], {}).get("category", "ERROR") if t else "ERROR"

    index = sorted(range(len(records)),
                   key=lambda i: (order.get(cat_of(records[i]), 99),
                                  records[i]["set"]))

    today = time.strftime("%Y-%m-%d")
    xlsx_path = os.path.join(out_dir, "photos.xlsx")
    n_rows = sum(1 + len(r["edit_is"]) for r in records)
    sheet = StreamedSheet(xlsx_path, ALL_COLUMNS, expected_rows=n_rows)

    debug_path = os.path.join(out_dir, "categorize_debug.log")
    counts = {}
    dbg = open(debug_path, "w", newline="", encoding="utf-8")
    dw = csv.writer(dbg)
    dw.writerow(["set_folder", "file_id", "role", "matched_original", "category",
                 "method", "confidence", "note", "edit_operation", "title",
                 "original_description", "final_description", "keywords",
                 "folder_operation", "op_sources", "model_edit_operation",
                 "description_used_for_prose", "op_review"])
    try:
        for out_i, ri in enumerate(index, 1):
            r = records[ri]
            # The set folder's own name IS the set id - the client's naming is
            # never renumbered or reformatted (SET_037501 stays SET_037501).
            set_id = r.get("set_id") or cat.format_set_id(out_i)
            infos = r["infos"]
            orig_path = r["files"][r["orig_i"]]
            orig_info = infos[r["orig_i"]]
            head = results.get(first_task[id(r)]["key"],
                               cat._error_label("not classified"))
            category = head.get("category", "ERROR")
            title = head.get("title", "")
            keywords = head.get("keywords", "")

            # The operation folder that holds this set (metadata/color_tonal/
            # SET_000001) is the client's own label, so it leads edit_operation
            # EXACTLY as written - including any ';' or ',' inside the name.
            folder_op = r.get("op") or ""

            if not r["edit_is"]:
                # A lone photo: one final row, no original row.
                counts[category] = counts.get(category, 0) + 1
                lone_model_ops = head.get("edit_operation", "")
                lone_prose = " ".join((head.get("final_description", ""), title))
                ops, sources, review = cat.explain_ops(
                    folder_op, lone_model_ops,
                    hint_text=lone_prose, fallback="none")
                for op, why in sources.items():
                    op_sources_seen[why] = op_sources_seen.get(why, 0) + 1
                if not (lone_model_ops or "").strip():
                    n_empty_from_model += 1
                if "prose" in sources.values():
                    n_filled_by_prose += 1
                if review:
                    reviews.append((set_id, os.path.basename(orig_path), review))
                if folder_op and not ops.startswith(folder_op):
                    violations.append((set_id, os.path.basename(orig_path),
                                       folder_op, ops))
                row = row_dict(os.path.basename(orig_path), set_id, "01", "final",
                               title, head.get("final_description", ""), keywords,
                               orig_info, category,
                               orig_info.get("date_modified"), ops,
                               head.get("edit_operation_note", ""))
                row["date_contributed"] = today
                sheet.add(row)
                dw.writerow([r["set"], os.path.basename(orig_path), "final", "",
                             category, r["method"], r["confidence"], r["why"],
                             ops, title, head.get("original_description", ""),
                             head.get("final_description", ""), keywords,
                             folder_op,
                             " ".join(f"{op}<-{why}"
                                      for op, why in sources.items()),
                             lone_model_ops,
                             lone_prose if "prose" in sources.values() else "",
                             review])
                continue

            # Row 00 — the camera original, describing the untouched scene.
            counts[category] = counts.get(category, 0) + 1
            row = row_dict(os.path.basename(orig_path), set_id, "00", "original",
                           title, head.get("original_description", ""), keywords,
                           orig_info, category, orig_info.get("date_created"))
            row["date_contributed"] = today
            sheet.add(row)
            dw.writerow([r["set"], os.path.basename(orig_path), "original", "",
                         category, r["method"], r["confidence"], r["why"], "",
                         title, head.get("original_description", ""), "",
                         keywords])

            # Rows 01..n — one per delivered edit, carrying the edit_operation.
            for n, ei in enumerate(r["edit_is"], 1):
                key = task_key(r["set"], r["files"][ei], orig_path)
                label = results.get(key, cat._error_label("not classified"))
                edit_info = infos[ei]
                # The cell reads: the client's operation folder(s) first (it is
                # their own label for the batch, and the order inside the folder
                # name says which is primary), then whatever else the model saw
                # in ITS order (the prompt asks most-important-first, i.e. by
                # confidence), then a crop the file metadata proves outright.
                measured = cat.geometry_ops(orig_info, edit_info)
                # then the label's own prose: a model that left edit_operation
                # empty still wrote "Same scene with the crowd removed and a
                # warmer colour grade", and those words ARE the answer. A pair
                # that truly shows no change says "none" instead of blank.
                model_ops = label.get("edit_operation", "")
                prose = " ".join((label.get("final_description", ""),
                                  label.get("title", ""), title))
                # Every term is traced to its source so the debug log can
                # explain the cell instead of just printing it.
                ops, sources, review = cat.explain_ops(
                    folder_op, model_ops, measured, hint_text=prose,
                    fallback="none")
                for op, why in sources.items():
                    op_sources_seen[why] = op_sources_seen.get(why, 0) + 1
                if not (model_ops or "").strip():
                    n_empty_from_model += 1
                if "prose" in sources.values():
                    n_filled_by_prose += 1
                if review:
                    reviews.append((set_id,
                                    os.path.basename(r["files"][ei]), review))
                # INVARIANT: a set inside an operation folder must lead with
                # that folder's own text. Anything else means the client's
                # label was lost, and is reported, never ignored.
                if folder_op and not ops.startswith(folder_op):
                    violations.append((set_id,
                                       os.path.basename(r["files"][ei]),
                                       folder_op, ops))
                edit_cat = label.get("category", category)
                counts[edit_cat] = counts.get(edit_cat, 0) + 1
                row = row_dict(os.path.basename(r["files"][ei]), set_id,
                               f"{n:02d}", "final",
                               label.get("title", "") or title,
                               label.get("final_description", ""),
                               label.get("keywords", "") or keywords,
                               edit_info, edit_cat,
                               edit_info.get("date_modified"), ops,
                               label.get("edit_operation_note", ""))
                row["date_contributed"] = today
                sheet.add(row)
                dw.writerow([r["set"], os.path.basename(r["files"][ei]), "final",
                             os.path.basename(orig_path), edit_cat, r["method"],
                             r["confidence"], r["why"], ops,
                             label.get("title", ""),
                             label.get("original_description", ""),
                             label.get("final_description", ""),
                             label.get("keywords", ""),
                             folder_op,
                             " ".join(f"{op}<-{why}"
                                      for op, why in sources.items()),
                             model_ops,
                             prose if "prose" in sources.values() else "",
                             review])
    finally:
        dbg.close()

    sheet.close()
    log(f"  excel -> {xlsx_path}  ({sheet.rows_written:,} row(s), "
        f"{len(ALL_COLUMNS)} columns)")
    log(f"  debug -> {debug_path}")

    # ---- how each final cell was built --------------------------------- #
    total_finals = sum(len(r["edit_is"]) for r in records)
    log("  edit_operation evidence (per final row):")
    for why in ("folder", "model", "measured", "prose", "fallback"):
        n = op_sources_seen.get(why, 0)
        if n:
            log(f"{n:>7,} term(s) came from {why}", indent=4)
    log(f"{total_finals:,} final row(s); the model's own answer was empty on "
        f"{n_empty_from_model:,}", indent=4)
    log(f"{n_filled_by_prose:,} cell(s) filled from the label's own "
        f"description (model left edit_operation blank)", indent=4)

    # ---- the hard invariant -------------------------------------------- #
    if violations:
        log(f"[FAIL] {len(violations):,} final row(s) do NOT lead with their "
            f"operation folder name - the client's label was lost:")
        for sid, fname, folder, cell in violations[:10]:
            log(f"{sid}/{fname}: folder {folder!r} -> cell {cell!r}", indent=4)
        if len(violations) > 10:
            log(f"... and {len(violations) - 10} more "
                f"(see categorize_debug.log)", indent=4)
    else:
        log("[ok] every set that sits in an operation folder leads its cell "
            "with that folder's name, verbatim")

    # ---- rows where the model and its own description disagree ---------- #
    # The only cells where "performed" is genuinely in doubt: the model named
    # one operation, its own sentence names another. Guessing was what the
    # client rejected, so these are listed instead - each is a two-second check
    # against the pair.
    if reviews:
        log(f"[REVIEW] {len(reviews):,} row(s) where the description names an "
            f"operation the model did not report - the cell keeps the model's "
            f"answer and is flagged for a human:")
        for sid, fname, why in reviews[:10]:
            log(f"{sid}/{fname}: {why}", indent=4)
        if len(reviews) > 10:
            log(f"... and {len(reviews) - 10} more "
                f"(filter op_review in categorize_debug.log)", indent=4)
    else:
        log("[ok] the model's answer and the description agree on every row")

    # ---- object_removal rows that still do not say WHAT went ------------ #
    # "object_removal" is useless on its own: the client needs to know what
    # disappeared. The prompt demands it and a miss is asked once more, so this
    # should be empty. When it is not, the rows are listed by name.
    n_removal = 0
    missing = []
    for r in records:
        for ei in r["edit_is"] or [r["orig_i"]]:
            key = task_key(r["set"], r["files"][ei],
                           "" if not r["edit_is"] else r["files"][r["orig_i"]])
            label = results.get(key) or {}
            ops = label.get("edit_operation", "")
            if "object_removal" not in cat.ops_ordered(ops):
                continue
            n_removal += 1
            if cat.removal_detail_missing(
                    ops, label.get("final_description", "")):
                missing.append((r.get("set_id", ""),
                                os.path.basename(r["files"][ei]),
                                label.get("final_description", "")))
    if n_removal:
        log(f"  object_removal rows: {n_removal:,}")
        if missing:
            log(f"[REVIEW] {len(missing):,} object_removal row(s) still do NOT "
                f"say what was removed:")
            for sid, fname, desc in missing[:10]:
                log(f"{sid}/{fname}: {desc}", indent=4)
            if len(missing) > 10:
                log(f"... and {len(missing) - 10} more "
                    f"(see categorize_debug.log)", indent=4)
        else:
            log("[ok] every object_removal row names what was removed")
    return counts, xlsx_path, debug_path


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main():
    args = build_parser().parse_args()

    root = os.path.abspath(args.sets_root)
    if not os.path.isdir(root):
        sys.exit(f"error: '{args.sets_root}' is not a directory")
    out_dir = os.path.abspath(args.out)
    os.makedirs(out_dir, exist_ok=True)

    print(flush=True)
    print(hr("="), flush=True)
    print(safe("  INDIAPIX - SET-FOLDER METADATA RUN"), flush=True)
    print(hr("="), flush=True)
    note(f"input    : {root}")
    note(f"output   : {out_dir}")
    note(f"layout   : one folder per set (camera original + edit side by side)")
    note(f"provider : {args.provider}"
         + ("   [DRY RUN - no API calls]" if args.dry_run else ""))
    print(flush=True)

    records, tasks, n_uncertain = stage_discover_and_pair(args, root, out_dir)

    if args.dry_run:
        print(flush=True)
        print("  " + hr("-", indent=2), flush=True)
        print(safe("  DRY RUN - stopped before any API call"), flush=True)
        note(f"sets to caption : {len(records):,}")
        note(f"API requests    : {len(tasks):,}")
        note(f"pairing report  : {os.path.join(out_dir, 'report.csv')}")
        note("review it (filter the confidence column), then re-run without "
             "--dry-run to write photos.xlsx")
        print("  " + hr("-", indent=2), flush=True)
        return

    results, cache_path = stage_classify(args, tasks, out_dir)
    counts, xlsx_path, _ = stage_write(records, tasks, results, out_dir)

    print(flush=True)
    print("  " + hr("-", indent=2), flush=True)
    print(safe("  RESULT - categories assigned"), flush=True)
    for c in cat.CATEGORIES:
        print(safe(f"    {c:22s}: {counts.get(c, 0):,}"), flush=True)
    if counts.get("ERROR"):
        print(safe(f"    {'ERROR':22s}: {counts['ERROR']:,}"), flush=True)
    print("  " + hr("-", indent=2), flush=True)
    note(f"sets           : {len(records):,}")
    note(f"low-confidence : {n_uncertain:,} pair(s) decided with no strong "
         f"signal (filter report.csv)")
    note(f"label cache    : {cache_path}")
    note(f"total time     : {fmt_eta(time.time() - _T0)}")
    note(f"open this      : {xlsx_path}")
    print("  " + hr("-", indent=2), flush=True)


if __name__ == "__main__":
    main()
