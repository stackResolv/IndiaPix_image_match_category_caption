#!/usr/bin/env python3
"""
categorize.py — classify each edited photo into one of four categories and write
a one-line caption, using a vision LLM, then export everything to an Excel sheet.

It joins to the originals using sorted_output/report.csv (produced by
sort_photos.py) so each row shows: original, edit, category, caption.

Categories (fixed):
  - Portrait             : a person/people as the main subject, non-white bg
  - Traditional/Festival : traditional/ethnic attire or cultural/festive setting
  - White Background     : plain white/studio seamless backdrop
  - Others               : everything else

Provider: Anthropic Claude (vision). Default model: claude-haiku-4-5 (cheap+fast).

Usage:
  python3 categorize.py --edit-dir "_sort images/Edit" \
                        --report sorted_output/report.csv \
                        --out sorted_output/photos.xlsx

The API key is read from (in order): --api-key, $ANTHROPIC_API_KEY, or a .env
file sitting next to this script. Nothing is hard-coded.
"""

import argparse
import base64
import concurrent.futures as cf
import csv
import io
import json
import os
import re
import sys
import threading
import time

import requests
from PIL import Image, ImageOps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


CATEGORIES = ["Portrait", "Traditional/Festival", "White Background", "Others"]
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"

# Longest side (px) of the image we send to the model. Small = cheap, still clear.
SEND_MAX = 640

PROMPT = (
    "You are tagging photos for a professional Indian photo-library business. "
    "Read the image and output two things.\n\n"
    "CATEGORY — choose exactly one. Apply this PRIORITY ORDER, first match wins:\n"
    "1. \"Portrait\" if a person/people are the main subject, shot as a normal "
    "portrait on a plain/neutral background (not a clearly traditional setting "
    "and not a plain white studio backdrop).\n"
    "2. \"Traditional/Festival\" if a person is wearing traditional or ethnic "
    "attire (e.g. saree, salwar, kurta, turban, sherwani, lehenga, heavy "
    "jewellery), OR the scene is a cultural/religious/wedding/festival setting.\n"
    "3. \"White Background\" if the subject is shot against a plain white / "
    "seamless studio backdrop (bright, featureless white behind the subject).\n"
    "4. \"Others\" for anything else (product, food, scenery, group without a "
    "clear main subject, etc.).\n\n"
    "CAPTION — one short sentence, at most 12 words, describing the photo "
    "(e.g. \"Man in a pink turban seated, warm beige backdrop\").\n\n"
    "Return ONLY compact JSON and nothing else: "
    '{"category": "...", "caption": "..."}'
)


# --------------------------------------------------------------------------- #
# Config / keys
# --------------------------------------------------------------------------- #
def load_dotenv(path):
    env = {}
    if os.path.isfile(path):
        for line in open(path):
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env


def resolve_key(arg_key):
    if arg_key:
        return arg_key
    if os.environ.get("ANTHROPIC_API_KEY"):
        return os.environ["ANTHROPIC_API_KEY"]
    here = os.path.dirname(os.path.abspath(__file__))
    return load_dotenv(os.path.join(here, ".env")).get("ANTHROPIC_API_KEY")


# --------------------------------------------------------------------------- #
# Image -> data
# --------------------------------------------------------------------------- #
def image_to_data_url(path, max_side=SEND_MAX):
    """Load, transpose by EXIF orientation, downscale, base64-encode as JPEG."""
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        w, h = im.size
        scale = min(1.0, max_side / max(w, h))
        if scale < 1.0:
            im = im.resize((int(w * scale), int(h * scale)))
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=80)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def parse_json_loose(text):
    """Extract a JSON object from model output that may be wrapped in ``` fences."""
    text = text.strip()
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            pass
    return {}


# --------------------------------------------------------------------------- #
# LLM call
# --------------------------------------------------------------------------- #
def classify_image(b64, api_key, model, retries=3):
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
        "content-type": "application/json",
    }
    payload = {
        "model": model,
        "max_tokens": 120,
        "temperature": 0,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": "image/jpeg", "data": b64},
                    },
                    {"type": "text", "text": PROMPT},
                ],
            }
        ],
    }
    delay = 1.0
    for attempt in range(retries):
        try:
            r = requests.post(ANTHROPIC_URL, headers=headers, json=payload, timeout=90)
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            content = r.json()["content"][0]["text"]
            data = parse_json_loose(content)
            cat = str(data.get("category", "")).strip()
            cap = str(data.get("caption", "")).strip()
            for c in CATEGORIES:
                if c.lower() == cat.lower():
                    cat = c
                    break
            else:
                cat = "Others"
            return cat, cap
        except Exception as e:
            if attempt == retries - 1:
                return "ERROR", f"({e})"
            time.sleep(delay)
            delay *= 2
    return "ERROR", ""


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def read_report(report_path):
    with open(report_path, newline="") as f:
        return list(csv.DictReader(f))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--edit-dir", required=True, help="folder containing the edits")
    ap.add_argument("--report", default="sorted_output/report.csv",
                    help="report.csv from sort_photos.py (edit -> original map)")
    ap.add_argument("--out", default="sorted_output/photos.xlsx", help="output Excel file")
    ap.add_argument("--model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--api-key", default=None, help="Anthropic key (else env/.env)")
    ap.add_argument("--workers", type=int, default=6, help="parallel requests")
    ap.add_argument("--limit", type=int, default=0, help="only process first N (testing)")
    ap.add_argument("--no-cache", action="store_true", help="ignore cached labels")
    args = ap.parse_args()

    api_key = resolve_key(args.api_key)
    if not api_key:
        sys.exit("error: no Anthropic API key (use --api-key, $ANTHROPIC_API_KEY, or .env)")

    rows = read_report(args.report)
    if args.limit:
        rows = rows[: args.limit]

    by_edit = {r["edit_file"]: r for r in rows}
    edits = [f for f in sorted(os.listdir(args.edit_dir))
             if f.lower().endswith((".jpg", ".jpeg", ".png")) and f in by_edit]

    print(f"Classifying {len(edits)} edit(s) with {args.model} ...")

    cache_path = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".", "labels_cache.json")
    cache = {}
    if not args.no_cache and os.path.isfile(cache_path):
        try:
            cache = json.load(open(cache_path))
        except Exception:
            cache = {}

    lock = threading.Lock()
    results = {}
    done = [0]

    def work(fname):
        if fname in cache and cache[fname][0] not in ("ERROR",):
            return fname, tuple(cache[fname])
        try:
            b64 = image_to_data_url(os.path.join(args.edit_dir, fname))
            cat, cap = classify_image(b64, api_key, args.model)
        except Exception as e:
            cat, cap = "ERROR", f"({e})"
        with lock:
            done[0] += 1
            print(f"  [{done[0]}/{len(edits)}] {fname}: {cat}")
        return fname, (cat, cap)

    with cf.ThreadPoolExecutor(max_workers=args.workers) as ex:
        for fname, res in ex.map(work, edits):
            results[fname] = res

    with lock:
        for fname, res in results.items():
            cache[fname] = list(res)
        json.dump(cache, open(cache_path, "w"), indent=0)

    order = {c: i for i, c in enumerate(CATEGORIES)}
    sorted_edits = sorted(results, key=lambda f: (order.get(results[f][0], 99), f))

    # ---- write Excel (clean: no debug columns) ----------------------------- #
    wb = Workbook()
    ws = wb.active
    ws.title = "Photos"
    headers = ["group_id", "category", "caption", "edit_file", "matched_original"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
        cell.alignment = Alignment(horizontal="center")

    cat_fills = {
        "Portrait": "D9E1F2",
        "Traditional/Festival": "FCE4D6",
        "White Background": "E2EFDA",
        "Others": "EDEDED",
        "ERROR": "FFC7CE",
    }

    for fname in sorted_edits:
        cat, cap = results[fname]
        r = by_edit[fname]
        ws.append([r.get("group_id", ""), cat, cap, fname,
                   r.get("matched_original", "")])
        fill = cat_fills.get(cat)
        if fill:
            ws.cell(row=ws.max_row, column=2).fill = PatternFill("solid", fgColor=fill)

    widths = [10, 22, 52, 18, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(args.out)

    # ---- write debug log (match diagnostics for troubleshooting) ---------- #
    log_path = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".",
                            "categorize_debug.log")
    with open(log_path, "w", newline="") as lf:
        lw = csv.writer(lf)
        lw.writerow(["group_id", "edit_file", "matched_original", "category",
                     "method", "confidence", "distance", "burst_margin", "caption"])
        for fname in sorted_edits:
            cat, cap = results[fname]
            r = by_edit[fname]
            lw.writerow([r.get("group_id", ""), fname, r.get("matched_original", ""),
                         cat, r.get("method", ""), r.get("confidence", ""),
                         r.get("distance", ""), r.get("burst_margin", ""), cap])
    print(f"Debug log written: {log_path}")

    from collections import Counter
    counts = Counter(results[f][0] for f in results)
    print("\n=== Category counts ===")
    for c in CATEGORIES:
        print(f"  {c:22s}: {counts.get(c, 0)}")
    if counts.get("ERROR"):
        print(f"  {'ERROR':22s}: {counts['ERROR']}")
    print(f"\nExcel written: {args.out}")


if __name__ == "__main__":
    main()