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

Providers (vision LLM), chosen with --provider {auto,anthropic,openai}:
  - anthropic (default): claude-haiku-4-5  -> ANTHROPIC_API_KEY
  - openai            : gpt-4o-mini       -> OPENAI_API_KEY
  - auto (default)    : pick whichever key is present (.env / environment)

Usage:
  python3 categorize.py --edit-dir "_sort images/Edit" \
                        --report sorted_output/report.csv \
                        --out sorted_output/photos.xlsx
  python3 categorize.py ... --provider openai

The API key is read from (in order): --api-key, the provider's env var, or a
.env file sitting next to this script. Nothing is hard-coded.
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
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from imageutil import IMG_EXTS, load_image


# Windows-safe console: force UTF-8 with a lossy fallback so non-ASCII filenames
# or captions never crash printing on a cp1252 console.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# Wall-clock start, used by log() so long runs never look "stuck".
_T0 = time.time()


def log(msg, indent=2):
    """Print a timestamped, flushed progress line."""
    print(f"{' ' * indent}[{time.time() - _T0:6.1f}s] {msg}", flush=True)


CATEGORIES = ["Portrait", "Traditional/Festival", "White Background", "Others"]
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"
OPENAI_URL = "https://api.openai.com/v1/chat/completions"
DEFAULT_ANTHROPIC_MODEL = "claude-haiku-4-5-20251001"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"

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
        for line in open(path, encoding="utf-8", errors="replace"):
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env


def provider_env_names(provider):
    return ["OPENAI_API_KEY"] if provider == "openai" else ["ANTHROPIC_API_KEY"]


def detect_provider(requested):
    """Return 'anthropic' or 'openai'. If 'auto', pick whichever key is present."""
    if requested in ("anthropic", "openai"):
        return requested
    here = os.path.dirname(os.path.abspath(__file__))
    env = load_dotenv(os.path.join(here, ".env"))
    if os.environ.get("ANTHROPIC_API_KEY") or env.get("ANTHROPIC_API_KEY"):
        return "anthropic"
    if os.environ.get("OPENAI_API_KEY") or env.get("OPENAI_API_KEY"):
        return "openai"
    return "anthropic"


def resolve_key(arg_key, provider):
    if arg_key:
        return arg_key
    here = os.path.dirname(os.path.abspath(__file__))
    env = load_dotenv(os.path.join(here, ".env"))
    for name in provider_env_names(provider):
        if os.environ.get(name):
            return os.environ[name]
        if env.get(name):
            return env[name]
    return None


# --------------------------------------------------------------------------- #
# Image -> data
# --------------------------------------------------------------------------- #
def image_to_data_url(path, max_side=SEND_MAX):
    """Load (raster or RAW), orient, downscale, base64-encode as JPEG.

    Uses imageutil.load_image so Canon RAW (.cr3/.cr2) and other RAW files are
    handled the same as JPEG/PNG.
    """
    im = load_image(path)
    if im is None:
        raise ValueError(f"could not load image: {path}")
    im = im.convert("RGB")
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
def _normalise(data):
    cat = str(data.get("category", "")).strip()
    cap = str(data.get("caption", "")).strip()
    for c in CATEGORIES:
        if c.lower() == cat.lower():
            cat = c
            break
    else:
        cat = "Others"
    return cat, cap


def _classify_anthropic(b64, api_key, model, retries=3):
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
            return _normalise(parse_json_loose(content))
        except Exception as e:
            if attempt == retries - 1:
                return "ERROR", f"({e})"
            time.sleep(delay)
            delay *= 2
    return "ERROR", ""


def _classify_openai(b64, api_key, model, retries=3):
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    data_url = f"data:image/jpeg;base64,{b64}"
    payload = {
        "model": model,
        "max_tokens": 120,
        "temperature": 0,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": PROMPT},
                    {"type": "image_url", "image_url": {"url": data_url, "detail": "low"}},
                ],
            }
        ],
        "response_format": {"type": "json_object"},
    }
    delay = 1.0
    for attempt in range(retries):
        try:
            r = requests.post(OPENAI_URL, headers=headers, json=payload, timeout=90)
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            return _normalise(parse_json_loose(content))
        except Exception as e:
            if attempt == retries - 1:
                return "ERROR", f"({e})"
            time.sleep(delay)
            delay *= 2
    return "ERROR", ""


def classify_image(b64, api_key, model, provider="anthropic", retries=3):
    if provider == "openai":
        return _classify_openai(b64, api_key, model, retries)
    return _classify_anthropic(b64, api_key, model, retries)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def read_report(report_path):
    with open(report_path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--edit-dir", required=True, help="folder containing the edits")
    ap.add_argument("--report", default="sorted_output/report.csv",
                    help="report.csv from sort_photos.py (edit -> original map)")
    ap.add_argument("--out", default="sorted_output/photos.xlsx", help="output Excel file")
    ap.add_argument("--provider", default="auto", choices=["auto", "anthropic", "openai"],
                    help="vision LLM provider (default: auto-detect from available keys)")
    ap.add_argument("--model", default=None,
                    help="model name (default: claude-haiku-4-5 for anthropic, gpt-4o-mini for openai)")
    ap.add_argument("--api-key", default=None, help="API key (else env/.env)")
    ap.add_argument("--workers", type=int, default=6, help="parallel requests")
    ap.add_argument("--limit", type=int, default=0, help="only process first N (testing)")
    ap.add_argument("--no-cache", action="store_true", help="ignore cached labels")
    args = ap.parse_args()

    provider = detect_provider(args.provider)
    api_key = resolve_key(args.api_key, provider)
    if not api_key:
        keyname = "OPENAI_API_KEY" if provider == "openai" else "ANTHROPIC_API_KEY"
        sys.exit(f"error: no {provider} API key (use --api-key, ${keyname}, or .env)")
    model = args.model or (DEFAULT_OPENAI_MODEL if provider == "openai" else DEFAULT_ANTHROPIC_MODEL)

    # make sure the output folder exists (categorize.py can be run standalone,
    # not only after sort_photos.py has created it)
    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    rows = read_report(args.report)
    if args.limit:
        rows = rows[: args.limit]

    by_edit = {r["edit_file"]: r for r in rows}
    edits = [f for f in sorted(os.listdir(args.edit_dir))
             if os.path.splitext(f)[1].lower() in IMG_EXTS and f in by_edit]

    print(f"Classifying {len(edits)} edit(s) with {provider} ({model}) ...", flush=True)

    cache_path = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".", "labels_cache.json")
    cache = {}
    if not args.no_cache and os.path.isfile(cache_path):
        try:
            cache = json.load(open(cache_path, encoding="utf-8"))
        except Exception:
            cache = {}

    n_cached = sum(1 for f in edits if f in cache and cache[f][0] not in ("ERROR",))
    n_todo = len(edits) - n_cached
    if n_cached:
        log(f"{n_cached} already cached, {n_todo} to classify "
            f"({args.workers} worker(s))")
    if n_todo == 0:
        log("nothing to do — all labels cached; writing outputs ...")

    lock = threading.Lock()
    results = {}
    done = [0]

    def work(fname):
        if fname in cache and cache[fname][0] not in ("ERROR",):
            with lock:
                done[0] += 1
            return fname, tuple(cache[fname])
        try:
            b64 = image_to_data_url(os.path.join(args.edit_dir, fname))
            cat, cap = classify_image(b64, api_key, model, provider)
        except Exception as e:
            cat, cap = "ERROR", f"({e})"
        with lock:
            done[0] += 1
            print(f"  [{done[0]}/{len(edits)}] {fname}: {cat} "
                  f"({time.time() - _T0:.0f}s)", flush=True)
        return fname, (cat, cap)

    # Heartbeat: while API requests are in flight (they can take many seconds
    # each) print a periodic "still working" line so the run never looks stuck.
    stop = threading.Event()

    def heartbeat():
        while not stop.wait(15):
            with lock:
                log(f"  ... still working: {done[0]}/{len(edits)} done, "
                    f"waiting on API responses ...")

    hb = threading.Thread(target=heartbeat, daemon=True)
    hb.start()
    try:
        with cf.ThreadPoolExecutor(max_workers=args.workers) as ex:
            for fname, res in ex.map(work, edits):
                results[fname] = res
    finally:
        stop.set()

    with lock:
        for fname, res in results.items():
            cache[fname] = list(res)
        json.dump(cache, open(cache_path, "w", encoding="utf-8"), indent=0)
    log(f"labels cached: {len(cache)} total")

    order = {c: i for i, c in enumerate(CATEGORIES)}
    sorted_edits = sorted(results, key=lambda f: (order.get(results[f][0], 99), f))

    log(f"writing Excel ({len(sorted_edits)} row(s)) -> {args.out} ...")
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
    log(f"writing debug log -> {log_path} ...")
    with open(log_path, "w", newline="", encoding="utf-8") as lf:
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
    print("\n=== Category counts ===", flush=True)
    for c in CATEGORIES:
        print(f"  {c:22s}: {counts.get(c, 0)}", flush=True)
    if counts.get("ERROR"):
        print(f"  {'ERROR':22s}: {counts['ERROR']}", flush=True)
    print(f"\nExcel written: {args.out}", flush=True)


if __name__ == "__main__":
    main()