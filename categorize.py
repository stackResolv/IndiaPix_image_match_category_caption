#!/usr/bin/env python3
"""
categorize.py — build the IndiaPix metadata sheet.

For every matched pair (original + its edited derivative) this writes TWO rows in
the IndiaPix template layout, exactly like IndiaPix_metadata_template.csv:

  row 1  edit_state=00, role=original ... the camera file
  row 2  edit_state=01, role=final    ... the delivered edit, carrying the
                                          edit_operation column

A vision LLM sees BOTH images side by side so it can name the real differences
(crop, colour grade, object removal, retouching, ...). Title, keywords and
category are shared by the pair; the description differs — the original row
describes the untouched scene, the final row describes the edit.

Output columns (IndiaPix_metadata_template.csv order):
  file_id, set_id, edit_state, role, edit_operation, edit_operation_note,
  title, description, keywords, file_format, resolution, color_space,
  file_size_mb, date_created, date_contributed, asset_category, ai_used

Filled: file_id (the file NAME, with extension), set_id (SET_000001 style),
edit_state, role, edit_operation (final row only), title, description,
keywords, file_format, resolution, color_space, file_size_mb, date_created,
date_contributed (the run date).
Left blank: edit_operation_note, asset_category, ai_used.

Those 17 columns come first, in that order, exactly as the template defines
them. The classifier's own category is appended as an 18th trailing column
(`category`) for review, and is also written to categorize_debug.log.

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
                        --originals-dir "_sort images/JPEG" \
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

from imageutil import IMG_EXTS, image_metadata, load_image


# Windows-safe console: force UTF-8 with a lossy fallback so non-ASCII filenames
# or captions never crash printing on a cp1252 console.
from consoleutil import hr, safe, setup_console, term_width, wrap

# Console prep: UTF-8 where possible, and the real width so the narration fits
# a narrow Windows PowerShell window instead of wrapping into nonsense.
setup_console()


# Wall-clock start, used by log() so long runs never look "stuck".
_T0 = time.time()


def log(msg, indent=2):
    """Print a timestamped progress line, wrapped to the terminal width."""
    stamp = f"[{time.time() - _T0:6.1f}s] "
    text = safe(str(msg))
    lines = wrap(text, indent=indent + len(stamp), subsequent_indent=" " * 4)
    print(" " * indent + stamp + lines[0], flush=True)
    for extra in lines[1:]:
        print(" " * (indent + len(stamp)) + extra.strip(), flush=True)


def step(n, text, total=5):
    """Announce a pipeline step so the log reads as a story, not just numbers."""
    print(flush=True)
    print("  " + hr("-"), flush=True)
    for line in wrap(safe(f"Step {n}/{total}: {text}"),
                     indent=2, subsequent_indent=" " * 4):
        print("  " + line, flush=True)
    print("  " + hr("-"), flush=True)


def note(text):
    """A plain explanatory line, wrapped and indented under the current step."""
    for line in wrap(safe(text), indent=4, subsequent_indent=" " * 4):
        print(" " * 4 + line, flush=True)


CATEGORIES = ["Portrait", "Traditional/Festival", "White Background", "Others"]
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"
OPENAI_URL = "https://api.openai.com/v1/chat/completions"
DEFAULT_ANTHROPIC_MODEL = "claude-haiku-4-5-20251001"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"

# Longest side (px) of the image we send to the model. Small = cheap, still clear.
SEND_MAX = 640

# Token budget: five fields, two of them descriptions plus a keyword list.
MAX_TOKENS = 500

# Operations used in edit_operation — the client's eight, exactly. Everything the
# model says (and everything an older cache holds) is mapped onto this list by
# canonical_op, so a cell never carries a term the client cannot filter on.
EDIT_OPS = [
    "color_tonal",
    "crop_perspective",
    "white_balance",
    "object_removal",
    "sharpen_denoise",
    "background",
    "skin_portrait",
    "compositing",
]

# One line per operation, written for a vision model: what to LOOK FOR, not just
# the label. This ships inside the prompt, which is why the eight are separable
# from a single original/edit pair.
_OP_GUIDE = (
    "- color_tonal — the light or colour mood changed: brightness, exposure, "
    "contrast, highlights/shadows, saturation, vibrance, curves, levels, colour "
    "grading, split-toning, a LUT or a vignette. Framing and objects are the "
    "same; the image simply looks brighter/darker/richer/flatter.\n"
    "- crop_perspective — the frame itself changed: a tighter or looser crop, a "
    "different aspect ratio, a straightened tilted horizon, a rotation, or "
    "corrected converging/keystoned lines (a building that leaned now upright).\n"
    "- white_balance — a colour-temperature or tint correction only: a blue/cool "
    "cast removed, an orange/warm cast added, a yellow-green cast neutralised so "
    "whites and greys read neutral. Contrast and saturation look unchanged.\n"
    "- object_removal — something clearly visible in IMAGE 1 is GONE in IMAGE 2: "
    "a person, animal, bin, cable, sign, dust spot or stray object has been made "
    "to disappear and the area behind it rebuilt.\n"
    "- sharpen_denoise — detail handling only: extra sharpening, clarity, "
    "texture or micro-contrast, or the opposite — noise/grain reduction and "
    "smoothing that removes speckle while keeping edges crisp.\n"
    "- background — the background specifically is replaced, removed to white or "
    "transparency, blurred, extended or swapped for a studio backdrop, while the "
    "main subject keeps its shape and position.\n"
    "- skin_portrait — portrait retouching on a person: skin smoothing, "
    "blemish/spot healing, teeth whitening, eye brightening, under-eye or "
    "stray-hair cleanup, jawline or figure slimming.\n"
    "- compositing — two or more elements are combined into one image: an object "
    "or person ADDED, a sky swapped, an element pasted in, a montage or double "
    "exposure, or text/graphics/watermark overlaid.\n"
)

_CATEGORY_RULES = (
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
)

_TITLE_RULE = (
    "TITLE — a real, readable sentence of at most 7 or 8 words (never more). "
    "Sentence case: capitalise only the first word and proper nouns "
    "(e.g. \"Woman buying spices at a Jaipur market\"). It must read as a "
    "sentence, not a keyword list — so no comma-separated fragments and NO "
    "trailing period.\n\n"
)

_KEYWORDS_RULE = (
    "KEYWORDS — 8 to 15 lowercase keywords (capitalise place and proper names, "
    "e.g. \"India\"), comma-separated with NO space after the commas, ordered "
    "from most to least specific. Include the place, the subject, the activity "
    "and the mood where relevant (e.g. \"market,spices,woman,India,street\"). "
    "No duplicates, no hashtags, no trailing comma.\n\n"
)

# ---- prompt used when BOTH the original and the final edit are available ---- #
PROMPT_PAIR = (
    "You are tagging photos for a professional Indian photo-library business. "
    "You are shown TWO images of the SAME shot.\n"
    "  IMAGE 1 = the ORIGINAL camera file (untouched).\n"
    "  IMAGE 2 = the FINAL delivered edit of that same shot.\n\n"
    + _CATEGORY_RULES + _TITLE_RULE + _KEYWORDS_RULE +
    "original_description — one sentence, max 25 words, describing exactly what "
    "is visible in IMAGE 1, the original (e.g. \"Woman selecting spices at a "
    "market stall, crowd in background\"). No trailing period.\n\n"
    "final_description — one sentence, max 25 words, describing IMAGE 2. If it "
    "is the same scene edited, START WITH \"Same scene with \" and name only "
    "what changed (e.g. \"Same scene with background crowd removed and warm "
    "color grade\"). If the framing or subject changed substantially, describe "
    "the result directly instead. No trailing period.\n\n"
    "edit_operation — one or more operations, semicolon-separated, most "
    "important first, chosen ONLY from this vocabulary: " +
    "; ".join(EDIT_OPS) + ".\n"
    "  What each one means — match what you SEE to these definitions, not just "
    "the word:\n" + _OP_GUIDE +
    "  Report EVERY operation you can actually see; most real edits are a "
    "combination (e.g. \"color_tonal;skin_portrait\"). Use \"none\" only if the "
    "two images are visually identical apart from pixel dimensions.\n"
    "  Disambiguation, in order: if an object disappeared it is object_removal; "
    "if the whole backdrop changed it is background; if a new element appeared "
    "it is compositing; if only a person's skin/face was cleaned it is "
    "skin_portrait; if only the colour cast shifted it is white_balance; if only "
    "detail/noise changed it is sharpen_denoise.\n\n"
    "Return ONLY compact JSON and nothing else:\n"
    '{"category": "...", "title": "...", "keywords": "...", '
    '"original_description": "...", "final_description": "...", '
    '"edit_operation": "..."}'
)

# ---- prompt used when an edit has NO matched original ---------------------- #
PROMPT_SINGLE = (
    "You are tagging photos for a professional Indian photo-library business. "
    "Read the image and output four fields.\n\n"
    + _CATEGORY_RULES + _TITLE_RULE + _KEYWORDS_RULE +
    "final_description — one sentence, max 25 words, describing what is "
    "actually visible: subject, action, setting, lighting and colours. No "
    "speculation, no camera or editing talk. No trailing period.\n\n"
    "Return ONLY compact JSON and nothing else:\n"
    '{"category": "...", "title": "...", "keywords": "...", '
    '"final_description": "..."}'
)

# Template column order, taken from IndiaPix_metadata_template.csv. These 17
# columns come first, in this order, exactly as the template defines them. The
# classifier's own category is not part of the template, so it is appended as a
# trailing 18th column (EXTRA_COLUMNS) for review.
TEMPLATE_COLUMNS = [
    "file_id", "set_id", "edit_state", "role", "edit_operation",
    "edit_operation_note", "title", "description", "keywords", "file_format",
    "resolution", "color_space", "file_size_mb", "date_created",
    "date_contributed", "asset_category", "ai_used",
]
# `category` is the classifier's own 4 subject types (Portrait / Traditional /
# White BG / Others), appended after the 17 template columns for review.
EXTRA_COLUMNS = ["category"]

# Longest keyword list we keep (the prompt asks for 8-15), so a chatty
# model cannot bloat a cell.
MAX_KEYWORDS = 15

# Fractional aspect-ratio change that counts as a crop rather than a resize.
CROP_AR_TOLERANCE = 0.02

# set_id shape used by the template: SET_ + six digits (e.g. SET_000045).
SET_ID_DIGITS = 6


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


def format_set_id(group_id):
    """'45' / '0045' / 'SET_45' -> 'SET_000045' (the template's set_id shape)."""
    g = str(group_id or "").strip()
    if not g:
        return ""
    if g.isdigit():
        return "SET_" + str(int(g)).zfill(SET_ID_DIGITS)
    if g.upper().startswith("SET_"):
        return g
    return "SET_" + g


# --------------------------------------------------------------------------- #
# Locating the matched original
# --------------------------------------------------------------------------- #
def find_original(name, edit_dir, originals_dir=""):
    """Return the full path of `name`'s matched original, or ''.

    Tries, in order: the explicit --originals-dir, the flat `<shoot>/match`
    folder created by sort_photos.py, then any sibling folder of the edit dir.
    """
    if not name:
        return ""
    if originals_dir:
        p = os.path.join(originals_dir, name)
        if os.path.isfile(p):
            return p
    parent = os.path.dirname(os.path.abspath(edit_dir))
    p = os.path.join(parent, "match", name)
    if os.path.isfile(p):
        return p
    try:
        here = os.path.abspath(edit_dir)
        for d in sorted(os.listdir(parent)):
            dp = os.path.join(parent, d)
            if not os.path.isdir(dp) or os.path.abspath(dp) == here:
                continue
            p = os.path.join(dp, name)
            if os.path.isfile(p):
                return p
    except Exception:
        pass
    return ""


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
# Normalising model output
# --------------------------------------------------------------------------- #
def _clean_text(value):
    """Single-line, whitespace-collapsed text (keeps the Excel tidy)."""
    return " ".join(str(value or "").split())


def _clean_sentence(value):
    """Like _clean_text, but with no trailing period (template style)."""
    t = _clean_text(value)
    while t.endswith(".") and not t.endswith(".."):
        t = t[:-1].rstrip()
    return t


# Longest title we write. The prompt asks for 7-8 words; this is the hard cap so
# a chatty answer cannot overflow the cell.
MAX_TITLE_WORDS = 8

# Words a truncated title must never end on: articles, conjunctions, prepositions
# and bare possessives all leave the sentence hanging.
_DANGLING_WORDS = {
    "", "a", "an", "the", "and", "with", "of", "in", "at", "on", "for", "to",
    "by", "or", "nor", "but", "as", "from", "into", "onto", "over", "under",
    "near", "her", "his", "their", "its", "my", "our", "your", "this", "that",
    "these", "those", "some", "any", "all",
}

# Words that keep a capital in sentence case: proper nouns and names that turn up
# in photo-library titles. Everything else is lowercased unless it starts the
# sentence.
_PROPER_NOUNS = {
    "india", "indian", "jaipur", "delhi", "mumbai", "kolkata", "chennai",
    "bengaluru", "bangalore", "hyderabad", "goa", "kerala", "rajasthan",
    "gujarat", "punjab", "tamil", "nadu", "uttar", "pradesh", "maharashtra",
    "karnataka", "varanasi", "agra", "taj", "mahal", "himalaya", "himalayas",
    "ganges", "ganga", "diwali", "holi", "eid", "christmas", "navratri",
    "onam", "pongal", "durga", "puja", "sikh", "hindu", "muslim", "buddhist",
    "jain", "punjabi", "marathi", "bengali", "rajasthani",
    "iphone", "photoshop", "lightroom", "adobe",
}


def _normalise_title(value):
    """A sentence of at most MAX_TITLE_WORDS words, in sentence case.

    The prompt asks the model for this, but a model can always drift: this makes
    it true in the cell regardless. Sentence case means "capitalise the first
    word and proper nouns", so a title that arrives in Title Case is folded down
    ("Woman At Spice Market, Jaipur" -> "Woman at spice market, Jaipur").
    """
    text = _clean_sentence(value)
    if not text:
        return ""

    # Drop a trailing ellipsis/comma and collapse any doubled spaces.
    text = text.rstrip(" ,;:-")

    # Cap the length. Prefer cutting at a natural phrase boundary (a comma)
    # within the cap, so the title reads as a whole phrase rather than a
    # half-finished clause; fall back to a hard trim that never leaves a
    # dangling comma or conjunction.
    words = text.split()
    if len(words) > MAX_TITLE_WORDS:
        head = words[:MAX_TITLE_WORDS]
        cut = None
        for i in range(len(head) - 1, 3, -1):      # keep at least 4 words
            if head[i].endswith(","):
                cut = i
                break
        if cut:
            words = head[:cut]
        else:
            words = head
            while words and words[-1].strip(",;:-").lower() in _DANGLING_WORDS:
                words.pop()
        text = " ".join(words).rstrip(" ,;:-")

    # Sentence case. A word keeps its capital when it starts the title, is a
    # known proper noun, or carries an inner capital (iPhone, McLeod).
    out = []
    for i, w in enumerate(text.split()):
        core = w.strip(".,;:!?\"'()")
        keep = (i == 0
                or core.lower() in _PROPER_NOUNS
                or any(c.isupper() for c in core[1:]))
        if keep:
            out.append(w[:1].upper() + w[1:] if w else w)
        else:
            out.append(w.lower())
    return " ".join(out)


def _normalise_keywords(value):
    """Return a clean 'a,b,c' keyword string, no spaces after the commas.

    Casing is preserved so proper names stay capitalised ("India"), exactly like
    the template's keyword style; duplicates are removed case-insensitively.
    """
    if isinstance(value, (list, tuple)):
        parts = [str(v) for v in value]
    else:
        parts = str(value or "").replace(";", ",").replace("#", "").split(",")
    out, seen = [], set()
    for p in parts:
        k = _clean_text(p).strip(" .,;")
        if not k:
            continue
        low = k.lower()
        if low in seen:
            continue
        seen.add(low)
        out.append(k)
        if len(out) >= MAX_KEYWORDS:
            break
    return ",".join(out)


# Every synonym/sloppy term we have seen (or that the previous vocabulary used)
# mapped onto the client's eight. Longest phrases first, so "object removal"
# is matched before the bare word "object".
_OP_ALIASES = {
    # colour / tone
    "color_tonal": [
        "color_tonal", "colour_tonal", "color", "colour", "tonal", "tone",
        "color_grade", "colour_grade", "color_grading", "colour_grading",
        "grade", "grading", "lut", "vignette", "exposure", "brightness",
        "contrast", "saturation", "vibrance", "curves", "levels",
        "highlights", "shadows", "black_and_white", "black_white",
        "bw", "b&w", "monochrome", "sepia", "hdr", "film_look", "matte",
        "clarity_color", "cinematic_grade", "color_tonal_correction",
        "colour_tonal_correction",
    ],
    # framing / geometry
    "crop_perspective": [
        "crop_perspective", "crop", "cropped", "cropping", "perspective",
        "perspective_correction", "straightening", "straighten", "straightened",
        "rotation", "rotate", "rotated", "resize", "resized", "reframe",
        "reframing", "aspect_ratio", "keystone", "keystoning", "lens_correction",
        "lense_correction", "lense", "lens", "distortion", "tilt", "horizon",
    ],
    # colour cast only
    "white_balance": [
        "white_balance", "whitebalance", "wb", "color_cast", "colour_cast",
        "color_temperature", "colour_temperature", "temperature", "tint",
        "warmth", "cool_cast", "warm_cast", "neutralise", "neutralize",
        "white_balence", "white_balancing",
    ],
    # something disappeared
    "object_removal": [
        "object_removal", "object_removed", "remove_object", "removal",
        "remove", "removed", "erase", "erased", "erasing", "cleanup",
        "clean_up", "spot_removal", "spot_healing", "healing", "heal",
        "clone", "cloning", "clone_stamp", "content_aware", "inpainting",
        "dust_removal", "dust", "blemish_removal", "object_removeal",
    ],
    # detail only
    "sharpen_denoise": [
        "sharpen_denoise", "sharpen", "sharpening", "sharpened", "denoise",
        "denoising", "denoize", "noise_reduction", "noise", "grain",
        "grain_reduction",
        "texture", "detail_enhancement", "unsharp_mask", "micro_contrast",
        "smoothing", "smooth",
    ],
    # the backdrop changed
    "background": [
        "background", "background_replacement", "background_replaced",
        "background_removal", "background_removed", "background_change",
        "background_blur", "background_swap", "bg", "backdrop", "backdrop_change",
        "cutout", "cut_out", "masking", "subject_isolation", "white_backdrop",
        "transparent_background", "remove_background", "replace_background",
    ],
    # skin / face only
    "skin_portrait": [
        "skin_portrait", "retouch", "retouching", "retouched", "skin",
        "skin_smoothing", "skin_retouch", "portrait_retouch", "beauty",
        "beauty_retouch", "facial_retouch", "face_retouch", "blemish",
        "blemish_removal_skin", "teeth_whitening", "eye_brightening",
        "eye_enhancement", "figure_slimming", "body_reshape", "liquify",
        "makeup", "hair_cleanup",
    ],
    # two things combined / something added
    "compositing": [
        "compositing", "composite", "composed", "montage", "collage",
        "object_addition", "object_added", "add_object", "addition", "added",
        "sky_replacement", "sky_replaced", "sky_swap", "element_added",
        "pasted", "paste", "double_exposure", "text_overlay", "text",
        "watermark", "graphic_overlay", "logo_overlay",
    ],
}

# phrase -> canonical op, longest phrase first (so multi-word matches win).
_OP_LOOKUP = sorted(
    ((alias, op) for op, aliases in _OP_ALIASES.items() for alias in aliases),
    key=lambda kv: -len(kv[0]),
)

# Canonical order for output: fixed, so two runs of the same pair produce
# byte-identical cells and a spreadsheet filter is stable.
_OP_ORDER = {op: i for i, op in enumerate(EDIT_OPS)}


def canonical_op(term):
    """Map any operation term onto one of the eight, or '' if unrecognisable.

    Matching is on the normalised term (lowercase, single underscores) and tries
    the longest known phrase first, so "object_removal" is not mistaken for a
    generic "remove" and "background_removal" is not read as "object_removal".
    """
    k = _clean_text(term).lower().strip(" .;")
    k = k.replace("-", "_").replace(" ", "_")
    while "__" in k:
        k = k.replace("__", "_")
    if not k or k in ("none", "no_change", "n_a", "na", "unchanged", "nil"):
        return ""
    if k in _OP_ORDER:
        return k
    for alias, op in _OP_LOOKUP:
        if k == alias:
            return op
    # Substring fallback for phrases like "colour grading and retouch".
    for alias, op in _OP_LOOKUP:
        if len(alias) >= 4 and alias in k:
            return op
    return ""


def _op_parts(value):
    """Split a model answer into operation terms, whatever separator it used."""
    if isinstance(value, (list, tuple)):
        return [str(v) for v in value]
    # Models answer with every separator going: "crop, colour grade",
    # "crop; colour grade", "crop and colour grade", "crop + colour grade".
    text = str(value or "")
    for sep in (",", ";", " and ", " & ", " + ", " plus "):
        text = text.replace(sep, ";")
    return text.split(";")


def _normalise_ops(value):
    """Return 'op_a;op_b' using the template's semicolon style.

    Every term is mapped onto the client's eight operations (synonyms included),
    de-duplicated, and emitted in the fixed EDIT_OPS order so a pair always
    produces the same cell. Unknown terms are dropped rather than written,
    because the client can only filter on the eight.
    """
    found = set()
    for p in _op_parts(value):
        op = canonical_op(p)
        if op:
            found.add(op)
    return ";".join(sorted(found, key=lambda o: _OP_ORDER.get(o, 99)))


def ops_ordered(value):
    """The operations named by `value`, in the order it stated them.

    ``_normalise_ops`` re-sorts onto the fixed EDIT_OPS order so a single cell is
    stable across runs. When several sources are merged — the operation folder
    the client filed the batch under, then the model's answer — the ORDER is the
    information (primary first, then by confidence), so this keeps the caller's
    order and only canonicalises the individual terms.
    """
    out = []
    for p in _op_parts(value):
        op = canonical_op(p)
        if op and op not in out:
            out.append(op)
    return out


def dominant_op(value):
    """The FIRST (most important) operation in the model's own answer.

    The prompt asks for the operations most-important-first, so the first term
    is the one to file the image under. `_normalise_ops` re-sorts for a stable
    cell, which loses that ordering — so the dominant value is captured here,
    before sorting.
    """
    for p in _op_parts(value):
        op = canonical_op(p)
        if op:
            return op
    return ""


def _error_label(msg=""):
    """Uniform label dict for a failed classification."""
    return {"category": "ERROR", "title": "", "keywords": "",
            "original_description": "", "final_description": "",
            "edit_operation": "", "note": str(msg or "")}


def _is_good_label(value):
    """True if a cached/labelled value is a usable result (not an error)."""
    return (isinstance(value, dict)
            and value.get("category") not in (None, "", "ERROR"))


def _sanitise_label(label):
    """Re-apply the cell formatting rules to a label read from cache.

    A cache written by an earlier build (or a model that added trailing periods
    and stray spaces) must still produce template-style cells, so cached values
    go through the same normalisers as fresh API output.
    """
    return _normalise({
        "category": label.get("category", ""),
        "title": label.get("title", ""),
        "keywords": label.get("keywords", ""),
        "original_description": label.get("original_description", ""),
        "final_description": label.get("final_description", ""),
        "edit_operation": label.get("edit_operation", ""),
    }, has_original=bool(label.get("original_description")
                         or label.get("edit_operation")))


def _normalise(data, has_original=True):
    """Turn raw model output into a clean dict of label fields."""
    cat = str(data.get("category", "")).strip()
    for c in CATEGORIES:
        if c.lower() == cat.lower():
            cat = c
            break
    else:
        cat = "Others"
    final_desc = _clean_sentence(data.get("final_description")
                                 or data.get("description"))
    ops = _normalise_ops(data.get("edit_operation")) if has_original else ""
    return {
        "category": cat,
        "title": _normalise_title(data.get("title")),
        "keywords": _normalise_keywords(data.get("keywords")),
        "original_description": (_clean_sentence(data.get("original_description"))
                                 if has_original else ""),
        "final_description": final_desc,
        "edit_operation": ops,
    }


def geometry_ops(orig_meta, edit_meta):
    """Operations the file metadata proves, as a list — currently a real crop.

    Independent of what the model saw: if the two files have a materially
    different aspect ratio, something was cropped or the frame was reshaped.
    """
    try:
        ow, oh = (int(v) for v in str(orig_meta.get("resolution", "")).split("x"))
        ew, eh = (int(v) for v in str(edit_meta.get("resolution", "")).split("x"))
    except Exception:
        return []
    if min(ow, oh) <= 0 or min(ew, eh) <= 0:
        return []
    o_small, o_big = sorted((ow, oh))
    e_small, e_big = sorted((ew, eh))
    ar_o = o_big / o_small
    ar_e = e_big / e_small
    if abs(ar_o - ar_e) / ar_o > CROP_AR_TOLERANCE:
        return ["crop_perspective"]
    return []


def add_geometry_ops(ops, orig_meta, edit_meta):
    """Append 'crop_perspective' when the aspect ratio genuinely changed.

    This is a *measured* signal, independent of what the model saw: if the two
    files have a materially different aspect ratio, something was cropped or the
    frame was reshaped, whether or not the model mentioned it.

    Orientation is ignored (a portrait derivative of a landscape original is a
    resize, not a crop), and a plain resolution change never adds anything —
    that is already obvious from the file_size_mb / resolution columns.
    """
    if not geometry_ops(orig_meta, edit_meta):
        return ops
    parts = [p for p in (ops.split(";") if ops else []) if p]
    if "crop_perspective" not in parts:
        parts.append("crop_perspective")
    # Keep the fixed canonical order so the cell is stable across runs.
    return ";".join(sorted(parts, key=lambda o: _OP_ORDER.get(o, 99)))


# --------------------------------------------------------------------------- #
# LLM call
# --------------------------------------------------------------------------- #
def _anthropic_content(orig_b64, edit_b64, prompt):
    """Build the Anthropic message content: original, edit, then instructions."""
    content = []
    if orig_b64:
        content.append({"type": "text", "text": "IMAGE 1 — ORIGINAL camera file:"})
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg",
                       "data": orig_b64},
        })
        content.append({"type": "text", "text": "IMAGE 2 — FINAL delivered edit:"})
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg",
                       "data": edit_b64},
        })
    else:
        content.append({"type": "image",
                        "source": {"type": "base64", "media_type": "image/jpeg",
                                   "data": edit_b64}})
    content.append({"type": "text", "text": prompt})
    return content


def _openai_content(orig_b64, edit_b64, prompt):
    """Build the OpenAI message content: original, edit, then instructions."""
    def img(b64):
        return {"type": "image_url",
                "image_url": {"url": f"data:image/jpeg;base64,{b64}",
                              "detail": "low"}}

    content = []
    if orig_b64:
        content.append({"type": "text", "text": "IMAGE 1 — ORIGINAL camera file:"})
        content.append(img(orig_b64))
        content.append({"type": "text", "text": "IMAGE 2 — FINAL delivered edit:"})
        content.append(img(edit_b64))
    else:
        content.append(img(edit_b64))
    content.append({"type": "text", "text": prompt})
    return content


def _classify_anthropic(orig_b64, edit_b64, api_key, model, retries=3):
    prompt = PROMPT_PAIR if orig_b64 else PROMPT_SINGLE
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
        "content-type": "application/json",
    }
    payload = {
        "model": model,
        "max_tokens": MAX_TOKENS,
        "temperature": 0,
        "messages": [
            {"role": "user",
             "content": _anthropic_content(orig_b64, edit_b64, prompt)},
        ],
    }
    delay = 1.0
    for attempt in range(retries):
        try:
            r = requests.post(ANTHROPIC_URL, headers=headers, json=payload,
                              timeout=120)
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            content = r.json()["content"][0]["text"]
            return _normalise(parse_json_loose(content), bool(orig_b64))
        except Exception as e:
            if attempt == retries - 1:
                return _error_label(e)
            time.sleep(delay)
            delay *= 2
    return _error_label("no response")


def _classify_openai(orig_b64, edit_b64, api_key, model, retries=3):
    prompt = PROMPT_PAIR if orig_b64 else PROMPT_SINGLE
    headers = {"Authorization": f"Bearer {api_key}",
               "Content-Type": "application/json"}
    payload = {
        "model": model,
        "max_tokens": MAX_TOKENS,
        "temperature": 0,
        "messages": [
            {"role": "user",
             "content": _openai_content(orig_b64, edit_b64, prompt)},
        ],
        "response_format": {"type": "json_object"},
    }
    delay = 1.0
    for attempt in range(retries):
        try:
            r = requests.post(OPENAI_URL, headers=headers, json=payload,
                              timeout=120)
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            return _normalise(parse_json_loose(content), bool(orig_b64))
        except Exception as e:
            if attempt == retries - 1:
                return _error_label(e)
            time.sleep(delay)
            delay *= 2
    return _error_label("no response")


def classify_image(orig_b64, edit_b64, api_key, model, provider="anthropic",
                   retries=3):
    if provider == "openai":
        return _classify_openai(orig_b64, edit_b64, api_key, model, retries)
    return _classify_anthropic(orig_b64, edit_b64, api_key, model, retries)


# --------------------------------------------------------------------------- #
# Excel rows
# --------------------------------------------------------------------------- #
def make_row(file_id, set_id, edit_state, role, title, description, keywords,
             meta, category, edit_operation="", date_contributed="",
             edit_operation_note=""):
    """One Excel row in template order (blank cells are simply empty strings)."""
    return {
        "file_id": file_id,
        "set_id": set_id,
        "edit_state": edit_state,
        "role": role,
        "edit_operation": edit_operation,
        "edit_operation_note": edit_operation_note,
        "title": title,
        "description": description,
        "keywords": keywords,
        "file_format": meta.get("file_format", ""),
        "resolution": meta.get("resolution", ""),
        "color_space": meta.get("color_space", ""),
        "file_size_mb": meta.get("file_size_mb", ""),
        "date_created": meta.get("date_created", ""),
        "date_contributed": date_contributed,
        "asset_category": "",
        "ai_used": "",
        "category": category,
    }


def _safe_metadata(path, prefer_modified=False):
    """image_metadata() that never raises, so one odd file cannot stop a run."""
    try:
        return image_metadata(path, prefer_modified=prefer_modified)
    except Exception:
        return {}


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def read_report(report_path):
    with open(report_path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--edit-dir", required=True, help="folder containing the edits")
    ap.add_argument("--originals-dir", default="",
                    help="folder containing the originals (default: auto-find "
                         "via <shoot>/match or a sibling folder)")
    ap.add_argument("--report", default="sorted_output/report.csv",
                    help="report.csv from sort_photos.py (edit -> original map)")
    ap.add_argument("--out", default="sorted_output/photos.xlsx",
                    help="output Excel file")
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
    model = args.model or (DEFAULT_OPENAI_MODEL if provider == "openai"
                           else DEFAULT_ANTHROPIC_MODEL)

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

    step(1, "loading the match report from step 1")
    note(f"report : {args.report}")
    note(f"edits  : {args.edit_dir}")
    note(f"rows   : {len(rows)} in report, {len(edits)} edit file(s) to label")
    dropped = len(by_edit) - len(edits)
    if dropped:
        note(f"note   : {dropped} report row(s) have no matching file in the "
             f"edits folder (or are not images) - they are skipped")

    # Resolve every edit's original up front so the log can report pair coverage.
    originals = {f: find_original(by_edit[f].get("matched_original", ""),
                                 args.edit_dir, args.originals_dir)
                 for f in edits}
    n_pairs = sum(1 for p in originals.values() if p)
    if edits and not n_pairs:
        note("warn   : no matched originals found - pass --originals-dir "
             "(or run sort_photos.py first) for the full two-row layout")
    note(f"pairs  : {n_pairs} edit(s) with an original, "
         f"{len(edits) - n_pairs} single")

    if not edits:
        sys.exit("error: no edits to classify - is --edit-dir pointing at the "
                 "folder named in report.csv?")

    step(2, f"preparing the vision model ({provider})")
    note(f"model   : {model}")
    note(f"workers : {args.workers} image(s) in flight at once")
    note("what    : the original and its edit are sent together, so the model "
         "can name the real difference (crop, colour grade, retouching, ...).")
    note("cost    : roughly $0.001-0.002 per edit with this model.")

    # v6 cache: v5 predates the 7-8 word sentence title, so its titles would
    # arrive in Title Case and far too long. Keeping the name in step with the
    # prompt shape means a stale cache is ignored, not silently mixed in.
    cache_path = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".",
                              "labels_cache_v6.json")
    cache = {}
    if not args.no_cache and os.path.isfile(cache_path):
        try:
            loaded = json.load(open(cache_path, encoding="utf-8"))
            if isinstance(loaded, dict):
                cache = {k: _sanitise_label(v) for k, v in loaded.items()
                         if _is_good_label(v)}
        except Exception:
            cache = {}

    n_cached = sum(1 for f in edits if f in cache)
    n_todo = len(edits) - n_cached
    step(3, "checking the cache so you never pay twice for the same image")
    note(f"cache   : {cache_path}")
    if args.no_cache:
        note("--no-cache given: every label will be recomputed")
    note(f"result  : {n_cached} already labelled, {n_todo} still to classify")
    if n_todo == 0:
        note("nothing to do - all labels cached; writing outputs ...")

    lock = threading.Lock()
    results = {}
    done = [0]

    def work(fname):
        if fname in cache:
            with lock:
                done[0] += 1
            return fname, dict(cache[fname])
        try:
            edit_b64 = image_to_data_url(os.path.join(args.edit_dir, fname))
            orig_path = originals.get(fname, "")
            orig_b64 = image_to_data_url(orig_path) if orig_path else None
            label = classify_image(orig_b64, edit_b64, api_key, model, provider)
        except Exception as e:
            label = _error_label(e)
        with lock:
            done[0] += 1
            pct = 100.0 * done[0] / len(edits)
            elapsed = time.time() - _T0
            rate = done[0] / elapsed if elapsed > 0 else 0.0
            eta = (len(edits) - done[0]) / rate if rate > 0 else 0.0
            tail = f", ~{eta:.0f}s left" if done[0] < len(edits) else ""
            print(safe(f"  [{done[0]}/{len(edits)}] {fname}"), flush=True)
            for i, line in enumerate(wrap(
                    safe(f"-> {label.get('category', 'ERROR')} | "
                         f"{label.get('title') or '(no title)'}"),
                    indent=12, subsequent_indent=" " * 4)):
                print(" " * 8 + ("-> " if i == 0 else "   ") + line, flush=True)
            print(safe(f"        ({pct:.0f}% done, {elapsed:.0f}s{tail})"),
                  flush=True)
        return fname, label

    # Heartbeat: while API requests are in flight (they can take many seconds
    # each) print a periodic "still working" line so the run never looks stuck.
    stop = threading.Event()

    def heartbeat():
        while not stop.wait(15):
            with lock:
                left = len(edits) - done[0]
                log(f"  ... still working: {done[0]}/{len(edits)} labelled, "
                    f"{left} to go - waiting on API responses ...")

    step(4, f"classifying {n_todo} edit(s) with the vision model")
    note("each result is cached immediately, so an interrupted run resumes "
         "where it stopped instead of paying for the same image again.")
    hb = threading.Thread(target=heartbeat, daemon=True)
    hb.start()
    t_classify = time.time()
    try:
        with cf.ThreadPoolExecutor(max_workers=args.workers) as ex:
            for fname, res in ex.map(work, edits):
                results[fname] = res
    finally:
        stop.set()
    log(f"  classification finished in {time.time() - t_classify:.1f}s")

    with lock:
        for fname, res in results.items():
            if _is_good_label(res):
                cache[fname] = res
        try:
            # Closed explicitly (a stray handle blocks later writes on Windows).
            # NB: do not name this `cf` - that shadows the concurrent.futures
            # import and turns `cf` into a local for the whole function.
            with open(cache_path, "w", encoding="utf-8") as fh:
                json.dump(cache, fh, indent=0)
        except Exception as e:
            log(f"[warn] could not write cache: {e}")
    log(f"labels cached: {len(cache)} total")

    order = {c: i for i, c in enumerate(CATEGORIES)}
    sorted_edits = sorted(
        results,
        key=lambda f: (order.get(results[f].get("category", "ERROR"), 99), f))

    step(5, "writing the Excel workbook and the debug log")
    note("rows are grouped by category so similar photos sit together")
    note("two rows per pair: edit_state=00 (original) then 01 (final)")

    # ---- build rows: ONE original row + one row per edit in the set ------- #
    # A set is a matched original plus every edit made from it. The template
    # numbers the versions inside a set: 00 for the camera file, then 01, 02,
    # ... for each delivered edit. Writing the original once per edit would
    # duplicate it and give two rows the same edit_state.
    today = time.strftime("%Y-%m-%d")
    meta_cache = {}

    def meta_of(path, prefer_modified=False):
        key = (path, prefer_modified)
        if key not in meta_cache:
            meta_cache[key] = _safe_metadata(path, prefer_modified=prefer_modified)
        return meta_cache[key]

    sets = {}
    set_order = []
    for fname in sorted_edits:
        key = by_edit[fname].get("group_id", "") or fname
        if key not in sets:
            sets[key] = []
            set_order.append(key)
        sets[key].append(fname)

    excel_rows = []
    n_singles = 0
    for key in set_order:
        members = sets[key]
        first = members[0]
        rec = by_edit[first]
        set_id = format_set_id(rec.get("group_id", ""))
        cat = results[first].get("category", "ERROR")
        title = results[first].get("title", "")
        keywords = results[first].get("keywords", "")
        orig_path = originals.get(first, "") or next(
            (originals.get(f, "") for f in members if originals.get(f, "")), "")

        if not orig_path:
            for fname in members:
                n_singles += 1
                label = results[fname]
                excel_rows.append(make_row(
                    fname, set_id, "01", "final", label.get("title", ""),
                    label.get("final_description", ""),
                    label.get("keywords", ""),
                    meta_of(os.path.join(args.edit_dir, fname),
                            prefer_modified=True),
                    label.get("category", "ERROR"), "", today))
            continue

        orig_meta = meta_of(orig_path)
        excel_rows.append(make_row(
            os.path.basename(orig_path), set_id, "00", "original",
            title, results[first].get("original_description", ""), keywords,
            orig_meta, cat, "", today))

        for n, fname in enumerate(members, 1):
            label = results[fname]
            edit_meta = meta_of(os.path.join(args.edit_dir, fname),
                                prefer_modified=True)
            ops = add_geometry_ops(label.get("edit_operation", ""),
                                   orig_meta, edit_meta)
            excel_rows.append(make_row(
                fname, set_id, f"{n:02d}", "final",
                label.get("title", "") or title,
                label.get("final_description", ""),
                label.get("keywords", "") or keywords,
                edit_meta, label.get("category", cat), ops, today))

    all_columns = TEMPLATE_COLUMNS + EXTRA_COLUMNS
    note(f"excel   : {args.out}  ({len(excel_rows)} row(s), "
         f"{len(sorted_edits)} edit(s))")

    wb = Workbook()
    ws = wb.active
    ws.title = "Photos"
    ws.append(all_columns)
    for c in range(1, len(all_columns) + 1):
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
    role_fills = {"original": "F2F2F2", "final": ""}
    role_col = all_columns.index("role") + 1
    state_col = all_columns.index("edit_state") + 1
    file_col = all_columns.index("file_id") + 1
    cat_col = all_columns.index("category") + 1

    for row in excel_rows:
        ws.append([row.get(col, "") for col in all_columns])
        last = ws.max_row
        # The category cell is colour-coded per category, and the file_id cell
        # carries the same tint so a pair is scannable from the first column.
        cat = row.get("category", "")
        if cat in cat_fills:
            ws.cell(row=last, column=cat_col).fill = PatternFill(
                "solid", fgColor=cat_fills[cat])
            ws.cell(row=last, column=file_col).fill = PatternFill(
                "solid", fgColor=cat_fills[cat])
        fill = role_fills.get(row.get("role", ""))
        if fill:
            ws.cell(row=last, column=role_col).fill = PatternFill(
                "solid", fgColor=fill)
            ws.cell(row=last, column=state_col).fill = PatternFill(
                "solid", fgColor=fill)

    widths = {
        "file_id": 26, "set_id": 12, "edit_state": 11, "role": 10,
        "edit_operation": 24, "edit_operation_note": 18,
        "title": 40, "description": 62, "keywords": 50,
        "file_format": 12, "resolution": 12, "color_space": 13,
        "file_size_mb": 13, "date_created": 14, "date_contributed": 16,
        "asset_category": 22, "ai_used": 9, "category": 24,
    }
    for i, col in enumerate(all_columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 14)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(all_columns))}{ws.max_row}"
    try:
        wb.save(args.out)
    except PermissionError:
        # Windows keeps a lock while the workbook is open in Excel, so the save
        # fails with a bare traceback unless we say what to do about it.
        sys.exit(f"error: cannot write '{args.out}' - it is open in Excel. "
                 f"Close the file (or the Excel window) and run again.")
    except OSError as e:
        sys.exit(f"error: cannot write '{args.out}': {e}")

    # ---- write debug log (match diagnostics for troubleshooting) ---------- #
    log_path = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".",
                            "categorize_debug.log")
    log(f"writing debug log -> {log_path} ...")
    with open(log_path, "w", newline="", encoding="utf-8") as lf:
        lw = csv.writer(lf)
        lw.writerow(["group_id", "edit_file", "matched_original", "category",
                     "method", "confidence", "distance", "burst_margin",
                     "original_found", "edit_operation",
                     "title", "original_description", "final_description",
                     "keywords"])
        for fname in sorted_edits:
            label = results[fname]
            r = by_edit[fname]
            lw.writerow([r.get("group_id", ""), fname,
                         r.get("matched_original", ""),
                         label.get("category", "ERROR"), r.get("method", ""),
                         r.get("confidence", ""), r.get("distance", ""),
                         r.get("burst_margin", ""),
                         "yes" if originals.get(fname) else "no",
                         label.get("edit_operation", ""),
                         label.get("title", ""),
                         label.get("original_description", ""),
                         label.get("final_description", ""),
                         label.get("keywords", "")])
    print(safe(f"Debug log written: {log_path}"))

    from collections import Counter
    counts = Counter(results[f].get("category", "ERROR") for f in results)
    print(flush=True)
    print("  " + hr("-"), flush=True)
    print(safe("  RESULT - categories assigned"), flush=True)
    for c in CATEGORIES:
        print(safe(f"    {c:22s}: {counts.get(c, 0)}"), flush=True)
    if counts.get("ERROR"):
        print(safe(f"    {'ERROR':22s}: {counts['ERROR']}"), flush=True)
    print("  " + hr("-"), flush=True)
    print(safe(f"  rows written : {len(excel_rows)} "
               f"({len(set_order)} set(s) + {n_singles} single edit(s)) "
               f"x {len(all_columns)} columns "
               f"({len(TEMPLATE_COLUMNS)} template + "
               f"{len(EXTRA_COLUMNS)} extra: {', '.join(EXTRA_COLUMNS)})"),
          flush=True)
    print(safe(f"  open this    : {args.out}"), flush=True)


if __name__ == "__main__":
    main()