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

# Token budget: five fields, two compact descriptions, a long title and a
# 13-18 term keyword list. Sized so the JSON never gets cut off.
MAX_TOKENS = 700

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
    "TITLE — a real, readable sentence of 9 to 10 MEANINGFUL words. Count only "
    "content words: articles (a, an, the) do NOT count towards the 9-10, so "
    "the title may run to 11-13 words in total but must carry at least 9 and "
    "at most 10 meaningful words. Sentence case: capitalise only the first "
    "word and proper nouns (e.g. \"Woman buying fresh spices at a busy Jaipur "
    "market stall\"). It must read as one grammatical sentence, not a "
    "keyword list — so no comma-separated fragments, no filler like "
    "\"image of\" or \"photograph of\", and NO trailing period.\n\n"
)

_KEYWORDS_RULE = (
    "KEYWORDS — 13 to 18 lowercase keywords minimum 13 (capitalise place and "
    "proper names, e.g. \"India\"), comma-separated with NO space after the "
    "commas, ordered from most to least specific. Go well past the obvious: "
    "include the place, the subject/who, the activity, clothing and props, "
    "the setting, the lighting, the colour palette, the mood/emotion, the "
    "season or festival, the image style and any broad commercial use-case "
    "terms (e.g. \"spice market,turmeric,vendor,woman,saree,India,Jaipur,"
    "street,crowd,daylight,warm tones,travel,culture,retail\"). No "
    "duplicates, no hashtags, no trailing comma.\n\n"
)

# ---- prompt used when BOTH the original and the final edit are available ---- #
PROMPT_PAIR = (
    "You are tagging photos for a professional Indian photo-library business. "
    "You are shown TWO images of the SAME shot.\n"
    "  IMAGE 1 = the ORIGINAL camera file (untouched).\n"
    "  IMAGE 2 = the FINAL delivered edit of that same shot.\n\n"
    "WORK IN THIS ORDER. Do not skip step 1.\n"
    "  STEP 1 — DIFF FIRST. Compare the two images and name, before writing "
    "anything else, everything present in IMAGE 1 that is ABSENT in IMAGE 2: "
    "wall- or floor-mounted fittings (wiper, brush, holder, towel rail, hook), "
    "bottles, buckets, bins, boxes, cables, signs, furniture, people, cloths, "
    "clutter, stains, dust, stray objects - anything present in 1 but absent "
    "in 2.\n"
    "  This is the most commonly missed change. Removing something makes a "
    "scene look \"cleaner\", \"brighter\" or \"better lit\" — that impression "
    "is very often a REMOVAL dressed up as a lighting change. If a bathroom "
    "looks tidier, ask yourself what was physically taken off the wall or "
    "floor before you write \"improved lighting\".\n"
    "  If anything is gone, edit_operation MUST include object_removal and "
    "final_description MUST name the object and where it was.\n"
    "  STEP 2 — write the fields below from that diff.\n\n"
    + _CATEGORY_RULES + _TITLE_RULE + _KEYWORDS_RULE +
    "original_description — ONE sentence of about 20 words (18 to 22, never "
    "more than 25), describing exactly what is visible in IMAGE 1, the "
    "original: the main subject, what they are doing, and the setting or light "
    "(e.g. \"Woman selecting spices at a busy Jaipur market stall, warm "
    "daylight, crowd behind her\"). Be concrete and visual, never speculative. "
    "No trailing period.\n\n"
    "final_description — ONE sentence of about 20 words (18 to 22, never more "
    "than 25), describing IMAGE 2. If it is the same scene edited, START WITH "
    "\"Same scene with \" and name only what changed (e.g. \"Same scene with "
    "the crowd removed and a warmer colour grade\"). If the framing or subject "
    "changed substantially, describe the result directly instead. No trailing "
    "period.\n"
    "  IMPORTANT — if edit_operation includes object_removal, this sentence "
    "MUST name the thing that disappeared, in plain words: \"the bin removed\", "
    "\"the cable gone\", \"the second chair missing\", \"the clutter cleared "
    "from the desk\". A description that says only \"the desk adjusted\" or "
    "\"items rearranged\" does not say WHAT went, and that row is unusable. If "
    "several things went, name the most obvious one (or two).\n\n"
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

# Appended to the pair prompt when a first answer used object_removal without
# naming anything that went. Asking once more, with the omission spelled out, is
# far cheaper than reviewing 8,000 rows by hand and fixes the cause rather than
# only reporting the symptom.
PROMPT_REMOVAL_DETAIL = (
    "Your previous answer used object_removal but did not say WHAT was removed, "
    "so the description cannot be used. Look at the two images again and answer "
    "the same JSON with final_description naming the missing thing explicitly.\n"
    "  Name the ACTUAL OBJECT you can see is gone — its real name, the word you "
    "would use to describe it to someone who has never seen the photo. Compare "
    "IMAGE 1 and IMAGE 2 side by side and list what is in 1 but not in 2.\n"
    "  WRONG (never do this): \"one decorative item removed\", \"a cleaning tool "
    "removed\", \"some items removed\", \"the clutter cleared\", \"objects "
    "removed\", \"the product removed\" — these name a CATEGORY, not an object, "
    "and the row is still unusable.\n"
    "  RIGHT: \"the wiper removed from the wall\", \"the blue bottle removed "
    "from the shelf\", \"the second chair removed\", \"the cable removed from "
    "the desk\" — the specific thing, in the specific place.\n"
    "  If several things went, name the most obvious one. If genuinely nothing "
    "was removed, change edit_operation to what actually happened instead."
)


# The removal SWEEP. A cheap, single-purpose question asked of pairs whose first
# answer claimed NO removal but described only light/colour - the exact shape of
# the miss the client reported: a wall-mounted wiper was taken off, the room then
# looked "cleaner and better lit", and the model wrote "improved lighting and
# enhanced color tones". That is a removal dressed up as a lighting change.
#
# A separate request on purpose: as its own question the model has one job (diff
# the objects), which is what it skips when asked five things at once.
PROMPT_REMOVAL_SWEEP = (
    "Look at the two images ONLY for things that DISAPPEARED.\n"
    "  IMAGE 1 = the original.  IMAGE 2 = the final edit.\n"
    "Compare them object by object and list everything present in IMAGE 1 that "
    "is NOT in IMAGE 2. Look specifically at walls, floors, ceilings, corners, "
    "shelves, counters and tables for fittings, fixtures, holders, rails, "
    "hooks, brackets, signs, labels, bottles, buckets, bins, bags, boxes, "
    "cables, cloths, stains, marks, dust, rubbish and stray objects.\n"
    "  A removal often leaves a cleaner or better-lit result, so do NOT dismiss "
    "a change as \"just lighting\" until you have checked what is physically "
    "absent.\n\n"
    "Answer ONLY this JSON:\n"
    '{"removed_objects": "...", "removed_from": "...", "still_present": "yes|no"}\n'
    "  removed_objects - the concrete things that are gone, named specifically "
    "(\"the wall-mounted wiper\", \"the blue bucket\", \"the towel rail\"). "
    "Write \"none\" if nothing was removed.\n"
    "  removed_from - where each was, in a few words (\"left wall\", \"shelf "
    "above the counter\"). Empty if nothing was removed.\n"
    "  still_present - \"no\" if at least one object was removed, else \"yes\"."
)


# ---- prompt used when an edit has NO matched original ---------------------- #
PROMPT_SINGLE = (
    "You are tagging photos for a professional Indian photo-library business. "
    "Read the image and output four fields.\n\n"
    + _CATEGORY_RULES + _TITLE_RULE + _KEYWORDS_RULE +
    "final_description — ONE sentence of about 20 words (18 to 22, never more "
    "than 25), describing what is actually visible: the subject, the action, "
    "the setting, the lighting and the colours. Be concrete and visual. No "
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

# Keyword list bounds. The prompt asks for 13-18; MAX stops a chatty model from
# bloating a cell, MIN is enforced by topping up from the title/description when
# the model comes back short (see _normalise_keywords).
MIN_KEYWORDS = 13
MAX_KEYWORDS = 18

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


# Description length. The prompt asks for one sentence of about 20 words; this
# is the hard cap so a chatty answer cannot overflow the cell.
DESCRIPTION_MAX_WORDS = 25


def _normalise_description(value):
    """One compact sentence of about 20 words (hard cap 25).

    The prompt asks for this, but a model can drift, so the cell is made
    compliant regardless: the first clause is kept and anything past the cap is
    dropped, cutting at a comma/semicolon boundary when one exists so the text
    still reads as a whole phrase rather than a clipped one.
    """
    text = _clean_sentence(value)
    if not text:
        return ""

    # The keyword rule ("no space after the commas") sometimes bleeds into the
    # prose, giving "saturation,improved contrast". Restore the missing spaces
    # so a description always reads as a sentence.
    text = re.sub(r",(?=[^\s\d])", ", ", text)

    # Keep the first sentence only — a second clause is what made descriptions
    # run long before.
    m = re.search(r"[.!?](\s|$)", text)
    if m:
        text = text[: m.start()].rstrip(" ,;:-")

    words = text.split()
    if len(words) > DESCRIPTION_MAX_WORDS:
        head = words[:DESCRIPTION_MAX_WORDS]
        cut = None
        for i in range(len(head) - 1, 5, -1):      # keep at least 6 words
            if head[i].strip().endswith((",", ";", ":")):
                cut = i + 1
                break
        words = head[:cut] if cut else head

        # Never end on a hanging joining word.
        while len(words) > 3 and \
                words[-1].strip(",;:-").lower() in _DANGLING_WORDS:
            words.pop()

    return " ".join(words).rstrip(" ,;:-")


# Title length. The prompt asks for 9-10 MEANINGFUL words; articles (a, an, the)
# are not counted towards that number, so the budget below is enforced on the
# meaningful count, with a looser total-word cap as a second guard.
TITLE_MIN_WORDS = 9
TITLE_MAX_WORDS = 10

# Total words allowed in the cell: TITLE_MAX_WORDS content words plus a few
# articles. Anything longer is trimmed so a chatty answer cannot overflow.
MAX_TITLE_WORDS = 14

# Articles: ignored when counting the 9-10 meaningful words.
_ARTICLES = {"a", "an", "the"}

# Openers the prompt forbids ("image of a woman ..."); stripped before counting
# so the title starts on the subject. Longest phrase first.
_FILLER_OPENERS = (
    "a photograph of ", "a picture of ", "an image of ", "a photo of ",
    "photograph of ", "picture of ", "image of ", "photo of ",
)

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


def _meaningful_count(words):
    """How many words count towards the 9-10 title budget (articles excluded)."""
    return sum(1 for w in words
               if w.strip(".,;:!?\"'()").lower() not in _ARTICLES)


def _normalise_title(value):
    """A 9-10 meaningful-word sentence, in sentence case.

    The prompt asks the model for this, but a model can always drift: this makes
    it true in the cell regardless. Articles (a, an, the) do not count towards
    the 9-10 budget, and a longer answer is trimmed from the end without ever
    leaving a dangling article, conjunction or preposition.

    Sentence case means "capitalise the first word and proper nouns", so a title
    that arrives in Title Case is folded down ("Woman At Spice Market, Jaipur"
    -> "Woman at spice market, Jaipur").
    """
    text = _clean_sentence(value)
    if not text:
        return ""

    # Drop a trailing ellipsis/comma and collapse any doubled spaces.
    text = text.rstrip(" ,;:-")

    # Strip a forbidden "image of / photograph of" opener, so the title starts
    # on the subject and the opener does not waste a word of the budget.
    low = text.lower()
    for opener in _FILLER_OPENERS:
        if low.startswith(opener):
            text = text[len(opener):].lstrip(" ,;:-")
            break

    # Trim from the end while over budget — never below the 9-word floor, since
    # a short-but-whole title reads better than a maimed one.
    words = text.split()
    while words and (_meaningful_count(words) > TITLE_MAX_WORDS
                     or len(words) > MAX_TITLE_WORDS):
        if _meaningful_count(words) <= TITLE_MIN_WORDS:
            break
        words.pop()

    # Never end on a comma or a hanging joining word.
    while len(words) > 1 and words[-1].strip(",;:-").lower() in _DANGLING_WORDS:
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


# Filler words the keyword top-up must never promote into a keyword.
_KEYWORD_STOPWORDS = {
    "a", "an", "the", "and", "or", "but", "with", "without", "of", "in", "on",
    "at", "to", "from", "by", "for", "as", "into", "onto", "over", "under",
    "near", "is", "are", "was", "were", "be", "been", "being", "has", "have",
    "had", "this", "that", "these", "those", "it", "its", "his", "her", "their",
    "very", "while", "same", "scene", "image", "photo", "picture", "shot",
    "there", "here", "which", "who", "what", "when", "where", "how", "also",
}


def _normalise_keywords(value, extra_text=""):
    """Return a clean 'a,b,c' keyword string, no spaces after the commas.

    Casing is preserved so proper names stay capitalised ("India"), exactly like
    the template's keyword style; duplicates are removed case-insensitively.

    The client requires a minimum number of keywords, so a short model answer is
    topped up from `extra_text` (the title and descriptions of the same label):
    those words describe the very same photo, so promoting them into keywords is
    honest rather than invented.
    """
    if isinstance(value, (list, tuple)):
        parts = [str(v) for v in value]
    else:
        parts = str(value or "").replace(";", ",").replace("#", "").split(",")
    out, seen = [], set()

    def add(k):
        k = _clean_text(k).strip(" .,;:'\"")
        if not k or len(k) < 3:
            return
        low = k.lower()
        if low in seen or low in _KEYWORD_STOPWORDS:
            return
        seen.add(low)
        out.append(k)

    for p in parts:
        add(p)
        if len(out) >= MAX_KEYWORDS:
            break

    # Top up to the minimum from the label's own prose (title + descriptions).
    # The words keep their original casing (so "Jaipur" stays capitalised) and
    # their order of appearance, and words under four letters are skipped: they
    # are almost always too generic to earn a keyword slot.
    if len(out) < MIN_KEYWORDS and extra_text:
        for raw in str(extra_text).split():
            cand = raw.strip(".,;:!?\"'()[]")
            if len(cand) < 4:
                continue
            add(cand)
            if len(out) >= MIN_KEYWORDS:
                break

    return ",".join(out[:MAX_KEYWORDS])


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


# --------------------------------------------------------------------------- #
# Prose cues — the model's own description of the change is evidence too
# --------------------------------------------------------------------------- #
# The client asks for two things at once, and they pull in opposite directions:
#
#   * never list an operation that was NOT performed  -> do not invent;
#   * never omit one that WAS performed               -> do not under-report.
#
# The only evidence available is the model's own output. Its final_description
# describes the change in plain English ("Same scene with the crowd removed and
# a warmer colour grade"), and vision models routinely under-report the short
# `edit_operation` enum while describing the work correctly in that sentence.
# So the sentence is read — but ONLY for words that state work was DONE:
#
#   _OP_CUES_CERTAIN  an explicit performed-work verb: removed, added, cropped,
#                     retouched, background replaced, greys neutralised, …
#                     Corroborates the model and fills gaps it left.
#   _OP_CUES_VAGUE    says the frame was worked on but not how: "adjusted",
#                     "slightly altered", "minor edits", "rearranged".
#                     Used ONLY to fill a cell that would otherwise be blank —
#                     mapping a vague word onto a specific operation is a guess,
#                     so it must never be added beside a real answer.
#
# Deliberately absent: STATE descriptors ("visible", "placed on", "positioned",
# "stacked", "in the background", "on the table"). They say what is in the
# frame, not what was done to it — reading them as work is exactly what used to
# report `compositing` for "dust visible on the desk" and "cup placed on a
# coaster". They are never used, at any tier.
#
# A phrase that says nothing happened suppresses every tier, so the cell is told
# the truth ("none") instead of being invented.
_NO_CHANGE_CUES = (
    "no visible change", "no visible changes", "no change", "no changes",
    "unchanged", "identical", "nothing changed", "no discernible",
    "no noticeable change", "pixel dimensions only",
)

_OP_CUES_CERTAIN = (
    # work that makes something disappear
    ("object_removal", (
        "removed", "remove", "removal", "erased", "erasing", "deleted",
        "disappeared", "gone", "no longer there", "taken out", "cloned out",
        "cleaned up", "cleaning up", "spot healed", "spot healed out",
        "dust removed", "dust spot removed", "dusted off",
    )),
    # work that puts something in, or combines frames
    ("compositing", (
        "added", "adding", "inserted", "pasted", "combined", "composite",
        "merged", "overlaid", "overlay", "watermark", "double exposure",
        "new element", "element added", "logo added", "text added",
        "graphic added", "number changed", "design changed", "label changed",
        "text changed", "sign changed", "pattern changed", "logo changed",
        "screen replaced", "sky replacement", "sky swapped", "introduced",
    )),
    # the backdrop itself was worked on
    ("background", (
        "background replaced", "background removed", "background cleaned",
        "background changed", "background blurred", "background extended",
        "background swapped", "background swapped out", "backdrop replaced",
        "backdrop changed", "cut out", "cutout", "isolated on white",
        "white background added",
    )),
    # work on a person's skin / face / body
    ("skin_portrait", (
        "skin retouched", "skin smoothed", "skin smoothing", "blemish healed",
        "blemish removed", "retouched", "retouch", "teeth whitened",
        "eyes brightened", "eye brightened", "complexion smoothed",
        "hair cleaned", "hair cleanup", "makeup applied", "figure slimmed",
        "body reshaped", "liquified", "liquify",
    )),
    # the frame itself was changed
    ("crop_perspective", (
        "cropped", "crop", "reframed", "re framed", "straightened", "rotated",
        "perspective corrected", "perspective correction", "aspect ratio",
        "tighter frame", "tighter crop", "wider frame", "lens corrected",
        "lens correction", "horizon straightened", "keystone corrected",
    )),
    # colour cast corrected
    ("white_balance", (
        "white balance", "colour cast removed", "color cast removed",
        "yellow cast", "blue cast", "warmer cast", "cooler cast",
        "colour temperature", "color temperature", "neutralised",
        "neutralized", "tint corrected", "white balanced",
    )),
    # detail handling
    ("sharpen_denoise", (
        "sharpened", "sharpen", "denoised", "noise reduced", "grain reduced",
        "clarity added", "micro contrast", "texture enhanced",
    )),
    # colour / tone grading
    ("color_tonal", (
        "colour graded", "color graded", "colour grade", "color grade",
        "colour grading", "color grading", "black and white", "monochrome",
        "sepia toned", "vignette added", "tone mapped", "lut applied",
    )),
)

# Vague, but they do say the frame was worked on. FILL-ONLY: mapping one of
# these onto a specific operation is a guess, so it is never added beside a
# real answer - only used when the cell would otherwise be empty.
_OP_CUES_VAGUE = (
    ("color_tonal", (
        "adjusted", "adjustment", "adjustments", "tweaked", "tweak",
        "altered", "changed", "minor edit", "minor edits", "edited",
        "polished", "enhanced", "brighter", "darker", "warmer", "cooler",
        "lighter", "more vivid", "saturation", "contrast", "vibrance",
        "tonal", "tone", "mood", "exposure", "colour", "color", "light",
    )),
    ("compositing", (
        "repositioned", "re arranged", "rearranged", "moved", "shifted",
        "different arrangement", "arranged differently", "swapped places",
    )),
)


def ops_from_text(text, generic=False):
    """The operations a piece of prose proves were performed.

    `generic=False` (the default) reads only the CERTAIN verbs — explicit work
    such as removed / added / cropped / retouched. That tier is safe to add
    beside a model answer, which is how a performed operation the model forgot
    to list still reaches the sheet.

    `generic=True` additionally reads the VAGUE tier ("slightly adjusted",
    "rearranged") and is therefore only used to fill a cell that is otherwise
    empty: a vague word cannot prove WHICH operation was performed.

    State descriptors ("visible", "placed on") are never read at either tier —
    they describe the scene, not the work done to it.
    """
    low = _clean_text(text).lower()
    if not low:
        return []
    low = low.replace("_", " ").replace("-", " ")
    while "  " in low:
        low = low.replace("  ", " ")
    if any(cue in low for cue in _NO_CHANGE_CUES):
        return []
    out = []
    tables = ((_OP_CUES_CERTAIN, _OP_CUES_VAGUE) if generic
              else (_OP_CUES_CERTAIN,))
    for table in tables:
        for op, cues in table:
            if op in out:
                continue
            if any(re.search(rf"\b{re.escape(cue)}", low) for cue in cues):
                out.append(op)
    return out


# Words that prove a description NAMES something that disappeared. Kept broad so
# "clutter minimized", "fewer items visible", "tidied away" all count - the point
# is that a reader learns WHAT went, not the exact verb the model chose.
# Multi-word cues, mapped onto a single head verb before the word loop runs.
# The loop matches one word at a time, so a two-word cue ("no longer there",
# "replaced by") can never be seen for what it is; rewriting the phrase to
# "removed" puts the object on the correct side of the cue again.
_REMOVAL_PHRASES = (
    ("no longer", "removed"),
    ("replaced by", "removed"),
    ("replaced with", "removed"),
    ("swapped for", "removed"),
    ("swapped with", "removed"),
    ("instead of", "removed"),
    ("taken out", "removed"),
    ("took out", "removed"),
    ("edited out", "removed"),
    ("cloned out", "removed"),
    ("cleaned up", "removed"),
    ("cleared away", "removed"),
    ("clearing away", "removed"),
    ("wiped off", "removed"),
    ("swept away", "removed"),
)

# Single-word cues. The noun is looked for BEFORE the cue for a past participle
# ("the cable REMOVED") and AFTER it for a quantifier ("FEWER bottles").
_REMOVAL_NAMED = (
    "removed", "remove", "removal", "erased", "deleted", "disappeared", "gone",
    "cleared", "cleaner", "tidied", "tidy", "minimized", "minimised", "fewer",
    "missing", "dropped", "swept", "wiped", "emptied", "without", "replaced",
    "swapped",
)


# Nouns that name a CATEGORY, not an object. "one decorative item removed" or
# "a cleaning tool removed" tells the reader nothing they can act on - the same
# words would fit a hundred different pairs. A description whose only noun
# before the removal cue is one of these is treated as missing detail, exactly
# like a sentence with no cue at all.
_GENERIC_NOUNS = {
    "item", "items", "object", "objects", "element", "elements", "thing",
    "things", "stuff", "clutter", "mess", "debris", "contents", "detail",
    "details", "area", "areas", "background", "surface", "surfaces", "tool",
    "tools", "product", "products", "decoration", "decorations", "decor",
    "piece", "pieces", "container", "containers", "equipment", "supplies",
    "supply", "material", "materials", "part", "parts", "bit", "bits",
    "aspect", "feature", "features", "component", "components", "good",
    "goods", "merchandise", "paraphernalia", "accessory", "accessories",
    "prop", "props", "fixture", "fixtures", "belonging", "belongings",
    "offering", "offerings", "set", "sets", "group", "groups", "range",
    "number", "amount", "everything", "anything", "something", "nothing",
}

# Words that are structure, not content: they can sit between the noun and the
# cue ("the SECOND chair MISSING", "a SMALL pile of BOTTLES was cleared").
_DETERMINERS = {
    "a", "an", "the", "one", "two", "three", "four", "five", "six", "seven",
    "eight", "nine", "ten", "some", "few", "several", "many", "most", "all",
    "any", "each", "every", "both", "other", "another", "more", "less",
    "small", "large", "larger", "smaller", "big", "bigger", "little", "extra",
    "additional", "single", "entire", "whole", "main", "minor", "major",
    "decorative", "cleaning", "various", "certain", "visible", "stray",
    "loose", "remaining", "unwanted", "one's", "its", "their", "his", "her",
}

# Cues that are followed by the thing that went ("fewer BOTTLES", "without the
# CABLE") rather than preceded by it ("the cable was removed").
# Cues that quantify rather than describe ("fewer BOTTLES", "without the
# CABLE") name their object AFTER the cue.
_CUES_FORWARD = {"fewer", "without"}

# Prepositions end a noun phrase. "fewer items ON the counter" means the counter
# is the LOCATION, not what disappeared, so the forward window stops there.
_PHRASE_END = {
    "on", "in", "at", "from", "to", "for", "of", "with", "by", "near",
    "behind", "under", "above", "below", "beside", "around", "across",
    "through", "over", "onto", "into", "than", "that", "which",
}


def _is_concrete(word):
    """True if `word` could name a real object in a photo."""
    w = word.strip(".,;:!?()[]\"'")
    if len(w) < 3:
        return False
    if w in _GENERIC_NOUNS or w in _DETERMINERS:
        return False
    if w.isdigit():
        return False
    if w in _DANGLING_WORDS or w in _KEYWORD_STOPWORDS:
        return False
    return True


def names_removal(text):
    """True if `text` names an OBJECT that disappeared.

    `object_removal` is only useful if the reader learns *what* went, and a
    category is not an object:

        "the cables removed"              -> names it
        "one decorative item removed"     -> does NOT (a category)
        "a cleaning tool removed"         -> does NOT (a category)
        "the desk adjusted"               -> does NOT (no removal at all)

    Multi-word cues are rewritten to a single verb first ("no longer there" and
    "replaced by" both become "removed"), because the word loop below matches one
    word at a time and those phrases name the object BEFORE the cue.

    Then each cue is checked for a concrete noun in the words around it - before
    it for "…removed", after it for "fewer…", "without…". This is the check
    behind `removal_detail_missing`, so the gap is found automatically instead of
    by reading 8,000 rows.
    """
    low = re.sub(r"[^a-z0-9' ]+", " ", _clean_text(text).lower())
    for phrase, verb in _REMOVAL_PHRASES:
        # Keep the word count intact so the noun window still looks right, but
        # collapse the phrase onto its head verb: "no longer" -> "removed".
        low = re.sub(rf"\b{re.escape(phrase)}\b", verb, low)
    words = low.split()
    if not words:
        return False
    for i, w in enumerate(words):
        cue = next((c for c in _REMOVAL_NAMED if w.startswith(c)), "")
        if not cue:
            continue
        if any(w.startswith(c) for c in _CUES_FORWARD):
            # "FEWER bottles", "WITHOUT the cable": the object follows the cue,
            # so take the words up to the next preposition. "fewer items ON the
            # counter" must not count "counter" as the thing that went.
            window = []
            for nxt in words[i + 1:i + 5]:
                if nxt in _PHRASE_END:
                    break
                window.append(nxt)
        else:
            window = words[max(0, i - 5):i]
        if any(_is_concrete(x) for x in window):
            return True
    return False


def removal_detail_missing(ops, description):
    """True when a row claims object_removal but never says what was removed.

    Those rows are the ones a reviewer cannot act on: the operation says
    something disappeared, the sentence does not say what, and nothing in the
    sheet lets anyone find out without opening the pair. They are counted,
    listed, and (when a model is available) asked once more.
    """
    if "object_removal" not in ops_ordered(ops):
        return False
    return not names_removal(description)


def complete_ops(model_ops, hint_text="", measured=None, fallback=""):
    """The edit_operation cell: model answer, then measured, then prose.

    Evidence order — what the vision model actually reported, then what the file
    metadata proves (a changed aspect ratio IS a crop), and **only if those are
    empty**, what the label's own description names.

    The prose is deliberately a LAST RESORT that only fills a blank cell. It
    must never extend a real answer: a description like "scissors added on the
    desk" is a model error that already said `object_removal`, and adding
    `compositing` on top would record an operation no model reported and no
    human confirmed. Blank is a defect; invented is worse.

    `fallback` is written when nothing at all was found. A pair whose two files
    really are alike has no operation to name, and the caller passes "none" so
    the client sees that stated instead of an empty cell that looks like a lost
    row.
    """
    parts = ops_ordered(model_ops)
    for op in measured or []:
        if op not in parts:
            parts.append(op)
    if not parts:
        # Blank cell -> read the label's own sentence. Never used to extend.
        parts = ops_from_text(hint_text, generic=True)
    return ";".join(parts) if parts else fallback


# Operations that describe the SAME aspect of an image, in mutually exclusive
# directions. "something was removed" and "something was added" cannot both be
# the story of one pair; neither can "colour graded" and "colour cast
# corrected". A description that names the opposite direction from the model is
# a contradiction to flag, NOT a second operation to append.
_OP_AXIS = {
    "object_removal": "objects",
    "compositing": "objects",
    "color_tonal": "colour",
    "white_balance": "colour",
    "background": "backdrop",
    "crop_perspective": "frame",
    "sharpen_denoise": "detail",
    "skin_portrait": "skin",
}


def _axis_taken(ops):
    """{axis: op} for the operations already in the cell."""
    out = {}
    for op in ops:
        axis = _OP_AXIS.get(op)
        if axis and axis not in out:
            out[axis] = op
    return out


def op_review(cell_ops, model_ops, hint_text="", conflict=()):
    """Explain any place the model and its own description name opposing work.

    `conflict` is the list of `(prose_op, model_op)` pairs that `explain_ops`
    refused to merge because they sit on the same axis pointing opposite ways -
    the classic case, seen in real data::

        model : object_removal
        prose : "Same scene with scissors added on the desk"

    One picture, one change. Calling it both a removal and an addition would be
    asserting two facts that cannot both hold, which is exactly the "lists an
    operation that was not performed" failure the client rejected. So the
    description's term is NOT written; the row is listed instead and a human
    settles it in seconds by looking at the pair.

    Returns ``''`` when nothing conflicts.
    """
    if not conflict:
        return ""
    parts = [f"description says {p} where the model says {m}"
             for p, m in conflict]
    return "; ".join(parts) + " - review"


def explain_ops(folder_op="", model_ops="", measured=None, hint_text="",
                fallback=""):
    """The same decision as `merge_ops`/`complete_ops`, with its evidence laid out.

    Returns ``(cell, sources, review)``:

      ``cell``    the value written to the spreadsheet, folder name first;
      ``sources`` which term came from where (`folder` / `model` / `measured` /
                  `prose` / `fallback`), for the debug log;
      ``review``  a non-empty reason when the model's answer and the label's own
                  description disagree, so the row can be settled by a human.

    The description may only ADD a term when the model produced nothing at all.
    Anything else it names beyond the model's answer becomes a `review` line
    instead of an extra term — which is how a performed operation is surfaced
    without inventing one.
    """
    sources = {}
    parts = []

    def push(op, why):
        if op and op not in parts:
            parts.append(op)
            sources[op] = why

    if folder_op:
        push(folder_op, "folder")
        already = set(ops_ordered(folder_op))
    else:
        already = set()

    for op in ops_ordered(model_ops):
        if op not in already:
            push(op, "model")
    for op in measured or []:
        if op not in already:
            push(op, "measured")

    # The CERTAIN tier ADDS: an explicit performed-work verb ("the crowd
    # removed", "cropped tighter") proves that operation was done, whether or
    # not the model listed it. This is how a performed operation the model
    # forgot still reaches the sheet. Nothing is guessed here - the sentence
    # states the work.
    #
    # The one exception is an OPPOSITE on the same axis: if the model already
    # says something was removed, a description saying something was added is
    # not extra work, it is a disagreement about the same change. Appending it
    # would state that both happened, which is the invention the client
    # rejected - so it is recorded as a conflict for a human instead.
    conflict = []
    axes = _axis_taken(parts)
    for op in ops_from_text(hint_text):
        if op in parts or op in already:
            continue
        axis = _OP_AXIS.get(op)
        if axis and axis in axes:
            conflict.append((op, axes[axis]))
            continue
        push(op, "prose")
        if axis:
            axes.setdefault(axis, op)

    if not parts:
        # Nothing anywhere -> fall back to the vague tier, then to the caller's
        # sentinel. A vague word cannot prove WHICH operation was performed, so
        # it is only ever used when there is no other evidence at all.
        for op in ops_from_text(hint_text, generic=True):
            push(op, "prose")
    if not parts:
        push(fallback or "none", "fallback")

    return ";".join(parts), sources, op_review(";".join(parts), model_ops,
                                               hint_text, conflict)


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
    final_desc = _normalise_description(data.get("final_description")
                                        or data.get("description"))
    orig_desc = (_normalise_description(data.get("original_description"))
                 if has_original else "")
    title = _normalise_title(data.get("title"))
    ops = _normalise_ops(data.get("edit_operation")) if has_original else ""
    return {
        "category": cat,
        "title": title,
        # The label's own prose is passed in so a short keyword list can be
        # topped up to the client's 13-keyword minimum.
        "keywords": _normalise_keywords(
            data.get("keywords"),
            extra_text=" ".join((title, orig_desc, final_desc))),
        "original_description": orig_desc,
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


def _classify_anthropic(orig_b64, edit_b64, api_key, model, retries=3,
                        extra_prompt=""):
    prompt = PROMPT_PAIR if orig_b64 else PROMPT_SINGLE
    if extra_prompt:
        prompt = prompt + "\n\n" + extra_prompt
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


def _classify_openai(orig_b64, edit_b64, api_key, model, retries=3,
                     extra_prompt=""):
    prompt = PROMPT_PAIR if orig_b64 else PROMPT_SINGLE
    if extra_prompt:
        prompt = prompt + "\n\n" + extra_prompt
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
                   retries=3, extra_prompt=""):
    if provider == "openai":
        return _classify_openai(orig_b64, edit_b64, api_key, model, retries,
                                extra_prompt)
    return _classify_anthropic(orig_b64, edit_b64, api_key, model, retries,
                               extra_prompt)


def classify_pair(orig_b64, edit_b64, api_key, model, provider="anthropic",
                  retries=3, sweep=True):
    """Label a pair, then ask follow-ups for the two documented misses.

    `sweep=True` (the default) also catches the case the client reported: a
    wall-mounted wiper was taken off, the room then looked cleaner, and the model
    wrote "improved lighting and enhanced color tones" without ever mentioning a
    removal. A second, single-purpose question fixes that.

    Both follow-ups can only ever IMPROVE the answer - a bad one leaves the first
    label standing - so neither can make the sheet worse.

    Returns ``(label, follow_up)`` where `follow_up` is ``''``, ``'detail'`` or
    ``'sweep'``, for the run log.
    """
    label = classify_image(orig_b64, edit_b64, api_key, model, provider,
                           retries)
    if not orig_b64:
        return label, ""

    # (a) object_removal with no object named -> ask for the noun.
    if remove_detail_needed(label):
        retry = classify_image(orig_b64, edit_b64, api_key, model, provider,
                               retries, extra_prompt=PROMPT_REMOVAL_DETAIL)
        if not remove_detail_needed(retry) and \
                (retry.get("edit_operation") or "").strip():
            return retry, "detail"
        return label, ""

    # (b) no removal claimed, but the description only talks about light/colour
    #     -> ask the single-purpose "what is gone?" question.
    if sweep and removal_sweep_needed(label):
        found = classify_image(orig_b64, edit_b64, api_key, model, provider,
                               retries, extra_prompt=PROMPT_REMOVAL_SWEEP)
        upgraded = fold_removal_sweep(label, found)
        if upgraded is not label:
            return upgraded, "sweep"
    return label, ""


def remove_detail_needed(label):
    """True if this label used object_removal without naming the removed thing."""
    if not isinstance(label, dict) or label.get("category") == "ERROR":
        return False
    return removal_detail_missing(label.get("edit_operation", ""),
                                  label.get("final_description", ""))


# Descriptions that read like "nothing was taken away". When a first answer is
# one of these AND claims no object_removal, the removal sweep is worth one cheap
# call - precisely the shape of the miss reported by the client.
_LIGHT_ONLY_CUES = (
    "improved lighting", "better lighting", "enhanced lighting",
    "improved light", "lighting and", "and enhanced color", "color tone",
    "colour tone", "color tones", "colour tones", "brighter appearance",
    "enhanced color", "enhanced colour", "enhanced brightness",
    "color adjustments", "colour adjustments", "tonal adjustments",
)


def removal_sweep_needed(label):
    """True if this label claims NO removal but only talks about light/colour.

    That combination is the documented blind spot: a removed fixture leaves the
    room looking cleaner and better lit, the model describes the *impression*
    ("Same scene with improved lighting and enhanced color tones"), and the
    removal never reaches the sheet.

    A row that already names a change beyond light/colour is left alone, and a
    row that already includes object_removal is handled by
    `remove_detail_needed` instead.
    """
    if not isinstance(label, dict) or label.get("category") == "ERROR":
        return False
    if "object_removal" in ops_ordered(label.get("edit_operation", "")):
        return False
    desc = _clean_text(label.get("final_description", "")).lower()
    if not desc:
        return False
    return any(cue in desc for cue in _LIGHT_ONLY_CUES)


def fold_removal_sweep(label, sweep):
    """Merge a removal-sweep answer into a label, when it found something.

    The sweep can only ADD a finding: if it says "none", or names a category
    rather than an object, the original label stands untouched, so a bad sweep
    can never damage a good row.

    The finding is appended in the client's existing style ("Same scene with
    improved lighting and the wall-mounted wiper removed"), and `object_removal`
    is put at the FRONT of the operation list, because it is now the most
    important thing that happened to the image.
    """
    if not isinstance(sweep, dict):
        return label
    removed = _clean_text(sweep.get("removed_objects", ""))
    where = _clean_text(sweep.get("removed_from", ""))
    if not removed or removed.lower().strip(".") in ("none", "nothing", "no", "nil"):
        return label
    # Only accept a CONCRETE object - the same rule the detail check enforces.
    if not names_removal(f"{removed} removed"):
        return label

    out = dict(label)
    # Word order matters for readability: "the wipers removed from the wall",
    # NOT "the wipers from the wall removed".
    phrase = (f"{removed} removed from the {where}" if where
              else f"{removed} removed")
    desc = _clean_sentence(out.get("final_description", ""))
    if removed.lower() not in desc.lower():
        desc = f"{desc} and {phrase}" if desc else phrase.capitalize()
        out["final_description"] = _normalise_description(desc)

    ops = ops_ordered(out.get("edit_operation", ""))
    if "object_removal" in ops:
        ops.remove("object_removal")
    ops.insert(0, "object_removal")
    out["edit_operation"] = ";".join(ops)
    return out


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

    # v8 cache: v7 holds multi-sentence 35-70 word descriptions, which are far
    # longer than the one-sentence ~20 words the template wants; v6 and earlier
    # also hold 7-8 word titles and short keyword lists. Keeping the name in
    # step with the prompt shape means a stale cache is ignored rather than
    # silently mixed in.
    cache_path = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".",
                              "labels_cache_v8.json")
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
    n_asked_again = [0]
    n_follow_ups = {"detail": 0, "sweep": 0}

    def work(fname):
        if fname in cache:
            with lock:
                done[0] += 1
            return fname, dict(cache[fname])
        try:
            edit_b64 = image_to_data_url(os.path.join(args.edit_dir, fname))
            orig_path = originals.get(fname, "")
            orig_b64 = image_to_data_url(orig_path) if orig_path else None
            # classify_pair asks once more when the model used object_removal
            # without naming what went, so the cell is actionable.
            label, follow_up = classify_pair(orig_b64, edit_b64, api_key,
                                             model, provider)
            if follow_up:
                with lock:
                    n_asked_again[0] += 1
                    n_follow_ups[follow_up] += 1
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
            # The cell reads: what the model reported, then a crop the file
            # metadata proves, then anything the label's own description names
            # ("Same scene with the crowd removed ..." is the answer, even when
            # the model left edit_operation empty). A pair that truly shows no
            # change says "none" rather than leaving the cell blank.
            ops = complete_ops(
                label.get("edit_operation", ""),
                hint_text=" ".join((label.get("final_description", ""),
                                    label.get("title", ""))),
                measured=geometry_ops(orig_meta, edit_meta),
                fallback="none")
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
    # ---- object_removal rows that still do not say WHAT went ------------ #
    # The prompt now demands it and a miss is asked once more, so this should be
    # empty. When it is not, the row is listed rather than shipped looking
    # finished: an object_removal whose description says only "the desk
    # adjusted" cannot be acted on.
    missing = []
    n_removal = 0
    for fname in sorted_edits:
        ops = results[fname].get("edit_operation", "")
        if "object_removal" not in ops_ordered(ops):
            continue
        n_removal += 1
        if removal_detail_missing(ops, results[fname].get("final_description", "")):
            missing.append(fname)

    print(safe(f"  rows written : {len(excel_rows)} "
               f"({len(set_order)} set(s) + {n_singles} single edit(s)) "
               f"x {len(all_columns)} columns "
               f"({len(TEMPLATE_COLUMNS)} template + "
               f"{len(EXTRA_COLUMNS)} extra: {', '.join(EXTRA_COLUMNS)})"),
          flush=True)
    if n_removal:
        print(safe(f"  object_removal rows          : {n_removal}"), flush=True)
        print(safe(f"  ... asked again to name it   : {n_asked_again[0]}"),
              flush=True)
    if n_follow_ups["sweep"]:
        print(safe(f"  ... removal sweep found one  : {n_follow_ups['sweep']} "
                   f"row(s) that only said lighting/colour but had something "
                   f"physically removed"), flush=True)
    if missing:
        print(safe(f"  [REVIEW] {len(missing)} object_removal row(s) still do "
                   f"NOT say what was removed - fix these by hand or re-run "
                   f"those pairs with --no-cache:"), flush=True)
        for fname in missing[:10]:
            print(safe(f"    {fname}: "
                       f"{results[fname].get('final_description', '')}"),
                  flush=True)
        if len(missing) > 10:
            print(safe(f"    ... and {len(missing) - 10} more"), flush=True)
    elif n_removal:
        print(safe("  [ok] every object_removal row names what was removed"),
              flush=True)
    print(safe(f"  open this    : {args.out}"), flush=True)


if __name__ == "__main__":
    main()