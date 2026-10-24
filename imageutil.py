#!/usr/bin/env python3
"""
imageutil.py — shared image-loading helpers for the photo pipeline.

Centralises two things that both sort_photos.py and categorize.py need:

  * which file extensions count as "images" (now including Canon RAW: .cr3,
    .cr2, plus the other common RAW formats), and
  * how to actually decode a file into a normal Pillow RGB/L image — using
    Pillow for ordinary raster files and ``rawpy`` (libraw) for RAW files.

Why this exists: Pillow cannot decode CR3 (Canon's BMFF/HEIF-based RAW) or the
other RAW formats. ``rawpy`` can. If ``rawpy`` is not installed, RAW files are
simply skipped with a one-time warning instead of crashing, so JPEG-only
workflows keep working with no extra dependencies.
"""

import os
import re
import sys

from PIL import Image, ImageOps

try:
    import rawpy  # optional: only needed for RAW/CR3 support
except Exception:  # pragma: no cover - rawpy genuinely optional
    rawpy = None


# --------------------------------------------------------------------------- #
# Supported extensions
# --------------------------------------------------------------------------- #
# Ordinary raster formats Pillow handles directly.
RASTER_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp", ".bmp"}

# Canon + other camera RAW formats. Decoded with rawpy (libraw).
RAW_EXTS = {
    ".cr3", ".cr2", ".crw",          # Canon (cr3 = BMFF/HEIF-based RAW)
    ".nef", ".nrw",                  # Nikon
    ".arw", ".srf", ".sr2",          # Sony
    ".raf",                          # Fujifilm
    ".orf",                          # Olympus
    ".rw2",                          # Panasonic
    ".pef", ".dng",                  # Pentax / generic
    ".raw", ".rwl",                  # Leica / generic
    ".srw",                          # Samsung
    ".3fr", ".fff",                  # Hasselblad
    ".iiq",                          # Phase One
    ".x3f",                          # Sigma
    ".mrw",                          # Minolta
    ".erf",                          # Epson
    ".kdc", ".dcr",                  # Kodak
}

IMG_EXTS = RASTER_EXTS | RAW_EXTS

# Which RAW ext we've already warned about a missing rawpy for.
_warned_rawpy = False


def is_raw(path):
    return os.path.splitext(path)[1].lower() in RAW_EXTS


def raw_supported():
    """True if RAW files can actually be decoded on this machine."""
    return rawpy is not None


def _warn_missing_rawpy(ext):
    global _warned_rawpy
    if not _warned_rawpy:
        _warned_rawpy = True
        print(f"  [warn] '{ext}' is a RAW format but 'rawpy' is not installed; "
              f"RAW files will be skipped. Install it with: "
              f"python -m pip install rawpy", file=sys.stderr)


def _load_raw_rgb(path):
    """Decode a RAW file with rawpy to a Pillow RGB image."""
    with rawpy.imread(path) as raw:
        # half_size=True decodes at half resolution — plenty for hashing and the
        # 640 px vision payload, and ~4x faster on big RAW files.
        rgb = raw.postprocess(half_size=True, use_camera_wb=True)
    return Image.fromarray(rgb, "RGB")


# --------------------------------------------------------------------------- #
# Public loader
# --------------------------------------------------------------------------- #
def load_image(path):
    """Open any supported image (raster or RAW) as a Pillow Image.

    Applies EXIF orientation for raster files. Returns None on failure (or when
    a RAW file is encountered but rawpy is unavailable).
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in RAW_EXTS:
        if rawpy is None:
            _warn_missing_rawpy(ext)
            return None
        try:
            return _load_raw_rgb(path)
        except Exception:
            return None
    try:
        im = Image.open(path)
        return ImageOps.exif_transpose(im)
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# EXIF capture time
# --------------------------------------------------------------------------- #
# DateTimeOriginal (36867), DateTimeDigitized (36868), DateTime (306)
_TIFF_DT_TAGS = (36867, 36868, 306)
_TIFF_EXIF_PTR = 0x8769
_TIFF_INTEROP_PTR = 0xA005


def _rawpy_exif_time(path):
    """Capture time via rawpy's EXIF parser, if this rawpy build has one.

    (Newer rawpy — e.g. 0.27 — removed ``extract_exif``; in that case we fall
    back to reading the TIFF/Exif block ourselves.)
    """
    if rawpy is None or not hasattr(rawpy.RawPy, "extract_exif"):
        return None
    try:
        with rawpy.imread(path) as raw:
            exif = raw.extract_exif()
    except Exception:
        return None
    if not exif:
        return None
    ts = exif.get("timestamp")
    if ts:
        try:
            import datetime
            # colon format to match Pillow's EXIF date strings exactly.
            return datetime.datetime.utcfromtimestamp(ts).strftime("%Y:%m:%d %H:%M:%S")
        except Exception:
            pass
    for key in ("DateTimeOriginal", "date_time_original", "DateTime"):
        if exif.get(key):
            return str(exif[key]).strip()
    return None


def _parse_tiff_tags(buf, base, want):
    """Return ``{tag: value}`` read from a TIFF/Exif structure at ``base``.

    ``buf`` may be bytes or mmap. Handles little/big endian, follows both the
    Exif sub-IFD pointer (0x8769) and the Interoperability pointer (0xA005), and
    returns text tags as stripped strings plus integer tags as ``int``.
    """
    import struct
    n = len(buf)
    if base < 0 or base + 8 > n:
        return {}
    endian = buf[base:base + 2]
    if endian == b"II":
        e = "<"
    elif endian == b"MM":
        e = ">"
    else:
        return {}
    try:
        if struct.unpack(e + "H", buf[base + 2:base + 4])[0] != 42:
            return {}
        ifd_off = struct.unpack(e + "I", buf[base + 4:base + 8])[0]
    except Exception:
        return {}

    type_size = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1, 9: 4, 10: 8}
    want = set(want)
    found = {}
    visited = set()

    def read_ifd(off):
        if off <= 0 or off in visited or base + off + 2 > n:
            return
        visited.add(off)
        try:
            cnt = struct.unpack(e + "H", buf[base + off:base + off + 2])[0]
        except Exception:
            return
        for i in range(cnt):
            ent = base + off + 2 + i * 12
            if ent + 12 > n:
                return
            tag, typ, count = struct.unpack(e + "HHI", buf[ent:ent + 8])
            size = type_size.get(typ, 1) * count
            if size <= 4:
                valpos = ent + 8
            else:
                valpos = base + struct.unpack(e + "I", buf[ent + 8:ent + 12])[0]
            if tag in want and tag not in found:
                if 0 <= valpos and valpos + size <= n:
                    if typ == 3 and count == 1:          # SHORT
                        found[tag] = struct.unpack(e + "H", buf[valpos:valpos + 2])[0]
                    elif typ == 4 and count == 1:        # LONG
                        found[tag] = struct.unpack(e + "I", buf[valpos:valpos + 4])[0]
                    else:                                # ASCII and the rest
                        raw = bytes(buf[valpos:valpos + size])
                        found[tag] = raw.rstrip(b"\x00").decode(
                            "ascii", "replace").strip()
            elif tag in (_TIFF_EXIF_PTR, _TIFF_INTEROP_PTR):
                sub = struct.unpack(e + "I", buf[ent + 8:ent + 12])[0]
                read_ifd(sub)
        nxt_pos = base + off + 2 + cnt * 12
        if nxt_pos + 4 <= n:
            try:
                nxt = struct.unpack(e + "I", buf[nxt_pos:nxt_pos + 4])[0]
            except Exception:
                nxt = 0
            if nxt:
                read_ifd(nxt)

    read_ifd(ifd_off)
    return found


def _parse_tiff_ifds(buf, base, want=_TIFF_DT_TAGS):
    """First of the ``want`` date tags present, as a string, or None."""
    found = _parse_tiff_tags(buf, base, want)
    for tag in want:
        if found.get(tag):
            return str(found[tag])
    return None


_TIFF_SIGS = (b"II*\x00", b"MM\x00*")

# First pass scans this much for a TIFF header (fast path). Canon CR3 keeps its
# metadata inside a 'CMT1' box in the moov box, which can sit well past the
# start of the file, so anything not found here triggers a full-file scan
# before we ever resort to the slow rawpy container parser.
_QUICK_SCAN = 1 << 20          # 1 MiB
_CMT1_LOOKAHEAD = 1 << 16      # 64 KiB after a CMT1 box


def _find_tiff_base(mm, size):
    """Locate the best TIFF/Exif header offset inside an mmap'd file.

    Prefers the header that follows a Canon CR3 ``CMT1`` box (that box *is* the
    Exif payload), then falls back to the earliest header found anywhere in the
    file. All searching happens in C via ``mmap.find``; no image is decoded.
    """
    # 1) Cheap scan of the head of the file.
    head = min(size, _QUICK_SCAN)
    base = -1
    for sig in _TIFF_SIGS:
        i = mm.find(sig, 0, head)
        if i != -1 and (base == -1 or i < base):
            base = i
    if base != -1:
        return base

    # 2) CR3: metadata lives in a 'CMT1' box — jump straight after it.
    cmt = mm.find(b"CMT1", 0)
    if cmt != -1:
        end = min(size, cmt + _CMT1_LOOKAHEAD)
        for sig in _TIFF_SIGS:
            i = mm.find(sig, cmt, end)
            if i != -1:
                return i

    # 3) Last resort: scan the whole file (still C-speed, never decodes).
    base = -1
    for sig in _TIFF_SIGS:
        i = mm.find(sig, 0)
        if i != -1 and (base == -1 or i < base):
            base = i
    return base


def _embedded_tiff_time(path):
    """Read date/time from the TIFF/Exif block embedded in a RAW file.

    Canon CR3 (and many other RAW formats) embed a standard TIFF/Exif structure
    in the file. We memory-map it, locate that header, and parse its IFDs — a
    few hundred microseconds per file, versus seconds for ``rawpy.imread``,
    which has to open and parse the whole RAW container.
    """
    import mmap
    try:
        size = os.path.getsize(path)
    except Exception:
        return None
    if size <= 0:
        return None
    try:
        f = open(path, "rb")
    except Exception:
        return None
    with f:
        try:
            mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        except Exception:
            return None
        with mm:
            try:
                base = _find_tiff_base(mm, size)
                if base == -1:
                    return None
                return _parse_tiff_ifds(mm, base)
            except Exception:
                return None


def _raw_capture_time(path):
    """Capture timestamp string from a RAW file, or None.

    Order matters for speed. Parsing the embedded TIFF/Exif block is a mmap
    plus a few struct unpacks (sub-millisecond). ``rawpy`` instead opens and
    parses the whole RAW container, which can take seconds per file — across a
    4,000-file shoot that is the difference between seconds and many minutes,
    and it is exactly what makes the progress counter look frozen. So: cheap
    path first, rawpy only as a fallback for files we cannot parse directly.
    """
    t = _embedded_tiff_time(path)
    if t:
        return t
    return _rawpy_exif_time(path)


def capture_time(path, fast=False):
    """Best-effort capture timestamp (DateTimeOriginal/DateTime) or None.

    Works for both ordinary raster files (via Pillow EXIF) and RAW files
    (via rawpy/libraw EXIF).

    ``fast=True`` skips the rawpy container parser. That parser can sit for
    seconds on a damaged or exotic file, and when you are indexing 4,000 files
    one such file stalls the whole run. In fast mode a RAW file only gets the
    bounded mmap/TIFF scan (microseconds), and anything unparseable simply
    yields None — it can still be matched later via perceptual hashing.
    """
    ext = os.path.splitext(path)[1].lower()

    if ext in RAW_EXTS:
        t = _embedded_tiff_time(path)
        if t is None and not fast:
            t = _rawpy_exif_time(path)
        if t:
            return t
        # Some RAW files are really TIFF underneath (e.g. DNG) — Pillow can read
        # their Exif without any RAW decoding.
        try:
            with Image.open(path) as im:
                exif = im.getexif()
                try:
                    ifd = exif.get_ifd(0x8769)
                except Exception:
                    ifd = {}
                dto = ifd.get(36867) or ifd.get(36868) or exif.get(306)
                if dto:
                    return str(dto)
        except Exception:
            pass
        return None

    try:
        with Image.open(path) as im:
            exif = im.getexif()
            # DateTimeOriginal (36867) and SubsecTimeOriginal live in the Exif IFD.
            try:
                ifd = exif.get_ifd(0x8769)  # ExifIFD tag
            except Exception:
                ifd = {}
            dto = ifd.get(36867) or ifd.get(36868)  # DateTimeOriginal / DateTimeDigitized
            if dto:
                return str(dto)
            dt = exif.get(306)  # DateTime
            if dt:
                return str(dt)
    except Exception:
        return None
    return None


def modified_time(path):
    """Best-effort 'when was this file produced' timestamp, or None.

    Editing software (Photoshop, Lightroom, Capture One) preserves
    ``DateTimeOriginal`` but rewrites ``DateTime`` (tag 306). So for a
    *derivative* the ``DateTime`` tag is the day it was edited, while for a
    camera original both stamps are identical.

    Order: EXIF ``DateTime`` (tag 306), else the filesystem mtime. RAW
    containers cannot be opened by Pillow, so tag 306 is read from the embedded
    TIFF block with the same mmap scanner capture_time uses.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in RAW_EXTS:
        return _raw_tag_value(path, 306) or _file_date(path) or None
    try:
        with Image.open(path) as im:
            dt = im.getexif().get(306)  # DateTime - rewritten by editors
            if dt:
                return str(dt)
    except Exception:
        pass
    return _file_date(path) or None


# --------------------------------------------------------------------------- #
# Deliverable metadata (file format / resolution / colour space / size / date)
# --------------------------------------------------------------------------- #
# Pillow spells some formats differently from how a photo library labels them.
_FORMAT_ALIASES = {"JPG": "JPEG", "TIF": "TIFF", "MPO": "JPEG"}

# Exif tag 40961 (ColorSpace): 1 = sRGB, 65535 = uncalibrated.
_TAG_COLOR_SPACE = 40961

_MODE_COLOR_SPACE = {
    "1": "Grayscale", "L": "Grayscale", "LA": "Grayscale", "I": "Grayscale",
    "I;16": "Grayscale", "F": "Grayscale",
    "RGB": "sRGB", "RGBA": "sRGB", "RGBX": "sRGB",
    "CMYK": "CMYK", "YCbCr": "sRGB", "LAB": "Lab",
}


def _normalise_format(fmt):
    """Normalise a Pillow/extension format name to a library-style label."""
    if not fmt:
        return ""
    f = str(fmt).upper().lstrip(".")
    return _FORMAT_ALIASES.get(f, f)


def _icc_color_space(im):
    """Colour space from the embedded ICC profile, if it names one."""
    try:
        icc = im.info.get("icc_profile") or b""
    except Exception:
        return ""
    if not icc:
        return ""
    low = icc.lower()
    if b"adobe rgb" in low or b"adobergb" in low:
        return "Adobe RGB"
    if b"prophoto" in low:
        return "ProPhoto RGB"
    if b"srgb" in low:
        return "sRGB"
    return ""


def _exif_color_space(exif):
    """Colour space from Exif tag 40961 (blank when uncalibrated/absent)."""
    try:
        v = exif.get(_TAG_COLOR_SPACE)
    except Exception:
        return ""
    if v == 1:
        return "sRGB"
    return ""


def _iso_date(value):
    """'2025:03:14 10:22:01' -> '2025-03-14' (or '' if unparseable)."""
    if not value:
        return ""
    s = str(value).strip()
    m = re.match(r"^(\d{4})[:\-/](\d{2})[:\-/](\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return ""


def _file_date(path):
    """Fallback date from the filesystem (local time), ISO formatted."""
    try:
        import datetime
        return datetime.date.fromtimestamp(os.path.getmtime(path)).isoformat()
    except Exception:
        return ""


def _file_size_mb(path):
    """File size in MiB, one decimal."""
    try:
        return round(os.path.getsize(path) / (1024 * 1024), 1)
    except Exception:
        return ""


def _raw_tag_value(path, tag):
    """Read one Exif tag from a RAW container's embedded TIFF block, or None.

    Uses the same mmap/TIFF scanner as capture_time, so it never decodes the
    image (a few hundred microseconds per file).
    """
    import mmap
    try:
        size = os.path.getsize(path)
    except Exception:
        return None
    if size <= 0:
        return None
    try:
        f = open(path, "rb")
    except Exception:
        return None
    with f:
        try:
            mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        except Exception:
            return None
        with mm:
            try:
                base = _find_tiff_base(mm, size)
                if base == -1:
                    return None
                return _parse_tiff_tags(mm, base, (tag,)).get(tag)
            except Exception:
                return None


def _raw_color_space(path):
    """Colour space of a RAW file, read from its embedded Exif block.

    RAW containers cannot be opened by Pillow, so this reads tag 40961
    (ColorSpace) directly. 65535 (uncalibrated) or a missing tag means the file
    does not state a colour space - the cell is left blank rather than claiming
    sRGB.
    """
    return "sRGB" if _raw_tag_value(path, _TAG_COLOR_SPACE) == 1 else ""


def _raw_dimensions(path):
    """Full post-processed pixel dimensions of a RAW file, or (None, None).

    Reads the RAW container header only (no full decode), so this stays cheap.
    """
    if rawpy is None:
        return None, None
    try:
        with rawpy.imread(path) as raw:
            s = raw.sizes
            w = getattr(s, "iwidth", 0) or getattr(s, "width", 0)
            h = getattr(s, "iheight", 0) or getattr(s, "height", 0)
            return (int(w) or None), (int(h) or None)
    except Exception:
        return None, None


def image_metadata(path, prefer_modified=False):
    """Deliverable metadata for one image file.

    Returns a dict with the keys used by the IndiaPix metadata template:
    ``file_format``, ``resolution``, ``color_space``, ``file_size_mb`` and
    ``date_created``. Missing/unreadable values come back as empty strings so a
    bad file never breaks a batch run.

    ``prefer_modified=True`` uses the *edit* timestamp (EXIF ``DateTime``, else
    file mtime) instead of the capture timestamp. That is what the template's
    ``date_created`` shows for a delivered derivative, while the original row
    keeps the capture date (e.g. 2025-03-14 vs 2026-09-02).
    """
    meta = {
        "file_format": "",
        "resolution": "",
        "color_space": "",
        "file_size_mb": _file_size_mb(path),
        "date_created": "",
    }
    ext = os.path.splitext(path)[1].lower()

    if ext in RAW_EXTS:
        meta["file_format"] = _normalise_format(ext)
        w, h = _raw_dimensions(path)
        if w and h:
            meta["resolution"] = f"{w}x{h}"
        meta["color_space"] = _raw_color_space(path)
        stamp = modified_time(path) if prefer_modified else capture_time(path)
        meta["date_created"] = _iso_date(stamp) or _file_date(path)
        return meta

    try:
        with Image.open(path) as im:
            meta["file_format"] = _normalise_format(im.format or ext)
            w, h = im.size
            meta["resolution"] = f"{w}x{h}"
            meta["color_space"] = (_icc_color_space(im)
                                   or _exif_color_space(im.getexif())
                                   or _MODE_COLOR_SPACE.get((im.mode or "").upper(), ""))
    except Exception:
        meta["file_format"] = _normalise_format(ext)

    if not meta["date_created"]:
        stamp = modified_time(path) if prefer_modified else capture_time(path)
        meta["date_created"] = _iso_date(stamp) or _file_date(path)
    return meta
