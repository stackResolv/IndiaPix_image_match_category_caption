#!/usr/bin/env python3
"""
consoleutil.py — make the pipeline's narration readable on *every* console,
including a default Windows PowerShell / cmd.exe window.

Two problems this solves:

  1. **Narrow windows.** A fresh Windows console is 80 columns wide, and
     PowerShell in VS Code is often no wider. Lines built for a 120-column
     macOS terminal wrap mid-sentence and become unreadable. We measure the
     real width and clamp every line to it.

  2. **Encoding.** Windows may still be on cp1252, where characters like the
     em dash (—) raise UnicodeEncodeError or print as garbage. We reconfigure
     the stream to UTF-8 (the launchers also set chcp 65001) and degrade any
     remaining unencodable character to a safe ASCII equivalent instead of
     crashing.

Usage from the pipeline scripts:

    from consoleutil import setup_console, term_width, hr, wrap, safe

    setup_console()                 # call once, at import time
    print(hr("="))                  # a full-width rule
    for line in wrap(long_text):    # line-broken to fit the window
        print(line)
"""

import os
import shutil
import sys
import textwrap

# Populated by setup_console().
_WIDTH = None
_UTF8_OK = True

# Widths narrower than this are not worth honouring (some CI logs report 0).
_MIN_WIDTH = 60
# Hard ceiling so ultra-wide terminals don't produce unreadable 300-char lines.
_MAX_WIDTH = 110


def setup_console(force_width=None):
    """Prepare stdout/stderr and remember the usable terminal width.

    Safe to call more than once. Never raises: a console that cannot be
    reconfigured simply keeps its defaults.
    """
    global _WIDTH, _UTF8_OK

    for stream in (sys.stdout, sys.stderr):
        try:
            # errors="replace" means an unencodable filename prints as '?' on a
            # cp1252 console instead of aborting the run.
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    # Some Windows setups accept the reconfigure but still cannot encode.
    enc = (getattr(sys.stdout, "encoding", "") or "").lower().replace("-", "")
    _UTF8_OK = enc in ("utf8", "utf8sig", "cp65001") or enc == ""

    _WIDTH = force_width or _detect_width()
    return _WIDTH


def _detect_width():
    """Best-effort usable column count, clamped to a sane range."""
    cols = 0

    # 1) Explicit override (handy for tests and for piping into a log viewer).
    env = os.environ.get("PIPELINE_WIDTH")
    if env:
        try:
            cols = int(env)
        except ValueError:
            cols = 0

    # 2) The real terminal size.
    if not cols:
        try:
            cols = shutil.get_terminal_size(fallback=(0, 0)).columns
        except Exception:
            cols = 0

    # 3) Windows console API — shutil can report 80 even when the window is wider.
    if not cols and os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.windll.kernel32
            handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
            buf = ctypes.create_string_buffer(22)
            if kernel32.GetConsoleScreenBufferInfo(handle, buf):
                # COORD at offset 0 (dwSize), then SMALL_RECT dwWindowRight at 10.
                right = wintypes.SHORT.from_buffer(buf, 10).value
                cols = right + 1
        except Exception:
            cols = 0

    # 4) Anything at all, else the classic 80.
    if not cols:
        cols = int(os.environ.get("COLUMNS", 0) or 0)
    if not cols:
        cols = 80

    return max(_MIN_WIDTH, min(_MAX_WIDTH, cols))


def term_width():
    """Usable width in columns (already clamped)."""
    return _WIDTH or _detect_width()


def safe(text):
    """Return `text` with characters this console cannot encode replaced.

    On UTF-8 consoles the text is returned untouched. On a legacy code page the
    em dash and friends degrade to their ASCII look-alikes, which keeps the
    narration readable instead of raising UnicodeEncodeError.
    """
    if _UTF8_OK:
        return text
    replacements = {
        "\u2014": "-",   # em dash
        "\u2013": "-",   # en dash
        "\u2192": "->",  # right arrow
        "\u2026": "...", # ellipsis
        "\u00b7": "-",   # middle dot
        "\u2022": "*",   # bullet
        "\u2265": ">=",
        "\u2264": "<=",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    enc = (getattr(sys.stdout, "encoding", "") or "ascii")
    try:
        return text.encode(enc, "replace").decode(enc, "replace")
    except Exception:
        return text.encode("ascii", "replace").decode("ascii")


def hr(char="=", width=None, indent=0):
    """A horizontal rule that fills the window, allowing for `indent` spaces.

    The indent is subtracted so the rule never wraps onto a second line, which
    is the usual give-away that a layout was designed for a wider terminal.
    """
    w = (width or term_width()) - indent - 1
    return char * max(10, w)


def wrap(text, indent=0, width=None, subsequent_indent=None):
    """Break `text` so no line exceeds the terminal width.

    Returns a list of lines (the caller prints them). Long single tokens such as
    file paths are broken rather than allowed to overflow.
    """
    w = width or term_width()
    body = w - indent
    if body < 20:                      # absurdly narrow: don't mangle the text
        return [text]
    return textwrap.wrap(
        text,
        width=body,
        subsequent_indent=subsequent_indent if subsequent_indent is not None
        else "  ",
        break_long_words=True,
        break_on_hyphens=False,
    ) or [text]


def say(text, indent=0, width=None, subsequent_indent=None):
    """Print `text` wrapped to the terminal width (and encoding-safe)."""
    for line in wrap(safe(text), indent=indent, width=width,
                    subsequent_indent=subsequent_indent):
        print(" " * indent + line, flush=True)
