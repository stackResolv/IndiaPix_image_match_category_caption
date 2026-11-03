#!/usr/bin/env python3
"""
excelutil.py — streaming Excel writer for very large metadata sheets.

The shoot workflow fits comfortably in a normal openpyxl workbook, but the
set-folder workflow can produce 30,000-40,000 sets, i.e. 60,000-80,000 rows
(one `00 original` row plus one `final` row per set). A regular Workbook keeps
every cell in memory — roughly a gigabyte at that size — and the run would die
just before writing anything.

openpyxl's **write-only mode** streams rows straight into the worksheet XML, so
memory stays flat regardless of row count. This module wraps it so the produced
file is indistinguishable from the one categorize.py writes:

  * the same header styling (white bold on dark blue),
  * the same per-category colour coding on the `category` and `file_id` cells,
  * the same grey shading on `original` rows,
  * the same column widths, frozen header and auto-filter.

Usage:

    from excelutil import StreamedSheet, ALL_COLUMNS

    sheet = StreamedSheet(path, ALL_COLUMNS, expected_rows=len(sets) * 2)
    for row in rows:            # row: dict keyed by column name
        sheet.add(row)
    sheet.close()               # writes the file
"""

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Template column order (17 IndiaPix columns) + the classifier's trailing
# `category`. Kept here so every writer in the project agrees.
TEMPLATE_COLUMNS = [
    "file_id", "set_id", "edit_state", "role", "edit_operation",
    "edit_operation_note", "title", "description", "keywords", "file_format",
    "resolution", "color_space", "file_size_mb", "date_created",
    "date_contributed", "asset_category", "ai_used",
]
# `category` is the classifier's own 4 subject types (Portrait / Traditional /
# White BG / Others), appended after the 17 template columns for review.
EXTRA_COLUMNS = ["category"]
ALL_COLUMNS = TEMPLATE_COLUMNS + EXTRA_COLUMNS

# One tint per category, shared by the `category` cell and the `file_id` cell so
# a set is scannable from the first column.
CAT_FILLS = {
    "Portrait": "D9E1F2",
    "Traditional/Festival": "FCE4D6",
    "White Background": "E2EFDA",
    "Others": "EDEDED",
    "ERROR": "FFC7CE",
}

# `original` rows are shaded grey; `final` rows keep the default fill.
ROLE_FILLS = {"original": "F2F2F2", "final": ""}

COLUMN_WIDTHS = {
    "file_id": 26, "set_id": 12, "edit_state": 11, "role": 10,
    "edit_operation": 24, "edit_operation_note": 18,
    "title": 40, "description": 62, "keywords": 50,
    "file_format": 12, "resolution": 12, "color_space": 13,
    "file_size_mb": 13, "date_created": 14, "date_contributed": 16,
    "asset_category": 22, "ai_used": 9, "category": 24,
}

_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill("solid", fgColor="305496")
_HEADER_ALIGN = Alignment(horizontal="center")

# Reuse one fill object per colour instead of allocating one per cell — at
# 80,000 rows x 2 tinted cells that is the difference between trivial and slow.
_FILL_CACHE = {}


def _fill(color):
    """Cached PatternFill for a hex colour ('' -> None)."""
    if not color:
        return None
    f = _FILL_CACHE.get(color)
    if f is None:
        f = PatternFill("solid", fgColor=color)
        _FILL_CACHE[color] = f
    return f


class StreamedSheet:
    """Write an arbitrarily large sheet with flat memory (openpyxl write-only).

    ``expected_rows`` is a pre-flight estimate only: the auto-filter range is
    finalised in ``close()`` from the rows actually written, so a run that ends
    up with more (or fewer) rows than planned still leaves the client a filter
    that covers every row.
    """

    def __init__(self, path, columns=None, expected_rows=0, title="Photos"):
        self.path = path
        self.columns = list(columns or ALL_COLUMNS)
        self.expected_rows = max(0, int(expected_rows or 0))
        self.rows_written = 0

        self.wb = Workbook(write_only=True)
        self.ws = self.wb.create_sheet(title=title)

        # Column widths, frozen panes and the auto-filter must all be set BEFORE
        # the first append: write-only mode emits the sheet header at that point.
        for i, col in enumerate(self.columns, 1):
            self.ws.column_dimensions[get_column_letter(i)].width = \
                COLUMN_WIDTHS.get(col, 14)
        self.ws.freeze_panes = "B2"
        if self.expected_rows:
            last = get_column_letter(len(self.columns))
            self.ws.auto_filter.ref = f"A1:{last}{self.expected_rows + 1}"

        self._cat_col = self._index_of("category")
        self._file_col = self._index_of("file_id")
        self._role_col = self._index_of("role")
        self._state_col = self._index_of("edit_state")

        # In write-only mode `append` returns None: styles must be attached to
        # WriteOnlyCell objects BEFORE the row is appended. (A normal Worksheet
        # keeps its cells, write-only streams them straight to disk.)
        header = []
        for name in self.columns:
            cell = WriteOnlyCell(self.ws, value=name)
            cell.font = _HEADER_FONT
            cell.fill = _HEADER_FILL
            cell.alignment = _HEADER_ALIGN
            header.append(cell)
        self.ws.append(header)

    def _index_of(self, name):
        """1-based column index of `name`, or 0 when the column is absent."""
        try:
            return self.columns.index(name) + 1
        except ValueError:
            return 0

    def add(self, row):
        """Append one data row (a dict keyed by column name)."""
        values = [row.get(col, "") for col in self.columns]
        cells = values
        self.rows_written += 1

        # Only wrap the handful of cells that need a fill: a WriteOnlyCell per
        # cell would be far slower across 80,000 rows.
        fills = {}
        cat_fill = _fill(CAT_FILLS.get(row.get("category", "")))
        if cat_fill is not None:
            for idx in (self._cat_col, self._file_col):
                if idx:
                    fills[idx - 1] = cat_fill
        role_fill = _fill(ROLE_FILLS.get(row.get("role", "")))
        if role_fill is not None:
            for idx in (self._role_col, self._state_col):
                if idx:
                    fills[idx - 1] = role_fill

        if fills:
            cells = list(values)
            for idx, fill in fills.items():
                cell = WriteOnlyCell(self.ws, value=cells[idx])
                cell.fill = fill
                cells[idx] = cell

        self.ws.append(cells)

    def close(self):
        """Save the workbook, turning the usual write failures into clear errors."""
        # write-only mode emits the sheet header at the first append, but the
        # auto-filter can still be re-stated here, and this is the only place
        # that knows how many rows were really written.
        if self.rows_written or self.expected_rows:
            last = get_column_letter(len(self.columns))
            self.ws.auto_filter.ref = f"A1:{last}{self.rows_written + 1}"
        try:
            self.wb.save(self.path)
        except PermissionError:
            raise SystemExit(
                f"error: cannot write '{self.path}' - it is open in Excel. "
                f"Close the file (or the Excel window) and run again.")
        except OSError as e:
            raise SystemExit(f"error: cannot write '{self.path}': {e}")
        return self.path