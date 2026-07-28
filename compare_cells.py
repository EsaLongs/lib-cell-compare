#!/usr/bin/env python3
"""Compare cell lists from two process libraries and write an Excel report.

Usage:
    python3 compare_cells.py \\
        --format-filter '"COT.*" "EEQMBD" "EEQMBC" "OPT" "A.{2}$"' \\
        --format-replace '"FOO : BAR" "X(\\d{1,2}) : D\\\\1"'

Reads NP1PP.list and C1Y.list, applies regex delete filters then regex
replacements to each full cell name in left-to-right order, re-groups by
the resulting display key, and writes Excel.

Quoted tokens are parsed with shlex. Delete filters run before replaces.
"""

from __future__ import annotations

import argparse
import os
import re
import shlex
import sys
from collections import defaultdict
from typing import Dict, List, Optional, Pattern, Sequence, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

# #### Constants
NP_FILE = "/mnt/c/Users/t00961128/Downloads/NP1PP.list"
C1_FILE = "/mnt/c/Users/t00961128/Downloads/C1Y.list"
OUTPUT = "/mnt/c/Users/t00961128/Downloads/NP1PP_vs_C1Y.xlsx"

SHEET_TITLE = "Cell Comparison"
COL_WIDTHS = {"A": 40, "B": 55, "C": 55}
FONT_SIZE = 16
MATCH_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
NORMAL_FONT = Font(size=FONT_SIZE)
BOLD_FONT = Font(size=FONT_SIZE, bold=True)
KEY_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
BORDER_SIDE = Side(style="thin", color="000000")
TOP_BOTTOM_BORDER = Border(top=BORDER_SIDE, bottom=BORDER_SIDE)
TOP_BORDER = Border(top=BORDER_SIDE)
BOTTOM_BORDER = Border(bottom=BORDER_SIDE)
REPLACE_SEP_RE = re.compile(r"\s*:\s*")

CellRow = Tuple[str, Optional[str], Optional[str]]
GroupItem = Tuple[str, bool, bool]
ReplaceRule = Tuple[Pattern[str], str]


# #### CLI
def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Compare NP1PP and C1Y cell lists after regex delete filters "
            "and optional regex replacements on display keys."
        ),
    )
    parser.add_argument(
        "--format-filter",
        required=True,
        metavar="TOKENS",
        dest="format_filter",
        help=(
            "Quote-delimited regex tokens to delete, applied left-to-right, "
            "e.g. '\"COT.*\" \"EEQMBD\" \"EEQMBC\" \"OPT\" \"A.{2}$\"'."
        ),
    )
    parser.add_argument(
        "--format-replace",
        default="",
        metavar="MAPPINGS",
        dest="format_replace",
        help=(
            "Quote-delimited regex replacements applied after deletes, "
            'e.g. \'"FOO : BAR" "X(\\d{1,2}) : D\\\\1"\'. '
            "Use ASCII ':' between pattern and replacement."
        ),
    )
    parser.add_argument(
        "--np-file",
        default=NP_FILE,
        help=f"Path to NP1PP list file (default: {NP_FILE})",
    )
    parser.add_argument(
        "--c1-file",
        default=C1_FILE,
        help=f"Path to C1Y list file (default: {C1_FILE})",
    )
    parser.add_argument(
        "--output",
        default=OUTPUT,
        help=f"Output Excel path (default: {OUTPUT})",
    )
    return parser.parse_args(argv)


def split_quoted_tokens(raw: str, option_name: str) -> List[str]:
    """Split a quote-delimited option value into tokens via shlex."""
    if not raw or not raw.strip():
        return []
    try:
        tokens = shlex.split(raw, posix=True)
    except ValueError as exc:
        print(f"Invalid {option_name} quoting: {exc}", file=sys.stderr)
        sys.exit(2)
    return [token for token in tokens if token]


def split_format_filter(raw: str) -> List[str]:
    """Split --format-filter into quote-delimited regex tokens."""
    return split_quoted_tokens(raw, "--format-filter")


def parse_replace_mapping(token: str) -> Tuple[str, str]:
    """Split one replace token into (pattern, replacement) on ASCII ':'."""
    match = REPLACE_SEP_RE.search(token)
    if match is None:
        print(
            "Invalid --format-replace token "
            f"(missing ':'): {token!r}",
            file=sys.stderr,
        )
        sys.exit(2)
    pattern = token[: match.start()].strip()
    replacement = token[match.end() :].strip()
    if not pattern:
        print(
            f"Invalid --format-replace token (empty pattern): {token!r}",
            file=sys.stderr,
        )
        sys.exit(2)
    return pattern, replacement


def parse_format_replace(raw: str) -> List[Tuple[str, str]]:
    """Parse --format-replace into ordered (pattern, replacement) pairs."""
    mappings: List[Tuple[str, str]] = []
    for token in split_quoted_tokens(raw, "--format-replace"):
        mappings.append(parse_replace_mapping(token))
    return mappings


# #### I/O
def read_list(filepath: str) -> List[str]:
    """Read a cell list file, skipping blanks and rg: header lines."""
    items: List[str] = []
    with open(filepath, encoding="utf-8") as handle:
        for line in handle:
            stripped = line.strip()
            if not stripped or stripped.startswith("rg:"):
                continue
            items.append(stripped)
    return items


# #### Name helpers
def compile_regex_filters(patterns: Sequence[str]) -> List[Pattern[str]]:
    """Compile regex filter patterns, exiting on invalid syntax."""
    compiled: List[Pattern[str]] = []
    for pattern in patterns:
        try:
            compiled.append(re.compile(pattern))
        except re.error as exc:
            print(
                f"Invalid regex in --format-filter: {pattern!r}: {exc}",
                file=sys.stderr,
            )
            sys.exit(2)
    return compiled


def compile_regex_replaces(
    mappings: Sequence[Tuple[str, str]],
) -> List[ReplaceRule]:
    """Compile regex replacement rules, exiting on invalid syntax."""
    compiled: List[ReplaceRule] = []
    for pattern, replacement in mappings:
        try:
            compiled.append((re.compile(pattern), replacement))
        except re.error as exc:
            print(
                f"Invalid regex in --format-replace: {pattern!r}: {exc}",
                file=sys.stderr,
            )
            sys.exit(2)
    return compiled


def apply_regex_filters(
    value: str,
    regex_filters: Sequence[Pattern[str]],
) -> str:
    """Apply delete regexes left-to-right; each runs until stable."""
    for pattern in regex_filters:
        while True:
            updated = pattern.sub("", value)
            if updated == value:
                break
            value = updated
    return value


def apply_regex_replaces(
    value: str,
    regex_replaces: Sequence[ReplaceRule],
) -> str:
    """Apply replace regexes left-to-right once each via re.sub."""
    for pattern, replacement in regex_replaces:
        value = pattern.sub(replacement, value)
    return value


def make_display_key(
    name: str,
    regex_filters: Sequence[Pattern[str]],
    regex_replaces: Sequence[ReplaceRule] | None = None,
) -> str:
    """Build a display key: delete filters first, then replacements."""
    value = apply_regex_filters(name, regex_filters)
    if regex_replaces:
        value = apply_regex_replaces(value, regex_replaces)
    return value


# #### Grouping
def group_cells(
    all_cells: Sequence[str],
    np_set: Set[str],
    c1_set: Set[str],
    regex_filters: Sequence[Pattern[str]],
    regex_replaces: Sequence[ReplaceRule] | None = None,
) -> Dict[str, List[GroupItem]]:
    """Group cells by display key with NP/C1 membership flags."""
    groups: Dict[str, List[GroupItem]] = defaultdict(list)
    for cell in all_cells:
        key = make_display_key(cell, regex_filters, regex_replaces)
        groups[key].append((cell, cell in np_set, cell in c1_set))
    return groups


def _partition_group(
    items: Sequence[GroupItem],
) -> Tuple[List[str], List[str], List[str]]:
    """Split a group into exact matches, NP-only, and C1-only lists."""
    exact_matches: List[str] = []
    np_only: List[str] = []
    c1_only: List[str] = []
    for cell, in_np, in_c1 in items:
        if in_np and in_c1:
            exact_matches.append(cell)
        elif in_np:
            np_only.append(cell)
        elif in_c1:
            c1_only.append(cell)
    exact_matches.sort()
    np_only.sort()
    c1_only.sort()
    return exact_matches, np_only, c1_only


def build_rows(groups: Dict[str, List[GroupItem]]) -> List[CellRow]:
    """Build Excel data rows from grouped cells.

    Only identical full cell names share a row (both NP and C1 filled).
    NP-only and C1-only cells each get their own row; they are never
    zip-paired across libraries.
    """
    rows: List[CellRow] = []
    for key in sorted(groups.keys()):
        exact_matches, np_only, c1_only = _partition_group(groups[key])
        for cell in exact_matches:
            rows.append((key, cell, cell))
        for cell in np_only:
            rows.append((key, cell, None))
        for cell in c1_only:
            rows.append((key, None, cell))
    return rows


# #### Excel writing
def _iter_display_key_groups(rows: Sequence[CellRow]) -> List[Tuple[int, int]]:
    """Return inclusive Excel row ranges for each contiguous display-key group."""
    if not rows:
        return []

    groups: List[Tuple[int, int]] = []
    group_start = 2
    current_key = rows[0][0]
    for index in range(1, len(rows)):
        if rows[index][0] != current_key:
            groups.append((group_start, index + 1))
            group_start = index + 2
            current_key = rows[index][0]
    groups.append((group_start, len(rows) + 1))
    return groups


def _merge_display_key_column(worksheet: Worksheet, rows: Sequence[CellRow]) -> None:
    """Merge contiguous identical display-key cells in column A."""
    for start_row, end_row in _iter_display_key_groups(rows):
        if end_row > start_row:
            worksheet.merge_cells(
                start_row=start_row,
                start_column=1,
                end_row=end_row,
                end_column=1,
            )


def _apply_row_border(
    worksheet: Worksheet,
    row_index: int,
    border: Border,
) -> None:
    """Apply a border style across columns A-C on one row."""
    for column in range(1, 4):
        worksheet.cell(row=row_index, column=column).border = border  # noqa


def _apply_section_borders(worksheet: Worksheet, rows: Sequence[CellRow]) -> None:
    """Add top/bottom borders for display-key groups; header keeps top only.

    Header row uses freeze panes instead of a bottom border. Group borders
    span columns A-C so each key block reads as a full-width section.
    """
    _apply_row_border(worksheet, 1, TOP_BORDER)

    for start_row, end_row in _iter_display_key_groups(rows):
        if start_row == end_row:
            _apply_row_border(worksheet, start_row, TOP_BOTTOM_BORDER)
            continue
        _apply_row_border(worksheet, start_row, TOP_BORDER)
        _apply_row_border(worksheet, end_row, BOTTOM_BORDER)


def _format_worksheet(
    worksheet: Worksheet,
    rows: Sequence[CellRow],
) -> None:
    """Apply fonts, fills, widths, key alignment, freeze, and borders."""
    row_count = len(rows)
    for column, width in COL_WIDTHS.items():
        worksheet.column_dimensions[column].width = width

    worksheet.freeze_panes = "A2"  # noqa

    for column in range(1, 4):
        worksheet.cell(row=1, column=column).font = BOLD_FONT  # noqa

    for row_index in range(2, row_count + 2):
        key_cell = worksheet.cell(row=row_index, column=1)
        key_cell.font = BOLD_FONT  # noqa
        key_cell.alignment = KEY_ALIGNMENT  # noqa

        np_cell = worksheet.cell(row=row_index, column=2)
        c1_cell = worksheet.cell(row=row_index, column=3)
        np_cell.font = NORMAL_FONT  # noqa
        c1_cell.font = NORMAL_FONT  # noqa

        np_val = np_cell.value
        c1_val = c1_cell.value
        if np_val and c1_val and np_val == c1_val:
            np_cell.fill = MATCH_FILL  # noqa
            c1_cell.fill = MATCH_FILL  # noqa

    _apply_section_borders(worksheet, rows)


def write_excel(output_path: str, rows: Sequence[CellRow]) -> None:
    """Write comparison rows to an Excel workbook."""
    if os.path.exists(output_path):
        os.remove(output_path)

    workbook = openpyxl.Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    worksheet = workbook.create_sheet(SHEET_TITLE)

    worksheet["A1"] = ""
    worksheet["B1"] = "NP1PP"
    worksheet["C1"] = "C1Y"

    for row_index, (key, np_val, c1_val) in enumerate(rows, start=2):
        worksheet.cell(row=row_index, column=1, value=key)
        if np_val:
            worksheet.cell(row=row_index, column=2, value=np_val)
        if c1_val:
            worksheet.cell(row=row_index, column=3, value=c1_val)

    _merge_display_key_column(worksheet, rows)
    _format_worksheet(worksheet, rows)
    workbook.save(output_path)


# #### Main
def main(argv: Optional[Sequence[str]] = None) -> int:
    """Run the cell comparison pipeline."""
    args = parse_args(argv)
    patterns = split_format_filter(args.format_filter)
    if not patterns:
        print("Error: --format-filter has no tokens after quote parsing.")
        return 2

    replace_mappings = parse_format_replace(args.format_replace)
    regex_filters = compile_regex_filters(patterns)
    regex_replaces = compile_regex_replaces(replace_mappings)
    print(f"Regex filters (ordered): {patterns}")
    print(f"Regex replaces (ordered): {replace_mappings}")

    np1_cells = read_list(args.np_file)
    c1y_cells = read_list(args.c1_file)

    np_set = set(np1_cells)
    c1_set = set(c1y_cells)
    all_cells = sorted(np_set | c1_set)
    print(
        f"NP1PP: {len(np1_cells)}, C1Y: {len(c1y_cells)}, "
        f"Combined unique: {len(all_cells)}"
    )

    groups = group_cells(
        all_cells,
        np_set,
        c1_set,
        regex_filters,
        regex_replaces,
    )
    rows = build_rows(groups)
    print(f"Output rows: {len(rows)}, Unique display keys: {len(groups)}")

    write_excel(args.output, rows)
    print(f"Done! Saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
