#!/usr/bin/env python3
"""Compare cell lists from two process libraries and write an Excel report.

Usage:
    python3 compare_cells.py \\
        --format-filter '"COT.*" "EEQMBD" "EEQMBC" "OPT" "A.{2}$"'

Reads NP1PP.list and C1Y.list, strips regex filters from each full cell
name in left-to-right order, re-groups by the resulting display key, and
writes Excel.

--format-filter takes one string of quote-delimited tokens (shell-style).
Each token is applied as a regex via re.sub, in the given order.
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
from openpyxl.styles import Alignment, Font, PatternFill
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

CellRow = Tuple[str, Optional[str], Optional[str]]
GroupItem = Tuple[str, bool, bool]


# #### CLI
def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Compare NP1PP and C1Y cell lists after stripping regex "
            "filters from display keys in left-to-right order."
        ),
    )
    parser.add_argument(
        "--format-filter",
        required=True,
        metavar="TOKENS",
        dest="format_filter",
        help=(
            "Quote-delimited regex tokens, applied left-to-right, "
            "e.g. '\"COT.*\" \"EEQMBD\" \"EEQMBC\" \"OPT\" \"A.{2}$\"'."
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


def split_format_filter(raw: str) -> List[str]:
    """Split --format-filter into quote-delimited regex tokens via shlex."""
    try:
        tokens = shlex.split(raw, posix=True)
    except ValueError as exc:
        print(f"Invalid --format-filter quoting: {exc}", file=sys.stderr)
        sys.exit(2)
    return [token for token in tokens if token]


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


def apply_regex_filters(
    value: str,
    regex_filters: Sequence[Pattern[str]],
) -> str:
    """Apply regexes left-to-right; each runs until the value stabilizes."""
    for pattern in regex_filters:
        while True:
            updated = pattern.sub("", value)
            if updated == value:
                break
            value = updated
    return value


def make_display_key(
    name: str,
    regex_filters: Sequence[Pattern[str]],
) -> str:
    """Build a display key by stripping regex filters from the full name."""
    return apply_regex_filters(name, regex_filters)


# #### Grouping
def group_cells(
    all_cells: Sequence[str],
    np_set: Set[str],
    c1_set: Set[str],
    regex_filters: Sequence[Pattern[str]],
) -> Dict[str, List[GroupItem]]:
    """Group cells by display key with NP/C1 membership flags."""
    groups: Dict[str, List[GroupItem]] = defaultdict(list)
    for cell in all_cells:
        key = make_display_key(cell, regex_filters)
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

    Exact matches occupy both NP and C1 columns on the same row.
    Remaining NP-only / C1-only cells are paired by index.
    """
    rows: List[CellRow] = []
    for key in sorted(groups.keys()):
        exact_matches, np_only, c1_only = _partition_group(groups[key])
        max_len = max(len(exact_matches), len(np_only), len(c1_only), 0)
        for index in range(max_len):
            if index < len(exact_matches):
                np_val: Optional[str] = exact_matches[index]
                c1_val: Optional[str] = exact_matches[index]
            else:
                np_val = np_only[index] if index < len(np_only) else None
                c1_val = c1_only[index] if index < len(c1_only) else None
            rows.append((key, np_val, c1_val))
    return rows


# #### Excel writing
def _merge_display_key_column(worksheet: Worksheet, rows: Sequence[CellRow]) -> None:
    """Merge contiguous identical display-key cells in column A."""
    if not rows:
        return

    current_val = rows[0][0]
    merge_start = 2
    for index in range(1, len(rows)):
        value = rows[index][0]
        row_num = index + 2
        if value != current_val:
            if row_num - merge_start > 1:
                worksheet.merge_cells(
                    start_row=merge_start,
                    start_column=1,
                    end_row=row_num - 1,
                    end_column=1,
                )
            merge_start = row_num
            current_val = value

    last_data_row = len(rows) + 1
    if last_data_row - merge_start + 1 > 1:
        worksheet.merge_cells(
            start_row=merge_start,
            start_column=1,
            end_row=last_data_row,
            end_column=1,
        )


def _format_worksheet(worksheet: Worksheet, row_count: int) -> None:
    """Apply fonts, fills, widths, and key-column alignment."""
    for column, width in COL_WIDTHS.items():
        worksheet.column_dimensions[column].width = width

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
    _format_worksheet(worksheet, len(rows))
    workbook.save(output_path)


# #### Main
def main(argv: Optional[Sequence[str]] = None) -> int:
    """Run the cell comparison pipeline."""
    args = parse_args(argv)
    patterns = split_format_filter(args.format_filter)
    if not patterns:
        print("Error: --format-filter has no tokens after quote parsing.")
        return 2

    regex_filters = compile_regex_filters(patterns)
    print(f"Regex filters (ordered): {patterns}")

    np1_cells = read_list(args.np_file)
    c1y_cells = read_list(args.c1_file)

    np_set = set(np1_cells)
    c1_set = set(c1y_cells)
    all_cells = sorted(np_set | c1_set)
    print(
        f"NP1PP: {len(np1_cells)}, C1Y: {len(c1y_cells)}, "
        f"Combined unique: {len(all_cells)}"
    )

    groups = group_cells(all_cells, np_set, c1_set, regex_filters)
    rows = build_rows(groups)
    print(f"Output rows: {len(rows)}, Unique display keys: {len(groups)}")

    write_excel(args.output, rows)
    print(f"Done! Saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
