#!/usr/bin/env python3
"""Compare cell lists from two process libraries and write an Excel report.

Usage:
    python3 compare_cells.py --format-filter TOKEN1 TOKEN2 ...
    python3 compare_cells.py --format-filter EEQMBD EEQMBC OPT \\
        --format-filter-regex 'A.{2}$'

Reads NP1PP.list and C1Y.list, groups cells by base name (COT prefix),
strips given tokens / regexes from the display key, re-groups, and writes
Excel.

Exact tokens use string replacement. Regex filters use re.sub.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict
from typing import Dict, List, Optional, Pattern, Sequence, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet

# #### Constants
NP_FILE = "/mnt/c/Users/t00961128/Downloads/NP1PP.list"
C1_FILE = "/mnt/c/Users/t00961128/Downloads/C1Y.list"
OUTPUT = "/mnt/c/Users/t00961128/Downloads/NP1PP_vs_C1Y.xlsx"

COT_PREFIX_RE = re.compile(r"^(.*?)COT")
SHEET_TITLE = "Cell Comparison"
COL_WIDTHS = {"A": 40, "B": 55, "C": 55}

CellRow = Tuple[str, Optional[str], Optional[str]]
GroupItem = Tuple[str, bool, bool]


# #### CLI
def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Compare NP1PP and C1Y cell lists after stripping tokens "
            "and/or regexes from display keys."
        ),
    )
    parser.add_argument(
        "--format-filter",
        nargs="+",
        default=[],
        metavar="TOKEN",
        dest="format_filter",
        help=(
            "Tokens to strip from display keys (exact string match), "
            "e.g. EEQMBD EEQMBC OPT."
        ),
    )
    parser.add_argument(
        "--format-filter-regex",
        nargs="+",
        default=[],
        metavar="PATTERN",
        dest="format_filter_regex",
        help=(
            "Regex patterns to strip from display keys via re.sub, "
            "e.g. 'A.{2}$' for a trailing A?? suffix."
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
def get_base(name: str) -> str:
    """Return the prefix before the first COT occurrence, else the name."""
    match = COT_PREFIX_RE.match(name)
    return match.group(1) if match else name


def sort_tokens(tokens: Sequence[str]) -> List[str]:
    """Sort tokens by length descending to avoid substring collisions."""
    return sorted(tokens, key=len, reverse=True)


def compile_regex_filters(patterns: Sequence[str]) -> List[Pattern[str]]:
    """Compile regex filter patterns, exiting on invalid syntax."""
    compiled: List[Pattern[str]] = []
    for pattern in patterns:
        try:
            compiled.append(re.compile(pattern))
        except re.error as exc:
            print(
                f"Invalid regex in --format-filter-regex: {pattern!r}: {exc}",
                file=sys.stderr,
            )
            sys.exit(2)
    return compiled


def _apply_regex_filters(
    value: str,
    regex_filters: Sequence[Pattern[str]],
) -> str:
    """Apply each regex repeatedly until the value stabilizes."""
    for pattern in regex_filters:
        while True:
            updated = pattern.sub("", value)
            if updated == value:
                break
            value = updated
    return value


def make_display_key(
    name: str,
    tokens: Sequence[str],
    regex_filters: Sequence[Pattern[str]] | None = None,
) -> str:
    """Build a display key by stripping exact tokens then regexes."""
    base = get_base(name)
    for token in tokens:
        base = base.replace(token, "")
    if regex_filters:
        base = _apply_regex_filters(base, regex_filters)
    return base


# #### Grouping
def group_cells(
    all_cells: Sequence[str],
    np_set: Set[str],
    c1_set: Set[str],
    tokens: Sequence[str],
    regex_filters: Sequence[Pattern[str]] | None = None,
) -> Dict[str, List[GroupItem]]:
    """Group cells by display key with NP/C1 membership flags."""
    groups: Dict[str, List[GroupItem]] = defaultdict(list)
    for cell in all_cells:
        key = make_display_key(cell, tokens, regex_filters)
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
    """Apply column widths and center alignment on the display-key column."""
    for column, width in COL_WIDTHS.items():
        worksheet.column_dimensions[column].width = width

    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index in range(2, row_count + 2):
        worksheet.cell(row=row_index, column=1).alignment = alignment


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
    tokens = sort_tokens(args.format_filter)
    regex_filters = compile_regex_filters(args.format_filter_regex)
    print(f"Tokens to strip: {tokens}")
    print(f"Regex filters: {args.format_filter_regex}")

    np1_cells = read_list(args.np_file)
    c1y_cells = read_list(args.c1_file)

    np_set = set(np1_cells)
    c1_set = set(c1y_cells)
    all_cells = sorted(np_set | c1_set)
    print(
        f"NP1PP: {len(np1_cells)}, C1Y: {len(c1y_cells)}, "
        f"Combined unique: {len(all_cells)}"
    )

    groups = group_cells(all_cells, np_set, c1_set, tokens, regex_filters)
    rows = build_rows(groups)
    print(f"Output rows: {len(rows)}, Unique display keys: {len(groups)}")

    write_excel(args.output, rows)
    print(f"Done! Saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
