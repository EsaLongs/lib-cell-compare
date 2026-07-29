#!/usr/bin/env python3
"""Compare cell lists from two process libraries and write an Excel report.

Usage:
    python3 compare_cells.py --function-key-zh-file preview/function_key_zh.txt

--function-key-zh-file is a KEY<TAB>中文 table: column 1 is match order,
column 2 is the Chinese label in Excel column A (edit the file to change).
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict
from typing import Dict, List, Optional, Sequence, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

# #### Constants
NP_FILE = "/mnt/c/Users/t00961128/Downloads/NP1PP.list"
C1_FILE = "/mnt/c/Users/t00961128/Downloads/C1Y.list"
OUTPUT = "/mnt/c/Users/t00961128/Downloads/NP1PP_vs_C1Y.xlsx"

SHEET_TITLE = "Cell Comparison"
COL_WIDTHS = {"A": 28, "B": 55, "C": 55}
FONT_SIZE = 16
MATCH_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
NORMAL_FONT = Font(size=FONT_SIZE)
BOLD_FONT = Font(size=FONT_SIZE, bold=True)
KEY_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
BORDER_SIDE = Side(style="thin", color="000000")
TOP_BOTTOM_BORDER = Border(top=BORDER_SIDE, bottom=BORDER_SIDE)
TOP_BORDER = Border(top=BORDER_SIDE)
BOTTOM_BORDER = Border(bottom=BORDER_SIDE)

# Mid-name process/optimization markers: cut from here to end (pre-COT).
FUNCTION_CUT_TOKENS = (
    "EEQMBD",
    "EEQMBC",
    "EEQ",
    "OPT",
    "SKFMDB",
    "SKRMDB",
    "SKND",
    "SKPD",
    "SKRD",
    "SKF",
    "SKR",
    "SKN",
    "SKP",
    "TWA",
)

# Layout / track suffixes peeled only when preceded by a digit.
_TRAILING_TAG_NAMES = (
    "CNWBAL",
    "CNW",
    "CWBAL",
    "CWRB",
    "BALRB",
    "BAL",
    "RBD",
    "AROAF",
    "3FINC",
    "3FIN",
    "SHXP",
    "DHXP",
    "XP",
    "VG",
    "SH",
    "DH",
    "SF",
    "BF",
    "FF",
    "AF",
    "AR2",
    "AR1",
    "ARP",
    "ARN",
    "ARQ",
    "ARO",
    "AAR",
    "AR",
    "APA",
    "APB",
    "APC",
    "APD",
    "APM",
    "APP",
    "APN",
    "APQ",
    "ANS",
    "AKP",
    "AKN",
    "ATA",
    "ATB",
    "ATC",
    "ATS",
    "ATF",
    "ALP",
    "AHB",
    "AHA",
    "AHQ",
    "AHC",
    "ASB",
    "ONE",
    "ME",
    "ICC",
    "LN",
    "TG",
    "P5",
    "SC4LP",
    "SC5LP",
    "COMB1",
)
TRAILING_TAG_RE = re.compile(
    r"(?<=\d)(?:"
    + "|".join(
        re.escape(tag)
        for tag in sorted(_TRAILING_TAG_NAMES, key=len, reverse=True)
    )
    + r")$"
)
DRIVE_RE = re.compile(r"[DX]\d+(?:P\d+)?$")
# Whole-name roots that still end with D\d+/X\d+ and must not peel.
# Protect single-digit topologies only (ND2/NR2/...), so CKND24 peels as
# CKN+D24 (clock inverter drive) instead of a fake 24-input NAND.
PROTECTED_WHOLE_RE = re.compile(
    r"(?:"
    r"(?:CK|G)?(?:ND|NR|IND|INR|IIND|IINR|GND|GNR|GNAND|GNOR)\d|"
    r"(?:G)?AND\d|"
    r"(?:CK|G)?MUX\d|"
    r"MX\d"
    r")$"
)
# Real compound *ND2+/*NR2+ (optional extra N: MUX2NND2); not ND1 drive.
COMPOUND_ND_NR_RE = re.compile(
    r"(?:AOI|OAI|XOR|XNR|MUX|AN|OR|AO|OA|ND|NR|INR|IND)\d+"
    r"N?(?:ND|NR)(?:[2-9]\d*)$"
)
# Non-digit-prefixed variants peeled after drive strength.
VARIANT_SUFFIX_RE = re.compile(
    r"(?:CCB|CCM|CCA|SNK|SRC|CW|CWBAL|CWRB|BALRB|BAL|DBA4|NOBCM|"
    r"TGAR|XNRAR|ARSP|VPPVBB|IW|V2|COM|XP|XN|FR)$"
)
# Device/ratio variants after topology digits, e.g. AOI21B1 / AOI21N2 / ND2N1.
BN_VARIANT_RE = re.compile(r"(?<=\d)[BN]\d$")
# Level-shifter domain side after LH/HL direction.
LVL_SIDE_RE = re.compile(r"(?<=(?:LH|HL))(?:CH|CL)$")
HD_SUFFIX_MIN_STEM = 4
LAYOUT_FAMILY_PREFIXES = (
    "BOUNDARY",
    "HDDICWY",
    "HDDID",
)
COT_AND_AFTER_RE = re.compile(r"COT.*$")
FF_STEMS = (
    "Y3SDFF",
    "Y2SDFF",
    "Y3SDF",
    "Y2SDF",
    "YSDF",
    "SDFSYNC1",
    "SEDF",
    "RSDF",
    "GSDF",
    "SD2FF",
    "SDFFE",
    "SDFF",
    "SDF",
    "GDF",
    "EDF",
    "DF",
)

CellRow = Tuple[str, Optional[str], Optional[str]]
GroupItem = Tuple[str, bool, bool]


# #### CLI
def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Compare NP1PP and C1Y cell lists by KEY<TAB>中文 table "
            "(ordered prefix match; first hit wins)."
        ),
    )
    parser.add_argument(
        "--function-key-zh-file",
        required=True,
        metavar="PATH",
        dest="function_key_zh_file",
        help=(
            "KEY<TAB>中文 table: column 1 = match order, column 2 = "
            "Chinese label for Excel column A "
            "(e.g. preview/function_key_zh.txt)."
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


def load_function_key_zh_file(
    filepath: str,
) -> Tuple[List[str], Dict[str, str]]:
    """Load ordered keys and Chinese map from a KEY<TAB>中文 file."""
    keys: List[str] = []
    mapping: Dict[str, str] = {}
    with open(filepath, encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, start=1):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "\t" not in stripped:
                print(
                    f"Invalid zh map in {filepath}:{line_number}: {stripped!r} "
                    "(expected KEY<TAB>中文)",
                    file=sys.stderr,
                )
                sys.exit(2)
            key, chinese = stripped.split("\t", 1)
            key = key.strip()
            chinese = chinese.strip()
            if not key or any(char.isspace() for char in key):
                print(
                    f"Invalid key in {filepath}:{line_number}: {stripped!r} "
                    "(KEY must be a single token)",
                    file=sys.stderr,
                )
                sys.exit(2)
            if not chinese:
                print(
                    f"Invalid zh map in {filepath}:{line_number}: empty Chinese",
                    file=sys.stderr,
                )
                sys.exit(2)
            keys.append(key)
            mapping[key] = chinese
    if not keys:
        print(
            f"Error: no KEY<TAB>中文 rows in {filepath}",
            file=sys.stderr,
        )
        sys.exit(2)
    return keys, mapping


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
def _cut_process_tail(name: str) -> str:
    """Truncate at the earliest mid-name process/optimization marker."""
    cut_at = len(name)
    for token in FUNCTION_CUT_TOKENS:
        index = name.find(token)
        if index != -1 and index < cut_at:
            cut_at = index
    return name[:cut_at]


def _peel_trailing_tags(name: str) -> str:
    """Peel layout/track suffixes that appear immediately after a digit."""
    while True:
        match = TRAILING_TAG_RE.search(name)
        if match is None:
            break
        name = name[: match.start()]
    return name


def _peel_variant_suffixes(name: str) -> str:
    """Peel known non-drive variant suffixes (CCA/CCB/SNK/HD/B1/N2/...)."""
    while True:
        match = VARIANT_SUFFIX_RE.search(name)
        if match is not None:
            name = name[: match.start()]
            continue
        if name.endswith("HD") and len(name) - 2 >= HD_SUFFIX_MIN_STEM:
            name = name[:-2]
            continue
        match = BN_VARIANT_RE.search(name)
        if match is not None:
            name = name[: match.start()]
            continue
        match = LVL_SIDE_RE.search(name)
        if match is not None:
            name = name[: match.start()]
            continue
        break
    return name


def _should_stop_drive_peel(name: str) -> bool:
    """True when trailing D/X digits are part of the function, not drive."""
    if PROTECTED_WHOLE_RE.fullmatch(name):
        return True
    if COMPOUND_ND_NR_RE.search(name):
        return True
    return False


def _normalize_layout_family(name: str) -> str:
    """Collapse sized layout cells to their family prefix key."""
    for prefix in LAYOUT_FAMILY_PREFIXES:
        if name.startswith(prefix):
            return prefix
    return name


def _normalize_function_aliases(name: str) -> str:
    """Merge remaining non-functional aliases into logic roots."""
    if name.startswith("ISOC") and len(name) >= 5 and name[4] in "HL":
        name = "ISO" + name[4:]
    if name.startswith("LVUFR"):
        name = "LVU" + name[5:]
    if re.match(r"^(?:CK)?LVL.*BUFF$", name):
        name = name[:-4]
    if name == "BUFF":
        name = "BUF"
    # Cross-library spelling aliases (C1Y full / NP1PP short).
    if name.startswith("GNAND"):
        name = "GND" + name[5:]
    elif name.startswith("GAND"):
        name = "GAN" + name[4:]
    if name.startswith("GNOR"):
        name = "GNR" + name[4:]
    return name


def _collapse_ff_tail(name: str) -> str:
    """Omit Q/RPQ/SNQ/... after DF/SDF/EDF/SDFF/... flip-flop stems."""
    best_index: Optional[int] = None
    best_end: Optional[int] = None
    for stem in FF_STEMS:
        index = name.find(stem)
        if index == -1:
            continue
        end = index + len(stem)
        if (
            best_index is None
            or index < best_index
            or (index == best_index and end > best_end)
        ):
            best_index = index
            best_end = end
    if best_end is None:
        return name
    return name[:best_end]


def extract_function_root(name: str) -> str:
    """Extract a logic-function root from a full cell name."""
    value = COT_AND_AFTER_RE.sub("", name)
    value = _cut_process_tail(value)
    while True:
        value = _peel_trailing_tags(value)
        value = _peel_variant_suffixes(value)
        if _should_stop_drive_peel(value):
            break
        match = DRIVE_RE.search(value)
        if match is None:
            break
        value = value[: match.start()]
    value = _peel_trailing_tags(value)
    value = _peel_variant_suffixes(value)
    value = _collapse_ff_tail(value)
    value = _normalize_layout_family(value)
    return _normalize_function_aliases(value)


def _prefix_key_matches(
    value: str,
    key: str,
    layout_families: Set[str],
) -> bool:
    """True if key is a qualified prefix of value (digit-boundary aware)."""
    if not value.startswith(key):
        return False
    rest = value[len(key) :]
    if rest and rest[0].isdigit() and key not in layout_families:
        return False
    return True


def match_function_key(
    name: str,
    function_keys: Sequence[str],
) -> Optional[str]:
    """Return the first matching function key for name, else None.

    Walk keys in table order. At each key, accept either an exact
    extract_function_root hit or a qualified prefix of the original name or
    root (so aliases like ISOCH→ISOH still match, without root bypassing
    longer keys listed earlier).
    """
    root = extract_function_root(name)
    layout_families = set(LAYOUT_FAMILY_PREFIXES)
    for key in function_keys:
        if key == root:
            return key
        if _prefix_key_matches(name, key, layout_families):
            return key
        if root != name and _prefix_key_matches(root, key, layout_families):
            return key
    return None


def format_column_a_label(function_key: str, zh_map: Dict[str, str]) -> str:
    """Build merged-cell text: function key plus Chinese from the zh table."""
    chinese = zh_map.get(function_key, "")
    if chinese:
        return f"{function_key}\n{chinese}"
    return function_key


def make_display_key(
    name: str,
    function_keys: Sequence[str],
) -> Tuple[str, bool]:
    """Build column-A key via root / ordered prefix match.

    Returns (key, matched). Unmatched cells keep the original name so gaps
    stay visible in the spreadsheet.
    """
    matched = match_function_key(name, function_keys)
    if matched is not None:
        return matched, True
    return name, False


# #### Grouping
def group_cells(
    all_cells: Sequence[str],
    np_set: Set[str],
    c1_set: Set[str],
    function_keys: Sequence[str],
) -> Tuple[Dict[str, List[GroupItem]], List[str]]:
    """Group cells by function key; return groups and unmatched names."""
    groups: Dict[str, List[GroupItem]] = defaultdict(list)
    unmatched: List[str] = []
    for cell in all_cells:
        key, matched = make_display_key(cell, function_keys)
        if not matched:
            unmatched.append(cell)
        groups[key].append((cell, cell in np_set, cell in c1_set))
    return groups, unmatched


def unmatched_source_label(
    cell: str,
    np_set: Set[str],
    c1_set: Set[str],
) -> str:
    """Return which input list(s) an unmatched cell comes from."""
    in_np = cell in np_set
    in_c1 = cell in c1_set
    if in_np and in_c1:
        return "NP1PP+C1Y"
    if in_np:
        return "NP1PP"
    if in_c1:
        return "C1Y"
    return "?"


def print_unmatched_cells(
    unmatched: Sequence[str],
    np_set: Set[str],
    c1_set: Set[str],
) -> None:
    """Print every unmatched cell with its source list (copy-paste friendly)."""
    print(
        f"Warning: {len(unmatched)} unmatched cells "
        "(kept as column-A keys). Full list:",
        file=sys.stderr,
    )
    print("# source\tcell", file=sys.stderr)
    for cell in unmatched:
        source = unmatched_source_label(cell, np_set, c1_set)
        print(f"{source}\t{cell}", file=sys.stderr)

    np_only = [c for c in unmatched if c in np_set and c not in c1_set]
    c1_only = [c for c in unmatched if c in c1_set and c not in np_set]
    both = [c for c in unmatched if c in np_set and c in c1_set]
    print(
        f"# summary: NP1PP-only={len(np_only)} "
        f"C1Y-only={len(c1_only)} both={len(both)}",
        file=sys.stderr,
    )


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


def build_rows(
    groups: Dict[str, List[GroupItem]],
    zh_map: Dict[str, str],
) -> List[CellRow]:
    """Build Excel data rows from grouped cells.

    Only identical full cell names share a row (both NP and C1 filled).
    NP-only and C1-only cells each get their own row.
    """
    rows: List[CellRow] = []
    for key in sorted(groups.keys()):
        display = format_column_a_label(key, zh_map)
        exact_matches, np_only, c1_only = _partition_group(groups[key])
        for cell in exact_matches:
            rows.append((display, cell, cell))
        for cell in np_only:
            rows.append((display, cell, None))
        for cell in c1_only:
            rows.append((display, None, cell))
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
        worksheet.cell(row=row_index, column=column).border = border


def _apply_section_borders(worksheet: Worksheet, rows: Sequence[CellRow]) -> None:
    """Add top/bottom borders for display-key groups; header keeps top only."""
    _apply_row_border(worksheet, 1, TOP_BORDER)

    for start_row, end_row in _iter_display_key_groups(rows):
        if start_row == end_row:
            _apply_row_border(worksheet, start_row, TOP_BOTTOM_BORDER)
            continue
        _apply_row_border(worksheet, start_row, TOP_BORDER)
        _apply_row_border(worksheet, end_row, BOTTOM_BORDER)


def _format_worksheet(worksheet: Worksheet, rows: Sequence[CellRow]) -> None:
    """Apply fonts, fills, widths, key alignment, freeze, and borders."""
    row_count = len(rows)
    for column, width in COL_WIDTHS.items():
        worksheet.column_dimensions[column].width = width

    worksheet.freeze_panes = "A2"

    for column in range(1, 4):
        worksheet.cell(row=1, column=column).font = BOLD_FONT

    for row_index in range(2, row_count + 2):
        key_cell = worksheet.cell(row=row_index, column=1)
        key_cell.font = BOLD_FONT
        key_cell.alignment = KEY_ALIGNMENT

        np_cell = worksheet.cell(row=row_index, column=2)
        c1_cell = worksheet.cell(row=row_index, column=3)
        np_cell.font = NORMAL_FONT
        c1_cell.font = NORMAL_FONT

        np_val = np_cell.value
        c1_val = c1_cell.value
        if np_val and c1_val and np_val == c1_val:
            np_cell.fill = MATCH_FILL
            c1_cell.fill = MATCH_FILL

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

    np1_cells = read_list(args.np_file)
    c1y_cells = read_list(args.c1_file)
    np_set = set(np1_cells)
    c1_set = set(c1y_cells)
    all_cells = sorted(np_set | c1_set)
    print(
        f"NP1PP: {len(np1_cells)}, C1Y: {len(c1y_cells)}, "
        f"Combined unique: {len(all_cells)}"
    )

    function_keys, zh_map = load_function_key_zh_file(args.function_key_zh_file)
    print(f"Key/Chinese table: {args.function_key_zh_file}")
    print(f"Function keys: {len(function_keys)}")

    groups, unmatched = group_cells(
        all_cells,
        np_set,
        c1_set,
        function_keys,
    )
    rows = build_rows(groups, zh_map)
    print(
        f"Output rows: {len(rows)}, Unique function keys used: {len(groups)}, "
        f"Unmatched cells: {len(unmatched)}"
    )
    if unmatched:
        print_unmatched_cells(unmatched, np_set, c1_set)

    write_excel(args.output, rows)
    print(f"Done! Saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
