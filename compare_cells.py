#!/usr/bin/env python3
"""Compare cell lists from two process libraries and write an Excel report.

Usage:
    python3 compare_cells.py --function-key-zh-file preview/function_key_zh.txt

--function-key-zh-file is a KEY<TAB>中文 table: column 1 is match order,
column 2 is the Chinese label shown in Excel column A (and with the KEY
in column B).
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
COL_WIDTHS = {"A": 16, "B": 28, "C": 55, "D": 55}
FONT_SIZE = 16
MATCH_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
NORMAL_FONT = Font(size=FONT_SIZE)
BOLD_FONT = Font(size=FONT_SIZE, bold=True)
KEY_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
BORDER_SIDE = Side(style="thin", color="000000")
TOP_BOTTOM_BORDER = Border(top=BORDER_SIDE, bottom=BORDER_SIDE)
TOP_BORDER = Border(top=BORDER_SIDE)
BOTTOM_BORDER = Border(bottom=BORDER_SIDE)
LAST_DATA_COLUMN = 4

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

CellRow = Tuple[str, str, Optional[str], Optional[str]]
GroupItem = Tuple[str, bool, bool]

# Column-B family merges within a coarse Chinese (A) group.
# Longer / more-specific patterns must come first.
# fine_zh is emitted only when this family is a proper subset of its A group.
KEY_FAMILY_RULES: Tuple[Tuple[Pattern[str], str, str], ...] = (
    (re.compile(r"^GDF$"), "GDF", "门阵列ECO"),
    (re.compile(r"^DF$"), "DF", "标准单元"),
    (re.compile(r"^GPULL\d+$"), "GPULL", ""),
    (re.compile(r"^GAOI\d+$"), "GAOI", "门阵列ECO"),
    (re.compile(r"^AOI\d+$"), "AOI", "标准单元"),
    (re.compile(r"^GOAI\d+$"), "GOAI", "门阵列ECO"),
    (re.compile(r"^OAI\d+$"), "OAI", "标准单元"),
    (re.compile(r"^GAO\d+$"), "GAO", "门阵列ECO"),
    (re.compile(r"^AO\d+$"), "AO", "标准单元"),
    (re.compile(r"^GOA\d+$"), "GOA", "门阵列ECO"),
    (re.compile(r"^OA\d+$"), "OA", "标准单元"),
    (re.compile(r"^GAN\d+$"), "GAN", "门阵列ECO"),
    (re.compile(r"^AN\d+$"), "AN", "标准单元"),
    (re.compile(r"^IIND\d+$"), "IIND", "双输入反相与非"),
    (re.compile(r"^IND\d+$"), "IND", "单输入反相与非"),
    (re.compile(r"^GND\d+$"), "GND", "门阵列ECO"),
    (re.compile(r"^ND\d+$"), "ND", "标准与非"),
    (re.compile(r"^IINR\d+$"), "IINR", "双输入反相或非"),
    (re.compile(r"^INR\d+$"), "INR", "单输入反相或非"),
    (re.compile(r"^GNR\d+$"), "GNR", "门阵列ECO"),
    (re.compile(r"^NR\d+$"), "NR", "标准或非"),
    (re.compile(r"^GOR\d+$"), "GOR", "门阵列ECO"),
    (re.compile(r"^OR\d+$"), "OR", "标准单元"),
    (re.compile(r"^GXOR\d+$"), "GXOR", "门阵列ECO"),
    (re.compile(r"^XOR\d+$"), "XOR", "标准单元"),
    (re.compile(r"^GXNR\d+$"), "GXNR", "门阵列ECO"),
    (re.compile(r"^XNR\d+$"), "XNR", "标准单元"),
    (re.compile(r"^GFILL\d+$"), "GFILL", "门阵列ECO"),
    (re.compile(r"^FILL\d+$"), "FILL", "标准单元"),
    (re.compile(r"^GDCAP\d+$"), "GDCAP", "门阵列ECO"),
    (re.compile(r"^DCAP\d+$"), "DCAP", "标准单元"),
    (re.compile(r"^GBUFF$"), "GBUFF", "门阵列ECO"),
    (re.compile(r"^APBUF$"), "APBUF", "AP系列"),
    (re.compile(r"^BUF$"), "BUF", "标准单元"),
    (re.compile(r"^GINV$"), "GINV", "门阵列ECO"),
    (re.compile(r"^APINV$"), "APINV", "AP系列"),
    (re.compile(r"^INV$"), "INV", "标准单元"),
    (re.compile(r"^APTIE[HL]$"), "APTIE", "AP系列"),
    (re.compile(r"^GTIE[HL]$"), "GTIE", "门阵列ECO"),
    (re.compile(r"^TIE[HL]$"), "TIE", "标准单元"),
    (re.compile(r"^GMUX2N$"), "GMUX2N", "门阵列ECO反相"),
    (re.compile(r"^GMUX\d+$"), "GMUX", "门阵列ECO"),
    (re.compile(r"^MUX\d+N$"), "MUXN", "标准反相"),
    (re.compile(r"^MUX\d+I$"), "MUXI", "标准反相"),
    (re.compile(r"^MXI\d+$"), "MXI", "标准反相"),
    (re.compile(r"^MUX\d+$"), "MUX", "标准单元"),
    (re.compile(r"^MX\d+$"), "MX", "标准单元"),
    (re.compile(r"^MUX2"), "MUX2CMP", "复合MUX2"),
    (re.compile(r"^ISO[HL]$"), "ISO", ""),
    (re.compile(r"^DEL[A-Z]$"), "DEL", ""),
    (re.compile(r"^FA"), "FA", ""),
    (re.compile(r"^HA(?:C|N)?1$"), "HA", ""),
    (re.compile(r"^BENC"), "BENC", ""),
    (re.compile(r"^CMPE\d+$"), "CMPE", ""),
    (re.compile(r"^SYNLH"), "SYNLH", ""),
    (re.compile(r"^SDFSYNC"), "SDFSYNC", ""),
    (re.compile(r"^MB\d+SRLSDF$"), "MBSDF", ""),
    # MB* = multi-bit D latch (VeriSilicon MBLAT); MCE* = clock-enable latch.
    (re.compile(r"^MB\d+"), "MBLAT", "多比特D锁存"),
    (re.compile(r"^MCE"), "MCE", "时钟使能锁存"),
    # Level shifters BEFORE LH/LN: names like LVLLH/LVLHL contain LH/HL.
    (re.compile(r"^CKLVL"), "CKLVL", "时钟电平转换"),
    (re.compile(r"^CKLVU"), "CKLVU", "时钟电平转换"),
    (re.compile(r"^LVU"), "LVU", "LVU电平转换"),
    (re.compile(r"^LVLHL"), "LVLHL", "高转低"),
    (re.compile(r"^LVLLH"), "LVLLH", "低转高"),
    (re.compile(r"^LVL"), "LVL", "LVL电平转换"),
    # Latches (incl. clock latches): only LH vs LN; LAH counted as high.
    (re.compile(r"LN"), "LN", "低电平透明"),
    (re.compile(r"LH|LAH"), "LH", "高电平透明"),
    # Scan FF: only Y-series vs non-Y.
    (re.compile(r"^Y"), "YSCAN", "Y系列"),
    (
        re.compile(r"^(?:G|R)?SDF$|^SDFF$|^SD2FF$|^SEDF$"),
        "SCAN",
        "非Y扫描",
    ),
    (re.compile(r"^DCCKB$"), "DCCKB", "DC时钟缓冲"),
    (re.compile(r"^DCCKN$"), "DCCKN", "DC时钟反相"),
    (re.compile(r"^CKB$"), "CKB", "时钟缓冲"),
    (re.compile(r"^CKN$"), "CKN", "时钟反相"),
    (re.compile(r"^CKAN"), "CKAN", ""),
    (re.compile(r"^CKND"), "CKND", ""),
    (re.compile(r"^CKOR"), "CKOR", ""),
    (re.compile(r"^CKNR"), "CKNR", ""),
    (re.compile(r"^CKXOR"), "CKXOR", ""),
    (re.compile(r"^CKMUX"), "CKMUX", ""),
    (re.compile(r"^GCKNQ"), "GCKNQ", ""),
    (re.compile(r"^BOUNDARY$"), "BOUNDARY", "边界单元"),
    (re.compile(r"^HDDICWY$"), "HDDICWY", "版图单元"),
    (re.compile(r"^HDDID$"), "HDDID", "版图单元"),
    # Special datapath FA (TCBN): FC* inverting-carry / FIICON carry-out select.
    (re.compile(r"^FC"), "FC", "反相进位全加"),
    (re.compile(r"^FI"), "FI", "进位选择全加"),
)


def key_family(function_key: str) -> Tuple[str, str]:
    """Return (family_id, fine_zh) for column-B merging."""
    for pattern, family_id, fine_zh in KEY_FAMILY_RULES:
        if pattern.search(function_key):
            return family_id, fine_zh
    # Compound / leftover: ignore 21/22-style digits, keep boolean skeleton.
    skeleton = re.sub(r"\d+", "", function_key)
    if skeleton and skeleton != function_key:
        return skeleton, f"复合·{skeleton}"
    return function_key, ""


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


def _normalize_fill_dcap_size(name: str) -> str:
    """Map GFILLD12 / GDCAPD10 / GFILLD2D → GFILL12 / GDCAP10 / GFILL2.

    C1Y sizes the cell as D<n> after FILL/DCAP; NP1PP uses FILL<n> directly.
    Optional trailing D (e.g. GFILLD2D) is a density variant, not a new function.
    """
    match = re.fullmatch(r"((?:G)?(?:FILL|DCAP))D(\d+)D?", name)
    if match is None:
        return name
    return f"{match.group(1)}{match.group(2)}"


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
    if name == "GBUF":
        name = "GBUFF"
    # GIND1/2/4/8 are G-inverter + drive (GIN+D), not multi-input NAND.
    # NP1PP names the same family GINV.
    if re.fullmatch(r"GIND\d+(?:P\d+)?", name):
        name = "GINV"
    # Cross-library spelling aliases (C1Y full / NP1PP short).
    if name.startswith("GNAND"):
        name = "GND" + name[5:]
    elif name.startswith("GAND"):
        name = "GAN" + name[4:]
    if name.startswith("GNOR"):
        name = "GNR" + name[4:]
    if name.startswith("GXNOR"):
        name = "GXNR" + name[5:]
    # Full-adder polarity / scan-output variants → one FA1 family.
    if name in {"FA1N", "FA1SN"}:
        name = "FA1"
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
        # After tags (SH/DH/...), map GFILLD12 / GDCAPD10 / GFILLD2D.
        value = _normalize_fill_dcap_size(value)
        if _should_stop_drive_peel(value):
            break
        match = DRIVE_RE.search(value)
        if match is None:
            break
        value = value[: match.start()]
    value = _peel_trailing_tags(value)
    value = _peel_variant_suffixes(value)
    value = _normalize_fill_dcap_size(value)
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


def chinese_for_key(function_key: str, zh_map: Dict[str, str]) -> str:
    """Return the coarse Chinese label for column A (may be empty)."""
    return zh_map.get(function_key, "")


def format_family_column_label(
    family_id: str,
    fine_zh: str,
    sole_family_in_chinese: bool,
) -> str:
    """Build column-B text: family English, plus fine Chinese when needed.

    Skip Chinese when this family alone fills the whole A-group, or when no
    fine label is defined.
    """
    if sole_family_in_chinese or not fine_zh:
        return family_id
    return f"{family_id}\n{fine_zh}"


def make_display_key(
    name: str,
    function_keys: Sequence[str],
) -> Tuple[str, bool]:
    """Build match key via root / ordered prefix match.

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
        "(kept as column-B keys). Full list:",
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


def _families_by_chinese(
    groups: Dict[str, List[GroupItem]],
    zh_map: Dict[str, str],
) -> Dict[str, Dict[str, Tuple[str, List[GroupItem]]]]:
    """Group keys into coarse Chinese -> family_id -> (fine_zh, items)."""
    by_chinese: Dict[str, Dict[str, Tuple[str, List[GroupItem]]]] = defaultdict(
        dict
    )
    for key, items in groups.items():
        chinese = chinese_for_key(key, zh_map)
        family_id, fine_zh = key_family(key)
        if family_id in by_chinese[chinese]:
            _old_zh, old_items = by_chinese[chinese][family_id]
            old_items.extend(items)
        else:
            by_chinese[chinese][family_id] = (fine_zh, list(items))
    return by_chinese


def collect_missing_fine_zh(
    groups: Dict[str, List[GroupItem]],
    zh_map: Dict[str, str],
) -> List[Tuple[str, str, str]]:
    """Cells whose B-family lacks fine_zh while sharing an A-group with others.

    Returns (coarse_zh, family_id, cell) rows for copy-paste reporting.
    """
    missing: List[Tuple[str, str, str]] = []
    by_chinese = _families_by_chinese(groups, zh_map)
    for chinese, families in by_chinese.items():
        if len(families) <= 1:
            continue
        coarse = chinese or "(空中文)"
        for family_id, (fine_zh, items) in families.items():
            if fine_zh:
                continue
            for cell, _in_np, _in_c1 in items:
                missing.append((coarse, family_id, cell))
    missing.sort()
    return missing


def print_missing_fine_zh(
    missing: Sequence[Tuple[str, str, str]],
    np_set: Set[str],
    c1_set: Set[str],
) -> None:
    """Print families that need fine_zh under a shared coarse Chinese."""
    print(
        f"Warning: {len(missing)} cells in multi-family A-groups "
        "have empty fine_zh (B shows English only). Full list:",
        file=sys.stderr,
    )
    print("# coarse_zh\tfamily\tsource\tcell", file=sys.stderr)
    families: Set[str] = set()
    for coarse, family_id, cell in missing:
        families.add(f"{coarse}/{family_id}")
        source = unmatched_source_label(cell, np_set, c1_set)
        print(f"{coarse}\t{family_id}\t{source}\t{cell}", file=sys.stderr)
    print(
        f"# summary: cells={len(missing)} family_slots={len(families)}",
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

    Rows are ordered by coarse Chinese (A), then B-family, so A/B merges
    stay contiguous. Column B may add a fine Chinese label when multiple
    families share one A-group.
    """
    by_chinese = _families_by_chinese(groups, zh_map)

    rows: List[CellRow] = []
    for chinese in sorted(by_chinese.keys(), key=lambda value: value or "\uffff"):
        families = by_chinese[chinese]
        sole = len(families) == 1
        for family_id in sorted(families.keys()):
            fine_zh, items = families[family_id]
            key_label = format_family_column_label(family_id, fine_zh, sole)
            exact_matches, np_only, c1_only = _partition_group(items)
            for cell in exact_matches:
                rows.append((chinese, key_label, cell, cell))
            for cell in np_only:
                rows.append((chinese, key_label, cell, None))
            for cell in c1_only:
                rows.append((chinese, key_label, None, cell))
    return rows


# #### Excel writing
def _iter_value_groups(
    rows: Sequence[CellRow],
    value_index: int,
) -> List[Tuple[int, int]]:
    """Return inclusive Excel row ranges for contiguous equal values."""
    if not rows:
        return []

    groups: List[Tuple[int, int]] = []
    group_start = 2
    current = rows[0][value_index]
    for index in range(1, len(rows)):
        if rows[index][value_index] != current:
            groups.append((group_start, index + 1))
            group_start = index + 2
            current = rows[index][value_index]
    groups.append((group_start, len(rows) + 1))
    return groups


def _merge_column_by_value(
    worksheet: Worksheet,
    rows: Sequence[CellRow],
    excel_column: int,
    value_index: int,
) -> None:
    """Merge contiguous identical values in one column."""
    for start_row, end_row in _iter_value_groups(rows, value_index):
        if end_row > start_row:
            worksheet.merge_cells(
                start_row=start_row,
                start_column=excel_column,
                end_row=end_row,
                end_column=excel_column,
            )


def _apply_row_border(
    worksheet: Worksheet,
    row_index: int,
    border: Border,
) -> None:
    """Apply a border style across data columns on one row."""
    for column in range(1, LAST_DATA_COLUMN + 1):
        worksheet.cell(row=row_index, column=column).border = border


def _apply_section_borders(worksheet: Worksheet, rows: Sequence[CellRow]) -> None:
    """Add top/bottom borders for KEY groups; header keeps top only."""
    _apply_row_border(worksheet, 1, TOP_BORDER)

    # Section borders still follow column-B KEY blocks.
    for start_row, end_row in _iter_value_groups(rows, 1):
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

    for column in range(1, LAST_DATA_COLUMN + 1):
        worksheet.cell(row=1, column=column).font = BOLD_FONT

    for row_index in range(2, row_count + 2):
        zh_cell = worksheet.cell(row=row_index, column=1)
        key_cell = worksheet.cell(row=row_index, column=2)
        zh_cell.font = BOLD_FONT
        key_cell.font = BOLD_FONT
        zh_cell.alignment = KEY_ALIGNMENT
        key_cell.alignment = KEY_ALIGNMENT

        np_cell = worksheet.cell(row=row_index, column=3)
        c1_cell = worksheet.cell(row=row_index, column=4)
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

    worksheet["A1"] = "中文"
    worksheet["B1"] = ""
    worksheet["C1"] = "NP1PP"
    worksheet["D1"] = "C1Y"

    for row_index, (chinese, key_label, np_val, c1_val) in enumerate(
        rows, start=2
    ):
        worksheet.cell(row=row_index, column=1, value=chinese)
        worksheet.cell(row=row_index, column=2, value=key_label)
        if np_val:
            worksheet.cell(row=row_index, column=3, value=np_val)
        if c1_val:
            worksheet.cell(row=row_index, column=4, value=c1_val)

    # A: merge same Chinese; B: merge same KEY label (unchanged rule).
    _merge_column_by_value(worksheet, rows, excel_column=1, value_index=0)
    _merge_column_by_value(worksheet, rows, excel_column=2, value_index=1)
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
    missing_fine = collect_missing_fine_zh(groups, zh_map)
    print(
        f"Output rows: {len(rows)}, Unique function keys used: {len(groups)}, "
        f"Unmatched cells: {len(unmatched)}, "
        f"Missing fine_zh cells: {len(missing_fine)}"
    )
    if unmatched:
        print_unmatched_cells(unmatched, np_set, c1_set)
    if missing_fine:
        print_missing_fine_zh(missing_fine, np_set, c1_set)

    write_excel(args.output, rows)
    print(f"Done! Saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
