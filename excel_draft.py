"""
Excel read/write for Hardball Dynasty amateur draft.
Maps between workbook sheets (Hitters, Pitchers) and draft order.
"""
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd


# Sheet names and header row indices (1-based in Excel)
HITTERS_SHEET = "Hitters"
PITCHERS_SHEET = "Pitchers"
HITTERS_HEADER_ROW = 6   # row 6 has: total, Rnk, Player, Pos, B, T, Age, ...
PITCHERS_HEADER_ROW = 5  # row 5 has: Overall Projection, Rnk, Player, Pos, B, T, Age, ...

# Column names we care about (as they appear in header)
RNK = "Rnk"
PLAYER = "Player"
POS = "Pos"
B = "B"
T = "T"
AGE = "Age"
# Score columns used to rank when interleaving hitters and pitchers
HITTERS_SCORE_COL = "Overall Projection"
PITCHERS_SCORE_COL = "Overall Projection"

# Hitting block: columns B–P (basic info + 9 hitting ratings from "Projected Hitting Ratings" view).
# Positional keys (Rating_N) always exist in scraped rows; they're the primary lookup.
HITTERS_HITTING_LAYOUT: list[tuple[int, str, list[str]]] = [
    (2, "Rnk", ["Rating_1", "Rnk", "Rank"]),
    (3, "Player", ["Rating_2", "Player", "Player Name"]),
    (4, "Pos", ["Rating_3", "Pos", "Position"]),
    (5, "B", ["Rating_4", "B", "Bats"]),
    (6, "T", ["Rating_5", "T", "Throws"]),
    (7, "Age", ["Rating_6", "Age"]),
    (8, "Contact", ["Rating_7"]),
    (9, "Power", ["Rating_8"]),
    (10, "vs L", ["Rating_9"]),
    (11, "vs R", ["Rating_10"]),
    (12, "Batting Eye", ["Rating_11"]),
    (13, "Baserunning", ["Rating_12"]),
    (14, "Arm", ["Rating_13"]),
    (15, "Bunt", ["Rating_14"]),
    (16, "Overall", ["Rating_15"]),
]

# Fielding block starts at column Q (17).
# Direct copy-paste of the "Projected Fielding/General Ratings" view from the website.
FIELDING_BLOCK_START_COL = 17

# Pitchers sheet: single block starting at column B (2).
# 6 basic columns + 13 pitching ratings, all using positional keys from the scrape.
PITCHERS_LAYOUT: list[tuple[int, str, list[str]]] = [
    (2, "Rank", ["Rating_1", "Rnk", "Rank"]),
    (3, "Player", ["Rating_2", "Player", "Player Name"]),
    (4, "Position", ["Rating_3", "Pos", "Position"]),
    (5, "B", ["Rating_4", "B", "Bats"]),
    (6, "T", ["Rating_5", "T", "Throws"]),
    (7, "Age", ["Rating_6", "Age"]),
    (8, "Durability", ["Rating_7"]),
    (9, "Stamina", ["Rating_8"]),
    (10, "Control", ["Rating_9"]),
    (11, "vsL", ["Rating_10"]),
    (12, "vsR", ["Rating_11"]),
    (13, "Velocity", ["Rating_12"]),
    (14, "Groundball/Flyball Tendency", ["Rating_13"]),
    (15, "Pitch 1", ["Rating_14"]),
    (16, "Pitch 2", ["Rating_15"]),
    (17, "Pitch 3", ["Rating_16"]),
    (18, "Pitch 4", ["Rating_17"]),
    (19, "Pitch 5", ["Rating_18"]),
    (20, "Overall", ["Rating_19"]),
]

# Display headers for fielding columns, in website order (Fielding_1 = first column, etc.).
FIELDING_HEADERS: list[str] = [
    "Rank", "Player", "Pos", "B", "T", "Age",
    "Range", "Glove", "Arm Strength", "Arm Accuracy",
    "Pitch Calling", "Durability", "Health", "Speed",
    "Patience", "Temper", "Makeup", "Overall",
]


def _row_value_for_keys(row: dict[str, Any], keys: list[str]) -> Any:
    """Return row[key] for the first key that exists (case-insensitive key match)."""
    row_lower = {str(k).strip().lower(): v for k, v in row.items() if k is not None and str(k).strip()}
    for k in keys:
        if k is None or not str(k).strip():
            continue
        if k in row:
            return row[k]
        kl = str(k).strip().lower()
        for rk, rv in row_lower.items():
            if rk == kl:
                return rv
    return None


def _header_to_col(ws: openpyxl.worksheet.worksheet.Worksheet, header_row: int, name: str) -> int | None:
    """Return 1-based column index for header name (exact match), or None."""
    for col in range(1, ws.max_column + 1):
        val = ws.cell(header_row, col).value
        if val is not None and str(val).strip() == name:
            return col
    return None


def _header_to_col_ic(ws: openpyxl.worksheet.worksheet.Worksheet, header_row: int, name: str) -> int | None:
    """Return 1-based column index for header name (case-insensitive), or None."""
    name_lower = name.lower()
    for col in range(1, ws.max_column + 1):
        val = ws.cell(header_row, col).value
        if val is not None and str(val).strip().lower() == name_lower:
            return col
    return None


def _parse_score(val: Any) -> float:
    """Parse a cell value as a numeric score; return 0 if missing or invalid."""
    if val is None:
        return 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def validate_template(path: str | Path) -> list[str]:
    """
    Check that the workbook has the expected structure. Returns list of error messages (empty if OK).
    """
    path = Path(path)
    errors: list[str] = []
    if not path.exists():
        return [f"File not found: {path}"]
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return [f"Cannot open workbook: {e}"]
    for sheet_name, header_row, score_col in [
        (HITTERS_SHEET, HITTERS_HEADER_ROW, HITTERS_SCORE_COL),
        (PITCHERS_SHEET, PITCHERS_HEADER_ROW, PITCHERS_SCORE_COL),
    ]:
        if sheet_name not in wb.sheetnames:
            errors.append(f"Missing sheet: {sheet_name}")
            continue
        ws = wb[sheet_name]
        if _header_to_col(ws, header_row, PLAYER) is None and _header_to_col_ic(ws, header_row, PLAYER) is None:
            errors.append(f"Sheet '{sheet_name}' has no 'Player' column in header row {header_row}")
        if _header_to_col(ws, header_row, score_col) is None and _header_to_col_ic(ws, header_row, score_col) is None:
            errors.append(f"Sheet '{sheet_name}' has no '{score_col}' column in header row {header_row}")
    wb.close()
    return errors


def get_draft_order_from_excel(path: str | Path) -> list[str]:
    """
    Read draft order from Excel: Hitters and Pitchers interleaved and ranked by score
    (Hitters by 'total', Pitchers by 'Overall Projection'), highest score first.
    Returns list of player names in that order.
    """
    path = Path(path)
    if not path.exists():
        return []

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    scored: list[tuple[float, str]] = []

    for sheet_name, header_row, score_col in [
        (HITTERS_SHEET, HITTERS_HEADER_ROW, HITTERS_SCORE_COL),
        (PITCHERS_SHEET, PITCHERS_HEADER_ROW, PITCHERS_SCORE_COL),
    ]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        player_col = _header_to_col(ws, header_row, PLAYER)
        score_col_idx = _header_to_col(ws, header_row, score_col) or _header_to_col_ic(ws, header_row, score_col)
        if player_col is None:
            continue
        data_start = header_row + 1
        for row in range(data_start, ws.max_row + 1):
            name = ws.cell(row, player_col).value
            if name is None or not str(name).strip():
                continue
            score = _parse_score(ws.cell(row, score_col_idx).value) if score_col_idx else 0.0
            scored.append((score, str(name).strip()))

    wb.close()
    # Sort by score descending (best first), then by name for ties
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [name for _, name in scored]


def _normalize_name(name: str) -> str:
    """Normalize for matching (strip, single spaces)."""
    return " ".join(str(name).strip().split())


def get_excel_row_ranges(path: str | Path) -> tuple[tuple[int, int], tuple[int, int]]:
    """Return (hitters_data_start, hitters_data_end), (pitchers_data_start, pitchers_data_end) (1-based)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    hitters = (HITTERS_HEADER_ROW + 1, 0)
    pitchers = (PITCHERS_HEADER_ROW + 1, 0)
    if HITTERS_SHEET in wb.sheetnames:
        ws = wb[HITTERS_SHEET]
        hitters = (HITTERS_HEADER_ROW + 1, ws.max_row)
    if PITCHERS_SHEET in wb.sheetnames:
        ws = wb[PITCHERS_SHEET]
        pitchers = (PITCHERS_HEADER_ROW + 1, ws.max_row)
    wb.close()
    return hitters, pitchers


def _cols_with_error_or_empty(ws: openpyxl.worksheet.worksheet.Worksheet, header_row: int) -> list[int]:
    """Return 1-based column indices where header is #VALUE!, empty, or looks like an error (for overwriting)."""
    result = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(header_row, col).value
        if val is None:
            result.append(col)
            continue
        s = str(val).strip()
        if s == "" or s == "#VALUE!" or "VALUE" in s.upper() or "REF" in s.upper():
            result.append(col)
    return result


def _write_hitters_sheet_fixed(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    header_row: int,
    rows: list[dict[str, Any]],
) -> None:
    """
    Write hitters data using the fixed template layout.
    Hitting block (B–P): basic info + 9 hitting ratings.
    Fielding block (Q onwards): direct copy-paste of the Fielding/General view (all columns in website order).
    """
    # Hitting block
    for col_idx, header_label, keys in HITTERS_HITTING_LAYOUT:
        ws.cell(header_row, col_idx, header_label)
        for i, row in enumerate(rows):
            val = _row_value_for_keys(row, keys)
            if val is not None:
                ws.cell(header_row + 1 + i, col_idx, val)

    # Fielding block: all Fielding_N keys in order, starting at column Q
    if not rows:
        return
    fielding_keys = sorted(
        [k for k in rows[0] if k.startswith("Fielding_")],
        key=lambda k: int(k.split("_")[1]),
    )
    if not fielding_keys:
        return
    start_col = FIELDING_BLOCK_START_COL
    for offset, fkey in enumerate(fielding_keys):
        col = start_col + offset
        header = FIELDING_HEADERS[offset] if offset < len(FIELDING_HEADERS) else fkey
        ws.cell(header_row, col, header)
        for i, row in enumerate(rows):
            val = row.get(fkey)
            if val is not None:
                ws.cell(header_row + 1 + i, col, val)


MASTER_LIST_SHEET = "Master List"


def _write_master_list(
    wb: openpyxl.Workbook,
    hitters_rows: list[dict[str, Any]],
    pitchers_rows: list[dict[str, Any]],
) -> None:
    """
    Create a Master List tab with all players (hitters + pitchers) sorted by raw Overall rating.
    Column A has a formula pointing to the source sheet's column A (the template's projection formula),
    so the real projection value appears when the file is opened in Excel.
    """
    if MASTER_LIST_SHEET in wb.sheetnames:
        del wb[MASTER_LIST_SHEET]
    ws = wb.create_sheet(MASTER_LIST_SHEET)

    players: list[tuple[float, str, str, str, str, int]] = []
    for i, row in enumerate(hitters_rows):
        excel_row = HITTERS_HEADER_ROW + 1 + i
        name = _row_value_for_keys(row, ["Rating_2", "Player", "Player Name"])
        pos = _row_value_for_keys(row, ["Rating_3", "Pos", "Position"])
        overall = _row_value_for_keys(row, ["Rating_15"]) or 0
        players.append((_parse_score(overall), str(name or ""), str(pos or ""), "Hitter", HITTERS_SHEET, excel_row))

    for i, row in enumerate(pitchers_rows):
        excel_row = PITCHERS_HEADER_ROW + 1 + i
        name = _row_value_for_keys(row, ["Rating_2", "Player", "Player Name"])
        pos = _row_value_for_keys(row, ["Rating_3", "Pos", "Position"])
        overall = _row_value_for_keys(row, ["Rating_19"]) or 0
        players.append((_parse_score(overall), str(name or ""), str(pos or ""), "Pitcher", PITCHERS_SHEET, excel_row))

    players.sort(key=lambda x: (-x[0], x[1]))

    ws.cell(1, 1, "Overall Projection")
    ws.cell(1, 2, "Player")
    ws.cell(1, 3, "Pos")
    ws.cell(1, 4, "Type")

    for i, (_, name, pos, ptype, src_sheet, src_row) in enumerate(players):
        out_row = i + 2
        ws.cell(out_row, 1, f"='{src_sheet}'!A{src_row}")
        ws.cell(out_row, 2, name)
        ws.cell(out_row, 3, pos)
        ws.cell(out_row, 4, ptype)


def write_draft_data_to_excel(
    path: str | Path,
    hitters_rows: list[dict[str, Any]],
    pitchers_rows: list[dict[str, Any]],
    output_path: str | Path | None = None,
) -> None:
    """
    Write scraped draft pool data into the Excel file.
    Hitters sheet: fixed layout (hitting block B–P, fielding block Q+).
    Pitchers sheet: fixed layout (single block B–T).
    If output_path is set, saves to that path instead of path (useful for saving to ./outputs/).
    """
    path = Path(path)
    wb = openpyxl.load_workbook(path)

    # Hitters: fixed column layout so output matches template
    if HITTERS_SHEET in wb.sheetnames and hitters_rows:
        _write_hitters_sheet_fixed(wb[HITTERS_SHEET], HITTERS_HEADER_ROW, hitters_rows)

    # Pitchers: fixed column layout
    if PITCHERS_SHEET in wb.sheetnames and pitchers_rows:
        ws = wb[PITCHERS_SHEET]
        for col_idx, header_label, keys in PITCHERS_LAYOUT:
            ws.cell(PITCHERS_HEADER_ROW, col_idx, header_label)
            for i, row in enumerate(pitchers_rows):
                val = _row_value_for_keys(row, keys)
                if val is not None:
                    ws.cell(PITCHERS_HEADER_ROW + 1 + i, col_idx, val)

    # Master List: all players combined, sorted by raw Overall rating.
    # Column A uses a formula referencing each player's source sheet column A,
    # so when opened in Excel the real projection value appears and the user can re-sort.
    _write_master_list(wb, hitters_rows, pitchers_rows)

    save_to = Path(output_path) if output_path else path
    save_to.parent.mkdir(parents=True, exist_ok=True)
    wb.save(save_to)
    wb.close()


def append_draft_order_sheet(path: str | Path, order: list[str], sheet_name: str = "DraftOrder") -> None:
    """Append a simple 'DraftOrder' sheet with one column of player names in order."""
    path = Path(path)
    wb = openpyxl.load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.cell(1, 1, "Order")
    ws.cell(2, 1, "Player")
    for i, name in enumerate(order, start=1):
        ws.cell(i + 2, 1, name)
    wb.save(path)
    wb.close()
