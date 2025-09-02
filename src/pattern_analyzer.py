from dataclasses import dataclass
from typing import Any, List, Optional, Tuple


@dataclass
class MergeBlock:
    row: int
    start_col: int
    end_col: int
    width: int


def _get_merge_area(cell: Any) -> Tuple[int, int, int, int]:
    """Return (top_row, left_col, num_rows, num_cols) for cell's merge area or the cell itself."""
    if bool(getattr(cell, "MergeCells", False)):
        area = cell.MergeArea
        return int(area.Row), int(area.Column), int(area.Rows.Count), int(area.Columns.Count)
    return int(cell.Row), int(cell.Column), 1, 1


def find_horizontal_merges_on_row(ws: Any, row: int, max_cols: int = 30) -> List[MergeBlock]:
    """Detect horizontal merge blocks on a given row. Only returns width > 1 blocks."""
    merges: List[MergeBlock] = []
    c = 1
    while c <= max_cols:
        cell = ws.Cells(row, c)
        top, left, nrows, ncols = _get_merge_area(cell)
        if nrows == 1 and ncols > 1 and top == row:
            merges.append(MergeBlock(row=row, start_col=left, end_col=left + ncols - 1, width=ncols))
            c = left + ncols
        else:
            c += 1
    return merges


def find_nearest_header_merge_ws(ws: Any, start_row: int, scan_up: int = 20, max_cols: int = 30) -> Optional[MergeBlock]:
    """Scan upwards to find the widest 1xN horizontal merge block (probable header)."""
    best: Optional[MergeBlock] = None
    r = max(1, start_row - scan_up)
    for row in range(start_row, r - 1, -1):
        blocks = find_horizontal_merges_on_row(ws, row, max_cols=max_cols)
        for b in blocks:
            if best is None or b.width > best.width:
                best = b
    return best


def find_vertical_merges_touching_row(ws: Any, row: int, max_scan_cols: int = 7) -> List[Tuple[int, int, int, int]]:
    """Return list of vertical merge areas that include the given row for first N columns.

    Returns tuples: (top_row, left_col, num_rows, num_cols)
    """
    areas: List[Tuple[int, int, int, int]] = []
    for c in range(1, max_scan_cols + 1):
        cell = ws.Cells(row, c)
        top, left, nrows, ncols = _get_merge_area(cell)
        if nrows > 1:  # vertical span
            areas.append((top, left, nrows, ncols))
    return areas


def is_header_like_row(ws: Any, row: int, used_cols: int, threshold_ratio: float = 0.5, min_width: int = 5) -> bool:
    """Heuristic: a row is header-like if it contains a horizontal 1xN merge that spans
    at least max(min_width, used_cols * threshold_ratio) columns.
    """
    blocks = find_horizontal_merges_on_row(ws, row, max_cols=used_cols)
    if not blocks:
        return False
    widest = max(b.width for b in blocks)
    return widest >= max(min_width, int(used_cols * threshold_ratio))


def find_nearest_data_row(ws: Any, start_row: int, used_cols: int, scan_distance: int = 25) -> Optional[int]:
    """Find the nearest non-header-like row around start_row.
    Prefer rows above to keep category style consistent with prior data.
    """
    # Scan upwards first
    for r in range(start_row - 1, max(1, start_row - scan_distance) - 1, -1):
        if not is_header_like_row(ws, r, used_cols):
            return r
    # Then scan downwards
    for r in range(start_row + 1, start_row + scan_distance + 1):
        try:
            _ = ws.Rows(r)  # ensure row exists
        except Exception:
            break
        if not is_header_like_row(ws, r, used_cols):
            return r
    return None


def detect_effective_max_cols(ws: Any, anchor_row: int, hard_cap: int = 50) -> int:
    """Estimate the effective table width starting from an anchor row.
    Prefers the widest horizontal merge on/above the row; otherwise scans rightward
    until the last cell with content, merge, or any border is found.
    """
    try:
        used_cols = int(ws.UsedRange.Columns.Count)
    except Exception:
        used_cols = hard_cap
    used_cols = min(used_cols, hard_cap)

    header = find_nearest_header_merge_ws(ws, start_row=anchor_row, max_cols=used_cols)
    if header:
        return min(header.end_col, used_cols)

    # Fallback: scan this row
    last = 1
    for c in range(1, used_cols + 1):
        cell = ws.Cells(anchor_row, c)
        top, left, nrows, ncols = _get_merge_area(cell)
        has_merge = (ncols > 1 or nrows > 1)
        has_value = str(getattr(cell, "Text", "") or getattr(cell, "Value", "")).strip() != ""
        has_border = False
        try:
            for idx in (1, 2, 3, 4):
                b = cell.Borders(idx)
                if getattr(b, "LineStyle", 0):
                    has_border = True
                    break
        except Exception:
            pass
        if has_merge:
            last = max(last, left + ncols - 1)
        elif has_value or has_border:
            last = max(last, c)
    return max(1, min(last, used_cols))


