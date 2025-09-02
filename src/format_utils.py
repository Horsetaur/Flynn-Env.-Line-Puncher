from typing import Any

def _range(ws: Any, r1: int, c1: int, r2: int, c2: int) -> Any:
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))


def _clear_row_borders(ws: Any, row: int, max_cols: int) -> None:
    for c in range(1, max_cols + 1):
        try:
            cell = ws.Cells(row, c)
            for idx in (1, 2, 3, 4):
                b = cell.Borders(idx)
                b.LineStyle = 0
        except Exception:
            continue


def copy_merge_and_borders_from_above(ws: Any, target_row: int, ref_row: int, max_cols: int = 30) -> None:
    """Lightweight format copy from ref_row to target_row for 1..max_cols.
    - Does NOT copy borders (handled separately to avoid outlines)
    - Copies basic font/alignment/number format only
    - Clears existing borders on the target row first
    """
    _clear_row_borders(ws, target_row, max_cols)
    for c in range(1, max_cols + 1):
        try:
            src = ws.Cells(ref_row, c)
            dst = ws.Cells(target_row, c)
            # Number format and alignment
            try:
                dst.NumberFormat = src.NumberFormat
            except Exception:
                pass
            try:
                dst.HorizontalAlignment = src.HorizontalAlignment
                dst.VerticalAlignment = src.VerticalAlignment
                dst.WrapText = src.WrapText
            except Exception:
                pass
            # Font basics
            try:
                dst.Font.Name = src.Font.Name
                dst.Font.Size = src.Font.Size
                dst.Font.Bold = src.Font.Bold
                dst.Font.Italic = src.Font.Italic
                dst.Font.Color = src.Font.Color
            except Exception:
                pass
            # Interior/shading
            try:
                dst.Interior.Color = src.Interior.Color
                dst.Interior.Pattern = src.Interior.Pattern
            except Exception:
                pass
        except Exception:
            continue


def apply_horizontal_merges_like_row(ws: Any, source_row: int, target_row: int, max_cols: int = 30) -> None:
    c = 1
    while c <= max_cols:
        cell = ws.Cells(source_row, c)
        try:
            if bool(getattr(cell, "MergeCells", False)):
                area = cell.MergeArea
                top = int(area.Row)
                left = int(area.Column)
                nrows = int(area.Rows.Count)
                ncols = int(area.Columns.Count)
                if nrows == 1 and top == source_row and ncols > 1:
                    _range(ws, target_row, left, target_row, left + ncols - 1).Merge()
                    c = left + ncols
                    continue
        except Exception:
            pass
        c += 1


def extend_vertical_merges_below(ws: Any, areas: list[tuple[int, int, int, int]]) -> None:
    for top, left, nrows, ncols in areas:
        try:
            # Extend by one row (after insertion)
            _range(ws, top, left, top + nrows, left + ncols - 1).Merge()
        except Exception:
            continue


def _copy_border_props(src: Any, dst: Any) -> None:
    # Border indices: 1-left, 2-top, 3-bottom, 4-right
    for idx in (1, 2, 3, 4):
        try:
            s = src.Borders(idx)
            d = dst.Borders(idx)
            d.LineStyle = s.LineStyle
            d.Weight = s.Weight
            d.Color = s.Color
        except Exception:
            continue


def apply_borders_like_row(ws: Any, source_row: int, target_row: int, max_cols: int = 30) -> None:
    """Copy perimeter border properties from source_row to target_row.
    Handles merged horizontal blocks by copying the merged range edge borders.
    """
    c = 1
    while c <= max_cols:
        cell = ws.Cells(source_row, c)
        try:
            if bool(getattr(cell, "MergeCells", False)):
                area = cell.MergeArea
                top = int(area.Row)
                left = int(area.Column)
                nrows = int(area.Rows.Count)
                ncols = int(area.Columns.Count)
                if nrows == 1 and top == source_row and ncols > 1:
                    src_rng = _range(ws, source_row, left, source_row, left + ncols - 1)
                    dst_rng = _range(ws, target_row, left, target_row, left + ncols - 1)
                    _copy_border_props(src_rng, dst_rng)
                    c = left + ncols
                    continue
        except Exception:
            pass

        # Non-merged cell: copy borders cell-to-cell
        try:
            src_cell = ws.Cells(source_row, c)
            dst_cell = ws.Cells(target_row, c)
            _copy_border_props(src_cell, dst_cell)
        except Exception:
            pass
        c += 1


def apply_neighbor_edge_borders(ws: Any, target_row: int, left_col: int, right_col: int) -> None:
    """Match the target row's top/bottom edges to neighbors above/below per column.
    This helps preserve dashed/solid patterns across inserted rows.
    """
    for c in range(left_col, right_col + 1):
        try:
            # Top edge from row above bottom edge
            if target_row > 1:
                above = ws.Cells(target_row - 1, c)
                curr = ws.Cells(target_row, c)
                curr.Borders(2).LineStyle = above.Borders(3).LineStyle  # top from above's bottom
                curr.Borders(2).Weight = above.Borders(3).Weight
                curr.Borders(2).Color = above.Borders(3).Color
        except Exception:
            pass
        try:
            # Bottom edge from row below top edge
            below = ws.Cells(target_row + 1, c)
            curr = ws.Cells(target_row, c)
            curr.Borders(3).LineStyle = below.Borders(2).LineStyle  # bottom from below's top
            curr.Borders(3).Weight = below.Borders(2).Weight
            curr.Borders(3).Color = below.Borders(2).Color
        except Exception:
            pass


def copy_font_from_cell(src_cell: Any, dst_cell: Any) -> None:
    try:
        dst_cell.Font.Name = src_cell.Font.Name
        dst_cell.Font.Size = src_cell.Font.Size
        dst_cell.Font.Bold = src_cell.Font.Bold
        dst_cell.Font.Italic = src_cell.Font.Italic
        dst_cell.Font.Color = src_cell.Font.Color
    except Exception:
        pass


