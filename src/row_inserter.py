from typing import Any
from pattern_analyzer import (
    find_nearest_header_merge_ws,
    find_vertical_merges_touching_row,
    find_nearest_data_row,
    detect_effective_max_cols,
)
from format_utils import (
    copy_merge_and_borders_from_above,
    apply_horizontal_merges_like_row,
    extend_vertical_merges_below,
    apply_borders_like_row,
    apply_neighbor_edge_borders,
)


class RowInserter:
    def __init__(self) -> None:
        pass

    def add_row_to_category(self, ws: Any, active_row: int) -> None:
        # Determine if the active row is the bottom of a vertical merge area.
        verticals = find_vertical_merges_touching_row(ws, active_row)
        is_bottom = False
        for top, _left, nrows, _ncols in verticals:
            if active_row == top + nrows - 1:
                is_bottom = True
                break

        # Preserve active column to restore selection after operations
        try:
            active_col = int(ws.Application.ActiveCell.Column)
        except Exception:
            active_col = 1

        ws.Rows(active_row + 1).Insert()
        used_cols = detect_effective_max_cols(ws, anchor_row=active_row)

        # If at bottom of a category block, copy from interior row and extend vertical merges
        ref_row = active_row if not is_bottom else max(1, active_row - 1)
        copy_merge_and_borders_from_above(ws, target_row=active_row + 1, ref_row=ref_row, max_cols=used_cols)
        apply_horizontal_merges_like_row(ws, source_row=ref_row, target_row=active_row + 1, max_cols=used_cols)
        apply_borders_like_row(ws, source_row=ref_row, target_row=active_row + 1, max_cols=used_cols)
        apply_neighbor_edge_borders(ws, target_row=active_row + 1, left_col=1, right_col=used_cols)
        if verticals:
            extend_vertical_merges_below(ws, verticals)

        # Restore selection
        try:
            ws.Cells(active_row + 1, active_col).Select()
        except Exception:
            pass

    def add_new_category(self, ws: Any, active_row: int) -> None:
        # Insert a spacer and a header-like row using nearest header merge
        try:
            active_col = int(ws.Application.ActiveCell.Column)
        except Exception:
            active_col = 1

        ws.Rows(active_row + 1).Insert()
        used_cols = detect_effective_max_cols(ws, anchor_row=active_row)
        header = find_nearest_header_merge_ws(ws, start_row=active_row)
        if header:
            copy_merge_and_borders_from_above(ws, target_row=active_row + 1, ref_row=header.row, max_cols=used_cols)
            apply_horizontal_merges_like_row(ws, source_row=header.row, target_row=active_row + 1, max_cols=used_cols)
            apply_borders_like_row(ws, source_row=header.row, target_row=active_row + 1, max_cols=used_cols)
            apply_neighbor_edge_borders(ws, target_row=active_row + 1, left_col=1, right_col=used_cols)
            # After creating a header row, immediately add a data-style row below using nearest data row as template
            data_template_row = find_nearest_data_row(ws, start_row=active_row, used_cols=used_cols)
            if data_template_row is not None:
                ws.Rows(active_row + 2).Insert()
                copy_merge_and_borders_from_above(ws, target_row=active_row + 2, ref_row=data_template_row, max_cols=used_cols)
                apply_horizontal_merges_like_row(ws, source_row=data_template_row, target_row=active_row + 2, max_cols=used_cols)
                apply_borders_like_row(ws, source_row=data_template_row, target_row=active_row + 2, max_cols=used_cols)
                apply_neighbor_edge_borders(ws, target_row=active_row + 2, left_col=1, right_col=used_cols)
                try:
                    ws.Cells(active_row + 2, active_col).Select()
                except Exception:
                    pass
        else:
            copy_merge_and_borders_from_above(ws, target_row=active_row + 1, ref_row=active_row, max_cols=used_cols)
            apply_borders_like_row(ws, source_row=active_row, target_row=active_row + 1, max_cols=used_cols)
            apply_neighbor_edge_borders(ws, target_row=active_row + 1, left_col=1, right_col=used_cols)
            try:
                ws.Cells(active_row + 1, active_col).Select()
            except Exception:
                pass


