import json
import os
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional, Tuple

try:
    import win32com.client as win32
except Exception:
    win32 = None  # Allows linting on non-Windows or without pywin32


@dataclass
class MergeAreaInfo:
    top: int
    left: int
    rows: int
    cols: int


@dataclass
class CellFormatInfo:
    address: str
    row: int
    col: int
    value_preview: str
    merge: Optional[MergeAreaInfo]
    borders: Dict[str, Dict[str, Any]]
    font: Dict[str, Any]


class ExcelPatternAnalyzer:
    def __init__(self, directory_path: str, include_borders: bool = False) -> None:
        self.directory_path = directory_path
        self.include_borders = include_borders

    def _list_excel_files(self) -> List[str]:
        allowed_ext = {".xls", ".xlsx", ".xlsm"}
        files: List[str] = []
        for root, _, filenames in os.walk(self.directory_path):
            for name in filenames:
                path = os.path.join(root, name)
                if os.path.isfile(path) and os.path.splitext(name)[1].lower() in allowed_ext:
                    files.append(path)
        return sorted(files)

    def analyze(self, max_cells_per_sheet: int = 2000) -> Dict[str, Any]:
        if win32 is None:
            raise RuntimeError("pywin32 is required to analyze Excel files on Windows.")

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        results: Dict[str, Any] = {"files": []}
        try:
            for file_path in self._list_excel_files():
                file_result: Dict[str, Any] = {"file": file_path, "sheets": []}
                try:
                    wb = excel.Workbooks.Open(os.path.abspath(file_path))
                except Exception as open_err:
                    file_result["error"] = f"open_failed: {open_err}"
                    results["files"].append(file_result)
                    continue

                try:
                    for sheet in wb.Worksheets:
                        sheet_info = self._analyze_sheet(sheet, max_cells_per_sheet)
                        file_result["sheets"].append(sheet_info)
                finally:
                    wb.Close(SaveChanges=False)

                results["files"].append(file_result)
        finally:
            excel.Quit()

        return results

    def _analyze_sheet(self, sheet: Any, max_cells_per_sheet: int) -> Dict[str, Any]:
        used_range = sheet.UsedRange
        rows = int(used_range.Rows.Count)
        cols = int(used_range.Columns.Count)

        # Sample cells from the used range to keep runtime bounded
        sampled_cells = self._sample_cells(rows, cols, max_cells_per_sheet)

        cells_info: List[Dict[str, Any]] = []
        for r, c in sampled_cells:
            cell = sheet.Cells(r, c)
            try:
                merge_info = None
                if bool(cell.MergeCells):
                    area = cell.MergeArea
                    merge_info = MergeAreaInfo(
                        top=int(area.Row),
                        left=int(area.Column),
                        rows=int(area.Rows.Count),
                        cols=int(area.Columns.Count),
                    )

                # Fast mode: optionally skip borders and font to reduce COM overhead
                if self.include_borders:
                    borders = self._extract_borders(cell)
                    font = self._extract_font(cell)
                else:
                    borders = {}
                    font = {}

                value = cell.Text if hasattr(cell, "Text") else str(cell.Value)

                cell_info = CellFormatInfo(
                    address=str(cell.Address),
                    row=int(r),
                    col=int(c),
                    value_preview=str(value)[:40],
                    merge=merge_info,
                    borders=borders,
                    font=font,
                )
                cells_info.append(asdict(cell_info))
            except Exception as err:
                cells_info.append({
                    "row": int(r),
                    "col": int(c),
                    "error": f"cell_inspect_failed: {err}",
                })

        merge_blocks_summary = self._summarize_merge_blocks(cells_info)

        return {
            "name": str(sheet.Name),
            "used_rows": rows,
            "used_cols": cols,
            "sampled_cell_count": len(sampled_cells),
            "merge_blocks_summary": merge_blocks_summary,
            "cells": cells_info,
        }

    def _sample_cells(self, rows: int, cols: int, max_cells: int) -> List[Tuple[int, int]]:
        coords: List[Tuple[int, int]] = []
        if rows <= 0 or cols <= 0:
            return coords

        total = rows * cols
        if total <= max_cells:
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    coords.append((r, c))
            return coords

        # Grid sampling when the sheet is large
        step = max(1, int((total / max_cells) ** 0.5))
        for r in range(1, rows + 1, step):
            for c in range(1, cols + 1, step):
                coords.append((r, c))
        return coords[:max_cells]

    def _extract_borders(self, cell: Any) -> Dict[str, Dict[str, Any]]:
        # Excel border index mapping
        # 1: xlEdgeLeft, 2: xlEdgeTop, 3: xlEdgeBottom, 4: xlEdgeRight, 5: xlInsideVertical, 6: xlInsideHorizontal
        sides = {
            1: "left",
            2: "top",
            3: "bottom",
            4: "right",
        }
        border_info: Dict[str, Dict[str, Any]] = {}
        for idx, name in sides.items():
            try:
                b = cell.Borders(idx)
                border_info[name] = {
                    "line_style": int(b.LineStyle) if b.LineStyle is not None else None,
                    "weight": int(b.Weight) if b.Weight is not None else None,
                    "color": int(b.Color) if b.Color is not None else None,
                }
            except Exception as _:
                border_info[name] = {"error": "border_read_failed"}
        return border_info

    def _extract_font(self, cell: Any) -> Dict[str, Any]:
        try:
            f = cell.Font
            return {
                "name": str(getattr(f, "Name", "")),
                "size": int(getattr(f, "Size", 0) or 0),
                "bold": bool(getattr(f, "Bold", False)),
                "italic": bool(getattr(f, "Italic", False)),
                "color": int(getattr(f, "Color", 0) or 0),
            }
        except Exception as _:
            return {"error": "font_read_failed"}

    def _summarize_merge_blocks(self, cells: List[Dict[str, Any]]) -> Dict[str, Any]:
        blocks: Dict[str, int] = {}
        for c in cells:
            merge = c.get("merge")
            if not merge:
                continue
            key = f"{merge['rows']}x{merge['cols']}"
            blocks[key] = blocks.get(key, 0) + 1
        return {
            "block_sizes": blocks,
            "distinct_block_count": len(blocks),
        }


def write_json_report(data: Dict[str, Any], out_path: str) -> None:
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


