from typing import Any, Optional, Tuple

try:
    import win32com.client as win32
except Exception:  # pragma: no cover
    win32 = None


class ExcelConnector:
    def __init__(self) -> None:
        if win32 is None:
            raise RuntimeError("pywin32 is required on Windows")
        self.app = win32.Dispatch("Excel.Application")

    def get_active_cell(self) -> Tuple[Any, Any, Any]:
        wb = self.app.ActiveWorkbook
        ws = self.app.ActiveSheet
        cell = self.app.ActiveCell
        if ws is None or cell is None:
            raise RuntimeError("No active worksheet or cell.")
        return wb, ws, cell

    def insert_row_below(self, ws: Any, row_index: int) -> None:
        ws.Rows(row_index + 1).Insert()

    def quit(self) -> None:
        try:
            self.app.Quit()
        except Exception:
            pass


