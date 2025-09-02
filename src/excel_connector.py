from typing import Any, Optional, Tuple

try:
    import win32com.client as win32
except Exception:  # pragma: no cover
    win32 = None


class ExcelConnector:
    def __init__(self) -> None:
        if win32 is None:
            raise RuntimeError("pywin32 is required on Windows")
        # Try to attach to a running Excel instance first; fall back to new
        try:
            self.app = win32.GetActiveObject("Excel.Application")
        except Exception:
            self.app = win32.Dispatch("Excel.Application")
        self.app.Visible = True

    def application(self) -> Any:
        return self.app

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


class ExcelPerformanceTuner:
    """Context manager to speed up COM operations and restore settings after."""

    def __init__(self, app: Any) -> None:
        self.app = app
        self._screen_updating: Optional[bool] = None
        self._enable_events: Optional[bool] = None
        self._display_alerts: Optional[bool] = None
        self._calculation: Optional[int] = None

    def __enter__(self) -> "ExcelPerformanceTuner":
        try:
            self._screen_updating = bool(self.app.ScreenUpdating)
            self._enable_events = bool(self.app.EnableEvents)
            self._display_alerts = bool(self.app.DisplayAlerts)
            self._calculation = int(self.app.Calculation)

            self.app.ScreenUpdating = False
            self.app.EnableEvents = False
            self.app.DisplayAlerts = False
            # xlCalculationManual = -4135
            self.app.Calculation = -4135
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            if self._screen_updating is not None:
                self.app.ScreenUpdating = self._screen_updating
            if self._enable_events is not None:
                self.app.EnableEvents = self._enable_events
            if self._display_alerts is not None:
                self.app.DisplayAlerts = self._display_alerts
            if self._calculation is not None:
                self.app.Calculation = self._calculation
            # Clear copy mode (marching ants)
            self.app.CutCopyMode = False
        except Exception:
            pass


