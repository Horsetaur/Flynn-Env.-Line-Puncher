from typing import Any, Dict, Optional


class RowInserter:
    def __init__(self) -> None:
        pass

    def add_row_to_category(self, ws: Any, active_row: int) -> None:
        # Placeholder: basic insert below active row
        ws.Rows(active_row + 1).Insert()

    def add_new_category(self, ws: Any, active_row: int) -> None:
        # Placeholder: insert an empty row and a separator row
        ws.Rows(active_row + 1).Insert()


