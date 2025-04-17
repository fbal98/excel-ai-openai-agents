

from __future__ import annotations

from typing import Any, Dict, List, Optional

import xlwings as xw


class LiveExcelManager:
    """Interact with an already‑running Excel instance via xlwings."""

    def __init__(self, file_path: Optional[str] = None, visible: bool = True) -> None:
        # Start (or connect to) Excel. A new app is safer than xw.App(visible=…) singleton.
        self.app = xw.App(visible=visible, add_book=not bool(file_path))
        self.book = (
            self.app.books.open(file_path)  # Open existing workbook
            if file_path
            else self.app.books.active  # New blank book already active
        )

    # ------------------------------------------------------------------ #
    #  Basic workbook info                                               #
    # ------------------------------------------------------------------ #
    def get_sheet_names(self) -> List[str]:
        return [s.name for s in self.book.sheets]

    def get_active_sheet_name(self) -> str:
        return self.book.sheets.active.name

    def get_sheet(self, sheet_name: str):
        try:
            return self.book.sheets[sheet_name]
        except (KeyError, ValueError):
            return None

    # ------------------------------------------------------------------ #
    #  Cell value helpers                                                #
    # ------------------------------------------------------------------ #
    def set_cell_value(self, sheet_name: str, cell_address: str, value: Any) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        sheet.range(cell_address).value = value

    def get_cell_value(self, sheet_name: str, cell_address: str) -> Any:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        return sheet.range(cell_address).value

    def set_cell_values(self, sheet_name: str, data: Dict[str, Any]) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        for addr, val in data.items():
            sheet.range(addr).value = val

    # ------------------------------------------------------------------ #
    #  Sheet management                                                  #
    # ------------------------------------------------------------------ #
    def create_sheet(self, sheet_name: str, index: Optional[int] = None) -> None:
        if sheet_name in self.get_sheet_names():
            raise ValueError(f"Sheet '{sheet_name}' already exists.")
        self.book.sheets.add(name=sheet_name, before=self.book.sheets[index] if index is not None else None)

    def delete_sheet(self, sheet_name: str) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found for deletion.")
        sheet.delete()

    # ------------------------------------------------------------------ #
    #  Misc (stubs for parity; extend as needed)                         #
    # ------------------------------------------------------------------ #
    def save_workbook(self, file_path: str) -> None:
        # In live mode we usually let the user hit Ctrl‑S, but provide save anyway.
        self.book.save(file_path)

    # The following methods are not yet implemented for xlwings but are kept
    # to satisfy the Agent tools. Implement them as needed.
    def set_range_style(self, *a, **kw):  # noqa: D401,E501
        raise NotImplementedError("set_range_style not yet supported in LiveExcelManager")

    def merge_cells_range(self, sheet_name: str, range_address: str) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        sheet.range(range_address).api.Merge()

    def unmerge_cells_range(self, sheet_name: str, range_address: str) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        sheet.range(range_address).api.UnMerge()

    def set_row_height(self, sheet_name: str, row_number: int, height: float) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        sheet.rows(row_number).row_height = height

    def set_column_width(self, sheet_name: str, column_letter: str, width: float) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        sheet.range(f"{column_letter}:{column_letter}").column_width = width

    def set_cell_formula(self, sheet_name: str, cell_address: str, formula: str) -> None:
        if not formula.startswith("="):
            formula = "=" + formula
        self.set_cell_value(sheet_name, cell_address, formula)

    # Style inspectors – return minimal info for now
    def get_cell_style(self, *a, **kw):  # noqa: D401
        raise NotImplementedError("Style inspection not yet supported in LiveExcelManager")

    def get_range_style(self, *a, **kw):  # noqa: D401
        raise NotImplementedError("Style inspection not yet supported in LiveExcelManager")