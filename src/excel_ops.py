"""
Unified ExcelManager: realtime xlwings backend with snapshot / undo support.
The public surface (method names / signatures) is preserved so existing tools
continue to work without modification.
"""

from __future__ import annotations

import os
import shutil
import tempfile
from typing import Any, Dict, List, Optional

import xlwings as xw


class ExcelManager:
    """Single realtime manager that always drives a visible Excel instance."""

    # ──────────────────────────────
    #  Construction / housekeeping
    # ──────────────────────────────
    def __init__(self, file_path: Optional[str] = None, visible: bool = True) -> None:
        """
        Launch Excel (or reuse an existing instance) and open *file_path*.
        If *file_path* is None a new blank workbook is created.
        """
        # Kill any existing Excel instances to avoid conflicts
        for app in xw.apps:
            try:
                app.quit()
            except:
                pass

        try:
            # Spawn a fresh Excel instance
            self.app = xw.App(visible=visible, add_book=False)
            
            # Create or open workbook
            if file_path:
                self.book = self.app.books.open(file_path)
            else:
                self.book = self.app.books.add()

            # Ensure there's at least one sheet
            if not self.book.sheets:
                self.book.sheets.add()

        except Exception as e:
            if hasattr(self, 'app'):
                try:
                    self.app.quit()
                except:
                    pass
            raise RuntimeError(f"Failed to initialize Excel: {e}")
            
        self._snapshot_path: Optional[str] = None  # path to last snapshot (temp file)

    def __del__(self):
        """Ensure Excel instance is cleaned up."""
        try:
            if hasattr(self, 'book') and self.book:
                try:
                    self.book.close()
                except:
                    pass
            if hasattr(self, 'app') and self.app:
                try:
                    self.app.quit()
                except:
                    pass
        except:
            pass

    # ──────────────────────────────
    #  Snapshot / undo helpers
    # ──────────────────────────────
    def snapshot(self) -> str:
        """Save a temp copy that can be rolled back to with `revert_to_snapshot()`."""
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(tmp_fd)
        self.book.save(tmp_path)
        self._snapshot_path = tmp_path
        return tmp_path

    def revert_to_snapshot(self) -> None:
        """Close current book and reopen the last snapshot (if any)."""
        if not self._snapshot_path or not os.path.exists(self._snapshot_path):
            raise RuntimeError("No snapshot available to revert to.")
        # Close without saving
        self.book.close(save_changes=False)
        # Open the snapshot
        self.book = self.app.books.open(self._snapshot_path)

    # ──────────────────────────────
    #  Explicit save helpers
    # ──────────────────────────────
    def save_workbook(self, file_path: str = None) -> None:
        """Save the current workbook. If no path is provided, save to a default location."""
        if not file_path:
            # Generate a default filename with timestamp
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"workbook_{timestamp}.xlsx"
        
        # Ensure the path has .xlsx extension
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'
            
        self.save_as(file_path)
        return file_path  # Return the path where the file was saved

    def save_as(self, file_path: str) -> None:
        """Save the workbook to the specified path, ensuring proper extension."""
        # Ensure the path has .xlsx extension
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'
            
        try:
            self.book.save(file_path)
        except Exception as e:
            raise RuntimeError(f"Failed to save workbook to {file_path}: {e}")

    # ──────────────────────────────
    #  New: open workbook helper
    # ──────────────────────────────
    def open_workbook(self, file_path: str) -> None:
        """Close the current book without saving and open the workbook at file_path."""
        try:
            if self.book:
                try:
                    self.book.close(save_changes=False)
                except:
                    pass
            self.book = self.app.books.open(file_path)
            # Ensure there's at least one sheet
            if not self.book.sheets:
                self.book.sheets.add()
        except Exception as e:
            raise RuntimeError(f"Failed to open workbook {file_path}: {e}")
        self._snapshot_path = None

    # ──────────────────────────────
    #  Basic workbook / sheet info
    # ──────────────────────────────
    def get_sheet_names(self) -> List[str]:
        return [s.name for s in self.book.sheets]

    def get_active_sheet_name(self) -> str:
        return self.book.sheets.active.name

    def get_sheet(self, sheet_name: str):
        try:
            return self.book.sheets[sheet_name]
        except (KeyError, ValueError):
            return None

    # ──────────────────────────────
    #  Cell value helpers
    # ──────────────────────────────
    def set_cell_value(self, sheet_name: str, cell_address: str, value: Any) -> None:
        sheet = self._require_sheet(sheet_name)
        sheet.range(cell_address).value = value

    def get_cell_value(self, sheet_name: str, cell_address: str) -> Any:
        sheet = self._require_sheet(sheet_name)
        return sheet.range(cell_address).value

    def set_cell_values(self, sheet_name: str, data: Dict[str, Any]) -> None:
        sheet = self._require_sheet(sheet_name)
        for addr, val in data.items():
            sheet.range(addr).value = val

    # ──────────────────────────────
    #  Range helpers
    # ──────────────────────────────
    def get_range_values(self, sheet_name: str, range_address: str) -> List[List[Any]]:
        sheet = self._require_sheet(sheet_name)
        vals = sheet.range(range_address).value
        # xlwings returns scalar for 1×1 range; list for others
        if not isinstance(vals, list):
            return [[vals]]
        # Normalise 1-D row or col to 2-D list-of-lists
        if vals and not isinstance(vals[0], list):
            return [vals]
        return vals

    # ──────────────────────────────
    #  Styles (minimal viable impl)
    # ──────────────────────────────
    def set_range_style(
        self, sheet_name: str, range_address: str, style: Dict[str, Any]
    ) -> None:
        """
        Currently supports:
            • 'font': {'bold': True/False}
            • 'fill': {'start_color': 'FFAABBCC' | '#AABBCC' | 'AABBCC'}
        Extend as needed.
        """
        sheet = self._require_sheet(sheet_name)
        rng = sheet.range(range_address)

        # Font → bold only for now
        if "font" in style and style["font"]:
            bold = style["font"].get("bold")
            if bold is not None:
                try:
                    rng.font.bold = bool(bold)
                except:
                    # Fallback to Excel API
                    rng.api.Font.Bold = bool(bold)

        # Fill
        if "fill" in style and style["fill"] and "start_color" in style["fill"]:
            rgb = style["fill"]["start_color"]
            rgb_tuple = _hex_argb_to_bgr_int(rgb)
            rng.color = rgb_tuple

    # ──────────────────────────────
    #  Sheet management
    # ──────────────────────────────
    def create_sheet(self, sheet_name: str, index: Optional[int] = None) -> None:
        if sheet_name in self.get_sheet_names():
            raise ValueError(f"Sheet '{sheet_name}' already exists.")
        before = self.book.sheets[index] if index is not None else None
        self.book.sheets.add(name=sheet_name, before=before)

    def delete_sheet(self, sheet_name: str) -> None:
        sheet = self._require_sheet(sheet_name)
        sheet.delete()

    # ──────────────────────────────
    #  Merge / unmerge
    # ──────────────────────────────
    def merge_cells_range(self, sheet_name: str, range_address: str) -> None:
        self._require_sheet(sheet_name).range(range_address).merge()

    def unmerge_cells_range(self, sheet_name: str, range_address: str) -> None:
        self._require_sheet(sheet_name).range(range_address).unmerge()

    # ──────────────────────────────
    #  Row / column sizing
    # ──────────────────────────────
    def set_row_height(self, sheet_name: str, row_number: int, height: float) -> None:
        self._require_sheet(sheet_name).rows(row_number).row_height = height

    def set_column_width(self, sheet_name: str, column_letter: str, width: float) -> None:
        rng = f"{column_letter}:{column_letter}"
        self._require_sheet(sheet_name).range(rng).column_width = width

    # ──────────────────────────────
    #  (Currently stub) advanced APIs
    # ──────────────────────────────
    def set_cell_formula(self, sheet_name: str, cell_address: str, formula: str) -> None:
        if not formula.startswith("="):
            formula = "=" + formula
        self.set_cell_value(sheet_name, cell_address, formula)

    # ──────────────────────────────
    #  Style inspectors
    # ──────────────────────────────
    def get_cell_style(self, sheet_name: str, cell_address: str) -> Dict[str, Any]:  # noqa: D401
        """Return a minimal style dict (bold + fill color) for a single cell."""
        sheet = self._require_sheet(sheet_name)
        rng = sheet.range(cell_address)

        style: Dict[str, Any] = {}

        # Font → bold
        bold = rng.api.Font.Bold
        if bold is not None:
            style["font"] = {"bold": bool(bold)}

        # Fill → start_color
        interior_color = rng.api.Interior.Color
        if interior_color not in (None, 0):  # 0 = no fill
            style["fill"] = {"start_color": _bgr_int_to_argb_hex(interior_color)}

        return style

    def get_range_style(self, sheet_name: str, range_address: str) -> Dict[str, Dict[str, Any]]:  # noqa: D401
        """
        Return {cell_address: style_dict} for every cell in the range (minimal style set).
        """
        sheet = self._require_sheet(sheet_name)
        rng = sheet.range(range_address)
        result: Dict[str, Dict[str, Any]] = {}
        for c in rng:
            addr = c.address.replace("$", "")
            font_bold = c.api.Font.Bold
            fill_color = c.api.Interior.Color
            cell_style: Dict[str, Any] = {}
            if font_bold is not None:
                cell_style["font"] = {"bold": bool(font_bold)}
            if fill_color not in (None, 0):
                cell_style.setdefault("fill", {})["start_color"] = _bgr_int_to_argb_hex(
                    fill_color
                )
            if cell_style:
                result[addr] = cell_style
        return result

    # Data-frame style dump for inspection / verification
    def get_sheet_dataframe(self, sheet_name: str, header: bool = True):
        values = self.get_range_values(sheet_name, _full_sheet_range(self._require_sheet(sheet_name)))
        if not values:
            return {"columns": [], "rows": []}
        if header:
            columns = [
                str(c) if c is not None else f"col_{i+1}"
                for i, c in enumerate(values[0])
            ]
            rows = values[1:]
        else:
            columns = [f"col_{i+1}" for i in range(len(values[0]))]
            rows = values
        return {"columns": columns, "rows": rows}

    # ──────────────────────────────
    #  Helpers
    # ──────────────────────────────
    def _require_sheet(self, sheet_name: str):
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        return sheet

    # ──────────────────────────────
    #  Table insertion
    # ──────────────────────────────
    def insert_table(
        self,
        sheet_name: str,
        start_cell: str,
        columns: List[Any],
        rows: List[List[Any]],
        table_name: Optional[str] = None,
        table_style: Optional[str] = None,
    ) -> None:
        """
        Inserts a formatted Excel table (ListObject) into the worksheet.
        """
        sheet = self._require_sheet(sheet_name)
        header_cell = sheet.range(start_cell)
        total_rows = 1 + len(rows)
        total_cols = len(columns)
        table_range = header_cell.resize(total_rows, total_cols)
        
        # Write header and data
        table_range.value = [columns] + rows
        
        try:
            # Try using the Excel API directly
            lo = sheet.api.ListObjects.Add(1, table_range.api, None, 1)
            if table_name:
                lo.Name = table_name
            if table_style:
                lo.TableStyle = table_style
        except:
            # Fallback: Just format as a regular range if table creation fails
            header_row = header_cell.resize(row_size=1, column_size=total_cols)
            try:
                header_row.api.Font.Bold = True
            except:
                header_row.font.bold = True
            
            try:
                # Light blue header background
                header_row.api.Interior.Color = _hex_argb_to_bgr_int("FFD9E1F2")
            except:
                header_row.color = _hex_argb_to_bgr_int("FFD9E1F2")
            
            # Add basic borders
            try:
                table_range.api.Borders.LineStyle = 1  # xlContinuous
            except:
                pass


# ╭────────────────────────── Helper functions ─────────────────────────╮
def _hex_argb_to_bgr_int(argb_or_rgb: str) -> tuple[int, int, int]:
    """
    Convert 'FFAABBCC', '#AABBCC', or 'AABBCC' → (r, g, b) tuple.
    Alpha is discarded.
    """
    s = argb_or_rgb.lstrip("#")
    if len(s) == 8:  # ARGB
        s = s[2:]
    if len(s) != 6:
        raise ValueError(f"Invalid RGB color '{argb_or_rgb}'")
    r_hex, g_hex, b_hex = s[0:2], s[2:4], s[4:6]
    r, g, b = int(r_hex, 16), int(g_hex, 16), int(b_hex, 16)
    return (r, g, b)


def _bgr_int_to_argb_hex(color_int: int) -> str:
    """
    Convert a BGR integer used by Excel back to an 8-digit ARGB hex string (FFRRGGBB).
    """
    # Mask to 24-bit and ensure 6-hex-digit string
    bgr_hex = f"{color_int & 0xFFFFFF:06X}"
    b, g, r = bgr_hex[0:2], bgr_hex[2:4], bgr_hex[4:6]
    return f"FF{r}{g}{b}"


def _full_sheet_range(sheet) -> str:
    """Return A1-style full-used-range of a sheet (simplistic)."""
    last_cell = sheet.used_range.last_cell
    return f"A1:{last_cell.address.replace('$', '')}"