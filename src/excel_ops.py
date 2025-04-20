"""
Unified ExcelManager: realtime xlwings backend with snapshot / undo support.
The public surface (method names / signatures) is preserved so existing tools
continue to work without modification.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import asyncio
import logging
from typing import Any, Dict, List, Optional, TYPE_CHECKING

import xlwings as xw

if TYPE_CHECKING:
    from .context import WorkbookShape # For type hinting


class ExcelManager:
    """Single realtime manager that always drives a visible Excel instance."""

    # ──────────────────────────────
    #  Construction / housekeeping
    # ──────────────────────────────
    def __init__(
        self,
        file_path: Optional[str] = None,
        visible: bool = True,
        *,
        kill_others: bool = False,
        attach_existing: bool = False,
    ) -> None:
        """
        Prepare an ExcelManager.

        Parameters
        ----------
        file_path:
            Workbook path to open. If *None*, a new blank workbook will be created
            when the context is entered.
        visible:
            Whether the Excel window should be visible.
        kill_others:
            If *True*, attempt to quit all running Excel instances *before* launching
            a fresh one.  Defaults to *False* (do not disturb other sessions).
        attach_existing:
            If *True* **and** an Excel instance is already running, re‑use the
            active instance instead of launching a new one.
        """
        # Configuration only – real work happens in ``__aenter__``.
        self._file_path = file_path
        self._visible = visible
        self._kill_others = kill_others
        self._attach_existing = attach_existing

        # Handles populated later
        self.app: Optional[xw.App] = None
        self.book: Optional[xw.Book] = None

        # Tracking for snapshot / undo helper
        self._snapshot_path: Optional[str] = None
    # ──────────────────────────────
    #  Async context‑manager helpers
    # ──────────────────────────────
    async def __aenter__(self) -> "ExcelManager":
        """Initialise Excel resources on entering an ``async with`` block."""
        if self.app is None:
            # Optionally close other Excel processes
            if self._kill_others:
                for _app in xw.apps:
                    try:
                        _app.quit()
                    except Exception:
                        pass

            # Optionally attach to an existing process
            if self._attach_existing and xw.apps:
                try:
                    self.app = xw.apps.active
                except Exception:
                    self.app = None

            # Otherwise start a new instance
            if self.app is None:
                self.app = xw.App(visible=self._visible, add_book=False)

            # Open or create workbook
            if self._file_path:
                self.book = self.app.books.open(self._file_path)
            else:
                self.book = self.app.books.add()

            # Ensure at least one sheet exists
            if not self.book.sheets:
                self.book.sheets.add()

        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        """Gracefully close workbook and guarantee the Excel process dies."""
        try:
            if self.book:
                try:
                    self.book.close()
                except Exception:
                    pass
        finally:
            if self.app:
                try:
                    # Use kill when available to prevent zombie COM hosts
                    if hasattr(self.app, "kill"):
                        self.app.kill()
                    else:
                        self.app.quit()
                except Exception:
                    pass
                self.app = None
                self.book = None

    # Optional synchronous helper for legacy call‑sites
    def close(self) -> None:
        """Explicitly dispose Excel handles (sync)."""
        if self.book:
            try:
                self.book.close()
            except Exception:
                pass
            self.book = None
        if self.app:
            try:
                self.app.quit()
            except Exception:
                pass
            self.app = None



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
    #  Ensure changes are applied
    # ──────────────────────────────
    async def ensure_changes_applied(self) -> None:
        """Asynchronously flush Excel UI and calculation pipelines.

        This method yields to the event loop for ≈0.5 s, preventing the hard
        stop caused by ``time.sleep`` while Excel finishes painting.
        """
        logger = logging.getLogger(__name__)
        try:
            # Force a visual and calculation refresh
            self.app.screen_updating = False
            self.app.screen_updating = True
            self.app.calculate()

            # Re‑activate active sheet to nudge UI
            active_sheet = self.book.sheets.active
            active_sheet.activate()

            # Give Excel a brief moment without blocking the loop
            await asyncio.sleep(0.5)
            logger.debug("Excel display refreshed.")
        except Exception as e:
            logger.debug(f"Could not refresh Excel display: {e}")
            
    async def save_with_confirmation(self, file_path: str | None = None) -> str:
        """Save the workbook and **return the full path**.

        This helper is now *async* so it can await
        :pyfunc:`ensure_changes_applied` before persisting.
        """
        logger = logging.getLogger(__name__)

        # Flush Excel changes first
        await self.ensure_changes_applied()

        if not file_path:
            from datetime import datetime
            file_path = f"workbook_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

        # Guarantee the .xlsx extension
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            self.book.save(file_path)
            logger.debug(f"Workbook saved to: {file_path}")
            return file_path
        except Exception as e:
            logger.debug(f"Primary save '{file_path}' failed: {e}")
            # Fallback to ~/Documents
            try:
                documents = os.path.expanduser("~/Documents")
                alt_path = os.path.join(documents, os.path.basename(file_path))
                self.book.save(alt_path)
                logger.debug(f"Workbook saved to fallback location: {alt_path}")
                return alt_path
            except Exception as e2:
                raise RuntimeError(f"All save attempts failed: {e2}") from e2

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

    def quick_scan_shape(self) -> "WorkbookShape":
        """
        Scans the current workbook state via xlwings and returns a WorkbookShape object.
        Raises exceptions if critical operations fail (e.g., accessing book).
        Logs warnings for non-critical issues (e.g., cannot read headers).
        """
        from .context import WorkbookShape # Avoid circular import at top level
        import logging
        logger = logging.getLogger(__name__)

        if not self.book:
            raise RuntimeError("Cannot scan shape: No active workbook found in ExcelManager.")

        shape = WorkbookShape()
        book = self.book

        # 1. Scan sheets for used range and headers
        for sheet in book.sheets:
            try:
                sheet_name = sheet.name
                # Get used range - handle potential errors if sheet is empty
                try:
                    # Use api for potentially more robust access, fallback to xlwings property
                    used_range_api = sheet.api.UsedRange
                    last_cell_addr = sheet.range((used_range_api.Row + used_range_api.Rows.Count - 1,
                                                  used_range_api.Column + used_range_api.Columns.Count - 1)).address.replace("$","")

                    # Handle truly empty sheet case where UsedRange might still return A1
                    if last_cell_addr == 'A1' and sheet.range('A1').value is None and len(book.sheets) > 1 : # Check value only if A1 reported
                         first_cell_val = sheet.range('A1').value
                         if first_cell_val is None:
                            # If A1 is the only cell and it's empty, consider the sheet effectively empty.
                            shape.sheets[sheet_name] = "A1:A1" # Represent as single cell
                         else:
                            shape.sheets[sheet_name] = f"A1:{last_cell_addr}"
                    else:
                         shape.sheets[sheet_name] = f"A1:{last_cell_addr}"

                except Exception as range_err:
                    # Could happen on completely empty sheets or COM errors
                    logger.warning(f"Could not determine used range for sheet '{sheet_name}': {range_err}. Defaulting to A1:A1.")
                    shape.sheets[sheet_name] = "A1:A1" # Fallback

                # Get headers (first row) - handle potential errors/empty rows
                try:
                    # Reading row 1 can be slow on huge sheets, optimize if needed later
                    header_values = sheet.range("1:1").value
                    if isinstance(header_values, list):
                        # Track the original length for logging
                        original_length = len(header_values)
                        # Remove trailing empty columns to reduce token usage
                        while header_values and (header_values[-1] is None or header_values[-1] == ""):
                            header_values.pop()
                        # Ensure all headers are strings, handle None
                        shape.headers[sheet_name] = [str(c) if c is not None else "" for c in header_values]
                        
                        # Log information about trimmed columns
                        retained = len(header_values)
                        trimmed = original_length - retained
                        logger.debug(f"Sheet '{sheet_name}': Headers trimmed from {original_length} to {retained} columns (removed {trimmed} empty trailing columns)")
                    elif header_values is not None: # Handle single-column sheet case
                        shape.headers[sheet_name] = [str(header_values)]
                    else: # Empty first row
                        shape.headers[sheet_name] = []
                except Exception as header_err:
                    logger.warning(f"Could not read headers for sheet '{sheet_name}': {header_err}. Defaulting to empty list.")
                    shape.headers[sheet_name] = [] # Fallback to empty list

            except Exception as sheet_err:
                logger.error(f"Error processing sheet '{getattr(sheet, 'name', 'unknown')}': {sheet_err}. Skipping sheet in shape.")
                continue # Skip this sheet on error

        # 2. Scan named ranges
        try:
            for name_obj in book.names:
                nm = name_obj.name
                try:
                    # Check if refers_to_range exists and retrieve address
                    refers_to = name_obj.refers_to
                    if refers_to.startswith("="): # It's likely a formula or constant
                         # Try to get refers_to_range, might fail if complex/external
                         addr = name_obj.refers_to_range.address.replace("$", "")
                         shape.names[nm] = addr
                    else: # Should be a direct range reference
                         addr = name_obj.refers_to_range.address.replace("$", "")
                         shape.names[nm] = addr

                except Exception as name_ref_err:
                    # Sometimes refers_to might be a constant or complex formula, not a range
                    logger.warning(f"Could not resolve address for named range '{nm}' (refers_to='{name_obj.refers_to}'): {name_ref_err}. Storing refers_to string.")
                    # Store the raw refers_to string if address fails
                    shape.names[nm] = name_obj.refers_to

        except Exception as names_err:
            logger.error(f"Error accessing named ranges: {names_err}. Skipping named ranges in shape.")
            # Continue without names if there's a general error

        shape.version = 0 # Base version, caller (AppContext) will manage incrementing
        return shape

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
            • 'font': {'bold': True/False, 'color': 'FFRRGGBB'}
            • 'fill': {'fill_type': 'solid'/'pattern'/'gradient', 'start_color': 'FFAABBCC', 'end_color': 'FFAABBCC'}
            • 'border': {'left': {'style': 'thin'}, 'right': {'style': 'thin'}, 'top': {'style': 'thin'}, 'bottom': {'style': 'thin'}}
        Extend as needed.
        """
        sheet = self._require_sheet(sheet_name)
        rng = sheet.range(range_address)

        # Font → bold, color
        if "font" in style and style["font"]:
            # Handle bold
            bold = style["font"].get("bold")
            if bold is not None:
                rng.font.bold = bool(bold)
                
            # Handle font color
            color = style["font"].get("color")
            if color is not None:
                try:
                    rgb_tuple = _hex_argb_to_bgr_int(color)
                    rng.font.color = rgb_tuple
                except:
                    pass

        # Fill
        if "fill" in style and style["fill"]:
            # Get fill type
            fill_type = style["fill"].get("fill_type", "solid")
            
            # Handle start color
            if "start_color" in style["fill"]:
                rgb = style["fill"]["start_color"]
                try:
                    color_int = _hex_argb_to_bgr_int(rgb)
                    rng.color = color_int
                except Exception as e:
                    print(f"Color application error: {e}")
        
        # Borders
        if "border" in style and style["border"]:
            try:
                border = style["border"]
                # Apply borders if specified
                if "left" in border:
                    rng.api.Borders(7).LineStyle = 1  # xlContinuous
                if "right" in border:
                    rng.api.Borders(10).LineStyle = 1  # xlContinuous
                if "top" in border:
                    rng.api.Borders(8).LineStyle = 1  # xlContinuous
                if "bottom" in border:
                    rng.api.Borders(9).LineStyle = 1  # xlContinuous
            except Exception as e:
                print(f"Border application error: {e}")
                
        # Force update the Excel application to show changes
        try:
            self.app.screen_updating = False
            self.app.screen_updating = True
        except:
            pass

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
        """Merge cells in the specified range."""
        sheet = self._require_sheet(sheet_name)
        try:
            # Direct API approach
            sheet.range(range_address).api.Merge()
        except Exception as e:
            try:
                # Alternative xlwings approach
                sheet.range(range_address).merge()
            except Exception as e2:
                print(f"Failed to merge cells: {e2}")

    def unmerge_cells_range(self, sheet_name: str, range_address: str) -> None:
        """Unmerge cells in the specified range."""
        sheet = self._require_sheet(sheet_name)
        try:
            # Direct API approach
            sheet.range(range_address).api.UnMerge()
        except Exception as e:
            try:
                # Alternative xlwings approach
                sheet.range(range_address).unmerge()
            except Exception as e2:
                print(f"Failed to unmerge cells: {e2}")

    # ──────────────────────────────
    #  Row / column sizing
    # ──────────────────────────────
    def set_row_height(self, sheet_name: str, row_number: int, height: float) -> None:
        """Set the height of a specific row in the given sheet."""
        sheet = self._require_sheet(sheet_name)
        try:
            # First attempt with direct row method
            sheet.api.Rows(row_number).RowHeight = height
        except Exception as e:
            try:
                # Alternative approach using range
                row_range = f"{row_number}:{row_number}"
                sheet.range(row_range).row_height = height
            except Exception as e2:
                print(f"Failed to set row height: {e2}")
                raise RuntimeError(f"Failed to set row height for row {row_number} in '{sheet_name}': {e2}")

    def set_column_width(self, sheet_name: str, column_letter: str, width: float) -> None:
        rng = f"{column_letter}:{column_letter}"
        self._require_sheet(sheet_name).range(rng).column_width = width

    # ──────────────────────────────
    #  Copy / Paste range helper
    # ──────────────────────────────
    def copy_paste_range(
        self,
        src_sheet_name: str,
        src_range: str,
        dst_sheet_name: str,
        dst_anchor: str,
        paste_opts: str = "values",
    ) -> None:
        """
        Clone *src_range* from *src_sheet_name* and paste into *dst_sheet_name*
        at *dst_anchor* in a single round‑trip.

        paste_opts:
            • "values"   → values only
            • "formulas" → formulas only
            • "formats"  → formats only
        """
        src_sheet = self._require_sheet(src_sheet_name)
        dst_sheet = self._require_sheet(dst_sheet_name)

        src_rng = src_sheet.range(src_range)
        rows = src_rng.rows.count
        cols = src_rng.columns.count
        dst_rng = dst_sheet.range(dst_anchor).resize(rows, cols)

        opts = paste_opts.lower()
        if opts == "values":
            dst_rng.value = src_rng.value
        elif opts == "formulas":
            dst_rng.formula = src_rng.formula
        elif opts == "formats":
            # xlPasteFormats = ‑4104
            src_rng.api.Copy()
            dst_rng.api.PasteSpecial(Paste=-4104)
        else:
            raise ValueError(
                f"Invalid paste_opts '{paste_opts}'. Use 'values', 'formulas', or 'formats'."
            )

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
def _hex_argb_to_bgr_int(argb: str) -> int:
    """
    Convert an **8‑digit ARGB** string (``'FFRRGGBB'`` or ``'#FFRRGGBB'``) to an
    integer in BGR byte order for the Excel COM API.

    The function now *requires* the alpha channel; sending a 6‑digit RGB code
    raises ``ValueError`` so callers cannot silently lose transparency
    information.
    """
    s = argb.lstrip("#")
    if len(s) != 8:
        raise ValueError(
            f"Color '{argb}' must be 8‑digit ARGB (e.g. 'FF3366CC' or '#FF3366CC')."
        )

    # Drop alpha then swap to BGR
    r, g, b = s[2:4], s[4:6], s[6:8]
    return int(f"{b}{g}{r}", 16)


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