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
import logging # Ensure logging is imported
import os # Ensure os is imported
from typing import Any, Dict, List, Optional, TYPE_CHECKING


import xlwings as xw
from xlwings.constants import LineStyle, BorderWeight, PasteType
import re
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

if TYPE_CHECKING:
    from .context import WorkbookShape # For type hinting


# ── Cross‑platform style probe ────────────────────────────────────────────────
def _safe_cell_style(cell) -> dict[str, Any]:
    """Return {font:{bold}, fill:{start_color}} safely on every platform."""
    style: dict[str, Any] = {}

    # Bold --------------------------------------------------------------------
    bold = None
    for probe in (
        lambda c: c.api.Font.Bold,   # fast on Windows
        lambda c: c.font.bold,       # works everywhere
    ):
        try:
            bold = probe(cell)
            break
        except Exception:
            pass
    if bold is not None:
        style.setdefault("font", {})["bold"] = bool(bold)

    # Fill --------------------------------------------------------------------
    color = None
    for probe in (
        lambda c: c.api.Interior.Color,
        lambda c: c.color,
    ):
        try:
            color = probe(cell)
            break
        except Exception:
            pass
    if color not in (None, 0):
        style.setdefault("fill", {})["start_color"] = _bgr_int_to_argb_hex(color)

    return style


class ExcelManager:
    def _normalise_rows(self, columns: list[Any], rows: list[list[Any]]) -> list[list[Any]]:
        """Pad or truncate each row so they're exactly as wide as `columns`."""
        width = len(columns)
        fixed: list[list[Any]] = []
        for r in rows:
            if len(r) < width:
                # pad short rows
                fixed.append(r + [None] * (width - len(r)))
            elif len(r) > width:
                # truncate long rows
                fixed.append(r[:width])
            else:
                fixed.append(r)
        return fixed
    """Single realtime manager that always drives a visible Excel instance."""

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Construction / housekeeping
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def __init__(
        self,
        file_path: Optional[str] = None,
        visible: bool = True,
        *,
        kill_others: bool = False,
        attach_existing: bool = False,
        single_workbook: bool = True,
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
            If *True* **and** an Excel instance is already running, re-use the
            active instance instead of launching a new one.
        single_workbook:
            If *True*, automatically close all other open workbooks in the same Excel instance,
            leaving only the one managed by this class.
        """
        # Configuration only â€“ real work happens in ``__aenter__``.
        self._file_path = file_path
        self._visible = visible
        self._kill_others = kill_others
        self._attach_existing = attach_existing
        self._single_workbook = single_workbook

        # Handles populated later
        self.app: Optional[xw.App] = None
        self.book: Optional[xw.Book] = None

        # Tracking for snapshot / undo helper
        self._snapshot_path: Optional[str] = None
        self._attached_mode: bool = False # Flag to track if we attached to existing Excel

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Async contextâ€‘manager helpers
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    async def __aenter__(self) -> "ExcelManager":
        """Initialise Excel resources on entering an ``async with`` block."""
        # Removed the premature check:
        # if not self.app or self.app.pid is None: return

        logger = logging.getLogger(__name__)
        logger.debug("Entering ExcelManager.__aenter__...") # Added debug log
        self.app = None
        self.book = None
        self._attached_mode = False # Reset flag on entry

        # Optionally close other Excel processes
        if self._kill_others:
            logger.info("Attempting to quit other running Excel instances...")
            killed_count = 0
            for _app in xw.apps:
                try:
                    _app.quit()
                    killed_count += 1
                except Exception as e:
                    logger.warning("Could not quit an Excel instance: %s", e)
            logger.info("Quit %d other Excel instance(s).", killed_count)

        # Optionally attach to an existing process
        if self._attach_existing and xw.apps:
            try:
                self.app = xw.apps.active
                self._attached_mode = True # Set flag
                logger.info("Attached to existing Excel instance (PID: %s)", self.app.pid)
            except Exception as e:
                logger.warning("Failed to attach to existing Excel instance: %s. Starting new instance.", e)
                self.app = None
                self._attached_mode = False # Ensure flag is False if attach fails

        # Start a new instance if needed
        if self.app is None:
            logger.info("Starting new Excel instance...")
            # Ensure add_book=False is critical here to avoid premature workbook creation
            self.app = xw.App(visible=self._visible, add_book=False)
            self._attached_mode = False # Explicitly false if we created a new app
            logger.info("New Excel instance started (PID: %s)", self.app.pid)


        # --- Workbook Handling ---
        target_file_name = os.path.basename(self._file_path) if self._file_path else None
        logger.debug("Target file name: %s, Attached mode: %s", target_file_name, self._attached_mode)

        if self._attached_mode:
            # We are attached to an existing app. Try to find the target book or use active/add new.
            found_book = None
            if target_file_name:
                logger.debug("Attached mode: Looking for workbook '%s' in %d open books.", target_file_name, len(self.app.books))
                for wb in self.app.books:
                    if wb.name == target_file_name:
                        found_book = wb
                        logger.info("Found target workbook '%s' already open in attached instance.", target_file_name)
                        break
                if not found_book:
                    # Book not found, try opening it
                    try:
                        logger.info("Opening specified workbook '%s' in attached instance...", self._file_path)
                        found_book = self.app.books.open(self._file_path)
                    except Exception as e:
                        logger.error("Failed to open specified workbook '%s' in attached instance: %s. Creating a new blank workbook instead.", self._file_path, e)
                        # Fallback to adding a new book if open fails
                        found_book = self.app.books.add()
                        logger.info("Added new blank workbook to attached instance as fallback.")
            else:
                # No file specified. Use active if available, otherwise add new.
                if self.app.books:
                    found_book = self.app.books.active # Use the currently active book
                    logger.info("Using active workbook '%s' in attached instance (no file specified).", found_book.name)
                else:
                    # No books open in the attached instance, add one.
                    logger.info("No workbooks open in attached instance. Adding a new blank workbook.")
                    found_book = self.app.books.add()

            self.book = found_book

        else:
            # We created a new app instance. Open file or add new book.
            if self._file_path:
                logger.info("Opening specified workbook '%s' in new instance...", self._file_path)
                try:
                    self.book = self.app.books.open(self._file_path)
                except Exception as e:
                    logger.error("Failed to open specified workbook '%s' in new instance: %s. Creating a new blank workbook instead.", self._file_path, e)
                    self.book = self.app.books.add() # Fallback
                    logger.info("Added new blank workbook to new instance as fallback.")

            else:
                # ── If we created a new Excel instance and the caller didn’t ask
                #    for a specific file, Excel may already have opened Book1.
                #    Using len(...) avoids a truthiness quirk in xlwings on macOS.
                if len(self.app.books) > 0:                  # reuse default workbook
                    try:
                        self.book = self.app.books.active    # prefer the active workbook
                    except Exception:
                        self.book = self.app.books[0]        # fall back to first workbook
                    logger.info("Re‑using default workbook %s", getattr(self.book, "name", "unknown"))
                else:                                        # Excel really is empty
                    logger.info("Adding new blank workbook to new instance…")
                    self.book = self.app.books.add()

        # Ensure the managed book is activated and visible
        if self.book:
            try:
                logger.debug("Activating workbook: %s", self.book.name)
                self.book.activate()
                logger.info("Managed workbook set to: %s", self.book.name)
                if self._visible and hasattr(self.app, 'activate'):
                    logger.debug("Activating Excel application window.")
                    self.app.activate(steal_focus=True)
            except Exception as e:
                logger.warning("Could not activate workbook '%s': %s", self.book.name, e)

            # If single_workbook is True, close all other workbooks
            if self._single_workbook:
                extra_count = 0
                for wb in list(self.app.books):
                    if wb != self.book:
                        try:
                            wb.close()
                            extra_count += 1
                        except Exception as e:
                            logger.warning("Could not close extra workbook: %s", e)
                if extra_count > 0:
                    logger.info("Closed %d extra workbook(s). Only one workbook remains.", extra_count)
                    # Ensure our handle is still valid after closing others
                    try:
                        self.book = self.app.books.active
                    except Exception:
                        try:
                            self.book = self.app.books[0]
                        except Exception:
                            logger.warning("Could not refresh managed workbook handle after cleanup.")
        else:
            # This case should ideally not happen if the logic above is correct
            raise RuntimeError("Failed to obtain a workbook handle within ExcelManager.__aenter__.")


        # Ensure at least one sheet exists in the managed book
        if self.book and not self.book.sheets:
            logger.info("Workbook '%s' has no sheets. Adding default sheet 'Sheet1'.", self.book.name)
            self.book.sheets.add(name="Sheet1") # Give it a default name

        logger.debug("ExcelManager.__aenter__ completed.")
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        """Gracefully close the managed workbook and potentially the Excel app."""
        logger = logging.getLogger(__name__)
        logger.debug("ExcelManager.__aexit__ started (Attached mode: %s)", self._attached_mode)
        was_attached = self._attached_mode # Capture state before cleanup

        try:
            # --- Close the managed Workbook ---
            if self.book:
                book_name = "Unknown"
                try:
                    book_name = self.book.name # Get name before attempting close
                    # Check if the book is still valid/open within the app before trying to close
                    # Need to handle potential app termination before book close
                    if self.app and self.book in self.app.books:
                        logger.info("Closing managed workbook: %s", book_name)
                        # Close without saving changes unless explicitly handled elsewhere (e.g., by save tool)
                        self.book.close(save_changes=False)
                    else:
                        logger.warning("Managed workbook '%s' seems to be already closed or app is unavailable.", book_name)
                except Exception as e:
                    logger.error("Error closing workbook '%s': %s", book_name, e)
                finally:
                    self.book = None # Clear handle regardless
            else:
                logger.debug("__aexit__: No workbook handle to close.")
        finally:
            # --- Quit/Kill the Excel Application ---
            if self.app and not was_attached:
                # Only quit/kill the app if *we* created it
                app_pid = self.app.pid
                logger.info("Quitting Excel instance (PID: %s) as it was created by this manager.", app_pid)
                try:
                    # Use kill when available to prevent zombie COM hosts
                    if hasattr(self.app, "kill"):
                        logger.debug("Using app.kill() for PID: %s", app_pid)
                        self.app.kill()
                    else:
                        logger.debug("Using app.quit() for PID: %s", app_pid)
                        self.app.quit()
                    logger.info("Excel instance (PID: %s) quit/killed.", app_pid)
                except Exception as e:
                    logger.error("Error quitting/killing Excel app (PID: %s): %s", app_pid, e)
                finally:
                    self.app = None # Clear handle regardless
            elif self.app and was_attached:
                logger.info("Leaving attached Excel instance (PID: %s) running.", self.app.pid)
                # Clear the handle, but don't quit the app
                self.app = None
            else:
                # App handle might already be None if creation failed or already cleaned up
                logger.debug("__aexit__: No app handle to clean up or already cleaned.")

            # Reset flag for potential reuse (though usually not done with async with)
            self._attached_mode = False
        logger.debug("ExcelManager.__aexit__ completed.")

    # Optional synchronous helper for legacy callâ€‘sites
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



    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Snapshot / undo helpers
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Ensure changes are applied
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    async def ensure_changes_applied(self) -> None:
        """Asynchronously flush Excel UI and calculation pipelines.

        This method yields to the event loop forÂ â‰ˆ0.5Â s, preventing the hard
        stop caused by ``time.sleep`` while Excel finishes painting.
        """
        logger = logging.getLogger(__name__)
        try:
            # Force a visual and calculation refresh
            self.app.screen_updating = False
            self.app.screen_updating = True
            try:
                self.app.calculate()
            except Exception as calc_err:
                logger.debug(f"self.app.calculate() failed (ignored): {calc_err}")

            # Reâ€‘activate active sheet to nudge UI
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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Explicit save helpers
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  New: open workbook helper
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Basic workbook / sheet info
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
                # Cross‑platform used‑range detection (works on both Windows COM and macOS)
                used = sheet.used_range      # xlwings Range (never None)
                last_cell = used.last_cell   # xlwings Range
                last_addr = last_cell.address.replace("$", "")
                shape.sheets[sheet_name] = f"A1:{last_addr}" if used.value else "A1:A1"

                # Get headers (first row) - handle potential errors/empty rows
                try:
                    # Fast path: fetch first row directly through COM to avoid many Range calls
                    used_range = sheet.api.UsedRange
                    # Bail out completely on extremely wide sheets (≫ token budget & very slow)
                    if used_range.Columns.Count > 2000:
                        logger.debug(f"Sheet '{sheet_name}': Skipping header scan — {used_range.Columns.Count} columns (>2000).")
                        header_values = []
                    else:
                        header_values = used_range.Rows(1).Value2 or []
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

    # The second quick_scan_shape (using sheet.used_range) has been removed.
    # The remaining one (above, using sheet.api.UsedRange) is now the active one.

    def get_sheet(self, sheet_name: str):
        try:
            return self.book.sheets[sheet_name]
        except (KeyError, ValueError):
            return None

    def fill_ranges(self, sheet_name: str, ranges: list[str], color_argb: str) -> None:
        """
        Apply a fill color to all listed ranges in one go,
        to reduce repeated COM calls.
        """
        sheet = self._require_sheet(sheet_name)
        for rng in ranges:
            # Convert 8‑digit ARGB hex (e.g. "FF3366CC") to BGR int for Excel
            sheet.range(rng).color = _hex_argb_to_bgr_int(color_argb)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Cell value helpers
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def set_cell_value(self, sheet_name: str, cell_address: str, value: Any) -> None:

        sheet = self._require_sheet(sheet_name)
        sheet.range(cell_address).value = value

    def get_cell_value(self, sheet_name: str, cell_address: str) -> Any:
        sheet = self._require_sheet(sheet_name)
        return sheet.range(cell_address).value

    def set_cell_values(self, sheet_name: str, data: Dict[str, Any]) -> None:
        sheet = self._require_sheet(sheet_name)
        num_cells = len(data)

        # Optimization: Try vectorized write for rectangular ranges > 1 cell
        if num_cells > 1:
            try:
                coords = []
                min_r, min_c = float('inf'), float('inf')
                max_r, max_c = 0, 0
                for addr in data.keys():
                    col_str, row_idx = coordinate_from_string(addr)
                    col_idx = column_index_from_string(col_str)
                    coords.append({'addr': addr, 'r': row_idx, 'c': col_idx})
                    min_r, max_r = min(min_r, row_idx), max(max_r, row_idx)
                    min_c, max_c = min(min_c, col_idx), max(max_c, col_idx)

                is_rectangular = (num_cells == (max_r - min_r + 1) * (max_c - min_c + 1))

                if is_rectangular:
                    # Build the 2D matrix in the correct order
                    rows_count = max_r - min_r + 1
                    cols_count = max_c - min_c + 1
                    matrix = [[None] * cols_count for _ in range(rows_count)]

                    # Map (row, col) to value
                    coord_map = { (item['r'], item['c']): data[item['addr']] for item in coords }

                    for r_offset in range(rows_count):
                        for c_offset in range(cols_count):
                            current_r = min_r + r_offset
                            current_c = min_c + c_offset
                            matrix[r_offset][c_offset] = coord_map.get((current_r, current_c))

                    start_addr = f"{get_column_letter(min_c)}{min_r}"
                    end_addr = f"{get_column_letter(max_c)}{max_r}"
                    range_address = f"{start_addr}:{end_addr}"

                    logging.debug(f"Using vectorized write for rectangular range: {sheet_name}!{range_address}")
                    sheet.range(range_address).value = matrix
                    return  # Vectorized write successful

            except Exception as e:
                logging.warning(f"Failed to apply vectorized optimization for set_cell_values: {e}. Falling back to iterative write.")

        # Fallback: non-rectangular or error during optimization
        logging.debug(f"Using iterative write for {num_cells} cells in {sheet_name}")
        for addr, val in data.items():
            sheet.range(addr).value = val

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Range helpers
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def get_range_values(self, sheet_name: str, range_address: str) -> List[List[Any]]:
        sheet = self._require_sheet(sheet_name)
        vals = sheet.range(range_address).value
        # xlwings returns scalar for 1Ã—1 range; list for others
        if not isinstance(vals, list):
            return [[vals]]
        # Normalise 1-D row or col to 2-D list-of-lists
        if vals and not isinstance(vals[0], list):
            return [vals]
        return vals

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Styles (minimal viable impl)
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def set_range_style(
        self, sheet_name: str, range_address: str, style: Dict[str, Any]
    ) -> None:
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
                    rng.font.color = _to_bgr(color)
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
                    rng.color = _to_bgr(rgb)
                except Exception as e:
                    print(f"Color application error: {e}")

        # Borders
        if "border" in style and style["border"]:
            try:
                border = style["border"]

                # If 'outline' is set, try using BorderAround as a fallback,
                # because rng.api.Borders(...) can fail on some Mac Excel versions.
                outline_requested = border.get("outline", False) is True

                if outline_requested:
                    # We'll apply an outside border using BorderAround
                    # Default to 'thin' if no style provided
                    border_style = border.get("style", "thin")
                    weight_map = {
                        "thin": 2,       # xlThin
                        "medium": -4138, # xlMedium
                        "thick": 4,      # xlThick
                    }
                    color_hex = border.get("color", "FF000000")  # default black
                    try:
                        rng.api.BorderAround(
                            Weight=weight_map.get(border_style, 2),
                            LineStyle=LineStyle.continuous,
                        )
                        rng.api.Borders.Color = _to_bgr(color_hex)
                    except Exception as e:
                        print(f"BorderAround application error: {e}")
                else:
                    # We'll attempt the Windows COM approach for each edge
                    # but it may fail on Mac. If it fails, we skip partial edges.
                    try:
                        edges = {
                            "left": 7,
                            "right": 10,
                            "top": 8,
                            "bottom": 9,
                        }

                        def _apply_edge(edge_key: str) -> None:
                            edge_style = border.get(edge_key, {})
                            xl_edge = rng.api.Borders(edges[edge_key])
                            xl_edge.LineStyle = LineStyle.continuous
                            weight_map = {
                                "thin": BorderWeight.thin,
                                "medium": BorderWeight.medium,
                                "thick": BorderWeight.thick,
                            }
                            xl_edge.Weight = weight_map.get(
                                str(edge_style.get("style", "thin")).lower(),
                                BorderWeight.thin,
                            )
                            if "color" in edge_style:
                                try:
                                    xl_edge.Color = _to_bgr(edge_style["color"])
                                except Exception:
                                    pass

                        for edge_name in ("left", "right", "top", "bottom"):
                            if edge_name in border:
                                _apply_edge(edge_name)
                    except Exception as e:
                        print(f"Border application error (edge-level): {e}")
            except Exception as e:
                print(f"Border application error: {e}")

        # Alignment
        if "alignment" in style and style["alignment"]:
            alignment = style["alignment"]
            horiz = alignment.get("horizontal")
            if horiz is not None:
                alignment_map = {
                    "left": -4131,     # xlLeft
                    "center": -4108,   # xlCenter
                    "right": -4152,    # xlRight
                    "justify": -4130,
                    "distributed": -4117,
                }
                rng.api.HorizontalAlignment = alignment_map.get(horiz.lower(), -4108)

            vert = alignment.get("vertical")
            if vert is not None:
                vertical_map = {
                    "top": -4160,      # xlTop
                    "center": -4108,   # xlCenter
                    "bottom": -4107,   # xlBottom
                    "justify": -4130,
                    "distributed": -4117,
                }
                rng.api.VerticalAlignment = vertical_map.get(vert.lower(), -4108)

            wrap = alignment.get("wrap_text")
            if wrap is not None:
                rng.api.WrapText = bool(wrap)

        # Force update the Excel application to show changes
        try:
            self.app.screen_updating = False
            self.app.screen_updating = True
        except:
            pass

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Sheet management
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def create_sheet(self, sheet_name: str, index: Optional[int] = None) -> None:
        if sheet_name in self.get_sheet_names():
            raise ValueError(f"Sheet '{sheet_name}' already exists.")
        before = self.book.sheets[index] if index is not None else None
        self.book.sheets.add(name=sheet_name, before=before)

    def delete_sheet(self, sheet_name: str) -> None:
        sheet = self._require_sheet(sheet_name)
        sheet.delete()

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Merge / unmerge
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Row / column sizing
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Copy / Paste range helper
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
        at *dst_anchor* in a single roundâ€‘trip.

        paste_opts:
            â€¢ "values"   â†’ values only
            â€¢ "formulas" â†’ formulas only
            â€¢ "formats"  â†’ formats only
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
            # xlPasteFormats = â€‘4104
            src_rng.api.Copy()
            dst_rng.api.PasteSpecial(Paste=-4104)
        else:
            raise ValueError(
                f"Invalid paste_opts '{paste_opts}'. Use 'values', 'formulas', or 'formats'."
            )

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  (Currently stub) advanced APIs
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def set_cell_formula(self, sheet_name: str, cell_address: str, formula: str) -> None:
        if not formula.startswith("="):
            formula = "=" + formula
        self.set_cell_value(sheet_name, cell_address, formula)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Style inspectors
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def get_cell_style(self, sheet_name: str, cell_address: str) -> Dict[str, Any]:  # noqa: D401
        """Return a minimal style dict (bold + fill color) for a single cell."""
        return _safe_cell_style(self._require_sheet(sheet_name).range(cell_address))

    def get_range_style(self, sheet_name: str, range_address: str) -> Dict[str, Dict[str, Any]]:  # noqa: D401
        """
        Return {cell_address: style_dict} for every cell in the range (minimal style set).
        """
        rng = self._require_sheet(sheet_name).range(range_address)
        return {
            c.address.replace("$", ""): _safe_cell_style(c)
            for c in rng
            if _safe_cell_style(c)
        }

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

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Helpers
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _require_sheet(self, sheet_name: str):
        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        return sheet

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    #  Table insertion
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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


# â•­â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ Helper functions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â•®
def append_table_rows(self, sheet_name: str, table_name: str, rows: List[List[Any]]) -> None:
        """
        Appends rows to an existing Excel table using COM ListRows.Add for incremental inserts.
        """
        sheet = self._require_sheet(sheet_name)
        # Locate the table by name
        tbl = None
        for lo in sheet.api.ListObjects:
            if lo.Name == table_name:
                tbl = lo
                break
        if tbl is None:
            raise KeyError(f"Table '{table_name}' not found on sheet '{sheet_name}'")
        # Append each row to the table
        for row_vals in rows:
            listrow = tbl.ListRows.Add()
            # Write values into the new row
            row_range = listrow.Range
            sheet.range(row_range.Address.replace('$', '')).value = row_vals
        # Refresh calculation if needed
        try:
            self.app.calculate()
        except:
            pass

def _hex_argb_to_bgr_int(argb: str) -> int:
    """
    Convert an **8â€‘digit ARGB** string (``'FFRRGGBB'`` or ``'#FFRRGGBB'``) to an
    integer in BGR byte order for the Excel COM API.

    The function now *requires* the alpha channel; sending a 6â€‘digit RGB code
    raises ``ValueError`` so callers cannot silently lose transparency
    information.
    """
    s = argb.lstrip("#")
    if len(s) != 8:
        raise ValueError(
            f"Color '{argb}' must be 8â€‘digit ARGB (e.g. 'FF3366CC' or '#FF3366CC')."
        )

    # Drop alpha then swap to BGR
    r, g, b = s[2:4], s[4:6], s[6:8]
    return int(f"{b}{g}{r}", 16)


# --------------------------------------------------------------------------
#  Cached colour converter
# --------------------------------------------------------------------------
_COLOR_CACHE: dict[str, int] = {}


def _to_bgr(argb: str) -> int:
    """
    Convert 8‑digit ARGB → BGR int with caching to avoid repeated
    `_hex_argb_to_bgr_int` calls inside tight loops.
    """
    return _COLOR_CACHE.setdefault(argb, _hex_argb_to_bgr_int(argb))


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