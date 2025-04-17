"""
Excel operations module for the Autonomous Excel Assistant.
Provides the ExcelManager class for manipulating Excel files using openpyxl.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet
from typing import Optional, Any, Dict, Union
import json

class ExcelManager:
    """
    Manages Excel workbook operations using openpyxl.
    """
    def __init__(self, file_path: Optional[str] = None):
        """
        Initialize the ExcelManager.
        If file_path is provided, load the workbook. Otherwise, create a new workbook.
        """
        try:
            if file_path:
                self.workbook = openpyxl.load_workbook(file_path)
            else:
                self.workbook = Workbook()
        except FileNotFoundError:
            raise FileNotFoundError(f"File not found: {file_path}")
        except Exception as e:
            raise RuntimeError(f"Error loading workbook: {e}")

    def save_workbook(self, file_path: str):
        """
        Save the workbook to the specified file path.
        Raises:
            ValueError: If file_path is empty.
            IOError: If saving fails.
        """
        if not file_path:
            raise ValueError("File path cannot be empty.")
        try:
            self.workbook.save(file_path)
            print(f"Workbook saved to {file_path}") # Keep print for confirmation
        except Exception as e:
            # Raise a more specific error for file operations
            raise IOError(f"Error saving workbook to '{file_path}': {e}")

    def get_sheet_names(self) -> list:
        """
        Return a list of sheet names in the workbook.
        """
        return self.workbook.sheetnames

    def get_active_sheet_name(self) -> str:
        """
        Return the name of the active sheet.
        """
        return self.workbook.active.title

    def get_sheet(self, sheet_name: str) -> Optional[Worksheet]:
        """
        Get a worksheet by name. Returns None if not found.
        """
        try:
            return self.workbook[sheet_name]
        except KeyError:
            return None

    def set_cell_value(self, sheet_name: str, cell_address: str, value: Any):
        """
        Set the value of a cell in the specified sheet.
        Raises:
            ValueError: If sheet_name or cell_address is empty.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not cell_address:
            raise ValueError("Cell address cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        try:
            sheet[cell_address].value = value
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error setting value for {sheet_name}!{cell_address}: {e}")

    def get_cell_value(self, sheet_name: str, cell_address: str) -> Any:
        """
        Get the value of a cell in the specified sheet.
        Returns the cell value, or None if the cell is empty.
        Raises:
            ValueError: If sheet_name or cell_address is empty.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not cell_address:
            raise ValueError("Cell address cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        try:
            # Return the value directly, None is valid for empty cells
            return sheet[cell_address].value
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error getting value for {sheet_name}!{cell_address}: {e}")

    def set_range_style(self, sheet_name: str, range_address: str, style_dict: Dict[str, Any]):
        """
        Apply styles to a range of cells or a single cell based on a style dictionary.
        style_dict: Dictionary adhering to CellStyle structure (keys: 'font', 'fill', 'border').
        Raises:
            ValueError: If sheet_name, range_address, or style_dict is empty/invalid.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not range_address:
            raise ValueError("Range address cannot be empty.")
        if not style_dict:
             raise ValueError("Style dictionary cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")

        try:
            # Create style objects only if they are defined in the style dict
            # Use .get() with default {} to avoid errors if key is missing
            font = Font(**style_dict.get('font', {})) if style_dict.get('font') else None
            fill = PatternFill(**style_dict.get('fill', {})) if style_dict.get('fill') else None
            border = Border(**style_dict.get('border', {})) if style_dict.get('border') else None

            target_cells = sheet[range_address]

            # Apply styles efficiently
            if isinstance(target_cells, openpyxl.cell.cell.Cell):
                # Single cell
                if font: target_cells.font = font
                if fill: target_cells.fill = fill
                if border: target_cells.border = border
            else:
                # Range
                for row in target_cells:
                    for cell in row:
                        if font: cell.font = font
                        if fill: cell.fill = fill
                        if border: cell.border = border
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error applying style to {sheet_name}!{range_address}: {e}")

    def create_sheet(self, sheet_name: str, index: Optional[int] = None):
        """
        Create a new sheet with the given name and optional index.
        Raises:
            ValueError: If sheet_name is empty.
            Exception: For other openpyxl errors (e.g., duplicate sheet name).
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        try:
            self.workbook.create_sheet(title=sheet_name, index=index)
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error creating sheet '{sheet_name}': {e}")

    def delete_sheet(self, sheet_name: str):
        """
        Delete the specified sheet from the workbook.
        Raises:
            ValueError: If sheet_name is empty.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        try:
            # This will raise KeyError if sheet doesn't exist
            sheet = self.workbook[sheet_name]
            self.workbook.remove(sheet)
        except KeyError:
            # Re-raise specifically for the tool
            raise KeyError(f"Sheet '{sheet_name}' not found for deletion.")
        except Exception as e:
            # Re-raise other exceptions
            raise Exception(f"Error deleting sheet '{sheet_name}': {e}")

    def merge_cells_range(self, sheet_name: str, range_address: str):
        """
        Merge a range of cells in the specified sheet.
        Raises:
            ValueError: If sheet_name or range_address is empty.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors (e.g., invalid range).
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not range_address:
            raise ValueError("Range address cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        try:
            sheet.merge_cells(range_address)
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error merging cells {sheet_name}!{range_address}: {e}")

    def unmerge_cells_range(self, sheet_name: str, range_address: str):
        """
        Unmerge a range of cells in the specified sheet.
        Raises:
            ValueError: If sheet_name or range_address is empty.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors (e.g., invalid range).
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not range_address:
            raise ValueError("Range address cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        try:
            sheet.unmerge_cells(range_address)
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error unmerging cells {sheet_name}!{range_address}: {e}")

    def set_row_height(self, sheet_name: str, row_number: int, height: float):
        """
        Set the height of a row in the specified sheet.
        Raises:
            ValueError: If sheet_name is empty, row_number is not positive, or height is negative.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not isinstance(row_number, int) or row_number <= 0:
            raise ValueError(f"Row number must be a positive integer (got {row_number}).")
        if not isinstance(height, (int, float)) or height < 0:
            raise ValueError(f"Height must be a non-negative number (got {height}).")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        try:
            sheet.row_dimensions[row_number].height = height
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error setting row height for row {row_number} in '{sheet_name}': {e}")

    def set_column_width(self, sheet_name: str, column_letter: str, width: float):
        """
        Set the width of a column in the specified sheet.
        Raises:
            ValueError: If sheet_name or column_letter is empty, or width is negative.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not column_letter or not isinstance(column_letter, str):
             raise ValueError("Column letter must be a non-empty string.")
        if not isinstance(width, (int, float)) or width < 0:
            raise ValueError(f"Width must be a non-negative number (got {width}).")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        try:
            # Ensure column letter is uppercase for consistency
            sheet.column_dimensions[column_letter.upper()].width = width
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error setting column width for column {column_letter.upper()} in '{sheet_name}': {e}")

    def set_cell_formula(self, sheet_name: str, cell_address: str, formula: str):
        """
        Set a formula in the specified cell.
        Raises:
            ValueError: If sheet_name, cell_address, or formula is empty.
            KeyError: If sheet_name does not exist.
            Exception: For other openpyxl errors.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not cell_address:
            raise ValueError("Cell address cannot be empty.")
        if not formula:
            raise ValueError("Formula cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")

        # Ensure formula starts with '='
        if not formula.startswith('='):
            formula = '=' + formula
        try:
            sheet[cell_address].value = formula
        except Exception as e:
            # Re-raise exception for the tool to handle
            raise Exception(f"Error setting formula for {sheet_name}!{cell_address}: {e}")

    def set_cell_values(self, sheet_name: str, data: Dict[str, Any]):
        """
        Set the values of multiple cells in the specified sheet from a dictionary.
        Keys are cell addresses (e.g., 'A1'), values are the data to set.
        Raises:
            ValueError: If sheet_name is empty or data dictionary is empty.
            KeyError: If sheet_name does not exist.
            Exception: If any cell address is invalid or another error occurs during setting values.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not data:
            raise ValueError("Data dictionary cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if not sheet:
            raise KeyError(f"Sheet '{sheet_name}' not found.")

        errors = {}
        try:
            for cell_address, value in data.items():
                if not cell_address: # Basic check for empty cell address key
                    errors["(empty_key)"] = "Cell address cannot be empty."
                    continue
                try:
                    sheet[cell_address].value = value
                except Exception as cell_e:
                    errors[cell_address] = str(cell_e) # Collect errors per cell

            if errors:
                # Raise a single exception summarizing the cell-specific errors
                raise Exception(f"Errors setting multiple cell values in '{sheet_name}': {errors}")

        except Exception as e:
            # Catch broader exceptions (like invalid sheet) or re-raise the aggregated error
            if errors: # If we already aggregated errors, raise that
                 raise Exception(f"Errors setting multiple cell values in '{sheet_name}': {errors}")
            else: # Otherwise, raise the general exception
                raise Exception(f"Error setting multiple cell values in '{sheet_name}': {e}")

    # ------------------------------------------------------------------ #
    #  New inspectors for style verification                              #
    # ------------------------------------------------------------------ #

    def get_cell_style(self, sheet_name: str, cell_address: str) -> Dict[str, Any]:
        """
        Return the font / fill / border style of a single cell as a serialisable
        dictionary. Keys that are None are omitted to keep the payload compact.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not cell_address:
            raise ValueError("Cell address cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")

        cell = sheet[cell_address]

        def _clean(d: Dict[str, Any]) -> Dict[str, Any]:
            """Return a copy of *d* with all None values removed."""
            return {k: v for k, v in d.items() if v is not None}

        font_dict = _clean(
            {
                "name": cell.font.name,
                "size": cell.font.sz,
                "bold": cell.font.bold,
                "italic": cell.font.italic,
                "underline": cell.font.underline,
                "strike": cell.font.strike,
                "color": cell.font.color.rgb if cell.font.color else None,
            }
        )

        fill_dict = _clean(
            {
                "fill_type": cell.fill.fill_type,
                "start_color": (
                    cell.fill.start_color.rgb if hasattr(cell.fill.start_color, "rgb") else None
                ),
                "end_color": (
                    cell.fill.end_color.rgb if hasattr(cell.fill.end_color, "rgb") else None
                ),
            }
        )

        border = cell.border
        border_dict = _clean(
            {
                "left": border.left.style,
                "right": border.right.style,
                "top": border.top.style,
                "bottom": border.bottom.style,
                "outline": border.outline,
                "vertical": border.vertical.style,
                "horizontal": border.horizontal.style,
            }
        )

        return {"font": font_dict, "fill": fill_dict, "border": border_dict}

    def get_range_style(self, sheet_name: str, range_address: str) -> Dict[str, Dict[str, Any]]:
        """
        Return a mapping of cell_address -> style_dict for every cell in the range.
        """
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        if not range_address:
            raise ValueError("Range address cannot be empty.")

        sheet = self.get_sheet(sheet_name)
        if sheet is None:
            raise KeyError(f"Sheet '{sheet_name}' not found.")

        styles: Dict[str, Dict[str, Any]] = {}
        for row in sheet[range_address]:
            # openpyxl returns tuples; iterate over individual cells
            for cell in (row if hasattr(row, "__iter__") else (row,)):
                address = cell.coordinate
                styles[address] = self.get_cell_style(sheet_name, address)

        return styles