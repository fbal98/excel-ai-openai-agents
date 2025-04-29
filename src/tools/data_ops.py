# src/tools/data_ops.py
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult, CellValue, CellValueMap, SetCellValuesResult, WriteVerifyResult
from typing import Any, Optional, List, Dict, TYPE_CHECKING
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string # Moved here
import asyncio # Import asyncio

if TYPE_CHECKING:
    from ..excel_ops import ExcelManager # Avoid circular import

@function_tool
def set_cell_value_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str, value: CellValue) -> ToolResult:
    """
    Sets the value of a single cell. **Use this tool only once per turn; if two+ cells need updates, call `set_cell_values_tool` instead.**

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.
        sheet_name (str): Name of the worksheet.
        cell_address (str): Cell address in A1 notation (e.g., 'B2').
        value (CellValue): The value to set in the cell (text, number, date, boolean, or formula).

    Returns:
        ToolResult: {'success': True} if the cell was updated successfully.
                    {'success': False, 'error': str} if an error occurred.
    """
    print(f"[TOOL] set_cell_value_tool: {sheet_name}!{cell_address} value={value}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_cell_value_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"success": False, "error": "Tool 'set_cell_value_tool' failed: 'cell_address' cannot be empty."}
    # Note: Validating 'value: Any' is complex; rely on underlying function for now.
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_cell_value(sheet_name, cell_address, value)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_cell_value_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_cell_value_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_value_tool: {e}")
        return {"success": False, "error": f"Exception setting cell value for {sheet_name}!{cell_address}: {e}"}

@function_tool
def get_cell_value_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str) -> ToolResult:
    """
    Retrieves the value from a single cell.

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.
        sheet_name (str): Name of the worksheet.
        cell_address (str): Cell address in A1 notation (e.g., 'C3').

    Returns:
        ToolResult: {'success': True, 'data': Any} The cell value (None if empty).
                    {'success': False, 'error': str} if an error occurred.
    """
    print(f"[TOOL] get_cell_value_tool: {sheet_name}!{cell_address}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'get_cell_value_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"success": False, "error": "Tool 'get_cell_value_tool' failed: 'cell_address' cannot be empty."}
    # --- End Validation ---
    try:
        value = ctx.context.excel_manager.get_cell_value(sheet_name, cell_address)
        return {"success": True, "data": value}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_cell_value_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] get_cell_value_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] get_cell_value_tool: {e}")
        return {"success": False, "error": f"Exception getting cell value for {sheet_name}!{cell_address}: {e}"}

@function_tool
def get_range_values_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str) -> ToolResult:
    """
    Retrieves values from a rectangular cell range. 

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.
        sheet_name (str): Name of the worksheet.
        range_address (str): Excel range in A1 notation (e.g., 'A1:C5').

    Returns:
        ToolResult: {'success': True, 'data': List[List[Any]]} 2-D array of values on success.
                    {'success': False, 'error': str} if an error occurred.
    """
    print(f"[TOOL] get_range_values_tool: {sheet_name}!{range_address}")
    # Input validation
    if not sheet_name:
        return {"success": False, "error": "Tool 'get_range_values_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"success": False, "error": "Tool 'get_range_values_tool' failed: 'range_address' cannot be empty."}
    try:
        values = ctx.context.excel_manager.get_range_values(sheet_name, range_address)
        return {"success": True, "data": values}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_range_values_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] get_range_values_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] get_range_values_tool: {e}")
        return {"success": False, "error": f"Exception getting range values for {sheet_name}!{range_address}: {e}"}

@function_tool(strict_mode=False) # Allow flexible dict structure for 'data'
def set_cell_values_tool(ctx: RunContextWrapper[AppContext],
                         sheet_name: str,
                         data: CellValueMap
                         ) -> SetCellValuesResult: # Use specific result type alias if desired
    """
    Sets the values of multiple cells in the specified sheet from a dictionary.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        data: Dictionary mapping cell addresses (e.g., 'A1') to values of any supported Excel type.

    Returns:
        SetCellValuesResult: {'success': True} if successful, or {'success': False, 'error': str} if failed.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_cell_values_tool' failed: 'sheet_name' cannot be empty."}
    if not data: # Check if data dictionary is empty
        print("[INFO] set_cell_values_tool: Received empty data dict – nothing to write.")
        return {
            "success": True,
            "error": None,
            "data": None,
            "hint": "No-op: empty mapping; nothing was written."
        }
    # Optional: Add validation for keys (cell addresses) and values if needed, though manager might handle it.
    # --- End Validation ---
    print(f"[TOOL] set_cell_values_tool: {sheet_name}, {len(data)} cells")
    try:
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
        return {"success": True}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_cell_values_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_cell_values_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_values_tool: {e}")
        return {"success": False, "error": f"Exception setting multiple cell values in '{sheet_name}': {e}"}

@function_tool
def set_table_tool(ctx: RunContextWrapper[AppContext],
                   sheet_name: str,
                   top_left: str,
                   rows: List[List[CellValue]]) -> ToolResult: # Changed Any to CellValue
    """
    Bulk-write a 2-D python list into *sheet_name* starting at *top_left* (e.g. 'A2').
    Treats input purely as data, does not create an Excel Table object.
    Saves potentially many single calls compared to set_cell_value_tool.

    Args:
        sheet_name: Worksheet name.
        top_left: Cell address for top-left corner of the data (e.g., 'A1').
        rows: 2D list of data (list of lists). Can include header as first row.

    Returns:
        ToolResult: {'success': True} on success, or {'success': False, 'error': str} on failure.
    """
    if not sheet_name or not top_left or not rows:
        return {"success": False, "error": "Tool 'set_table_tool' failed: sheet_name, top_left, and non-empty rows are required"}
    # Check if it's a list of lists. Type checking CellValue elements is complex at runtime,
    # rely on the type hint and API validation primarily.
    if not isinstance(rows, list) or not all(isinstance(row, list) for row in rows):
         return {"success": False, "error": "Tool 'set_table_tool' failed: 'rows' must be a list of lists."}

    print(f"[TOOL] set_table_tool: Writing {len(rows)} rows starting at {sheet_name}!{top_left}")
    try:
        col_letter, r0 = coordinate_from_string(top_left)
        c0_idx = column_index_from_string(col_letter)
        data = {
            f"{get_column_letter(c0_idx + c)}{r0 + r}": v
            for r, row in enumerate(rows)
            for c, v in enumerate(row)
        }
        # Use set_cell_values for the actual writing
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
        return {"success": True}
    except ExcelConnectionError as ce: # Catch if set_cell_values fails connection
        print(f"[TOOL ERROR] set_table_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_table_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_table_tool: {e}")
        return {"success": False, "error": f"Bulk write table in '{sheet_name}' starting at '{top_left}' failed: {e}"}

@function_tool
def insert_table_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    start_cell: str,
    columns: List[str], # Expect list of strings for headers
    rows: List[List[CellValue]], # Expect list of lists for data rows
    table_name: Optional[str] = None,
    table_style: Optional[str] = None, # e.g., 'TableStyleMedium2'
) -> ToolResult:
    """
    Inserts data and formats it as a true Excel table (ListObject).

    Args:
        ctx: Agent context containing the ExcelManager.
        sheet_name: Worksheet to insert the table into.
        start_cell: Top-left cell for the table header in A1 notation (e.g., 'A1').
        columns: Header names for each column.
        rows: Data rows matching the headers.
        table_name: Optional name for the Excel table. Auto-generated if None.
        table_style: Optional Excel table style (e.g., 'TableStyleMedium9').

    Returns:
        ToolResult: {'success': True} if table inserted successfully.
                    {'success': False, 'error': str} if an error occurred.
    """
    # Basic input validation
    if not sheet_name:
        return {"success": False, "error": "Tool 'insert_table_tool' failed: 'sheet_name' cannot be empty."}
    if not start_cell:
        return {"success": False, "error": "Tool 'insert_table_tool' failed: 'start_cell' cannot be empty."}
    if not columns:
        return {"success": False, "error": "Tool 'insert_table_tool' failed: 'columns' list cannot be empty."}
    # Rows can be empty, resulting in a table with only a header row.

    print(f"[TOOL] insert_table_tool: Into {sheet_name}!{start_cell} Name='{table_name}' Style='{table_style}' ({len(columns)} cols, {len(rows)} rows)")
    try:
        ctx.context.excel_manager.insert_table(sheet_name, start_cell, columns, rows, table_name, table_style)
        return {"success": True}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] insert_table_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] insert_table_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] insert_table_tool: {e}")
        # Make error more specific if possible (e.g., catch specific COM errors if manager raises them)
        return {"success": False, "error": f"Failed to insert table '{table_name}': {e}"}

@function_tool
def set_rows_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, start_row: int, rows: List[List[CellValue]]) -> ToolResult:
    """
    Writes a 2-D Python list *rows* into *sheet_name* beginning at **column A**
    and **start_row**. Each inner list is written to consecutive columns.

    Args:
        sheet_name (str): Worksheet name.
        start_row  (int): First row number (1-based).
        rows (List[List[CellValue]]): List of rows to write.

    Returns:
        ToolResult: {'success': True} on success, or {'success': False, 'error': str} on failure.
    """
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_rows_tool' failed: 'sheet_name' cannot be empty."}
    if not isinstance(start_row, int) or start_row <= 0:
        return {"success": False, "error": "Tool 'set_rows_tool' failed: 'start_row' must be a positive integer."}
    if not rows:
        return {"success": False, "error": "Tool 'set_rows_tool' failed: 'rows' cannot be empty."}
    if not isinstance(rows, list) or not all(isinstance(row, list) for row in rows):
         return {"success": False, "error": "Tool 'set_rows_tool' failed: 'rows' must be a list of lists."}

    print(f"[TOOL] set_rows_tool: Writing {len(rows)} rows starting at {sheet_name}!A{start_row}")
    try:
        data = {}
        for r_idx, row_vals in enumerate(rows):
            row_no = start_row + r_idx
            for c_idx, val in enumerate(row_vals):
                addr = f"{get_column_letter(c_idx + 1)}{row_no}" # Start at column A (index 1)
                data[addr] = val
        # Use set_cell_values for the actual writing
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
        return {"success": True}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_rows_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_rows_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_rows_tool: {e}")
        return {"success": False, "error": f"Bulk write rows in '{sheet_name}' starting at row {start_row} failed: {e}"}

@function_tool
def set_columns_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, start_col: str, cols: List[List[CellValue]]) -> ToolResult:
    """
    Writes a 2-D Python list *cols* into *sheet_name* beginning at **row 1**
    and **start_col** (column letter). Each inner list becomes a column written downward.

    Args:
        sheet_name (str): Worksheet name.
        start_col  (str): Column letter where the first column should be written (e.g., 'A').
        cols (List[List[CellValue]]): List of columns; each inner list contains values for one column.

    Returns:
        ToolResult: {'success': True} on success, or {'success': False, 'error': str} on failure.
    """
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_columns_tool' failed: 'sheet_name' cannot be empty."}
    if not start_col or not isinstance(start_col, str):
        return {"success": False, "error": "Tool 'set_columns_tool' failed: 'start_col' must be a column letter."}
    if not cols:
        return {"success": False, "error": "Tool 'set_columns_tool' failed: 'cols' cannot be empty."}
    if not isinstance(cols, list) or not all(isinstance(col, list) for col in cols):
         return {"success": False, "error": "Tool 'set_columns_tool' failed: 'cols' must be a list of lists."}

    print(f"[TOOL] set_columns_tool: Writing {len(cols)} columns starting at {sheet_name}!{start_col}1")
    try:
        c0_idx = column_index_from_string(start_col.upper())
        data = {}
        for c_idx, col_vals in enumerate(cols):
            col_letter = get_column_letter(c0_idx + c_idx)
            for r_idx, val in enumerate(col_vals):
                addr = f"{col_letter}{r_idx + 1}" # Start at row 1
                data[addr] = val
        # Use set_cell_values for the actual writing
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
        return {"success": True}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_columns_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_columns_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_columns_tool: {e}")
        return {"success": False, "error": f"Bulk write columns in '{sheet_name}' starting at column {start_col} failed: {e}"}

@function_tool(strict_mode=False) # Allow flexible dict for rows
def append_table_rows_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    table_name: str,
    rows: List[List[CellValue]],
) -> ToolResult:
    """
    Appends one or more *rows* below the last row of the specified Excel table *table_name* on *sheet_name*.
    Uses `ListRows.Add` if possible, which helps maintain table formatting and formulas.

    Args:
        sheet_name: The name of the sheet containing the table.
        table_name: The name of the existing Excel table (ListObject).
        rows: A list of rows to append. Each row is a list of cell values.

    Returns:
        ToolResult: {'success': True} if rows appended successfully.
                    {'success': False, 'error': str} if an error occurred (e.g., table not found).
    """
    print(f"[TOOL] append_table_rows_tool: {sheet_name}!{table_name} (+{len(rows)} rows)")
    if not sheet_name:
        return {"success": False, "error": "Tool 'append_table_rows_tool' failed: 'sheet_name' cannot be empty."}
    if not table_name:
        return {"success": False, "error": "Tool 'append_table_rows_tool' failed: 'table_name' cannot be empty."}
    if not rows:
        print("[INFO] append_table_rows_tool: No rows provided to append.")
        return {"success": True, "data": "No rows provided to append."}  # Success, but did nothing.
    if not isinstance(rows, list) or not all(isinstance(row, list) for row in rows):
        return {"success": False, "error": "Tool 'append_table_rows_tool' failed: 'rows' must be a list of lists."}

    try:
        ctx.context.excel_manager.append_table_rows(sheet_name, table_name, rows)
        return {"success": True}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] append_table_rows_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke:  # Catch table/sheet not found specifically
        print(f"[TOOL ERROR] append_table_rows_tool: Table or Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] append_table_rows_tool: {e}")
        return {"success": False, "error": f"Exception appending rows to table '{table_name}' on sheet '{sheet_name}': {e}"}


@function_tool(strict_mode=False) # Allow flexible dict structure for 'data'
def write_and_verify_range_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    data: Dict[str, CellValue], # More specific type hint for values
) -> WriteVerifyResult:
    """
    Writes multiple cells and verifies the write by reading back the values immediately.

    Args:
        ctx: Agent context containing the ExcelManager.
        sheet_name: Name of the worksheet.
        data: Mapping of cell addresses (e.g., 'A1') to expected values.

    Returns:
        WriteVerifyResult: {'success': True} if all values match after writing.
                         {'success': False, 'diff': dict} where 'diff' maps each mismatched cell
                         to {'expected': Any, 'actual': Any} or provides an error message under the 'error' key within diff.
    """
    if not sheet_name:
        return {"success": False, "diff": {"error": "'sheet_name' cannot be empty."}}
    if not data:
        return {"success": False, "diff": {"error": "'data' dictionary cannot be empty."}}

    print(f"[TOOL] write_and_verify_range_tool: Writing and verifying {len(data)} cells in {sheet_name}")

    # 1. Write
    try:
        # Use set_cell_values which handles connection checks and potential optimizations
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] write_and_verify_range_tool (Write Phase): Connection Error - {ce}")
        return {"success": False, "diff": {"error": f"Write failed (Connection Error): {ce}"}}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] write_and_verify_range_tool (Write Phase): Sheet not found - {ke}")
        return {"success": False, "diff": {"error": f"Write failed (Sheet not found): {ke}"}}
    except Exception as e:
        print(f"[TOOL ERROR] write_and_verify_range_tool (Write Phase): {e}")
        return {"success": False, "diff": {"error": f"Write failed: {e}"}}

    # 2. Verify
    diff: Dict[str, Any] = {}
    for addr, expected in data.items():
        try:
            # Use get_cell_value which handles connection checks
            actual = ctx.context.excel_manager.get_cell_value(sheet_name, addr)
            # Perform comparison (handle potential type differences if necessary, e.g., float precision)
            # Simple comparison for now:
            if actual != expected:
                print(f"[VERIFY FAIL] {sheet_name}!{addr}: Expected '{expected}' ({type(expected).__name__}), got '{actual}' ({type(actual).__name__})")
                diff[addr] = {"expected": expected, "actual": actual}
        except ExcelConnectionError as ce:
             print(f"[TOOL ERROR] write_and_verify_range_tool (Verify Phase): Connection Error reading {addr} - {ce}")
             diff[addr] = {"expected": expected, "actual": f"(read-error: Connection Error: {ce})"}
             # Optionally stop verification on connection error? For now, log and continue verifying others.
        except KeyError: # Sheet not found during verify (shouldn't happen if write succeeded, but safety check)
            print(f"[TOOL ERROR] write_and_verify_range_tool (Verify Phase): Sheet '{sheet_name}' not found reading {addr}")
            diff[addr] = {"expected": expected, "actual": "(read-error: Sheet not found)"}
        except Exception as e:
            print(f"[TOOL ERROR] write_and_verify_range_tool (Verify Phase): Error reading {addr} - {e}")
            diff[addr] = {"expected": expected, "actual": f"(read-error: {e})"}

    if diff:
        print(f"[TOOL] write_and_verify_range_tool: Verification failed for {len(diff)} cells.")
        return {"success": False, "diff": diff}

    print(f"[TOOL] write_and_verify_range_tool: Verification successful for {len(data)} cells.")
    return {"success": True}


@function_tool
def find_row_by_value_tool(ctx: RunContextWrapper[AppContext],
                           sheet_name: str,
                           column_letter: str,
                           value: CellValue) -> ToolResult:
    """
    Finds the first 1-based row index in a specific column that matches the given value.
    Search is case-insensitive for strings.

    Args:
        sheet_name (str): Worksheet to search.
        column_letter (str): Column to scan (e.g. "A").
        value (CellValue): Value to look for (text, number, bool, etc.).

    Returns:
        ToolResult: {'success': True, 'data': int} Row number (1-based) if found, or 0 if not found.
                    {'success': False, 'error': str} if an error occurred.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'find_row_by_value_tool' failed: 'sheet_name' cannot be empty."}
    if not column_letter or not isinstance(column_letter, str):
        return {"success": False, "error": "Tool 'find_row_by_value_tool' failed: 'column_letter' must be a non-empty string."}
    # Value can be None, so no specific check here
    # --- End Validation ---

    print(f"[TOOL] find_row_by_value_tool: Searching {sheet_name}!{column_letter} for value '{value}'")
    try:
        # Define the range for the entire column (potential performance issue on huge sheets)
        # Consider limiting the scan range if performance becomes an issue, e.g., sheet.used_range.last_cell.row
        # full_column_range = f"{column_letter}:{column_letter}" # This might be slow
        # Alternative: Get used range and calculate column range within it
        sheet = ctx.context.excel_manager._require_sheet(sheet_name) # Get sheet object first
        used_range = sheet.used_range
        last_row = used_range.last_cell.row
        scan_range = f"{column_letter}1:{column_letter}{last_row}"
        print(f"[DEBUG] Scanning range {scan_range} for find_row_by_value")

        # Get values for the calculated range
        col_values_2d = ctx.context.excel_manager.get_range_values(sheet_name, scan_range)

        # Flatten the 2D list returned by get_range_values (it's a single column)
        col_vals = [row[0] if row else None for row in col_values_2d]

        search_value_str = str(value).strip().lower() if value is not None else ""

        for idx, cell_value in enumerate(col_vals):
            # Normalize cell value for comparison
            current_value_str = str(cell_value).strip().lower() if cell_value is not None else ""
            if current_value_str == search_value_str:
                found_row = idx + 1 # Convert 0-based index to 1-based row number
                print(f"[TOOL] find_row_by_value_tool: Found value '{value}' at row {found_row}")
                return {"success": True, "data": found_row}

        print(f"[TOOL] find_row_by_value_tool: Value '{value}' not found in column {column_letter}.")
        return {"success": True, "data": 0} # Return 0 if not found

    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] find_row_by_value_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] find_row_by_value_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] find_row_by_value_tool: {e}")
        return {"success": False, "error": f"Failed to find row in {sheet_name}!{column_letter}: {e}"}