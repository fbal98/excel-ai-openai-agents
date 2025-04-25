# src/tools/style_ops.py
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult, CellStyle # Import result and style types
from typing import Any, Optional, Dict, List

# The SDK automatically handles JSON conversion from LLM to the Pydantic/TypedDict model (CellStyle).
@function_tool(strict_mode=False) # Allow CellStyle dict to have missing optional keys
def set_range_style_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str, style: CellStyle) -> ToolResult:
    """
    Applies formatting styles (font, fill, border, alignment, number_format) to a cell range.

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.
        sheet_name (str): Name of the worksheet.
        range_address (str): Excel range in A1 notation (e.g., 'A1:B2').
        style (CellStyle): Dictionary defining style attributes. See CellStyle definition.

    Returns:
        ToolResult: {'success': True} if styles applied successfully.
                    {'success': False, 'error': str} if an error occurred.
    """
    print(f"[TOOL] set_range_style_tool: {sheet_name}!{range_address} style={style}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_range_style_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"success": False, "error": "Tool 'set_range_style_tool' failed: 'range_address' cannot be empty."}
    if not style or not isinstance(style, dict): # Check if the style dictionary itself is empty or not a dict
        return {"success": False, "error": "Tool 'set_range_style_tool' failed: 'style' must be a non-empty dictionary."}
    # --- End Validation ---
    try:
        # Delegate to ExcelManager which handles applying sub-styles (font, fill etc.)
        ctx.context.excel_manager.set_range_style(sheet_name, range_address, style)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_range_style_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_range_style_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_range_style_tool: {e}")
        # Include more detail in error if possible (e.g., invalid color format caught by manager)
        return {"success": False, "error": f"Exception applying style to {sheet_name}!{range_address}: {e}"}

@function_tool(strict_mode=False) # Allow CellStyle dict to have missing optional keys
def set_cell_style_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    cell_address: str,
    style: CellStyle,
) -> ToolResult:
    """
    Applies formatting styles (font, fill, border, alignment, number_format) to a single cell.

    Args:
        ctx: Agent context.
        sheet_name: Worksheet name.
        cell_address: Cell address (e.g. "B4").
        style: Dictionary defining style attributes. See CellStyle definition.

    Returns
        ToolResult: {'success': True} on success, or {'success': False, 'error': str} on failure.
    """
    print(f"[TOOL] set_cell_style_tool: {sheet_name}!{cell_address} style={style}")
    # Basic validation
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_cell_style_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"success": False, "error": "Tool 'set_cell_style_tool' failed: 'cell_address' cannot be empty."}
    if not style or not isinstance(style, dict):
        return {"success": False, "error": "Tool 'set_cell_style_tool' failed: 'style' must be a non-empty dictionary."}
    try:
        # Re‑use the existing range‑style helper for a single‑cell address
        ctx.context.excel_manager.set_range_style(sheet_name, cell_address, style)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_cell_style_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_cell_style_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_style_tool: {e}")
        return {"success": False, "error": f"Exception applying cell style to {sheet_name}!{cell_address}: {e}"}

@function_tool
def get_cell_style_tool(
    ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str
) -> ToolResult:
    """
    Return the style dict (currently font bold + fill color) for a single cell.

    Returns:
        ToolResult: {'success': True, 'data': CellStyle} or {'success': False, 'error': str}
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'get_cell_style_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"success": False, "error": "Tool 'get_cell_style_tool' failed: 'cell_address' cannot be empty."}
    # --- End Validation ---
    print(f"[TOOL] get_cell_style_tool: {sheet_name}!{cell_address}")
    try:
        style_dict = ctx.context.excel_manager.get_cell_style(sheet_name, cell_address)
        return {"success": True, "data": style_dict}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_cell_style_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] get_cell_style_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] get_cell_style_tool: {e}")
        return {"success": False, "error": f"Failed to get cell style for {sheet_name}!{cell_address}: {e}"}

@function_tool
def get_range_style_tool(
    ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str
) -> ToolResult:
    """
    Return a mapping of cell_address -> style_dict for cells with non-default
    styles (currently font bold + fill color) within a rectangular range.

    Returns:
        ToolResult: {'success': True, 'data': Dict[str, CellStyle]} or {'success': False, 'error': str}
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'get_range_style_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"success": False, "error": "Tool 'get_range_style_tool' failed: 'range_address' cannot be empty."}
    # --- End Validation ---
    print(f"[TOOL] get_range_style_tool: {sheet_name}!{range_address}")
    try:
        styles_map = ctx.context.excel_manager.get_range_style(sheet_name, range_address)
        return {"success": True, "data": styles_map}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_range_style_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] get_range_style_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] get_range_style_tool: {e}")
        return {"success": False, "error": f"Failed to get range style for {sheet_name}!{range_address}: {e}"}


@function_tool
def merge_cells_range_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str) -> ToolResult:
    """
    Merges a range of cells in the specified sheet.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        range_address: Range to merge (e.g., 'A1:B2').

    Returns:
        ToolResult: {'success': True} if successful, or {'success': False, 'error': str} on failure.
    """
    print(f"[TOOL] merge_cells_range_tool: {sheet_name}!{range_address}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'merge_cells_range_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"success": False, "error": "Tool 'merge_cells_range_tool' failed: 'range_address' cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.merge_cells_range(sheet_name, range_address)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] merge_cells_range_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] merge_cells_range_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] merge_cells_range_tool: {e}")
        return {"success": False, "error": f"Exception merging cells {sheet_name}!{range_address}: {e}"}

@function_tool
def unmerge_cells_range_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str) -> ToolResult:
    """
    Unmerges any merged cells within the specified range.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        range_address: Range potentially containing merged cells (e.g., 'A1:B2').

    Returns:
        ToolResult: {'success': True} if successful, or {'success': False, 'error': str} on failure.
    """
    print(f"[TOOL] unmerge_cells_range_tool: {sheet_name}!{range_address}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'unmerge_cells_range_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"success": False, "error": "Tool 'unmerge_cells_range_tool' failed: 'range_address' cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.unmerge_cells_range(sheet_name, range_address)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] unmerge_cells_range_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] unmerge_cells_range_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] unmerge_cells_range_tool: {e}")
        return {"success": False, "error": f"Exception unmerging cells {sheet_name}!{range_address}: {e}"}


@function_tool
def set_row_height_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, row_number: int, height: Optional[float]) -> ToolResult:
    """
    Sets the height of a specific row in the specified sheet. Set height to None to autofit.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        row_number: Row number (1-based).
        height: Height in points, or None to autofit. Negative/zero height hides the row.

    Returns:
        ToolResult: {'success': True} if successful, or {'success': False, 'error': str} on failure.
    """
    print(f"[TOOL] set_row_height_tool: {sheet_name} row {row_number} height={height}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_row_height_tool' failed: 'sheet_name' cannot be empty."}
    if not isinstance(row_number, int) or row_number <= 0:
        return {"success": False, "error": f"Tool 'set_row_height_tool' failed: 'row_number' must be a positive integer (got {row_number})."}
    if height is not None and not isinstance(height, (int, float)):
         return {"success": False, "error": f"Tool 'set_row_height_tool' failed: 'height' must be a number or None (got {type(height)})."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_row_height(sheet_name, row_number, height)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_row_height_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_row_height_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_row_height_tool: {e}")
        return {"success": False, "error": f"Exception setting row height for row {row_number} in '{sheet_name}': {e}"}

@function_tool
def set_column_width_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, column_letter: str, width: Optional[float]) -> ToolResult:
    """
    Sets the width of a specific column in the specified sheet. Set width to None to autofit.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        column_letter: Column letter (e.g., 'A', 'AB').
        width: Width in characters, or None to autofit. Negative/zero width might hide column (implementation dependent).

    Returns:
        ToolResult: {'success': True} if successful, or {'success': False, 'error': str} on failure.
    """
    print(f"[TOOL] set_column_width_tool: {sheet_name} column {column_letter} width={width}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_column_width_tool' failed: 'sheet_name' cannot be empty."}
    if not column_letter or not isinstance(column_letter, str):
        return {"success": False, "error": "Tool 'set_column_width_tool' failed: 'column_letter' must be a non-empty string."}
    if width is not None and not isinstance(width, (int, float)):
        return {"success": False, "error": f"Tool 'set_column_width_tool' failed: 'width' must be a number or None (got {type(width)})."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_column_width(sheet_name, column_letter.upper(), width)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_column_width_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_column_width_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_column_width_tool: {e}")
        return {"success": False, "error": f"Exception setting column width for column {column_letter.upper()} in '{sheet_name}': {e}"}

@function_tool(strict_mode=False) # Allow flexible dict for widths
def set_columns_widths_tool(ctx: RunContextWrapper[AppContext],
                             sheet_name: str,
                             widths: Dict[str, Optional[float]]) -> ToolResult:
    """
    Set multiple column widths in one call. Maps column letters to widths (float) or None (autofit).

    Args:
        sheet_name: The name of the sheet.
        widths: A dictionary mapping column letters (e.g., "A", "BC") to desired widths or None for autofit.

    Returns:
        ToolResult: {'success': True} if all widths set successfully.
                    {'success': False, 'error': str} if any width setting failed.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_columns_widths_tool' failed: 'sheet_name' cannot be empty."}
    if not widths or not isinstance(widths, dict):
        return {"success": False, "error": "Tool 'set_columns_widths_tool' failed: 'widths' must be a non-empty dictionary."}
    # --- End Validation ---
    print(f"[TOOL] set_columns_widths_tool: Setting {len(widths)} column widths in {sheet_name}")
    errors = []
    try:
        # Check connection once before the loop
        ctx.context.excel_manager._validate_connection() # Let it raise ExcelConnectionError if needed

        for col, w in widths.items():
            try:
                # Use the single column width tool internally
                ctx.context.excel_manager.set_column_width(sheet_name, col, w)
            except Exception as e:
                error_msg = f"Failed for column '{col}': {e}"
                print(f"[TOOL ERROR] set_columns_widths_tool: {error_msg}")
                errors.append(error_msg)
                # Continue trying other columns

        if errors:
            return {"success": False, "error": f"Some column widths failed to set: {'; '.join(errors)}"}
        else:
            return {"success": True}

    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_columns_widths_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from initial connection check
        print(f"[TOOL ERROR] set_columns_widths_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e: # Catch unexpected errors during the process
        print(f"[TOOL ERROR] set_columns_widths_tool: {e}")
        return {"success": False, "error": f"Exception setting column widths in '{sheet_name}': {e}"}