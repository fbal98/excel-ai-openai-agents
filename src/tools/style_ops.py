# src/tools/style_ops.py
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult, CellStyle # Import result and style types
from typing import Any, Optional, Dict, List

# The SDK automatically handles JSON conversion from LLM to the Pydantic/TypedDict model (CellStyle).
@function_tool(strict_mode=False) # Allow CellStyle dict to have missing optional keys
def set_range_style_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str, style: CellStyle) -> ToolResult:
    """Applies various formatting styles to all cells within a specified range.

    Sets formatting properties like font (name, size, bold, italic, color),
    fill (background color, pattern), borders (style, color), alignment
    (horizontal, vertical, wrap text), and number format for every cell
    within the given `range_address` on the `sheet_name`.

    The `style` argument is a dictionary conforming to the `CellStyle` structure,
    allowing specification of multiple style aspects in one call. Only the provided
    style attributes are modified; unspecified attributes remain unchanged.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        range_address: The cell range in A1 notation (e.g., 'A1:B10', 'C2') to apply the style to.
        style: A dictionary specifying the desired style attributes (font, fill,
               border, alignment, number_format). Refer to `CellStyle` definition for details.
               Can contain subsets of style attributes (e.g., only font bold and fill color).

    Returns:
        ToolResult: {'success': True} if the styles were applied successfully.
                    {'success': False, 'error': str} if an error occurred (e.g., sheet not found, invalid style format, connection error).
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
    """Applies various formatting styles to a single specified cell.

    This is a convenience tool that applies styles (font, fill, border, etc.)
    to a single cell identified by `cell_address` using the same mechanism as
    `set_range_style_tool`.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        cell_address: The address of the single cell to style (e.g., 'B4', 'A1').
        style: A dictionary specifying the desired style attributes (font, fill,
               border, alignment, number_format). Refer to `CellStyle` definition.

    Returns:
        ToolResult: {'success': True} if the style was applied successfully.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid style format, connection error).
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
    """Retrieves the formatting style attributes of a single specified cell.

    Reads the style properties (like font, fill, border, alignment, number format)
    of the cell identified by `cell_address` on the `sheet_name`.

    Note: The specific style attributes retrieved might evolve. Currently focuses on commonly used ones.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        cell_address: The address of the single cell whose style is needed (e.g., 'B4').

    Returns:
        ToolResult: {'success': True, 'data': CellStyle} where 'data' is a dictionary
                    conforming to the `CellStyle` structure, containing the retrieved
                    style attributes of the cell.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, connection error).
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
    """Retrieves styles for cells within a range that have non-default formatting.

    Scans the specified `range_address` on the `sheet_name` and returns a mapping
    of cell addresses to their style properties (like font, fill, etc.). Only cells
    that have formatting different from the default style are included in the result.

    Note: The specific style attributes retrieved might evolve.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        range_address: The cell range (e.g., 'A1:C10') to scan for styles.

    Returns:
        ToolResult: {'success': True, 'data': Dict[str, CellStyle]} where 'data' is a
                    dictionary mapping cell addresses (e.g., 'A1', 'B3') within the
                    range to their respective `CellStyle` dictionaries, but only including
                    cells with non-default styles. Returns an empty dict if no non-default
                    styles are found.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid range, connection error).
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
    """Merges a rectangular range of cells into a single larger cell.

    Combines all cells within the specified `range_address` (e.g., 'A1:B2') on the
    `sheet_name` into one merged cell. The top-left cell's value and formatting
    are typically retained for the merged cell.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        range_address: The rectangular range of cells to merge (e.g., 'A1:B2', 'C3:E3'). Must cover multiple cells.

    Returns:
        ToolResult: {'success': True} if the cells were merged successfully.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid range, connection error).
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
    """Separates any merged cells found within a specified range back into individual cells.

    Scans the `range_address` on the `sheet_name` and unmerges any merged cells
    that fall entirely or partially within that range.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        range_address: The range (e.g., 'A1:C5') within which to unmerge cells.

    Returns:
        ToolResult: {'success': True} if the operation completed successfully (even if no cells were unmerged).
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid range, connection error).
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
    """Sets the height for a specific row on a worksheet.

    Adjusts the height of the row specified by `row_number` (1-based index) on
    the `sheet_name`.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        row_number: The 1-based index of the row whose height needs adjustment.
        height: The desired height in points (a positive number). Provide `None`
                to enable autofitting the row height based on content. Providing
                0 or a negative value typically hides the row.

    Returns:
        ToolResult: {'success': True} if the row height was set successfully.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid row number, connection error).
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
    """Sets the width for a specific column on a worksheet.

    Adjusts the width of the column specified by `column_letter` (e.g., 'A', 'AB')
    on the `sheet_name`.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        column_letter: The letter(s) identifying the column (e.g., 'A', 'AB'). Case-insensitive.
        width: The desired width in character units (a positive number). Provide
               `None` to enable autofitting the column width based on content.
               Providing 0 or a negative value typically hides the column.

    Returns:
        ToolResult: {'success': True} if the column width was set successfully.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid column letter, connection error).
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
    """Sets the widths for multiple columns simultaneously using a dictionary mapping.

    Applies column widths based on the provided `widths` dictionary, where keys are
    column letters (e.g., 'A', 'BC') and values are the desired widths (float for
    specific width, None for autofit). This is more efficient than calling
    `set_column_width_tool` repeatedly.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        widths: A dictionary mapping column letters (str, case-insensitive) to their
                desired widths (float) or `None` for autofit. Example:
                `{'A': 15.5, 'B': None, 'D': 20}`

    Returns:
        ToolResult: {'success': True} if all specified column widths were set successfully.
                    {'success': False, 'error': str} if any error occurred during application
                    (e.g., sheet not found, invalid column/width, connection error). If some
                    columns succeed and others fail, an error summarizing the failures is returned.
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