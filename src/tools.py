from agents import RunContextWrapper
from .context import AppContext
from typing import Any, Dict, Optional, List
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from typing_extensions import TypedDict

# NOTE: ExcelManager methods return None on success; tool wrappers must not treat None as failure.
#       Always return True (or success dict) when no exception is raised.

# Define specific types for style components
class FontStyle(TypedDict, total=False):
    name: Optional[str]
    size: Optional[float]
    bold: Optional[bool]
    italic: Optional[bool]
    vertAlign: Optional[str]
    underline: Optional[str]
    strike: Optional[bool]
    color: Optional[str]

class FillStyle(TypedDict, total=False):
    fill_type: Optional[str] # e.g., 'solid'
    start_color: Optional[str]
    end_color: Optional[str]

class BorderStyleDetails(TypedDict, total=False):
    style: Optional[str] # e.g., 'thin', 'medium'
    color: Optional[str]

class BorderStyle(TypedDict, total=False):
    left: Optional[BorderStyleDetails]
    right: Optional[BorderStyleDetails]
    top: Optional[BorderStyleDetails]
    bottom: Optional[BorderStyleDetails]
    diagonal: Optional[BorderStyleDetails]
    diagonal_direction: Optional[int]
    outline: Optional[bool]
    vertical: Optional[BorderStyleDetails]
    horizontal: Optional[BorderStyleDetails]


# Define the main style structure using TypedDict
class CellStyle(TypedDict, total=False):
    """Defines the structure for cell styling options."""
    font: Optional[FontStyle]
    fill: Optional[FillStyle]
    border: Optional[BorderStyle]

# All tool functions are ready for @function_tool decoration.

# Tool: Get all sheet names
def get_sheet_names_tool(ctx: RunContextWrapper[AppContext]) -> Any:
    """
    Returns a list of all sheet names in the workbook.
    Args:
        ctx: Agent context (injected automatically).
    Returns:
        List of sheet names, or an error message.
    """
    try:
        return ctx.context.excel_manager.get_sheet_names()
    except Exception as e:
        # Use print for server-side logging, return dict for agent
        print(f"[TOOL ERROR] get_sheet_names_tool: {e}")
        return {"error": f"Failed to get sheet names: {e}"}

# Tool: Get active sheet name
def get_active_sheet_name_tool(ctx: RunContextWrapper[AppContext]) -> Any:
    """
    Returns the name of the active sheet in the workbook.
    Args:
        ctx: Agent context (injected automatically).
    Returns:
        Active sheet name, or an error message.
    """
    try:
        return ctx.context.excel_manager.get_active_sheet_name()
    except Exception as e:
        print(f"[TOOL ERROR] get_active_sheet_name_tool: {e}")
        return {"error": f"Failed to get active sheet name: {e}"}

# Tool: Set cell value
def set_cell_value_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str, value: str) -> Any:
    """
    Sets the value of a single cell in the specified sheet. The value will be passed as a string.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        cell_address: Cell address (e.g., 'A1').
        value: Value to set (as a string).
    Returns:
        True if successful, or an error dictionary.
    """
    print(f"[TOOL] set_cell_value_tool: {sheet_name}!{cell_address} value={value}")
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'set_cell_value_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"error": "Tool 'set_cell_value_tool' failed: 'cell_address' cannot be empty."}
    # Note: Validating 'value: Any' is complex; rely on underlying function for now.
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_cell_value(sheet_name, cell_address, value)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_value_tool: {e}")
        return {"error": f"Exception setting cell value for {sheet_name}!{cell_address}: {e}"}

# Tool: Get cell value
def get_cell_value_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str) -> Any:
    print(f"[TOOL] get_cell_value_tool: {sheet_name}!{cell_address}")
    """
    Gets the value of a cell in the specified sheet.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        cell_address: Cell address (e.g., 'A1').
    Returns:
        Cell value, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'get_cell_value_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"error": "Tool 'get_cell_value_tool' failed: 'cell_address' cannot be empty."}
    # --- End Validation ---
    try:
        value = ctx.context.excel_manager.get_cell_value(sheet_name, cell_address)
        # Allow returning None if cell is genuinely empty, but catch exceptions
        return value
    except Exception as e:
        print(f"[TOOL ERROR] get_cell_value_tool: {e}")
        return {"error": f"Exception getting cell value for {sheet_name}!{cell_address}: {e}"}

# Tool: Get range of cell values
def get_range_values_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str) -> Any:
    """
    Retrieves values for a rectangular range of cells in the specified sheet.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        range_address: Excel range (e.g., 'A1:C5').
    Returns:
        A dict with 'values' as a list of rows (each row is a list of cell values),
        or an 'error' message if failed.
    """
    print(f"[TOOL] get_range_values_tool: {sheet_name}!{range_address}")
    # Validation
    if not sheet_name:
        return {"error": "Tool 'get_range_values_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"error": "Tool 'get_range_values_tool' failed: 'range_address' cannot be empty."}
    try:
        ws = ctx.context.excel_manager.get_sheet(sheet_name)
        if ws is None:
            return {"error": f"Sheet '{sheet_name}' not found."}
        # Fetch cells in range
        cells = ws[range_address]
        # Normalize to list of rows
        # openpyxl returns a single cell or tuple of tuples
        if not hasattr(cells, '__iter__') or isinstance(cells, ctx.context.excel_manager.workbook.__class__):
            # Single cell case
            return {"values": [[cells.value]]}
        rows = []
        for row in cells:
            # row may be a tuple for range or a single cell
            if hasattr(row, '__iter__'):
                rows.append([cell.value for cell in row])
            else:
                rows.append([row.value])
        return {"values": rows}
    except Exception as e:
        print(f"[TOOL ERROR] get_range_values_tool: {e}")
        return {"error": f"Exception getting range values for {sheet_name}!{range_address}: {e}"}

# Tool: Set range style
# The SDK automatically handles JSON conversion to the Pydantic/TypedDict model.
def set_range_style_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str, style: CellStyle) -> Any:
    print(f"[TOOL] set_range_style_tool: {sheet_name}!{range_address} style={style}")
    """
    Applies styles to a range of cells based on a style description dictionary.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        range_address: Excel range (e.g., 'A1:B2').
        style_json: Style dictionary adhering to CellStyle structure (keys: 'font', 'fill', 'border').
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'set_range_style_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"error": "Tool 'set_range_style_tool' failed: 'range_address' cannot be empty."}
    if not style: # Check if the style dictionary itself is empty
        return {"error": "Tool 'set_range_style_tool' failed: 'style' dictionary cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_range_style(sheet_name, range_address, style)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] set_range_style_tool: {e}")
        return {"error": f"Exception applying cell style to {sheet_name}!{range_address}: {e}"}

# Tool: Create sheet
def create_sheet_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, index: Optional[int] = None) -> Any:
    print(f"[TOOL] create_sheet_tool: sheet_name={sheet_name}, index={index}")
    """
    Creates a new sheet with the given name and optional index.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the new sheet.
        index: Optional position for the new sheet.
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'create_sheet_tool' failed: 'sheet_name' cannot be empty."}
    # Optional: Add check for invalid characters in sheet names if needed, though openpyxl might handle this.
    # --- End Validation ---
    try:
        ctx.context.excel_manager.create_sheet(sheet_name, index)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] create_sheet_tool: {e}")
        return {"error": f"Exception creating sheet '{sheet_name}': {e}"}

# Tool: Delete sheet
def delete_sheet_tool(ctx: RunContextWrapper[AppContext], sheet_name: str) -> Any:
    print(f"[TOOL] delete_sheet_tool: sheet_name={sheet_name}")
    """
    Deletes the specified sheet from the workbook.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet to delete.
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'delete_sheet_tool' failed: 'sheet_name' cannot be empty."}
    # --- End Validation ---
    try:
        result = ctx.context.excel_manager.delete_sheet(sheet_name)
        if not result:
            return {"error": f"Failed to delete sheet '{sheet_name}' (Operation returned false)"}
        return True
    except Exception as e:
        print(f"[TOOL ERROR] delete_sheet_tool: {e}")
        return {"error": f"Exception deleting sheet '{sheet_name}': {e}"}

# Tool: Merge cells
def merge_cells_range_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str) -> Any:
    print(f"[TOOL] merge_cells_range_tool: {sheet_name}!{range_address}")
    """
    Merges a range of cells in the specified sheet.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        range_address: Range to merge (e.g., 'A1:B2').
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'merge_cells_range_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"error": "Tool 'merge_cells_range_tool' failed: 'range_address' cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.merge_cells_range(sheet_name, range_address)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] merge_cells_range_tool: {e}")
        return {"error": f"Exception merging cells {sheet_name}!{range_address}: {e}"}

# Tool: Unmerge cells
def unmerge_cells_range_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str) -> Any:
    print(f"[TOOL] unmerge_cells_range_tool: {sheet_name}!{range_address}")
    """
    Unmerges a range of cells in the specified sheet.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        range_address: Range to unmerge (e.g., 'A1:B2').
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'unmerge_cells_range_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"error": "Tool 'unmerge_cells_range_tool' failed: 'range_address' cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.unmerge_cells_range(sheet_name, range_address)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] unmerge_cells_range_tool: {e}")
        return {"error": f"Exception unmerging cells {sheet_name}!{range_address}: {e}"}

# Tool: Set row height
def set_row_height_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, row_number: int, height: float) -> Any:
    print(f"[TOOL] set_row_height_tool: {sheet_name} row {row_number} height={height}")
    """
    Sets the height of a row in the specified sheet.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        row_number: Row number (1-based).
        height: Height in points.
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'set_row_height_tool' failed: 'sheet_name' cannot be empty."}
    if not isinstance(row_number, int) or row_number <= 0:
        return {"error": f"Tool 'set_row_height_tool' failed: 'row_number' must be a positive integer (got {row_number})."}
    if not isinstance(height, (int, float)) or height < 0:
         # Excel might allow 0 height, but negative is invalid.
        return {"error": f"Tool 'set_row_height_tool' failed: 'height' must be a non-negative number (got {height})."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_row_height(sheet_name, row_number, height)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] set_row_height_tool: {e}")
        return {"error": f"Exception setting row height for row {row_number} in '{sheet_name}': {e}"}

# Tool: Set column width
def set_column_width_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, column_letter: str, width: float) -> Any:
    print(f"[TOOL] set_column_width_tool: {sheet_name} column {column_letter} width={width}")
    """
    Sets the width of a column in the specified sheet.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        column_letter: Column letter (e.g., 'A').
        width: Width in points.
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'set_column_width_tool' failed: 'sheet_name' cannot be empty."}
    if not column_letter or not isinstance(column_letter, str):
        return {"error": "Tool 'set_column_width_tool' failed: 'column_letter' must be a non-empty string."}
    if not isinstance(width, (int, float)) or width < 0:
        # Excel might allow 0 width, but negative is invalid.
        return {"error": f"Tool 'set_column_width_tool' failed: 'width' must be a non-negative number (got {width})."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_column_width(sheet_name, column_letter.upper(), width)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] set_column_width_tool: {e}")
        return {"error": f"Exception setting column width for column {column_letter.upper()} in '{sheet_name}': {e}"}

# Tool: Set cell formula
def set_cell_formula_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str, formula: str) -> Any:
    print(f"[TOOL] set_cell_formula_tool: {sheet_name}!{cell_address} formula={formula}")
    """
    Sets a formula in the specified cell. Formula must start with '='.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        cell_address: Cell address (e.g., 'A1').
        formula: The formula string (with or without '=' prefix).
    Returns:
        True if successful, or an error message.
    """
    # --- Input Validation ---
    if not sheet_name:
        return {"error": "Tool 'set_cell_formula_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"error": "Tool 'set_cell_formula_tool' failed: 'cell_address' cannot be empty."}
    if not formula: # Formula string cannot be empty
        return {"error": "Tool 'set_cell_formula_tool' failed: 'formula' cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_cell_formula(sheet_name, cell_address, formula)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_formula_tool: {e}")
        return {"error": f"Exception setting cell formula for {sheet_name}!{cell_address}: {e}"}

# Tool: Set multiple cell values
class CellValueMap(TypedDict):
    """A mapping from cell addresses (e.g., 'A1') to values (as strings)."""
    # This TypedDict represents a dictionary like: {"A1": "value1", "B2": "value2"}
    # It doesn't need explicit fields defined here because it acts as a type hint
    # for arbitrary key-value pairs where keys are strings (cell addresses)
    # and values are strings (cell values). The actual validation happens
    # during runtime or type checking based on how it's used.
    pass # Use pass instead of ellipsis for an empty body


class SetCellValuesResult(TypedDict, total=False):
    success: bool
    error: str


def set_cell_values_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    data: Dict[str, str]
) -> SetCellValuesResult:
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_cell_values_tool' failed: 'sheet_name' cannot be empty."}
    if not data: # Check if data dictionary is empty
        print("[TOOL ERROR] set_cell_values_tool called with empty data dictionary.")
        return {"success": False, "error": "Tool 'set_cell_values_tool' failed: The 'data' dictionary cannot be empty. Provide cell addresses and values."}
    # Optional: Add validation for keys (cell addresses) and values if needed, though manager might handle it.
    # --- End Validation ---
    print(f"[TOOL] set_cell_values_tool: {sheet_name}, {len(data)} cells")
    """
    Sets the values of multiple cells in the specified sheet from a dictionary.
    Args:
        ctx: Agent context (injected automatically).
        sheet_name: Name of the sheet.
        data: Dictionary mapping cell addresses (e.g., 'A1') to string values to set.
    Returns:
        A result dict with 'success' True if successful, or 'error' message if failed.
    """
    try:
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
        return {"success": True}
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_values_tool: {e}")
        return {"success": False, "error": f"Exception setting multiple cell values in '{sheet_name}': {e}"}

# ------------------------------------------------------------------ #
#  Bulk helper tools                                                 #
# ------------------------------------------------------------------ #

def set_table_tool(ctx: RunContextWrapper[AppContext],
                   sheet_name: str,
                   top_left: str,
                   rows: List[List[Any]]) -> Any:
    """
    Bulk‑write a 2‑D python list into *sheet_name* starting at *top_left* (e.g. 'A2').
    Saves ≥30 single calls on header+data tables.
    """
    if not sheet_name or not top_left or not rows:
        return {"error": "sheet_name, top_left, rows are required"}
    try:
        col_letter, r0 = coordinate_from_string(top_left)
        c0_idx = column_index_from_string(col_letter)
        data = {
            f"{get_column_letter(c0_idx + c)}{r0 + r}": v
            for r, row in enumerate(rows)
            for c, v in enumerate(row)
        }
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
        return True
    except Exception as e:
        return {"error": str(e)}

# Tool: Insert a formatted Excel table
def insert_table_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    start_cell: str,
    columns: List[Any],
    rows: List[List[Any]],
    table_name: Optional[str] = None,
    table_style: Optional[str] = None,
) -> Any:
    """
    Inserts a formatted Excel table with headers and data in one call.
    """
    try:
        ctx.context.excel_manager.insert_table(sheet_name, start_cell, columns, rows, table_name, table_style)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] insert_table_tool: {e}")
        return {"error": f"Exception in insert_table_tool: {e}"}

def set_columns_widths_tool(ctx: RunContextWrapper[AppContext],
                            sheet_name: str,
                            widths: Dict[str, float]) -> Any:
    """Set multiple column widths in one call (openpyxl column_dimensions)."""
    try:
        for col, w in widths.items():
            ctx.context.excel_manager.set_column_width(sheet_name, col, w)
        return True
    except Exception as e:
        return {"error": str(e)}

def set_range_formula_tool(ctx: RunContextWrapper[AppContext],
                           sheet_name: str,
                           range_address: str,
                           template: str) -> Any:
    """
    Apply *template* row‑wise to the leftmost cell in each row of *range_address*.
    Example: template='=SUM(B{row}:E{row})' on F2:F6.
    """
    try:
        ws = ctx.context.excel_manager.get_sheet(sheet_name)
        for row in ws[range_address]:
            r = row[0].row
            row[0].value = template.format(row=r)
        return True
    except Exception as e:
        return {"error": str(e)}

# ------------------------------------------------------------------ #
#  Composite write‑and‑verify tool                                   #
# ------------------------------------------------------------------ #

class WriteVerifyResult(TypedDict, total=False):
    success: bool
    diff: Dict[str, Any]

def write_and_verify_range_tool(
    ctx: RunContextWrapper[AppContext],
    sheet_name: str,
    data: Dict[str, str],
) -> WriteVerifyResult:
    """
    Bulk‑write the provided cell→value mapping, then immediately read the same
    cells back and compare. Returns {"success": True} on a perfect match, or
    {"success": False, "diff": {...}} where diff maps each mismatching cell to
    {"expected": value, "actual": actual_value}.
    """
    if not sheet_name:
        return {"success": False, "diff": {"error": "'sheet_name' cannot be empty."}}
    if not data:
        return {"success": False, "diff": {"error": "'data' dictionary cannot be empty."}}

    # 1. Write
    try:
        ctx.context.excel_manager.set_cell_values(sheet_name, data)
    except Exception as e:
        return {"success": False, "diff": {"error": f"Write failed: {e}"}}

    # 2. Verify
    diff: Dict[str, Any] = {}
    for addr, expected in data.items():
        try:
            actual = ctx.context.excel_manager.get_cell_value(sheet_name, addr)
            if actual != expected:
                diff[addr] = {"expected": expected, "actual": actual}
        except Exception as e:
            diff[addr] = {"expected": expected, "actual": f"(read‑error: {e})"}

    if diff:
        return {"success": False, "diff": diff}
    return {"success": True}

# ------------------------------------------------------------------ #
#  Style‑inspection tools                                            #
# ------------------------------------------------------------------ #

def get_cell_style_tool(
    ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str
) -> Any:
    """
    Return the style dict (font/fill/border) for a single cell.
    """
    try:
        return ctx.context.excel_manager.get_cell_style(sheet_name, cell_address)
    except Exception as e:
        return {"error": str(e)}

def get_range_style_tool(
    ctx: RunContextWrapper[AppContext], sheet_name: str, range_address: str
) -> Any:
    """
    Return a mapping of cell_address -> style_dict for a rectangular range.
    """
    try:
        return ctx.context.excel_manager.get_range_style(sheet_name, range_address)
    except Exception as e:
        return {"error": str(e)}

# ------------------------------------------------------------------ #
#  (Existing) Save workbook tool                                     #
# ------------------------------------------------------------------ #
# Tool: Save workbook
def save_workbook_tool(ctx: RunContextWrapper[AppContext], file_path: str) -> Any:
    """
    Saves the workbook to the specified file path.
    Args:
        ctx: Agent context (injected automatically).
        file_path: Path to save the workbook.
    Returns:
        True if successful, or an error message.
    """
    print(f"[TOOL] save_workbook_tool: path={file_path}")
    # --- Input Validation ---
    if not file_path:
        return {"error": "Tool 'save_workbook_tool' failed: 'file_path' cannot be empty."}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.save_workbook(file_path)
        return True
    except Exception as e:
        print(f"[TOOL ERROR] save_workbook_tool: {e}")
        return {"error": f"Exception saving workbook to '{file_path}': {e}"}