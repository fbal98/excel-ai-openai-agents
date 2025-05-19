# src/tools/formula_ops.py
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult
from openpyxl.utils.cell import coordinate_from_string # Import from correct submodule

@function_tool
def set_cell_formula_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, cell_address: str, formula: str) -> ToolResult:
    """Sets an Excel formula in a single specified cell.

    Writes the provided `formula` string into the cell identified by `cell_address`
    on the `sheet_name`. The formula MUST begin with an equals sign ('=').

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        cell_address: The cell address in A1 notation (e.g., 'A1', 'D5') where the formula should be set.
        formula: The Excel formula string to set (e.g., '=SUM(B2:B10)', '=VLOOKUP(A1,Sheet2!A:B,2,FALSE)').
                 Must start with '='.

    Returns:
        ToolResult: {'success': True} if the formula was set successfully.
                    {'success': False, 'error': str} on failure (e.g., sheet not found, invalid formula syntax, connection error).
    """
    print(f"[TOOL] set_cell_formula_tool: {sheet_name}!{cell_address} formula={formula}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_cell_formula_tool' failed: 'sheet_name' cannot be empty."}
    if not cell_address:
        return {"success": False, "error": "Tool 'set_cell_formula_tool' failed: 'cell_address' cannot be empty."}
    if not formula: # Formula string cannot be empty
        return {"success": False, "error": "Tool 'set_cell_formula_tool' failed: 'formula' cannot be empty."}
    if not formula.startswith('='):
         # Enforce starting with '='
         return {"success": False, "error": f"Tool 'set_cell_formula_tool' failed: 'formula' must start with '='. Received: '{formula}'"}
    # --- End Validation ---
    try:
        ctx.context.excel_manager.set_cell_formula(sheet_name, cell_address, formula)
        return {"success": True} # Explicit success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_cell_formula_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] set_cell_formula_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] set_cell_formula_tool: {e}")
        return {"success": False, "error": f"Exception setting cell formula for {sheet_name}!{cell_address}: {e}"}


@function_tool
def set_range_formula_tool(ctx: RunContextWrapper[AppContext],
                          sheet_name: str,
                          range_address: str,
                          template: str) -> ToolResult:
    """Applies a formula template to each cell within a specified range, adjusting for row number.

    Iterates through each cell in the `range_address` (e.g., 'F2:F6') on the `sheet_name`.
    For each cell, it substitutes the placeholder '{row}' within the `template` formula
    with the cell's current row number, and then sets the resulting formula in that cell.
    The template MUST begin with an equals sign ('=').

    Example:
    If `range_address` is 'F2:F4' and `template` is '=SUM(B{row}:E{row})', this tool will set:
    - F2: '=SUM(B2:E2)'
    - F3: '=SUM(B3:E3)'
    - F4: '=SUM(B4:E4)'

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the target worksheet.
        range_address: The target range (e.g., 'F2:F6', 'A1:A10') where the formula template will be applied to each cell.
        template: The formula string template. It MUST start with '=' and can contain the
                  placeholder '{row}', which will be replaced by the 1-based row number of the cell being processed.

    Returns:
        ToolResult: {'success': True} if the template was applied successfully to all cells in the range.
                    {'success': False, 'error': str} if any error occurred during application (e.g., sheet not found, invalid range, template issues, connection error). If some cells succeed and others fail, an error summarizing the failures is returned.
    """
    # Input validation
    if not sheet_name:
        return {"success": False, "error": "Tool 'set_range_formula_tool' failed: 'sheet_name' cannot be empty."}
    if not range_address:
        return {"success": False, "error": "Tool 'set_range_formula_tool' failed: 'range_address' cannot be empty."}
    if not template:
        return {"success": False, "error": "Tool 'set_range_formula_tool' failed: 'template' formula cannot be empty."}
    if not template.startswith('='):
        return {"success": False, "error": f"Tool 'set_range_formula_tool' failed: 'template' formula must start with '='. Received: '{template}'"}

    print(f"[TOOL] set_range_formula_tool: Applying template '{template}' to {sheet_name}!{range_address}")
    errors = []
    try:
        # Check connection once before loop
        sheet = ctx.context.excel_manager._require_sheet(sheet_name) # Raises if sheet not found or connection lost

        # Get the range object
        target_range = sheet.range(range_address)

        # Iterate through each cell in the range
        for cell in target_range:
            try:
                current_row = cell.row
                # Format the template with the current row number
                formula_instance = template.format(row=current_row)
                # Set formula for the individual cell
                cell.formula = formula_instance # Use xlwings direct attribute
                # Alternative: Call manager's single cell function (might be less efficient)
                # ctx.context.excel_manager.set_cell_formula(sheet_name, cell.address.replace('$', ''), formula_instance)
            except Exception as cell_err:
                 error_msg = f"Failed for cell '{cell.address.replace('$', '')}': {cell_err}"
                 print(f"[TOOL ERROR] set_range_formula_tool: {error_msg}")
                 errors.append(error_msg)
                 # Continue applying to other cells

        if errors:
            return {"success": False, "error": f"Some formulas failed to apply in range {range_address}: {'; '.join(errors)}"}
        else:
            return {"success": True}

    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_range_formula_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from initial check
        print(f"[TOOL ERROR] set_range_formula_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e: # Catch errors like invalid range address
        print(f"[TOOL ERROR] set_range_formula_tool: {e}")
        return {"success": False, "error": f"Exception applying range formula for {sheet_name}!{range_address}: {e}"}