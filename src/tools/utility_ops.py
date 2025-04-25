# src/tools/utility_ops.py
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult # Import the standard result type
from typing import Any, Optional, Dict

@function_tool
def copy_paste_range_tool(
    ctx: RunContextWrapper[AppContext],
    src_sheet: str,
    src_range: str,
    dst_sheet: str,
    dst_anchor: str, # Top-left cell of the destination paste area
    paste_opts: str, # 'values', 'formulas', 'formats', 'all', 'column_widths', etc.
) -> ToolResult:
    """
    Copies a source range and paste-special into a destination sheet starting at an anchor cell.

    Args:
        src_sheet (str): Source worksheet name.
        src_range (str): Source range (A1 style, e.g., 'A1:B10').
        dst_sheet (str): Destination worksheet name.
        dst_anchor (str): Top-left destination cell (A1 style, e.g., 'D1').
        paste_opts (str): Paste type ('values', 'formulas', 'formats', 'all', 'column_widths', 'values_number_formats').

    Returns:
        ToolResult: {'success': True} on success or {'success': False, 'error': str}.
    """
    # --- Input Validation ---
    valid_paste_opts = {"values", "formulas", "formats", "all", "column_widths", "values_number_formats"}
    if not all([src_sheet, src_range, dst_sheet, dst_anchor]):
        return {"success": False, "error": "Tool 'copy_paste_range_tool' failed: All sheet/range/anchor parameters are required."}
    opts_lower = paste_opts.lower()
    if opts_lower not in valid_paste_opts:
        return {"success": False, "error": f"Tool 'copy_paste_range_tool' failed: paste_opts must be one of {valid_paste_opts}. Got '{paste_opts}'."}
    # --- End Validation ---

    print(f"[TOOL] copy_paste_range_tool: From {src_sheet}!{src_range} To {dst_sheet}!{dst_anchor} (Paste: {opts_lower})")
    try:
        ctx.context.excel_manager.copy_paste_range(
            src_sheet, src_range, dst_sheet, dst_anchor, opts_lower
        )
        return {"success": True}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] copy_paste_range_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet calls within manager
        print(f"[TOOL ERROR] copy_paste_range_tool: Source or Destination Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] copy_paste_range_tool: {e}")
        return {"success": False, "error": f"Failed to copy/paste range: {e}"}


@function_tool(strict_mode=False) # Allow flexible dict structure for 'mapping'
def set_named_ranges_tool(ctx: RunContextWrapper[AppContext],
                          # sheet_name: str, # Sheet name often not strictly needed for workbook-level names
                          mapping: Dict[str, str]) -> ToolResult:
    """
    Creates or updates one or more workbook-level named ranges.

    Args:
        ctx: Agent context.
        mapping: Dictionary where keys are the desired names (e.g., "SalesData")
                 and values are the range references they should point to
                 (e.g., "Sheet1!A1:B10", "='Constants'!$C$1:$C$5").
                 The reference string should include the sheet name if it's sheet-specific.
                 Prefix formulas with '=' (e.g., "=OFFSET(...)").

    Returns:
        ToolResult: {'success': True} on success or {'success': False, 'error': str} on failure.
    """
    # --- Input Validation ---
    if not mapping or not isinstance(mapping, dict):
        return {"success": False, "error": "Tool 'set_named_ranges_tool' failed: 'mapping' must be a non-empty dictionary."}
    # --- End Validation ---
    print(f"[TOOL] set_named_ranges_tool: Setting {len(mapping)} named ranges: {list(mapping.keys())}")
    errors = []
    try:
        app, book = ctx.context.excel_manager._validate_connection() # Get validated book

        for nm, refers_to in mapping.items():
            try:
                # Attempt to add or update the name
                # xlwings handles create/update via item access/assignment
                book.names.add(name=nm, refers_to=refers_to) # xlwings handles the '=' prefix if needed for refs
                print(f"[DEBUG] Set named range '{nm}' to '{refers_to}'")
            except Exception as e:
                error_msg = f"Failed setting name '{nm}' to '{refers_to}': {e}"
                print(f"[TOOL ERROR] set_named_ranges_tool: {error_msg}")
                errors.append(error_msg)
                # Continue with the next name

        if errors:
             return {"success": False, "error": f"Some named ranges failed to set: {'; '.join(errors)}"}
        else:
             return {"success": True}

    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] set_named_ranges_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except Exception as e: # Catch unexpected errors during the process
        print(f"[TOOL ERROR] set_named_ranges_tool: {e}")
        return {"success": False, "error": f"Failed to set named ranges: {e}"}