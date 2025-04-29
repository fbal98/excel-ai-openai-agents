# src/tools/sheet_ops.py
import asyncio
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult # Import the standard result type
from typing import Any, Optional, List

@function_tool
def get_sheet_names_tool(ctx: RunContextWrapper[AppContext]) -> ToolResult:
    """Retrieves the names of all worksheets currently in the workbook.

    Args:
        ctx: Agent context (injected automatically).

    Returns:
        ToolResult: {'success': True, 'data': List[str]} where 'data' is a list
                    containing the names of all worksheets in their current order.
                    Returns an empty list if the workbook has no sheets.
                    {'success': False, 'error': str} if an error occurred (e.g., connection error).
    """
    try:
        names = ctx.context.excel_manager.get_sheet_names()
        return {"success": True, "data": names}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_sheet_names_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except Exception as e:
        # Use print for server-side logging, return dict for agent
        print(f"[TOOL ERROR] get_sheet_names_tool: {e}")
        return {"success": False, "error": f"Failed to get sheet names: {e}"}

@function_tool
def get_active_sheet_name_tool(ctx: RunContextWrapper[AppContext]) -> ToolResult:
    """Retrieves the name of the worksheet that is currently active (selected) in Excel.

    Args:
        ctx: Agent context (injected automatically).

    Returns:
        ToolResult: {'success': True, 'data': str} where 'data' is the name of the
                    currently active worksheet.
                    {'success': False, 'error': str} if an error occurred or if no sheet
                    could be identified as active (e.g., workbook closing, no sheets exist).
    """
    try:
        name = ctx.context.excel_manager.get_active_sheet_name()
        if name:
             return {"success": True, "data": name}
        else:
             # This case might occur if the workbook is closing or in an odd state
             return {"success": False, "error": "Could not determine active sheet (might be closing or no sheets exist)."}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_active_sheet_name_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except Exception as e:
        print(f"[TOOL ERROR] get_active_sheet_name_tool: {e}")
        return {"success": False, "error": f"Failed to get active sheet name: {e}"}

@function_tool
async def create_sheet_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, index: Optional[int] = None) -> ToolResult:
    """Creates a new worksheet in the current workbook.

    Adds a new sheet with the specified `sheet_name`. Optionally, the sheet
    can be inserted at a specific position using the 0-based `index`. If the
    name already exists, an error is returned. Updates the workbook shape context.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The desired name for the new worksheet. Must be unique.
        index: (Optional) The 0-based position where the new sheet should be
               inserted (e.g., 0 for the first sheet, 1 for the second). If omitted,
               the sheet is typically added at the end.

    Returns:
        ToolResult: {'success': True, 'data': str} where 'data' is the `sheet_name`
                    of the successfully created sheet. The new sheet becomes the 'current_sheet'.
                    {'success': False, 'error': str} on failure (e.g., name already exists, connection error).
    """
    print(f"[TOOL] create_sheet_tool: sheet_name={sheet_name}, index={index}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'create_sheet_tool' failed: 'sheet_name' cannot be empty."}
    # Optional: Add check for invalid characters in sheet names if needed, though xlwings might handle this.
    # --- End Validation ---
    try:
        # Run the synchronous Excel manager method in a thread to avoid blocking
        await asyncio.to_thread(ctx.context.excel_manager.create_sheet, sheet_name, index)
        # After creating, force a shape update
        await asyncio.to_thread(ctx.context.update_shape)
        # Set the newly created sheet as the current context sheet
        ctx.context.state["current_sheet"] = sheet_name
        return {"success": True, "data": sheet_name} # Return sheet name on success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] create_sheet_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except ValueError as ve: # Catch specific error for existing sheet name
        print(f"[TOOL ERROR] create_sheet_tool: {ve}")
        return {"success": False, "error": str(ve)}
    except Exception as e:
        # Catch potential underlying errors
        print(f"[TOOL ERROR] create_sheet_tool: {e}")
        return {"success": False, "error": f"Exception creating sheet '{sheet_name}': {e}"}

@function_tool
def delete_sheet_tool(ctx: RunContextWrapper[AppContext], sheet_name: str) -> ToolResult:
    """Deletes a specified worksheet from the workbook.

    Removes the sheet identified by `sheet_name`. Note that Excel typically
    prevents the deletion of the very last remaining sheet in a workbook.
    Updates the workbook shape context and resets 'current_sheet' if the deleted sheet was active.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the worksheet to delete.

    Returns:
        ToolResult: {'success': True, 'data': str} where 'data' is the `sheet_name`
                    of the successfully deleted sheet.
                    {'success': False, 'error': str} on failure (e.g., sheet not found,
                    attempting to delete the last sheet, connection error).
    """
    print(f"[TOOL] delete_sheet_tool: sheet_name={sheet_name}")
    # --- Input Validation ---
    if not sheet_name:
        return {"success": False, "error": "Tool 'delete_sheet_tool' failed: 'sheet_name' cannot be empty."}
    # --- End Validation ---
    try:
        # Delegate deletion and propagate its return (usually None).
        ctx.context.excel_manager.delete_sheet(sheet_name)
         # If the deleted sheet was the current one, reset current_sheet in state
        if ctx.context.state.get("current_sheet") == sheet_name:
            ctx.context.state.pop("current_sheet", None)
            print(f"[INFO] Reset current_sheet context after deleting '{sheet_name}'.")
        # Force shape update after deleting
        ctx.context.update_shape()
        return {"success": True, "data": sheet_name} # Return sheet name on success
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] delete_sheet_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except ValueError as ve: # Catch specific error for trying to delete last sheet
        print(f"[TOOL ERROR] delete_sheet_tool: {ve}")
        return {"success": False, "error": str(ve)}
    except KeyError as ke: # Catch specific error for sheet not found
        print(f"[TOOL ERROR] delete_sheet_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] delete_sheet_tool: {e}")
        return {"success": False, "error": f"Exception deleting sheet '{sheet_name}': {e}"}


@function_tool
def get_dataframe_tool(ctx: RunContextWrapper[AppContext], sheet_name: str, header: Optional[bool] = None) -> ToolResult:
    """Reads the used range of a sheet and returns it as structured data (columns and rows).

    Retrieves the content of the 'used range' (the rectangular area containing data)
    from the specified `sheet_name`. The result is formatted as a dictionary containing
    a list of column headers and a list of data rows.

    Args:
        ctx: Agent context (injected automatically).
        sheet_name: The name of the worksheet to read from.
        header: (Optional) Boolean indicating whether the first row of the used range
                should be treated as column headers. If `True` or omitted/None, the first
                row becomes the 'columns' list in the result, and 'rows' contains the
                subsequent data. If `False`, the 'columns' list will be empty or represent
                Excel column letters, and 'rows' will contain all data including the first row.
                Defaults internally to True if not specified.

    Returns:
        ToolResult:
            {'success': True, 'data': {'columns': List[str], 'rows': List[List[CellValue]]}}
            On success, 'data' contains the extracted headers ('columns') and data ('rows').
            {'success': False, 'error': str} if an error occurred (e.g., sheet not found, connection error).
    """
    print(f"[TOOL] get_dataframe_tool: sheet={sheet_name}, header={header}")
    if not sheet_name:
        return {"success": False, "error": "Tool 'get_dataframe_tool' failed: 'sheet_name' cannot be empty."}

    # Default header to True if not provided or explicitly set to None by the agent
    use_header = True if header is None else header
    print(f"[TOOL] get_dataframe_tool: sheet={sheet_name}, header={use_header} (Original input: {header})") # Updated log

    try:
        # Pass the resolved header value to the underlying manager function
        dataframe_dict = ctx.context.excel_manager.get_sheet_dataframe(sheet_name, use_header)
        return {"success": True, "data": dataframe_dict}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] get_dataframe_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except KeyError as ke: # Catch sheet not found from _require_sheet
        print(f"[TOOL ERROR] get_dataframe_tool: Sheet not found - {ke}")
        return {"success": False, "error": str(ke)}
    except Exception as e:
        print(f"[TOOL ERROR] get_dataframe_tool: {e}")
        return {"success": False, "error": f"Exception dumping sheet '{sheet_name}': {e}"}