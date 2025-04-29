# src/tools/workbook_ops.py
import asyncio
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult # Import the standard result type
from typing import Any

@function_tool
async def open_workbook_tool(ctx: RunContextWrapper[AppContext], file_path: str) -> ToolResult:
    """Opens an existing Excel workbook file or attaches to an already open instance.

    Loads the workbook specified by `file_path` into the Excel application managed
    by the agent. If the underlying `ExcelManager` is configured to attach to
    existing instances, it might connect to an already open workbook matching
    the path instead of opening a new one. After opening, it updates the workbook shape context.

    Args:
        ctx: Agent context (injected automatically).
        file_path: The full path to the Excel workbook file (.xlsx, .xls, .xlsm, etc.) to open.

    Returns:
        ToolResult: {'success': True, 'data': str} A message indicating success and if shape was updated.
                    {'success': False, 'error': str} on failure (e.g., file not found, password protected, connection error).
    """
    try:
        # Delegate to ExcelManager (synchronous)
        ctx.context.excel_manager.open_workbook(file_path)
        # After opening, force a shape update to populate context immediately
        # Run the sync update_shape in a thread to avoid blocking if it becomes async later
        shape_updated = ctx.context.update_shape()
        # Even if shape update fails, opening the book itself succeeded at this point.
        # The shape update failure might be logged within update_shape.
        return {"success": True, "data": f"Workbook '{file_path}' opened. Shape updated: {shape_updated}"}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] open_workbook_tool: Connection Error - {ce}")
        # No need to wrap in _ensure_toolresult here as we manually return ToolResult
        return {"success": False, "error": f"Connection Error: {ce}"}
    except FileNotFoundError as fnf:
        print(f"[TOOL ERROR] open_workbook_tool: File Not Found - {fnf}")
        return {"success": False, "error": f"File not found: {fnf}"}
    except Exception as e:
        print(f"[TOOL ERROR] open_workbook_tool: {e}")
        return {"success": False, "error": f"Failed to open workbook '{file_path}': {e}"}

@function_tool
def snapshot_tool(ctx: RunContextWrapper[AppContext]) -> ToolResult:
    """Saves the current state of the workbook to a temporary file for potential restoration.

    Creates a temporary copy of the entire workbook in its current state. This allows
    the agent to revert back to this state later using `revert_snapshot_tool` if
    subsequent operations need to be undone. Only one snapshot is stored at a time;
    calling this again overwrites the previous snapshot.

    Args:
        ctx: Agent context (injected automatically).

    Returns:
        ToolResult: {'success': True, 'data': {'snapshot_path': str}} where 'data'
                    contains the file path of the saved temporary snapshot.
                    {'success': False, 'error': str} if snapshot creation failed (e.g., connection error, disk space issue).
    """
    try:
        path = ctx.context.excel_manager.snapshot()
        return {"success": True, "data": {"snapshot_path": path}}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] snapshot_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except Exception as e:
        print(f"[TOOL ERROR] snapshot_tool: {e}")
        return {"success": False, "error": f"Failed to take snapshot: {e}"}

@function_tool
def revert_snapshot_tool(ctx: RunContextWrapper[AppContext]) -> ToolResult:
    """Restores the workbook to the state saved by the last call to `snapshot_tool`.

    Closes the current workbook without saving changes made since the last snapshot
    and re-opens the temporary snapshot file created by the `snapshot_tool`. If no
    snapshot has been taken previously in the session, this tool will fail. After
    reverting, it updates the workbook shape context.

    Args:
        ctx: Agent context (injected automatically).

    Returns:
        ToolResult: {'success': True, 'data': str} A message indicating successful revert and shape refresh.
                    {'success': False, 'error': str} if reverting failed (e.g., no snapshot exists, file error, connection error).
    """
    try:
        # Propagate manager return value (None on success).
        ctx.context.excel_manager.revert_to_snapshot()
        # Force shape update after reverting
        ctx.context.update_shape()
        return {"success": True, "data": "Reverted to snapshot and refreshed shape."}
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] revert_snapshot_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except RuntimeError as rte: # Catch specific "No snapshot" error from manager
         print(f"[TOOL ERROR] revert_snapshot_tool: {rte}")
         return {"success": False, "error": str(rte)}
    except Exception as e:
        print(f"[TOOL ERROR] revert_snapshot_tool: {e}")
        return {"success": False, "error": f"Failed to revert to snapshot: {e}"}

@function_tool
def save_workbook_tool(ctx: RunContextWrapper[AppContext], file_path: str) -> ToolResult:
    """Saves the currently active workbook to a specified file path.

    Writes the current state of the workbook (all sheets, data, formatting, etc.)
    to the location specified by `file_path`. If the file already exists, it will
    be overwritten without warning.

    Args:
        ctx: Agent context (injected automatically).
        file_path: The full destination path (including filename and extension,
                   e.g., '/Users/me/Documents/report.xlsx') where the workbook should be saved.

    Returns:
        ToolResult: {'success': True, 'data': str} where 'data' is the absolute path
                    to which the workbook was successfully saved.
                    {'success': False, 'error': str} on failure (e.g., invalid path, permissions error, connection error).
    """
    print(f"[TOOL] save_workbook_tool: path={file_path}")
    # --- Input Validation ---
    if not file_path:
        return {"success": False, "error": "Tool 'save_workbook_tool' failed: 'file_path' cannot be empty."}
    # --- End Validation ---
    try:
        # save_workbook returns the saved path on success
        saved_path = ctx.context.excel_manager.save_workbook(file_path)
        return {"success": True, "data": saved_path} # Return the absolute path
    except ExcelConnectionError as ce:
        print(f"[TOOL ERROR] save_workbook_tool: Connection Error - {ce}")
        return {"success": False, "error": f"Connection Error: {ce}"}
    except Exception as e:
        print(f"[TOOL ERROR] save_workbook_tool: {e}")
        return {"success": False, "error": f"Exception saving workbook to '{file_path}': {e}"}

# Note: The _wrap_tool_result decorator will be applied in __init__.py
#       so we don't need to manually call _ensure_toolresult here unless we want to be explicit.
#       The current implementation returns the required dict structure directly.