# src/tools/workbook_ops.py
import asyncio
from agents import RunContextWrapper, function_tool
from ..context import AppContext
from ..excel_ops import ExcelConnectionError
from .core_defs import ToolResult # Import the standard result type
from typing import Any

@function_tool
async def open_workbook_tool(ctx: RunContextWrapper[AppContext], file_path: str) -> ToolResult:
    """
    Opens or attaches to an Excel workbook at the given path.

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.
        file_path (str): Path to the Excel workbook to open or attach.

    Returns:
        ToolResult: {'success': True} if the workbook was opened successfully,
                    {'success': False, 'error': str} if an error occurred.
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
    """
    Saves a temporary snapshot of the current workbook state.

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.

    Returns:
        ToolResult: {'success': True, 'data': {'snapshot_path': str}} path to the saved snapshot file.
                    {'success': False, 'error': str} if an error occurred.
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
    """
    Reverts the workbook to the last snapshot taken. Fails if no snapshot exists.
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
    """
    Saves the current workbook to the given file path.

    Args:
        ctx (RunContextWrapper[AppContext]): Agent context containing the ExcelManager.
        file_path (str): Destination file path where the workbook should be saved.

    Returns:
        ToolResult: {'success': True, 'data': str} saved path on success.
                    {'success': False, 'error': str} if an error occurred.
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