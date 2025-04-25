# src/tools/__init__.py
import inspect
import sys

# Import individual tool functions from their respective modules
from .core_defs import (
    ToolResult, SetCellValuesResult, CellValue, FontStyle, FillStyle,
    BorderStyleDetails, BorderStyle, AlignmentStyle, CellStyle, CellValueMap,
    WriteVerifyResult
)
from .workbook_ops import (
    open_workbook_tool,
    save_workbook_tool,
    snapshot_tool,
    revert_snapshot_tool,
)
from .sheet_ops import (
    get_sheet_names_tool,
    get_active_sheet_name_tool,
    create_sheet_tool,
    delete_sheet_tool,
    get_dataframe_tool,
)
from .data_ops import (
    set_cell_value_tool,
    get_cell_value_tool,
    get_range_values_tool,
    set_cell_values_tool,
    set_table_tool,
    insert_table_tool,
    set_rows_tool,
    set_columns_tool,
    append_table_rows_tool,
    write_and_verify_range_tool,
    find_row_by_value_tool,
)
from .style_ops import (
    set_range_style_tool,
    set_cell_style_tool,
    get_cell_style_tool,
    get_range_style_tool,
    merge_cells_range_tool,
    unmerge_cells_range_tool,
    set_row_height_tool,
    set_column_width_tool,
    set_columns_widths_tool,
)
from .formula_ops import (
    set_cell_formula_tool,
    set_range_formula_tool,
)
from .utility_ops import (
    copy_paste_range_tool,
    set_named_ranges_tool,
)

# --- Apply ToolResult wrapper ---
# Import the wrapper and necessary types
from .core_defs import _wrap_tool_result, _ensure_toolresult
from agents import FunctionTool as _FunctionTool # Alias to check instance type

# Get all members of the current module
_current_module = sys.modules[__name__]
_all_members = inspect.getmembers(_current_module)

# Filter for functions ending in '_tool' that haven't already been wrapped (or are FunctionTool instances)
_unwrapped_tools = []
for _name, _obj in _all_members:
    if _name.endswith("_tool") and callable(_obj):
         # Check if it's a FunctionTool instance (already processed by @function_tool)
         # Or if it has already been wrapped by checking for the wrapper's signature characteristics (less reliable)
         # Safest is usually to let the wrapper be idempotent or skip FunctionTool instances.
         if not isinstance(_obj, _FunctionTool):
             # Check if it looks like our wrapper already applied it (e.g., check internal flags if set by wrapper)
             # For simplicity, assume if it's not a FunctionTool, it needs wrapping.
             # The _wrap_tool_result decorator itself handles sync/async correctly.
             _unwrapped_tools.append((_name, _obj))

# Apply the wrapper to the identified tool functions
for _name, _func in _unwrapped_tools:
    setattr(_current_module, _name, _wrap_tool_result(_func))
    print(f"[DEBUG] Applied ToolResult wrapper to: {_name}")


# --- Explicitly define __all__ for export ---
# This controls what `from src.tools import *` imports
# List all the tool functions intended for agent use
__all__ = [
    # Core Types (optional, but useful for type hinting outside)
    'ToolResult', 'SetCellValuesResult', 'CellValue', 'CellStyle', 'CellValueMap', 'WriteVerifyResult',
    'FontStyle', 'FillStyle', 'BorderStyleDetails', 'BorderStyle', 'AlignmentStyle',
    # Tool Functions
    'open_workbook_tool',
    'save_workbook_tool',
    'snapshot_tool',
    'revert_snapshot_tool',
    'get_sheet_names_tool',
    'get_active_sheet_name_tool',
    'create_sheet_tool',
    'delete_sheet_tool',
    'get_dataframe_tool',
    'set_cell_value_tool',
    'get_cell_value_tool',
    'get_range_values_tool',
    'set_cell_values_tool',
    'set_table_tool',
    'insert_table_tool',
    'set_rows_tool',
    'set_columns_tool',
    'append_table_rows_tool',
    'write_and_verify_range_tool',
    'find_row_by_value_tool',
    'set_range_style_tool',
    'set_cell_style_tool',
    'get_cell_style_tool',
    'get_range_style_tool',
    'merge_cells_range_tool',
    'unmerge_cells_range_tool',
    'set_row_height_tool',
    'set_column_width_tool',
    'set_columns_widths_tool',
    'set_cell_formula_tool',
    'set_range_formula_tool',
    'copy_paste_range_tool',
    'set_named_ranges_tool',
    # Helper functions (maybe not needed in __all__ unless used directly elsewhere)
    # '_ensure_toolresult',
    # '_wrap_tool_result',
    # '_to_bgr',
    # '_hex_argb_to_bgr_int',
    # '_bgr_int_to_argb_hex',
]

# Cleanup namespace - remove helper variables if desired
# del _current_module, _all_members, _unwrapped_tools, _name, _obj, _func, _FunctionTool, inspect, sys