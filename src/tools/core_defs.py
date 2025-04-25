# src/tools/core_defs.py
import logging
from typing import Any, Dict, Optional, List, Union
from typing_extensions import TypedDict  # Use typing_extensions for compatibility
import asyncio
from agents import FunctionTool as _FunctionTool # Use alias to avoid potential name clash

# Define a custom exception for connection issues (though defined in excel_ops, good to have awareness)
# from ..excel_ops import ExcelConnectionError # Relative import if needed elsewhere

# --- Type Definitions ---

class ToolResult(TypedDict, total=False):
    """Standard schema every Excel agent tool must now return."""
    success: bool           # always present – True/False
    error: Optional[str]    # present on failure, None on success
    data: Optional[Any]     # optional payload (lists, scalars, etc.)

class SetCellValuesResult(ToolResult, total=False):
    """Backward‑compat alias kept for type hints."""
    pass

# Define a union type for a single cell's data
CellValue = Union[str, int, float, bool, None]

class FontStyle(TypedDict, total=False):
    name: Optional[str]
    size: Optional[float]
    bold: Optional[bool]
    italic: Optional[bool]
    vertAlign: Optional[str] # Note: Openpyxl uses 'vertAlign', xlwings uses 'vertical_alignment' potentially
    underline: Optional[str]
    strike: Optional[bool]
    color: Optional[str] # Expects ARGB Hex like 'FFRRGGBB'

class FillStyle(TypedDict, total=False):
    fill_type: Optional[str] # e.g., 'solid'
    start_color: Optional[str] # Expects ARGB Hex like 'FFRRGGBB'
    end_color: Optional[str] # Expects ARGB Hex like 'FFRRGGBB'

class BorderStyleDetails(TypedDict, total=False):
    style: Optional[str] # e.g., 'thin', 'medium'
    color: Optional[str] # Expects ARGB Hex like 'FFRRGGBB'

class BorderStyle(TypedDict, total=False):
    left: Optional[BorderStyleDetails]
    right: Optional[BorderStyleDetails]
    top: Optional[BorderStyleDetails]
    bottom: Optional[BorderStyleDetails]
    diagonal: Optional[BorderStyleDetails]
    diagonal_direction: Optional[int]
    outline: Optional[bool] # Or BorderStyleDetails for complex outline
    vertical: Optional[BorderStyleDetails]
    horizontal: Optional[BorderStyleDetails]

class AlignmentStyle(TypedDict, total=False):
    horizontal: Optional[str]
    vertical: Optional[str]
    wrap_text: Optional[bool]

class CellStyle(TypedDict, total=False):
    """Defines the structure for cell styling options for tools."""
    font: Optional[FontStyle]
    fill: Optional[FillStyle]
    border: Optional[BorderStyle]
    alignment: Optional[AlignmentStyle]
    number_format: Optional[str] # Added number format

class CellValueMap(TypedDict):
    """A mapping from cell addresses (e.g., 'A1') to values of any supported type."""
    # This acts as a type hint for dictionaries like {"A1": "value1", "B2": 123}
    pass

class WriteVerifyResult(TypedDict, total=False):
    """Result structure for write_and_verify_range_tool."""
    success: bool
    diff: Dict[str, Any] # Maps cell address to {'expected': Any, 'actual': Any} or {'error': str}

# --- Helper Functions ---

def _hex_argb_to_bgr_int(argb: str) -> int:
    """
    Convert an 8‑digit ARGB string ('FFRRGGBB' or '#FFRRGGBB') to an
    integer in BGR byte order for the Excel COM API.

    Requires the alpha channel; attempts to fix 6-digit RGB. Raises ValueError on other invalid formats.
    """
    s = str(argb).lstrip("#").upper()
    if len(s) != 8:
        # Try to handle 6-digit by adding FF alpha
        if len(s) == 6:
            s = "FF" + s
        else:
            raise ValueError(
                f"Color '{argb}' must be 8‑digit ARGB (e.g. 'FF3366CC' or '#FF3366CC') or 6-digit RGB."
            )

    # Drop alpha then swap R and B to get BGR
    r, g, b = s[2:4], s[4:6], s[6:8]
    return int(f"{b}{g}{r}", 16)


# Cached colour converter
_COLOR_CACHE: dict[str, int] = {}

def _to_bgr(argb: str) -> int:
    """
    Convert 8‑digit ARGB → BGR int with caching to avoid repeated
    `_hex_argb_to_bgr_int` calls inside tight loops.
    Handles potential errors during conversion, defaulting to black (0).
    """
    logger = logging.getLogger(__name__)
    if argb in _COLOR_CACHE:
        return _COLOR_CACHE[argb]
    try:
        bgr_int = _hex_argb_to_bgr_int(argb)
        _COLOR_CACHE[argb] = bgr_int
        return bgr_int
    except ValueError as e:
        logger.warning(f"Invalid color format '{argb}' for BGR conversion: {e}. Using black (0).")
        _COLOR_CACHE[argb] = 0 # Cache 0 for invalid format
        return 0 # Default to black

def _bgr_int_to_argb_hex(color_int: Optional[int]) -> str:
    """
    Convert a BGR integer used by Excel back to an 8-digit ARGB hex string (FFRRGGBB).
    Returns 'FF000000' (black) if input is None or invalid.
    """
    if color_int is None or not isinstance(color_int, int) or color_int < 0:
        return "FF000000" # Default to opaque black for invalid input

    try:
        # Mask to 24-bit and ensure 6-hex-digit string, padded with zeros
        bgr_hex = f"{color_int & 0xFFFFFF:06X}"
        b, g, r = bgr_hex[0:2], bgr_hex[2:4], bgr_hex[4:6]
        # Assume full opacity (FF)
        return f"FF{r}{g}{b}"
    except Exception:
        return "FF000000" # Fallback on any conversion error


def _normalise_rows(columns: list[Any], rows: list[list[Any]]) -> list[list[Any]]:
    """Pad or truncate each row so they're exactly as wide as `columns`."""
    width = len(columns)
    fixed: list[list[Any]] = []
    for r in rows:
        if not isinstance(r, list): # Handle cases where a row might not be a list
            r = [r] + [None] * (width - 1) if width > 0 else [] # Convert non-list row
        row_len = len(r)
        if row_len < width:
            # pad short rows
            fixed.append(r + [None] * (width - row_len))
        elif row_len > width:
            # truncate long rows
            fixed.append(r[:width])
        else:
            fixed.append(r)
    return fixed

# --- Tool Result Unification ---

def _ensure_toolresult(res: Any) -> ToolResult:
    """Normalize arbitrary returns → ToolResult."""
    if isinstance(res, dict) and res.get("success") is not None:
        # Check if 'error' should be None when success is True
        if res.get("success") is True and "error" not in res:
             res["error"] = None
        # Check if 'error' should exist when success is False
        elif res.get("success") is False and "error" not in res:
            res["error"] = "Operation failed without explicit error message."
        return res  # already compliant or fixed

    if isinstance(res, dict):  # legacy {'error': '...'} or other dicts without 'success'
        err_msg = res.get("error")
        if err_msg:
             # Check if 'data' exists, otherwise set to None
             data_payload = {k: v for k, v in res.items() if k != "error"} or None
             return {"success": False, "error": str(err_msg), "data": data_payload}
        else:
             # If it's a dict without 'success' and without 'error', treat as successful data
             return {"success": True, "error": None, "data": res}

    # Handle non-dict results
    if res in (True, None):
        return {"success": True, "error": None, "data": res}
    if res is False:
        return {"success": False, "error": "Operation returned False", "data": None}

    # Treat any other non-dict, non-bool, non-None result as successful data
    return {"success": True, "error": None, "data": res}


def _wrap_tool_result(func):
    """Decorator that enforces ToolResult on any sync/async tool."""
    if asyncio.iscoroutinefunction(func):
        async def _async_wrapper(*args, **kwargs):
            result = await func(*args, **kwargs)
            return _ensure_toolresult(result)
        _async_wrapper.__name__ = func.__name__
        _async_wrapper.__doc__ = func.__doc__ # Preserve docstring
        # Expose .name for SDK compatibility if not already a FunctionTool instance
        if not isinstance(func, _FunctionTool):
             _async_wrapper.name = func.__name__
        return _async_wrapper
    else:
        def _sync_wrapper(*args, **kwargs):
            result = func(*args, **kwargs)
            return _ensure_toolresult(result)
        _sync_wrapper.__name__ = func.__name__
        _sync_wrapper.__doc__ = func.__doc__ # Preserve docstring
        # Expose .name for SDK compatibility if not already a FunctionTool instance
        if not isinstance(func, _FunctionTool):
             _sync_wrapper.name = func.__name__
        return _sync_wrapper