"""
Agent‑level hooks for memory, progressive‑summary, and workbook shape tracking.
"""

from __future__ import annotations

import copy # Import copy for deepcopy
import logging
import re
from typing import Any, Dict, Optional, TYPE_CHECKING

from agents import Agent, AgentHooks, RunContextWrapper, Tool, FunctionTool, Usage

# Import AppContext only for type hinting if necessary, avoid runtime circular dependency
if TYPE_CHECKING:
    from .context import AppContext

from .conversation_context import ConversationContext # Import the new helper
from .constants import WRITE_TOOLS
from .debounce_constants import SHAPE_SCAN_EVERY_N_WRITES, STRUCTURAL_WRITE_TOOLS, MAX_CONSECUTIVE_ERRORS
from .costs import dollars_for_usage
from .context import AppContext # WorkbookShape is implicitly available via AppContext

logger = logging.getLogger(__name__)

# ----------------------------------------------------------
#  Shared helper: decide if a tool result means "success”
# ----------------------------------------------------------
# Import from the new location in the tools package
from .tools.core_defs import _ensure_toolresult, ToolResult

def _is_result_ok(res: Any) -> bool: 
    """
    Examine result to determine if it indicates success, failure, or warning.
    
    This function strictly enforces the ToolResult format with explicit success/failure indicators.

    Success:
        • res is a dict with {"success": True} (explicit success)
    
    Failure:
        • res is a dict with {"success": False} (explicit failure)
    
    Any result that doesn't conform to the ToolResult format will be normalized first.
    Non-dict returns (None, str, list, etc.) are converted to {"success": True, "data": res}
    """
    # First, ensure res is properly normalized to ToolResult format
    tool_result = _ensure_toolresult(res)
    
    # Explicitly check for success/failure based on the 'success' key
    # This requires tools to explicitly signal their status
    if "success" in tool_result:
        return bool(tool_result["success"])
    
    # If we reach here, something went wrong with _ensure_toolresult
    # Log a warning since this shouldn't happen after normalization
    logger.warning(f"Result missing 'success' key after normalization: {res}. Assuming failure.")
    return False  # Safer to assume failure if status is unclear


def append_summary_line(app_ctx: "AppContext", line: str, max_lines: int = 15) -> None:
    """
    Append *line* to ``app_ctx.state["summary"]`` keeping only the last
    *max_lines* entries to bound prompt size.
    """
    prev = app_ctx.state.get("summary", "")
    lines = (prev.splitlines() + [line])[-max_lines:]
    app_ctx.state["summary"] = "\n".join(lines) # Keep for backward compat

    # Now also add to conversation history using the new context helper
    try:
        ConversationContext.emit_progress_line(app_ctx, line)
        # Pruning happens immediately after emitting a progress line
        ConversationContext.maybe_prune(app_ctx)
    except Exception as e:
        logger.error(f"Error emitting progress line or pruning: {e}", exc_info=True)


class SummaryHooks(AgentHooks):
    """
    - Emits progress lines and tool failures to conversation history.
    - Emits workbook shape deltas after write operations.
    - Refreshes ``ctx.shape`` snapshot via ``ctx.update_shape()`` after WRITE_TOOLS.
    - If shape refresh succeeds, dump the state (shape + agent_state) to JSON.
    - Tracks consecutive errors to prevent loops.
    - (Legacy) Appends short summary lines to ``ctx.state["summary"]``.
    - After calls to WRITE_TOOLS, refresh the ``ctx.shape`` snapshot via ``ctx.update_shape()``.
    - If shape refresh succeeds, dump the state (shape + agent_state) to JSON.
    """

    def _get_tool_name(self, tool: Tool) -> str:  # noqa: D401
        """
        Return a stable, human-readable identifier for *tool*.

        Parameters
        ----------
        tool : Union[Tool, str]
            The tool instance (FunctionTool, computer tool, etc.) or
            its plain-string name.

        Returns
        -------
        str
            The best-effort name to use in logs, summaries, and state.
        """
        # Case 1 – already a string (e.g., CLI pseudo-tool)
        if isinstance(tool, str):
            return tool

        # Case 2 – wrapped with @function_tool → exposes .name
        name = getattr(tool, "name", None)
        if name:
            return name

        # Case 3 – fall back to the underlying callable’s __name__
        return getattr(tool, "__name__", str(tool))
    async def on_tool_end(
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
            result: Any,
    ) -> None:
        tool_name = self._get_tool_name(tool)
        app_ctx = context.context # Get AppContext instance
        logger.debug(f"HOOK: on_tool_end - tool type: {type(tool)}, tool: {tool}, tool_name: {tool_name}, result: {result}")

        # Capture the shape before update
        old_shape_before = None
        if app_ctx.shape:
            try:
                old_shape_before = copy.deepcopy(app_ctx.shape)
                logger.debug("Captured old shape (v%s) before tool execution.", old_shape_before.version)
            except Exception as e:
                 logger.error(f"Failed to deepcopy old shape: {e}", exc_info=True)
                 # Proceed without old_shape if copy fails

        args = app_ctx.state.pop("_last_args", {})
        ok = _is_result_ok(result)

        app_ctx.record_action(
            tool=tool_name,
            args=args,
            result=result,
            ok=ok,
        )
        # --- Track current sheet for future turns ---
        if ok and tool_name == "create_sheet_tool":
            sheet_created = args.get("sheet_name")
            if sheet_created:
                context.context.state["current_sheet"] = sheet_created
        elif ok and tool_name == "get_active_sheet_name_tool" and isinstance(result, str):
            context.context.state["current_sheet"] = result
        # Call parent (SummaryHooks) to retain summary logic *before* handling shape update
        # This will append the summary line using the safe tool_name getter.
        await super().on_tool_end(context, agent, tool, result)

        app_ctx = context.context  # Re‑use the underlying AppContext instance

        # ── Debounced workbook‑shape refresh ─────────────────────────
        if tool_name in WRITE_TOOLS:
            app_ctx.pending_write_count += 1
            # --- Corrected Indentation Starts Here ---
            should_scan = (
                tool_name in STRUCTURAL_WRITE_TOOLS
                or app_ctx.pending_write_count >= SHAPE_SCAN_EVERY_N_WRITES
            )
            if should_scan:
                logger.debug(
                    "Refreshing workbook shape (tool=%s, pending=%s)…",
                    tool_name,
                    app_ctx.pending_write_count,
                )
                try:
                    # Store result of update_shape to check if it *actually* updated
                    shape_updated = app_ctx.update_shape(tool_name=tool_name) # Pass tool_name here

                    if shape_updated:
                        # Emit shape delta only if update_shape returned True (meaning change detected)
                        try:
                             ConversationContext.emit_shape_delta(app_ctx, old_shape_before, app_ctx.shape)
                             # Pruning might be needed after emitting potentially large shape deltas
                             ConversationContext.maybe_prune(app_ctx)
                        except Exception as e:
                             logger.error(f"Error emitting shape delta or pruning: {e}", exc_info=True)

                        app_ctx.pending_write_count = 0 # Reset counter only on successful scan
                        app_ctx.dump_state_to_json() # Dump state after successful update and potential delta emission
                    else:
                        # Update_shape might return False if scan failed or was skipped or no change detected.
                        # Keep the counter incrementing if the scan didn't actually happen or failed.
                        # If update_shape returned False due to an error, the error is logged within update_shape.
                        if app_ctx.shape: # Log current state if scan failed but shape exists
                            logger.debug(f"Shape update skipped or failed for tool {tool_name}. Current shape v{app_ctx.shape.version}, pending writes: {app_ctx.pending_write_count}")
                        else:
                            logger.debug(f"Shape update skipped or failed for tool {tool_name}. No current shape. Pending writes: {app_ctx.pending_write_count}")

                except Exception as e:
                    logger.error(f"Error during workbook shape refresh or state dump for tool {tool_name}: {e}")
                    # Decide if you want to reset the counter even on error, or let it keep growing.
                    # Resetting might prevent scans for a while if errors persist.
                    # Not resetting might trigger scans too often if errors are transient.
                    # Let's keep the counter growing for now to potentially retry scanning later.
            # The 'else' for 'if should_scan' is implicitly handled now: if it wasn't a write tool OR should_scan was False, this block is skipped.
        # --- Corrected Indentation Ends Here ---
        else: # This 'else' corresponds to 'if tool_name in WRITE_TOOLS:'
            logger.debug(
                f"Read tool '{tool_name}' executed. Skipping workbook shape refresh check."
            )

        # ── Self‑regulation: abort on repeated failures ──────────────────
        is_failure = not ok # Use the result of _is_result_ok to determine actual failure
        error_msg = ""
        tool_result = _ensure_toolresult(result) # Ensure we have ToolResult format

        # Extract error message ONLY if it was an actual failure
        if is_failure:
            # Ensure error_msg is a string representation
            error_msg = str(tool_result.get("error", f"Operation failed with result: {result}"))

            # Emit the failure to conversation history
            try:
                ConversationContext.emit_tool_failure(app_ctx, tool_name, error_msg)
                # Prune after emitting failure message as well
                ConversationContext.maybe_prune(app_ctx)
            except Exception as e:
                logger.error(f"Error emitting tool failure or pruning: {e}", exc_info=True)

            logger.debug(f"Tool '{tool_name}' failed. Error: {error_msg}. Consecutive count: {app_ctx.consecutive_errors + 1}")

            # Extract core error message by removing variable parts
            # This helps match similar errors with slight differences
            core_error = self._extract_core_error(error_msg)
            error_key = (tool_name, core_error)  # Use tool name + core error as the key
            
            # Check if this is the same type of error as the previous one
            last_tool, last_core_error = getattr(app_ctx, "last_error_key", ("", ""))
            
            # Check if error is similar to previous error (same tool, similar error message)
            if tool_name == last_tool and core_error == last_core_error and core_error:
                app_ctx.consecutive_errors += 1
                logger.debug(f"Consecutive error count increased to {app_ctx.consecutive_errors} for tool '{tool_name}'")
            else:
                # Different tool failed or different error, reset counter to 1
                app_ctx.consecutive_errors = 1
                app_ctx.last_error_key = error_key # Update the error key
                logger.debug(f"New error type detected, reset consecutive error count to 1 for tool '{tool_name}'")
        else:
             # Any successful tool call resets the error loop completely
             if app_ctx.consecutive_errors > 0:
                 logger.debug(f"Successful tool '{tool_name}' reset consecutive error count from {app_ctx.consecutive_errors}.")
             app_ctx.consecutive_errors = 0
             app_ctx.last_error_key = ("", "") # Reset key

        # Check if the consecutive error limit has been exceeded AFTER potentially incrementing
        if app_ctx.consecutive_errors > MAX_CONSECUTIVE_ERRORS:
            # Retrieve the error message associated with the last_error_key that triggered the limit
            last_fail_tool, last_fail_core_error = app_ctx.last_error_key
            logger.error(
                f"Aborting run: Tool '{last_fail_tool}' failed {app_ctx.consecutive_errors} times consecutively with similar errors."
            )
            # Raise specific exception rather than generic RuntimeError
            from agents.exceptions import MaxTurnsExceeded # Or a more specific error if available
            raise MaxTurnsExceeded(
                 f"Aborting run: Tool '{last_fail_tool}' failed {app_ctx.consecutive_errors} times consecutively with similar errors."
            )
            
    def _extract_core_error(self, error_msg: str) -> str:
        """
        Extract the core part of an error message by removing variable elements.
        This helps match similar errors that differ only in specific details.
        
        Examples:
        - "Sheet 'Sales' not found" -> "Sheet not found"
        - "Cannot find cell A3" -> "Cannot find cell"
        - "File 'data.xlsx' not found" -> "File not found"
        
        Args:
            error_msg: The full error message
            
        Returns:
            A simplified error message with variable parts removed
        """
        if not error_msg:
            return ""
            
        # Common patterns to normalize
        common_patterns = [
            # Replace quoted names
            (r"'[^']*'", "'NAME'"),
            (r'"[^"]*"', "'NAME'"),
            
            # Replace cell references
            (r"\b[A-Z]+\d+\b", "CELL"),
            (r"\b[A-Z]+\d+:[A-Z]+\d+\b", "RANGE"),
            
            # Replace numbers
            (r"\b\d+\b", "NUM"),
            
            # Replace paths
            (r"(?:/[^/\s]+)+", "PATH"),
            (r"(?:\\[^\\]+)+", "PATH"),
        ]
        
        # Apply each pattern
        normalized = error_msg
        for pattern, replacement in common_patterns:
            normalized = re.sub(pattern, replacement, normalized)
            
        # Trim whitespace and convert to lowercase for better matching
        normalized = normalized.strip().lower()
        
        # If normalization removed too much, use first 30 chars of original
        if len(normalized) < 5 and len(error_msg) > 0:
            return error_msg[:30].strip().lower()
            
        return normalized