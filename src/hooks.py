"""
Agent‑level hooks for memory, progressive‑summary, and workbook shape tracking.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from agents import Agent, AgentHooks, RunContextWrapper, Tool, FunctionTool, Usage

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
    Treat everything as success unless it explicitly signals failure.

    Success:
        • res is a dict with {"success": True}
        • res is a dict with no "error"
        • res is any non‑dict value (None, str, list, int, …)

    Failure:
        • res is a dict containing a truthy "error"
        • res is dict with {"success": False}
    """
    # Ensure res is converted to ToolResult format first
    tool_result = _ensure_toolresult(res)
    # Now check the 'success' key, defaulting to True if somehow missing (though _ensure should add it)
    return tool_result.get("success", True)


def append_summary_line(app_ctx: "AppContext", line: str, max_lines: int = 15) -> None:
    """
    Append *line* to ``app_ctx.state["summary"]`` keeping only the last
    *max_lines* entries to bound prompt size.
    """
    prev = app_ctx.state.get("summary", "")
    lines = (prev.splitlines() + [line])[-max_lines:]
    app_ctx.state["summary"] = "\n".join(lines)


class SummaryHooks(AgentHooks):
    """
    - After every tool call, append a short line to ``ctx.state["summary"]``.
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
        logger.debug(f"HOOK: on_tool_end - tool type: {type(tool)}, tool: {tool}, tool_name: {tool_name}, result: {result}")

        args = context.context.state.pop("_last_args", {})
        ok = _is_result_ok(result)

        context.context.record_action(
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
                    if app_ctx.update_shape(tool_name=tool_name): # Pass tool_name here
                        app_ctx.pending_write_count = 0 # Reset counter only on successful scan
                        app_ctx.dump_state_to_json()
                    else:
                        # Update_shape might return False if scan failed or was skipped.
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
            error_msg = tool_result.get("error", f"Operation failed with result: {result}")
            logger.debug(f"Tool '{tool_name}' failed. Error: {error_msg}. Consecutive count: {app_ctx.consecutive_errors + 1}")
            key = (tool_name, error_msg) # Key includes tool name and specific error msg
            # Increment counter only if the *same tool* fails with the *same error message* consecutively
            if key == app_ctx.last_error_key and error_msg:
                app_ctx.consecutive_errors += 1
            else:
                # Different tool failed, or different error from the same tool, reset counter to 1
                app_ctx.consecutive_errors = 1
                app_ctx.last_error_key = key # Update the last error key
        else:
             # Any successful tool call resets the error loop completely
             if app_ctx.consecutive_errors > 0:
                 logger.debug(f"Successful tool '{tool_name}' reset consecutive error count from {app_ctx.consecutive_errors}.")
             app_ctx.consecutive_errors = 0
             app_ctx.last_error_key = ("", "") # Reset key

        # Check if the consecutive error limit has been exceeded AFTER potentially incrementing
        if app_ctx.consecutive_errors > MAX_CONSECUTIVE_ERRORS:
            # Retrieve the error message associated with the last_error_key that triggered the limit
            last_fail_tool, last_fail_msg = app_ctx.last_error_key
            logger.error(
                f"Aborting run: Tool '{last_fail_tool}' failed {app_ctx.consecutive_errors} times consecutively with the same error: {last_fail_msg}"
            )
            # Raise specific exception rather than generic RuntimeError
            from agents.exceptions import MaxTurnsExceeded # Or a more specific error if available
            raise MaxTurnsExceeded(
                 f"Aborting run: Tool '{last_fail_tool}' failed {app_ctx.consecutive_errors} times consecutively."
            )