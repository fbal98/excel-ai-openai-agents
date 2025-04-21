"""
Agent‑level hooks for memory, progressive‑summary, and workbook shape tracking.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from agents import Agent, AgentHooks, RunContextWrapper, Tool, FunctionTool

from .constants import WRITE_TOOLS
from .debounce_constants import SHAPE_SCAN_EVERY_N_WRITES, STRUCTURAL_WRITE_TOOLS, MAX_CONSECUTIVE_ERRORS
from .context import AppContext # WorkbookShape is implicitly available via AppContext

logger = logging.getLogger(__name__)

# ----------------------------------------------------------
#  Shared helper: decide if a tool result means "success”
# ----------------------------------------------------------
from .tools import _ensure_toolresult

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
    res = _ensure_toolresult(res)
    return bool(res.get("success", True))


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

    async def on_tool_end(  # noqa: D401
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        result: Any,
    ) -> None:
        # --- 1. Update Summary ---
        ok = _is_result_ok(result)
        outcome = "ok" if ok else "error"
        # Safely get tool name using the new helper
        tool_name = self._get_tool_name(tool)
        line = f"{tool_name} → {outcome}"

        app_ctx = context.context  # Ensure app_ctx is defined before use
        state = app_ctx.state

        # Shape update logic is handled by the inheriting class (ActionLoggingHooks)
        # --- 2. Update summary lines ---
        append_summary_line(app_ctx, line)
        # --- 4. (Optional) Log tool result for debugging ---
        # logger.debug(f"Tool '{tool_name}' result: {result}")

    def _get_tool_name(self, tool: Any) -> str:
        """Safely gets the name of a tool, handling raw functions."""
        if isinstance(tool, (Tool, FunctionTool)):
            return tool.name
        elif callable(tool):
            return getattr(tool, '__name__', str(tool))
        else:
            return str(tool)


class ActionLoggingHooks(SummaryHooks):
    """
    Extends SummaryHooks with a rolling action history saved in AppContext.
    Each tool call is logged with success/failure so the agent can
    see its recent behaviour. Also handles debounced workbook shape refresh.
    """

    async def on_tool_start(
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        args: Optional[Dict[str, Any]] = None,
    ) -> None:
        """Cache real arguments so we can log them accurately later."""
        tool_name = self._get_tool_name(tool) # Use helper here too
        logger.debug(f"HOOK: on_tool_start - tool type: {type(tool)}, tool: {tool}, tool_name: {tool_name}, args: {args}")

        # Store args only if provided by the SDK
        if args is not None:
            context.context.state["_last_args"] = args
        else:
            # Ensure _last_args exists even if SDK fails to provide args
            context.context.state["_last_args"] = {}
            logger.debug(
                f"SDK did not provide args to on_tool_start for tool '{tool_name}'. Logging args as empty."
            )

    async def on_tool_end(  # noqa: D401
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        result: Any,
    ) -> None:
        tool_name = self._get_tool_name(tool) # Use helper
        logger.debug(f"HOOK: on_tool_end - tool type: {type(tool)}, tool: {tool}, tool_name: {tool_name}, result: {result}")

        args = context.context.state.pop("_last_args", {})
        ok = _is_result_ok(result)

        context.context.record_action(
            tool=tool_name,
            args=args,
            result=result,
            ok=ok,
        )
        # Call parent (SummaryHooks) to retain summary logic *before* handling shape update
        # This will append the summary line using the safe tool_name getter.
        await super().on_tool_end(context, agent, tool, result)

        app_ctx = context.context  # Re‑use the underlying AppContext instance

        # ── Debounced workbook‑shape refresh ─────────────────────────
        if tool_name in WRITE_TOOLS:
            app_ctx.pending_write_count += 1
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
        else:
            logger.debug(
                f"Read tool '{tool_name}' executed. Skipping workbook shape refresh check."
            )


        # ── Self‑regulation: abort on repeated failures ──────────────────
        is_err = not ok # Use the result of _is_result_ok
        error_msg = ""
        if is_err and isinstance(result, dict):
             error_msg = result.get("error", "Unknown error")
        elif is_err:
             error_msg = f"Operation failed with result: {result}"


        if is_err:
            key = (tool_name, error_msg)
            if key == app_ctx.last_error_key and error_msg: # Only count consecutive if error msg is same
                app_ctx.consecutive_errors += 1
            else:
                app_ctx.consecutive_errors = 1
                app_ctx.last_error_key = key
        else:
            # Any successful tool call resets the error loop completely
            app_ctx.consecutive_errors = 0
            app_ctx.last_error_key = ("", "") # Reset key

        if app_ctx.consecutive_errors > MAX_CONSECUTIVE_ERRORS:
            logger.error(
                 f"Aborting run: Tool '{tool_name}' failed {app_ctx.consecutive_errors} times consecutively with error: {error_msg}"
             )
            # Raise specific exception rather than generic RuntimeError
            from agents.exceptions import MaxTurnsExceeded # Or a more specific error if available
            raise MaxTurnsExceeded(
                f"Aborting run: Tool '{tool_name}' failed {app_ctx.consecutive_errors} times consecutively."
            )