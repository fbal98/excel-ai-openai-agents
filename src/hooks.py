"""
Agent‑level hooks for memory, progressive‑summary, and workbook shape tracking.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from agents import Agent, AgentHooks, RunContextWrapper, Tool

from .constants import WRITE_TOOLS
from .context import AppContext # WorkbookShape is implicitly available via AppContext

logger = logging.getLogger(__name__)


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
        outcome = "error" if isinstance(result, dict) and result.get("error") else "ok"
        # Safely get tool name, falling back to function name or string representation
        tool_name = getattr(tool, "name", getattr(tool, "__name__", str(tool)))
        line = f"{tool_name} → {outcome}"

        app_ctx = context.context  # Ensure app_ctx is defined before use
        state = app_ctx.state

        if tool_name in WRITE_TOOLS:
            # Pass tool_name to update_shape - the check is now inside update_shape
            update_succeeded = app_ctx.update_shape(tool_name=tool_name)
            # Note: update_succeeded will be True even if the scan was skipped inside update_shape
            # We only dump if a *real* update happened, which update_shape doesn't currently signal separately.
            # For now, we rely on update_shape's internal logging and dump if no error occurred.
        else:
            logger.debug(f"Read tool '{tool_name}' executed. Skipping workbook shape refresh.")
            # No need to call update_shape here, the check is inside it.
        # Shape update logic is now handled by the inheriting class (ActionLoggingHooks)
        # --- 2. Update summary lines ---
        prev_summary = state.get("summary", "")
        lines = (prev_summary.splitlines() + [line])[-25:]
        state["summary"] = "\n".join(lines)
        # --- 4. (Optional) Log tool result for debugging ---
        # logger.debug(f"Tool '{tool_name}' result: {result}")

class ActionLoggingHooks(SummaryHooks):
    """
    Extends SummaryHooks with a rolling action history saved in AppContext.
    Each tool call is logged with success/failure so the agent can
    see its recent behaviour.
    """

    async def on_tool_start(
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        args: Optional[Dict[str, Any]] = None,
    ) -> None:
        """Cache real arguments so we can log them accurately later."""
        # Store args only if provided by the SDK
        if args is not None:
            context.context.state["_last_args"] = args
        else:
            # Ensure _last_args exists even if SDK fails to provide args
            context.context.state["_last_args"] = {}
            logger.debug(
                # Safely get tool name for logging
                f"SDK did not provide args to on_tool_start for tool '{getattr(tool, 'name', getattr(tool, '__name__', str(tool)))}'. Logging args as empty."
            )

    async def on_tool_end(  # noqa: D401
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        result: Any,
    ) -> None:
        args = context.context.state.pop("_last_args", {})
        ok = not (isinstance(result, dict) and result.get("error"))
        # Safely get tool name for recording
        tool_name = getattr(tool, "name", getattr(tool, "__name__", str(tool)))
        context.context.record_action(
            tool=tool_name,
            args=args,
            result=result,
            ok=ok,
        )
        # Call parent (SummaryHooks) to retain summary logic *before* handling shape update
        # Note: SummaryHooks.on_tool_end itself doesn't handle shape update anymore.
        # We'll keep the shape update logic directly within ActionLoggingHooks.on_tool_end below the summary update.
        # Update Summary (from parent logic, simplified here)
        outcome = "error" if isinstance(result, dict) and result.get("error") else "ok"
        line = f"{tool_name} → {outcome}"
        app_ctx = context.context # Get the underlying AppContext instance
        state = app_ctx.state  # Access the state dictionary
        prev_summary = state.get("summary", "")
        lines = (prev_summary.splitlines() + [line])[-25:] # Keep summary bounded
        state["summary"] = "\n".join(lines)

        # Now handle shape update based on tool_name
        if tool_name in WRITE_TOOLS:
            logger.debug(f"Write tool '{tool_name}' executed. Refreshing workbook shape.")
            update_succeeded = app_ctx.update_shape(tool_name=tool_name)

            if update_succeeded:
                app_ctx.dump_state_to_json()
            else:
                logger.warning(f"State not dumped due to shape update failure after tool '{tool_name}'.")
        else:
            logger.debug(f"Read tool '{tool_name}' executed. Skipping workbook shape refresh.")