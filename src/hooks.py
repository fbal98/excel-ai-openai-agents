"""
Agent‑level hooks for memory, progressive‑summary, and workbook shape tracking.
"""

from __future__ import annotations

import logging
from typing import Any, Dict

from agents import Agent, AgentHooks, RunContextWrapper, Tool

from .context import AppContext # WorkbookShape is implicitly available via AppContext

logger = logging.getLogger(__name__)

# Define tools that modify the workbook structure or content significantly enough to warrant a shape refresh
# Note: Style changes are included as they can impact layout perception. Reads are excluded.
WRITE_TOOLS = {
    "open_workbook_tool",       # Opens a different book
    "set_cell_value_tool",
    "set_range_style_tool",
    "set_cell_style_tool",
    "create_sheet_tool",
    "delete_sheet_tool",
    "merge_cells_range_tool",
    "unmerge_cells_range_tool",
    "set_row_height_tool",
    "set_column_width_tool",
    "set_columns_widths_tool",
    "set_cell_formula_tool",
    "set_cell_values_tool",
    "set_table_tool",
    "insert_table_tool",
    "set_rows_tool",
    "set_columns_tool",
    "set_named_ranges_tool",    # Modifies named ranges definitions
    "copy_paste_range_tool",    # Can paste values, formulas, or formats
    "write_and_verify_range_tool",
    "revert_snapshot_tool",     # Reverts to a previous state
    # Exclude save_workbook_tool as it doesn't change the structure/content itself
}

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
        line = f"{tool.name} → {outcome}"

        app_ctx = context.context # Get the underlying AppContext instance
        state = app_ctx.state     # Access the state dictionary

        prev_summary = state.get("summary", "")
        # Keep summary bounded
        lines = (prev_summary.splitlines() + [line])[-25:]
        state["summary"] = "\n".join(lines)

        # --- 2. Update Workbook Shape and Dump State if a write tool was used ---
        if tool.name in WRITE_TOOLS:
            logger.debug(f"Write tool '{tool.name}' executed. Refreshing workbook shape.")
            # Use the helper method on AppContext for updating shape
            update_succeeded = app_ctx.update_shape()

            # --- 3. Dump state to JSON if shape update succeeded ---
            if update_succeeded:
                app_ctx.dump_state_to_json() # Use the helper method
            else:
                logger.warning(f"State not dumped due to shape update failure after tool '{tool.name}'.")

        # --- 4. (Optional) Log tool result for debugging ---
        # logger.debug(f"Tool '{tool.name}' result: {result}")

class ActionLoggingHooks(SummaryHooks):
    """
    Extends SummaryHooks with a rolling action history saved in AppContext.
    Each tool call is logged with success/failure so the agent can
    see its recent behaviour.
    """

    async def on_tool_start(  # noqa: D401
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        args: Dict[str, Any],
    ) -> None:
        """Cache real arguments so we can log them accurately later."""
        context.context.state["_last_args"] = args

    async def on_tool_end(  # noqa: D401
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        result: Any,
    ) -> None:
        args = context.context.state.pop("_last_args", {})
        ok = not (isinstance(result, dict) and result.get("error"))
        context.context.record_action(
            tool=tool.name,
            args=args,
            result=result,
            ok=ok,
        )
        # Call parent to retain summary + shape‑update behaviour
        await super().on_tool_end(context, agent, tool, result)