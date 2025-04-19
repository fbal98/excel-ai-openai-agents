"""
Agent‑level hooks for memory & progressive‑summary support.
"""

from __future__ import annotations

from typing import Any

from agents import Agent, AgentHooks, RunContextWrapper, Tool

from .context import AppContext


class SummaryHooks(AgentHooks):
    """
    After every tool call, append a *very* short line to ``ctx.state["summary"]`` so
    later turns can be given a bounded recap (“progressive summarisation”).
    """

    async def on_tool_end(  # noqa: D401
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        tool: Tool,
        result: Any,
    ) -> None:
        # Build one‑liner: e.g.  set_cell_value_tool → ok   /   set_range_style_tool → error
        outcome = "error" if isinstance(result, dict) and result.get("error") else "ok"
        line = f"{tool.name} → {outcome}"

        state = context.context.state  # the scratch‑pad requested by the user
        prev = state.get("summary", "")
        # Keep it bounded to the last ~25 lines (≈ 1‑2 k‑tokens max)
        lines = (prev.splitlines() + [line])[-25:]
        state["summary"] = "\n".join(lines)