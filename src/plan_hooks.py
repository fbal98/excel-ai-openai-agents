"""Plan-capture hook for small-context models.

Adds the first line of the assistant’s final reply to ctx.state["summary"]
so that the plan is threaded back into the next prompt.
"""

from __future__ import annotations

import logging
from typing import Any

from agents import Agent, RunContextWrapper, AgentHooks
from .hooks import SummaryHooks, append_summary_line  # reuse existing helpers
from .context import AppContext

logger = logging.getLogger(__name__)


class PlanCaptureHooks(SummaryHooks):
    """
    Extends SummaryHooks by also recording the first line of the model’s
    final reply at the end of every turn.
    """

    async def on_agent_end(  # type: ignore[override]
        self,
        context: RunContextWrapper[AppContext],
        agent: Agent,
        result: Any,
    ) -> None:
        # Call parent (no-op in SummaryHooks) first, in case it grows later.
        if hasattr(super(), "on_agent_end"):
            await super().on_agent_end(context, agent, result)  # pytype: disable=attribute-error

        try:
            if isinstance(result, str) and result.strip():
                first_line = result.strip().splitlines()[0][:200]  # hard cap length
                append_summary_line(context.context, first_line, max_lines=30)
                logger.debug("PlanCaptureHooks: stored plan line -> %s", first_line)
        except Exception as e:  # pragma: no cover
            logger.warning("PlanCaptureHooks failed to capture plan line: %s", e)