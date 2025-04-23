"""excel_ai.costs
===================

Centralised pricing helpers for OpenAI token-billing.

* Keep **all** model pricing in **one** place so Finance/Ops can update a
  single file when OpenAI changes their prices.
* Remains cycle-free: **only** depends on the public `agents.Usage` dataclass
  (no local-project imports).

The cost calculation follows OpenAIʼs public billing model:
    cost = (input_tokens / 1000) × price_per_1k_prompt_tokens
          + (output_tokens / 1000) × price_per_1k_completion_tokens
"""

from __future__ import annotations

from agents import Usage


# ---------------------------------------------------------------------------#
#  Public price table (USD per 1 000 tokens)
#  Extend / adjust as OpenAI publishes new price lists.
# ---------------------------------------------------------------------------#
PRICES: dict[str, dict[str, float]] = {
    # Default model for this project (used as "safe fallback”)
    "gpt-4.1-mini": {"input": 0.01, "output": 0.03},
    # Other common models – figures copied from OpenAI price list (Apr 2025)
    "gpt-4o-2025-04-09": {"input": 0.005, "output": 0.015},
    "gpt-4-turbo": {"input": 0.01, "output": 0.03},
    "gpt-4": {"input": 0.03, "output": 0.06},
    "gpt-3.5-turbo": {"input": 0.001, "output": 0.002},
}


def _get_rates(model: str) -> tuple[float, float]:
    """Return *(input_rate, output_rate)* for *model*, falling back to default."""
    rates = PRICES.get(model, PRICES["gpt-4.1-mini"])
    return float(rates["input"]), float(rates["output"])


# ---------------------------------------------------------------------------#
#  Public helper
# ---------------------------------------------------------------------------#
def dollars_for_usage(usage: "Usage", model: str) -> float:
    """Convert an :class:`agents.Usage` instance to **USD** for *model*.

    Parameters
    ----------
    usage
        The token usage record returned by *agents*.
    model
        The model name used for the calls.  Unknown names fall back to
        ``PRICES['gpt-4.1-mini']`` so that callers never crash.

    Returns
    -------
    float
        The calculated cost in US dollars.
    """
    input_rate, output_rate = _get_rates(model)

    # Defensive: ensure attributes exist (older SDK versions may differ).
    in_tokens = getattr(usage, "input_tokens", 0) or 0
    out_tokens = getattr(usage, "output_tokens", 0) or 0

    cost = (in_tokens / 1000.0) * input_rate + (out_tokens / 1000.0) * output_rate
    return round(cost, 6)  # Sane default precision: micro-dollar resolution