"""excel_ai.costs
===================

Cost calculation using litellm based on the active provider configuration.
Relies on litellm's internal pricing data.
"""

from __future__ import annotations
import logging
import os
from agents import Usage
from litellm import completion_cost # Use litellm function

# Import helper to get config, avoiding direct dependency on active state module if possible
# However, we might need the model name string associated with the run.
from .model_config import get_provider_config, get_active_provider

logger = logging.getLogger(__name__)

# No longer needed
# PRICES: dict[str, dict[str, float]] = { ... }
# def _get_rates(model: str) -> tuple[float, float]: ...

def get_model_name_for_costing(model_name_from_agent: str | None = None) -> str | None:
    """
    Gets the model name ready for litellm costing.
    Primarily uses the name passed from the agent, which might already be prefixed.
    Falls back to config only if agent name is missing.
    """
    if model_name_from_agent:
        # Assume the model name from agent (e.g., "gpt-4..." or "litellm/gemini/...")
        # is the correct identifier to pass to litellm.completion_cost
        # litellm should handle prefixed names directly for cost lookup.
        logger.debug(f"Using model name directly from agent for costing: '{model_name_from_agent}'")
        return model_name_from_agent

    # Fallback: Get default model from config if agent didn't provide one (less ideal)
    logger.warning("No model name provided from agent context, falling back to active provider's default model for costing.")
    target_provider = get_active_provider()
    config = get_provider_config(target_provider)
    model_name = config.get("model")

    if not model_name:
         logger.error(f"Fallback failed: Could not get default model name for provider '{target_provider}'")
         return None

    # Format the fallback name if necessary (similar to get_model_string logic)
    if target_provider == "openai":
        return model_name
    else:
        # Basic prefixing for fallback, may need refinement as in get_model_string
        formatted_model_name = f"litellm/{target_provider}/{model_name}"
        if target_provider == "openrouter" and ":" in model_name:
             model_base = model_name.split(':')[0]
             if model_base == "llama3.1":
                 formatted_model_name = "litellm/openrouter/meta-llama/Llama-3.1-8B-Instruct"
                 logger.warning(f"Using mapped fallback for OpenRouter '{model_name}' -> '{formatted_model_name}' for costing.")
             else:
                  formatted_model_name = f"litellm/openrouter/{model_name}"
        logger.debug(f"Using formatted fallback model name for costing: '{formatted_model_name}'")
        return formatted_model_name


def dollars_for_usage(usage: Usage, model_name_from_agent: str | None = None) -> float:
    """Convert an agents.Usage instance to USD using litellm.

    Parameters
    ----------
    usage
        The token usage record from `agents`.
    model_name_from_agent : Optional[str]
        The specific model name string used in the run (e.g., from agent.model.model).
        This is preferred for accurate costing.

    Returns
    -------
    float
        The calculated cost in US dollars, or 0.0 if calculation fails.
    """
    # Defensive: ensure attributes exist.
    prompt_tokens = getattr(usage, "input_tokens", 0) or 0
    completion_tokens = getattr(usage, "output_tokens", 0) or 0

    # Determine the model name to use for litellm costing
    # Pass the name directly from the agent instance
    model_for_costing = get_model_name_for_costing(model_name_from_agent=model_name_from_agent)

    if not model_for_costing:
        logger.error("Could not determine model name for cost calculation.")
        return 0.0

    logger.debug(f"Calculating cost for model: '{model_for_costing}' (Input: {prompt_tokens}, Output: {completion_tokens})")

    try:
        # Use litellm.completion_cost
        cost = completion_cost(
            model=model_for_costing,
            prompt_tokens=prompt_tokens,
            completion_tokens=completion_tokens,
        )

        # completion_cost returns None if model pricing is not found
        if cost is None:
             logger.warning(f"LiteLLM could not find pricing data for model '{model_for_costing}'. Cost assumed to be $0.00.")
             return 0.0

        logger.debug(f"Calculated cost via litellm: ${cost:.6f}")
        return round(cost, 6) # Sane default precision: micro-dollar resolution

    except Exception as e:
        # Catch potential errors during litellm lookup (e.g., model not recognized)
        logger.error(f"Error calculating cost for model '{model_for_costing}' using litellm: {e}", exc_info=False) # Set exc_info=True for full trace
        return 0.0 # Return 0 on error