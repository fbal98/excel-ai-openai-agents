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

# Override rates for special models or newer models not yet in litellm:
EXTRA_RATES = {
    # Gemini models
    "gemini-2.5-flash-preview-04-17": (0.000075 / 1000, 0.00030 / 1000),
    
    # OpenAI models (fallbacks if litellm pricing fails)
    "gpt-4o": (0.01 / 1000, 0.03 / 1000),        # $0.01 per 1K input, $0.03 per 1K output
    "gpt-4-turbo": (0.01 / 1000, 0.03 / 1000),   # $0.01 per 1K input, $0.03 per 1K output
    "gpt-4": (0.03 / 1000, 0.06 / 1000),         # $0.03 per 1K input, $0.06 per 1K output
    "gpt-3.5-turbo": (0.0005 / 1000, 0.0015 / 1000)  # $0.0005 per 1K input, $0.0015 per 1K output
}

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

    # Determine the initial model name to use for litellm costing
    raw_model_name = get_model_name_for_costing(model_name_from_agent=model_name_from_agent)

    if not raw_model_name:
        logger.error("Could not determine model name for cost calculation.")
        return 0.0

    # Handle model name formatting for litellm cost calculation
    model_for_costing = raw_model_name
    
    # Special handling for Gemini models
    if "gemini" in raw_model_name.lower():
        # For Gemini models, litellm might need the full model name
        # Check if the name is already prefixed
        if '/' in raw_model_name and "gemini" in raw_model_name.lower():
            parts = raw_model_name.split('/')
            if len(parts) > 1:
                # Extract just the model name without provider prefix
                model_for_costing = parts[-1]
                logger.debug(f"Using Gemini model name '{model_for_costing}' for litellm costing.")
        logger.debug(f"Using Gemini model format for costing: '{model_for_costing}'")
    # Handle other models with prefixes
    elif '/' in raw_model_name:
        parts = raw_model_name.split('/')
        if len(parts) > 1:
            model_for_costing = parts[-1] # Take the last part after the last '/'
            logger.debug(f"Stripped provider prefix from '{raw_model_name}', using '{model_for_costing}' for litellm costing.")
        else:
             logger.warning(f"Model name '{raw_model_name}' contains '/' but couldn't extract base name properly. Using original.")

    logger.debug(f"Calculating cost for model: '{model_for_costing}' (Input: {prompt_tokens}, Output: {completion_tokens})")

    try:
        # Use litellm.completion_cost with the potentially stripped name
        cost = completion_cost(
            model=model_for_costing,
            input_tokens=prompt_tokens,  # litellm expects input_tokens, not prompt_tokens
            output_tokens=completion_tokens,  # litellm might expect output_tokens instead of completion_tokens
        )

        # completion_cost returns None if model pricing is not found
        if cost is None:
            logger.warning(f"LiteLLM could not find pricing data for model '{model_for_costing}'. Trying fallback pricing.")
            
            # Try to use our predefined fallback rates
            model_key = None
            
            # Check for exact match first
            if model_for_costing in EXTRA_RATES:
                model_key = model_for_costing
            # Then try base model name (without version specifiers)
            else:
                # For names like "gpt-4-1106-preview", extract the base model "gpt-4"
                base_name = model_for_costing.split('-')[0:2]  # Take first two parts, e.g., ["gpt", "4"]
                if len(base_name) >= 2:
                    base_model = '-'.join(base_name)  # Create "gpt-4"
                    # Check if we have rates for the base model
                    if base_model in EXTRA_RATES:
                        model_key = base_model
                
            # Use fallback pricing if we found a match
            if model_key:
                try:
                    input_rate, output_rate = EXTRA_RATES[model_key]
                    cost = (prompt_tokens * input_rate) + (completion_tokens * output_rate)
                    logger.info(f"Used fallback pricing for model '{model_for_costing}' (matched to '{model_key}'): ${cost:.6f}")
                    return round(cost, 6)
                except Exception as e:
                    logger.error(f"Fallback pricing calculation failed: {e}")
            
            # If all else fails, use generic estimates based on model family
            try:
                if "gpt-4" in model_for_costing.lower():
                    # Generic GPT-4 family rate
                    input_rate, output_rate = 0.01 / 1000, 0.03 / 1000
                elif "gpt-3" in model_for_costing.lower():
                    # Generic GPT-3.5 family rate
                    input_rate, output_rate = 0.0005 / 1000, 0.0015 / 1000
                elif "gemini" in model_for_costing.lower():
                    # Generic Gemini rate
                    input_rate, output_rate = 0.000075 / 1000, 0.00030 / 1000
                else:
                    # If we can't identify the model family, assume it's a basic model
                    logger.warning(f"Using generic pricing for unknown model type: {model_for_costing}")
                    input_rate, output_rate = 0.0005 / 1000, 0.0015 / 1000
                
                cost = (prompt_tokens * input_rate) + (completion_tokens * output_rate)
                logger.info(f"Used generic fallback pricing for {model_for_costing}: ${cost:.6f}")
                return round(cost, 6)
            except Exception as e:
                logger.error(f"Generic fallback pricing calculation failed: {e}")
                
            # If all fallback strategies fail    
            return 0.0

        logger.info(f"Calculated cost via litellm: ${cost:.6f}")
        return round(cost, 6) # Sane default precision: micro-dollar resolution

    except Exception as e:
        # Catch potential errors during litellm lookup (e.g., model not recognized)
        logger.error(f"Error calculating cost for model '{model_for_costing}' using litellm: {e}", exc_info=False) # Set exc_info=True for full trace
        return 0.0 # Return 0 on error