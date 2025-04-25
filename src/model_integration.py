# src/model_integration.py
import os
import logging
from agents import Agent # For type hinting in factory function
# Removed: OpenAIChatCompletionsModel, set_tracing_disabled, set_tracing_export_api_key, AsyncOpenAI
from .model_config import get_active_provider, get_provider_config
from .context import AppContext # For type hinting Agent[AppContext]

logger = logging.getLogger(__name__)

def get_model_string() -> str:
    """
    Gets the model identifier string formatted for the SDK's native LiteLLM integration
    (e.g., "litellm/gemini/gemini-pro").
    """
    provider = get_active_provider()
    config = get_provider_config(provider)

    model_name = config.get("model")

    logger.info(f"Configuring model string for provider: {provider}")

    if not model_name:
        raise ValueError(f"Model name for provider '{provider}' not found in environment variables.")

    # For OpenAI, use the model name directly (SDK handles it)
    if provider == "openai":
        logger.info(f"Using direct model name for OpenAI: {model_name}")
        # Note: Ensure OPENAI_API_KEY is set for this to work.
        # Tracing should work automatically if OPENAI_API_KEY is set.
        return model_name
    else:
        # For other providers, use the "litellm/" prefix
        # LiteLLM requires specific formatting like provider/model
        # We might need more robust mapping here depending on .env values vs litellm expectations
        # Example: if GEMINI_MODEL is "gemini-2.5-pro-preview-03-25", litellm might expect "gemini/gemini-2.5-pro-preview-03-25"
        # Example: if OPENROUTER_MODEL is "llama3.1:3b", litellm expects "openrouter/meta-llama/Llama-3.1-8B-Instruct" (needs mapping)

        # Basic prefixing for now, might need refinement
        formatted_model_name = f"litellm/{provider}/{model_name}"

        # Potential OpenRouter Mapping (simple example, may need improvement)
        if provider == "openrouter" and ":" in model_name:
             # Try a common pattern: replace ':' with '/' or use known mappings
             # This is fragile - a proper mapping dict would be better.
             model_base = model_name.split(':')[0]
             if model_base == "llama3.1": # User's specific example
                 formatted_model_name = "litellm/openrouter/meta-llama/Llama-3.1-8B-Instruct" # Example mapping
                 logger.warning(f"Attempting mapping for OpenRouter '{model_name}' -> '{formatted_model_name}'. Verify this is correct.")
             else:
                 # Default prefixing if no specific mapping found
                 formatted_model_name = f"litellm/openrouter/{model_name}"

        logger.info(f"Using LiteLLM model string: {formatted_model_name}")
        # Ensure the relevant API key (e.g., GEMINI_API_KEY) is set in the environment
        # for litellm to pick up automatically.
        return formatted_model_name

def create_excel_assistant_agent() -> Agent[AppContext]:
    """Creates the Excel Assistant agent instance with the currently configured model string."""
    from .agent_core import _dynamic_instructions, _validated_agent_tools
    from .plan_hooks import PlanCaptureHooks  # Use enhanced hooks
    from .context import AppContext

    logger.debug("Attempting to create Excel Assistant Agent instance...")
    try:
        model_string = get_model_string() # Get the configured model identifier string
        agent = Agent[AppContext](
            name="Excel Assistant",
            instructions=_dynamic_instructions,
            hooks=PlanCaptureHooks(),
            tools=_validated_agent_tools,
            model=model_string # Use the model string directly
        )
        logger.info(f"Excel Assistant Agent created successfully with model string: '{model_string}' (Provider: {get_active_provider()})")
        return agent
    except Exception as e:
        logger.critical(f"FATAL: Failed to create Excel Assistant Agent: {e}", exc_info=True)
        raise RuntimeError("Could not create the core Excel Assistant Agent.") from e