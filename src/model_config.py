# src/model_config.py
import os
import logging

logger = logging.getLogger(__name__)

_DEFAULT_PROVIDER = "gemini"
# Read the default provider from .env, falling back to "gemini"
_active_provider = os.getenv("DEFAULT_MODEL_PROVIDER", _DEFAULT_PROVIDER).lower()
_SUPPORTED_PROVIDERS = ["openai", "gemini", "openrouter"]

# Validate initial provider selection
if _active_provider not in _SUPPORTED_PROVIDERS:
    logger.warning(f"DEFAULT_MODEL_PROVIDER ('{_active_provider}') in .env is not supported. Using default '{_DEFAULT_PROVIDER}'.")
    _active_provider = _DEFAULT_PROVIDER


def _is_provider_configured(provider_name: str) -> bool:
    """Check if essential env vars are set for a provider."""
    provider_name = provider_name.lower()
    key_var = f"{provider_name.upper()}_API_KEY"
    model_var = f"{provider_name.upper()}_MODEL"
    # Base URL is often optional or defaults, primarily check key and model
    key_exists = bool(os.getenv(key_var))
    model_exists = bool(os.getenv(model_var))

    if not key_exists:
        logger.debug(f"Provider '{provider_name}': Missing environment variable '{key_var}'.")
    if not model_exists:
        logger.debug(f"Provider '{provider_name}': Missing environment variable '{model_var}'.")

    return key_exists and model_exists

def list_available_providers() -> dict[str, bool]:
    """Returns a dictionary of supported providers and their configuration status."""
    return {p: _is_provider_configured(p) for p in _SUPPORTED_PROVIDERS}

def set_active_provider(provider_name: str):
    """Sets the active LLM provider."""
    global _active_provider
    provider_name = provider_name.lower()
    if provider_name not in _SUPPORTED_PROVIDERS:
        raise ValueError(f"Unsupported provider: '{provider_name}'. Supported: {', '.join(_SUPPORTED_PROVIDERS)}")

    if not _is_provider_configured(provider_name):
         # Allow setting even if not fully configured, but warn
         logger.warning(f"Setting provider to '{provider_name}', but it may not be fully configured in .env. Operations might fail.")

    _active_provider = provider_name
    logger.info(f"Active model provider set to: {_active_provider}")

def get_active_provider() -> str:
    """Gets the currently active LLM provider."""
    # Re-validate just in case state became inconsistent
    global _active_provider
    if _active_provider not in _SUPPORTED_PROVIDERS:
        logger.warning(f"Active provider '{_active_provider}' is invalid/unsupported. Resetting to default '{_DEFAULT_PROVIDER}'.")
        _active_provider = _DEFAULT_PROVIDER

    return _active_provider

def get_provider_config(provider_name: str | None = None) -> dict[str, str | None]:
    """Gets API key, model, and base URL for a given provider or the active one."""
    target_provider = provider_name if provider_name else get_active_provider()
    target_provider = target_provider.lower()

    if target_provider not in _SUPPORTED_PROVIDERS:
        logger.error(f"Attempted to get config for unsupported provider: {target_provider}")
        return {"api_key": None, "model": None, "base_url": None}

    api_key = os.getenv(f"{target_provider.upper()}_API_KEY")
    model = os.getenv(f"{target_provider.upper()}_MODEL")
    base_url = os.getenv(f"{target_provider.upper()}_BASE_URL") # May be None

    # Add default base URLs if missing and applicable
    if target_provider == "openai" and not base_url:
        base_url = "https://api.openai.com/v1"
    elif target_provider == "openrouter" and not base_url:
        base_url = "https://openrouter.ai/api/v1"
    # Note: Gemini base URL might point to Google AI Studio or Vertex AI, less standard default.
    # The one in the user's .env might be specific.

    return {"api_key": api_key, "model": model, "base_url": base_url}


# Initial check for the active provider on import
if not _is_provider_configured(_active_provider):
    logger.warning(f"Initial active provider '{_active_provider}' is not fully configured in .env. Check { _active_provider.upper()}_API_KEY and {_active_provider.upper()}_MODEL.")

logger.info(f"Model config initialized. Active provider: {_active_provider}")