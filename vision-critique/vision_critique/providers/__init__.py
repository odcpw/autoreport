"""
Vision Provider Implementations

Pluggable vision model providers following a common interface.
Supports multiple backends: Anthropic Claude, OpenAI GPT-4V, Local LLMs.
"""

from .base import VisionProvider
from .anthropic import AnthropicProvider
from .openai import OpenAIProvider
from .local import LocalProvider

__all__ = [
    "VisionProvider",
    "AnthropicProvider",
    "OpenAIProvider",
    "LocalProvider",
]


def get_provider(provider_name: str, config) -> VisionProvider:
    """
    Factory function to get configured vision provider.

    Args:
        provider_name: One of "anthropic", "openai", or "local"
        config: Configuration object with API keys

    Returns:
        Configured vision provider instance

    Raises:
        ValueError: If provider name is unknown or not configured

    Example:
        provider = get_provider("anthropic", config)
        result = await provider.critique(screenshot_path, context)
    """
    if provider_name == "anthropic":
        if not config.has_anthropic():
            raise ValueError(
                "Anthropic API key not configured. "
                "Set ANTHROPIC_API_KEY in .env file"
            )
        return AnthropicProvider(api_key=config.anthropic_api_key)

    elif provider_name == "openai":
        if not config.has_openai():
            raise ValueError(
                "OpenAI API key not configured. "
                "Set OPENAI_API_KEY in .env file"
            )
        return OpenAIProvider(api_key=config.openai_api_key)

    elif provider_name == "local":
        return LocalProvider(
            host=config.ollama_host,
            model=config.ollama_model
        )

    else:
        raise ValueError(
            f"Unknown provider: {provider_name}. "
            f"Choose from: anthropic, openai, local"
        )
