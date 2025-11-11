"""
Configuration Management

Loads configuration from .env files and provides typed config objects.
Handles API keys, provider settings, and default values.
"""

import os
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

from .models import Config


def load_config(env_file: Optional[Path] = None) -> Config:
    """
    Load configuration from .env file and environment variables.

    Searches for .env file in:
    1. Provided env_file path
    2. Current directory
    3. User's home directory

    Environment variables override .env file values.

    Args:
        env_file: Optional path to .env file

    Returns:
        Config object with all settings

    Example:
        config = load_config()
        if config.has_anthropic():
            provider = AnthropicProvider(config.anthropic_api_key)
    """
    # Load .env file
    if env_file and env_file.exists():
        load_dotenv(env_file)
    elif Path(".env").exists():
        load_dotenv(".env")
    elif (Path.home() / ".env").exists():
        load_dotenv(Path.home() / ".env")

    # Build config from environment
    config = Config(
        anthropic_api_key=os.getenv("ANTHROPIC_API_KEY"),
        openai_api_key=os.getenv("OPENAI_API_KEY"),
        ollama_host=os.getenv("OLLAMA_HOST", "http://localhost:11434"),
        ollama_model=os.getenv("OLLAMA_MODEL", "llava"),
        vision_provider=os.getenv("VISION_PROVIDER", "anthropic"),
        viewport_width=int(os.getenv("VIEWPORT_WIDTH", "1920")),
        viewport_height=int(os.getenv("VIEWPORT_HEIGHT", "1080"))
    )

    return config
