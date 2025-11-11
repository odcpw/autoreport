"""
Vision Critique - UI Quality Analysis Tool

A vision-based UI quality critique tool that coding agents can use
to get feedback on visual design, accessibility, and user experience.

Supports multiple vision providers:
- Anthropic Claude
- OpenAI GPT-4V
- Local LLMs (Ollama/LLaVA)
"""

from .models import CritiqueResult, Issue, Scores
from .scorer import UIScorer
from .capture import ScreenshotCapturer

__version__ = "0.1.0"
__all__ = ["CritiqueResult", "Issue", "Scores", "UIScorer", "ScreenshotCapturer"]
