"""
Data Models for Vision Critique

Type-safe Pydantic models for all data structures.
LLM-readable with clear documentation and validation.
"""

from datetime import datetime
from pathlib import Path
from typing import Literal, Optional

from pydantic import BaseModel, Field, field_validator


class Issue(BaseModel):
    """
    A specific UI quality issue identified by vision or automated checks.

    Attributes:
        dimension: Which aspect of UI quality this issue affects
        severity: How important this issue is to fix
        description: Clear explanation of the problem
        location: Where in the UI this issue occurs (optional)
        suggestion: Actionable recommendation for fixing the issue
    """

    dimension: Literal[
        "visual_hierarchy",  # Layout, spacing, focal points
        "color_typography",  # Colors, fonts, contrast
        "ux_interaction",    # Usability, interaction patterns
        "accessibility"      # WCAG compliance, a11y
    ]
    severity: Literal["high", "medium", "low"]
    description: str = Field(min_length=10)
    location: Optional[str] = Field(default=None, description="CSS selector or description of location")
    suggestion: str = Field(min_length=10, description="Actionable fix recommendation")

    def __str__(self) -> str:
        """Human-readable string representation"""
        severity_emoji = {"high": "🔴", "medium": "🟡", "low": "🟢"}
        return f"{severity_emoji[self.severity]} [{self.dimension}] {self.description}"


class Scores(BaseModel):
    """
    Comprehensive UI quality scores across multiple dimensions.

    All scores are 0-100, where:
    - 90-100: Excellent
    - 75-89: Good
    - 60-74: Acceptable
    - 40-59: Needs improvement
    - 0-39: Poor

    Attributes:
        visual_hierarchy: Layout, spacing, visual flow, element sizing
        color_typography: Color harmony, contrast, font choices, readability
        ux_interaction: Usability, interaction patterns, cognitive load
        accessibility: WCAG compliance, a11y best practices
        overall: Weighted average of all dimensions
    """

    visual_hierarchy: float = Field(ge=0, le=100, description="Layout and visual flow quality")
    color_typography: float = Field(ge=0, le=100, description="Color and typography quality")
    ux_interaction: float = Field(ge=0, le=100, description="UX and interaction quality")
    accessibility: float = Field(ge=0, le=100, description="Accessibility compliance")
    overall: float = Field(ge=0, le=100, description="Overall weighted score")

    @field_validator("*")
    @classmethod
    def round_scores(cls, v: float) -> float:
        """Round all scores to 1 decimal place for readability"""
        return round(v, 1)

    def get_grade(self) -> str:
        """Get letter grade for overall score"""
        if self.overall >= 90:
            return "A"
        elif self.overall >= 75:
            return "B"
        elif self.overall >= 60:
            return "C"
        elif self.overall >= 40:
            return "D"
        else:
            return "F"


class CritiqueResult(BaseModel):
    """
    Complete vision critique result with scores, issues, and recommendations.

    This is the main output format that coding agents will consume.

    Attributes:
        scores: Quantitative scores across all dimensions
        issues: List of specific problems identified
        suggestions: Prioritized list of actionable improvements
        screenshot_path: Path to the analyzed screenshot
        timestamp: When this critique was generated
        provider: Which vision model was used
        context: Additional context about what was analyzed
    """

    scores: Scores
    issues: list[Issue] = Field(default_factory=list)
    suggestions: list[str] = Field(default_factory=list, description="Prioritized improvement suggestions")
    screenshot_path: str = Field(description="Path to analyzed screenshot")
    timestamp: str = Field(default_factory=lambda: datetime.now().isoformat())
    provider: str = Field(default="unknown", description="Vision model provider used")
    context: dict = Field(default_factory=dict, description="Analysis context metadata")

    @property
    def critical_issues(self) -> list[Issue]:
        """Get only high-severity issues"""
        return [issue for issue in self.issues if issue.severity == "high"]

    @property
    def has_critical_issues(self) -> bool:
        """Check if there are any high-severity issues"""
        return len(self.critical_issues) > 0

    def summary(self) -> str:
        """Generate a human-readable summary"""
        grade = self.scores.get_grade()
        total_issues = len(self.issues)
        critical = len(self.critical_issues)

        summary = f"Grade: {grade} ({self.scores.overall}/100)\n"
        summary += f"Issues: {total_issues} total ({critical} critical)\n"

        if self.suggestions:
            summary += "\nTop suggestions:\n"
            for i, suggestion in enumerate(self.suggestions[:3], 1):
                summary += f"  {i}. {suggestion}\n"

        return summary


class Config(BaseModel):
    """
    Configuration for vision critique tool.

    Loaded from .env file and optional YAML config.

    Attributes:
        anthropic_api_key: Anthropic API key (optional)
        openai_api_key: OpenAI API key (optional)
        ollama_host: Ollama server URL for local LLMs (optional)
        ollama_model: Model name for Ollama (default: llava)
        vision_provider: Which provider to use by default
        viewport_width: Screenshot viewport width
        viewport_height: Screenshot viewport height
    """

    anthropic_api_key: Optional[str] = None
    openai_api_key: Optional[str] = None
    ollama_host: str = "http://localhost:11434"
    ollama_model: str = "llava"
    vision_provider: Literal["anthropic", "openai", "local"] = "anthropic"
    viewport_width: int = Field(default=1920, ge=800, le=3840)
    viewport_height: int = Field(default=1080, ge=600, le=2160)

    def has_anthropic(self) -> bool:
        """Check if Anthropic is configured"""
        return self.anthropic_api_key is not None and len(self.anthropic_api_key) > 0

    def has_openai(self) -> bool:
        """Check if OpenAI is configured"""
        return self.openai_api_key is not None and len(self.openai_api_key) > 0
