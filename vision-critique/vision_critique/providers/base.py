"""
Base Vision Provider Interface

Abstract base class defining the contract for vision model providers.
All providers must implement this interface for consistent behavior.
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional

from ..models import CritiqueResult


class VisionProvider(ABC):
    """
    Abstract base class for vision model providers.

    All vision providers (Anthropic, OpenAI, Local) must implement
    this interface to ensure consistent behavior and easy swapping.

    Subclasses must implement:
    - critique(): Analyze screenshot and return structured feedback
    - is_available(): Check if provider is configured and ready
    - name: Property returning provider name
    """

    @property
    @abstractmethod
    def name(self) -> str:
        """
        Provider name for logging and identification.

        Returns:
            Provider name (e.g., "anthropic", "openai", "local")
        """
        pass

    @abstractmethod
    async def critique(
        self,
        screenshot_path: Path,
        context: Optional[dict] = None
    ) -> CritiqueResult:
        """
        Analyze screenshot and return structured UI quality critique.

        This is the main method that vision providers implement.
        It sends the screenshot to the vision model with a structured
        prompt and parses the response into a CritiqueResult.

        Args:
            screenshot_path: Path to screenshot PNG file
            context: Optional context about what's being analyzed:
                    - project_name: Name of the project
                    - tab_name: Which tab/view is shown
                    - description: Purpose of this view
                    - user_goals: What users are trying to accomplish

        Returns:
            CritiqueResult with scores, issues, and suggestions

        Raises:
            RuntimeError: If API call fails or response is invalid

        Example:
            context = {
                "project_name": "AutoBericht",
                "tab_name": "photosorter",
                "description": "Photo tagging interface",
                "user_goals": "Tag photos efficiently with keyboard shortcuts"
            }
            result = await provider.critique(screenshot_path, context)
        """
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """
        Check if provider is configured and ready to use.

        Validates that:
        - API keys are set (for cloud providers)
        - Service is reachable (for local providers)

        Returns:
            True if provider can be used, False otherwise

        Example:
            if not provider.is_available():
                print(f"{provider.name} is not available")
        """
        pass

    def _build_critique_prompt(self, context: Optional[dict] = None) -> str:
        """
        Build standardized vision critique prompt.

        This prompt is designed to elicit structured, actionable feedback
        from vision models. Can be overridden by subclasses for
        provider-specific optimizations.

        Args:
            context: Optional context dictionary

        Returns:
            Structured prompt text optimized for vision models
        """
        ctx = context or {}
        project_name = ctx.get("project_name", "the application")
        tab_name = ctx.get("tab_name", "this view")
        description = ctx.get("description", "")
        user_goals = ctx.get("user_goals", "")

        prompt = f"""You are an expert UI/UX evaluator analyzing {project_name}.

**Current View:** {tab_name}
"""

        if description:
            prompt += f"**Purpose:** {description}\n"

        if user_goals:
            prompt += f"**User Goals:** {user_goals}\n"

        prompt += """
**Your Task:** Analyze this screenshot and provide structured feedback on UI quality.

**Evaluation Dimensions:**

1. **Visual Hierarchy** (0-100)
   - Is the focal point clear?
   - Does the layout guide the eye naturally?
   - Are element sizes appropriate for their importance?
   - Is spacing consistent and purposeful?

2. **Color & Typography** (0-100)
   - Do text elements meet WCAG AA contrast (4.5:1 minimum)?
   - Is there a coherent color palette?
   - Are font sizes readable (minimum 14px for body text)?
   - Is typography consistent and hierarchical?

3. **UX & Interaction** (0-100)
   - Are touch targets at least 44×44px?
   - Is the information hierarchy clear at a glance?
   - Are related controls grouped logically?
   - Is the interface intuitive?

4. **Accessibility** (0-100)
   - Would this work for users with visual impairments?
   - Are there sufficient visual cues?
   - Is the interface keyboard-navigable (visual indicators)?

**Output Format:**

Provide your analysis in JSON format:

```json
{
  "scores": {
    "visual_hierarchy": <0-100>,
    "color_typography": <0-100>,
    "ux_interaction": <0-100>,
    "accessibility": <0-100>,
    "overall": <weighted average>
  },
  "issues": [
    {
      "dimension": "<dimension name>",
      "severity": "<high|medium|low>",
      "description": "<clear explanation of the problem>",
      "location": "<where in the UI>",
      "suggestion": "<actionable fix recommendation>"
    }
  ],
  "suggestions": [
    "<prioritized improvement 1>",
    "<prioritized improvement 2>",
    "<prioritized improvement 3>"
  ]
}
```

**Guidelines:**
- Be specific and actionable
- Focus on the most impactful issues
- Prioritize high-severity problems
- Provide concrete CSS/design recommendations
- Consider WCAG 2.1 AA standards for accessibility

Analyze the screenshot now:"""

        return prompt
