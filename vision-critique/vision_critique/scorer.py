"""
UI Scorer Orchestrator

Main orchestration module that coordinates screenshot capture,
vision analysis, and automated checks to produce comprehensive UI critique.
"""

import asyncio
from pathlib import Path
from typing import Optional

from .capture import ScreenshotCapturer
from .checks import check_contrast, check_sizing
from .models import CritiqueResult, Issue, Config
from .providers.base import VisionProvider


class UIScorer:
    """
    Orchestrates complete UI quality scoring workflow.

    Coordinates:
    1. Screenshot capture (via Playwright)
    2. Vision model analysis (via provider)
    3. Automated accessibility checks
    4. Result aggregation and scoring

    Example:
        config = load_config()
        provider = get_provider("anthropic", config)
        scorer = UIScorer(provider, config)

        result = await scorer.score(
            url="file:///path/to/index.html",
            tab="photosorter"
        )

        print(f"Score: {result.scores.overall}/100")
    """

    def __init__(self, provider: VisionProvider, config: Config):
        """
        Initialize UI scorer.

        Args:
            provider: Configured vision provider (Anthropic/OpenAI/Local)
            config: Configuration with viewport settings, etc.
        """
        self.provider = provider
        self.config = config
        self.capturer = ScreenshotCapturer(
            viewport={
                "width": config.viewport_width,
                "height": config.viewport_height
            }
        )

    async def score(
        self,
        url: str,
        tab: Optional[str] = None,
        selector: Optional[str] = None,
        wait_for: Optional[str] = None,
        css_path: Optional[Path] = None
    ) -> CritiqueResult:
        """
        Perform complete UI quality scoring.

        Workflow:
        1. Capture screenshot of specified view
        2. Send to vision model for analysis
        3. Run automated accessibility checks (parallel)
        4. Combine results into final critique
        5. Return structured result

        Args:
            url: URL to analyze (file:// or http(s)://)
            tab: Optional tab name for context
            selector: Optional CSS selector to click before capture
            wait_for: Optional CSS selector to wait for
            css_path: Optional CSS file path for automated checks

        Returns:
            CritiqueResult with scores, issues, and suggestions

        Raises:
            RuntimeError: If capture or analysis fails

        Example:
            result = await scorer.score(
                url="file:///home/user/project/index.html",
                tab="photosorter",
                selector="[data-tab='photosorter']",
                wait_for=".photosorter__main",
                css_path=Path("project/css/main.css")
            )
        """
        # 1. Capture screenshot
        screenshot_path = await self.capturer.capture(
            url=url,
            selector=selector,
            wait_for=wait_for
        )

        # 2. Build context for vision model
        context = self._build_context(url, tab)

        # 3. Run vision analysis + automated checks in parallel
        vision_task = self.provider.critique(screenshot_path, context)
        checks_task = self._run_automated_checks(css_path)

        vision_result, checks_result = await asyncio.gather(
            vision_task,
            checks_task
        )

        # 4. Enhance result with automated check findings
        enhanced_result = self._enhance_with_checks(vision_result, checks_result)

        return enhanced_result

    async def score_multiple(
        self,
        url: str,
        tabs: list[dict]
    ) -> dict[str, CritiqueResult]:
        """
        Score multiple tabs/views efficiently.

        Args:
            url: Base URL
            tabs: List of tab configurations, each with:
                 - name: Tab identifier
                 - selector: CSS selector to click
                 - wait_for: CSS selector to wait for (optional)

        Returns:
            Dictionary mapping tab names to CritiqueResults

        Example:
            results = await scorer.score_multiple(
                url="file:///project/index.html",
                tabs=[
                    {"name": "import", "selector": "[data-tab='import']"},
                    {"name": "photosorter", "selector": "[data-tab='photosorter']"},
                ]
            )

            for tab_name, result in results.items():
                print(f"{tab_name}: {result.scores.overall}/100")
        """
        results = {}

        for tab in tabs:
            result = await self.score(
                url=url,
                tab=tab["name"],
                selector=tab.get("selector"),
                wait_for=tab.get("wait_for")
            )
            results[tab["name"]] = result

        return results

    async def _run_automated_checks(self, css_path: Optional[Path]) -> dict:
        """
        Run automated accessibility checks.

        Runs checks in parallel for speed.

        Args:
            css_path: Optional CSS file to analyze

        Returns:
            Dictionary with check results
        """
        # Run checks (these are synchronous but we await for consistency)
        contrast_result = await asyncio.to_thread(check_contrast, css_path)
        sizing_result = await asyncio.to_thread(check_sizing, css_path)

        return {
            "contrast": contrast_result,
            "sizing": sizing_result
        }

    def _build_context(self, url: str, tab: Optional[str]) -> dict:
        """
        Build context dictionary for vision model.

        Extracts useful context from URL and tab name to help
        vision model understand what it's looking at.

        Args:
            url: Page URL
            tab: Optional tab name

        Returns:
            Context dictionary
        """
        context = {
            "project_name": "AutoBericht" if "AutoBericht" in url else "Application",
        }

        if tab:
            context["tab_name"] = tab
            context["description"] = self._get_tab_description(tab)

        return context

    def _get_tab_description(self, tab: str) -> str:
        """
        Get description of tab purpose based on name.

        Provides context to vision model about user goals.

        Args:
            tab: Tab name

        Returns:
            Description string
        """
        descriptions = {
            "import": "Data upload and configuration interface",
            "photosorter": "Photo tagging and categorization tool",
            "autobericht": "Report editor with markdown editing",
            "export": "Export options and session management"
        }

        return descriptions.get(tab.lower(), f"{tab} interface")

    def _enhance_with_checks(
        self,
        vision_result: CritiqueResult,
        checks_result: dict
    ) -> CritiqueResult:
        """
        Enhance vision result with automated check findings.

        Combines vision model analysis with automated accessibility checks
        to produce more comprehensive critique.

        Strategy:
        - If automated checks find violations, add them as high-severity issues
        - Adjust accessibility score based on automated results
        - Keep vision model's scores for other dimensions

        Args:
            vision_result: Result from vision provider
            checks_result: Results from automated checks

        Returns:
            Enhanced CritiqueResult
        """
        enhanced_issues = list(vision_result.issues)

        # Add automated check violations as issues
        contrast = checks_result.get("contrast", {})
        if not contrast.get("pass", True) and contrast.get("violations"):
            for violation in contrast["violations"][:2]:  # Top 2 contrast issues
                enhanced_issues.append(Issue(
                    dimension="accessibility",
                    severity="high",
                    description=violation["description"],
                    location="CSS",
                    suggestion=violation["suggestion"]
                ))

        sizing = checks_result.get("sizing", {})
        if not sizing.get("pass", True) and sizing.get("violations"):
            for violation in sizing["violations"][:2]:  # Top 2 sizing issues
                enhanced_issues.append(Issue(
                    dimension="accessibility",
                    severity="high",
                    description=violation["description"],
                    location="CSS",
                    suggestion=violation["suggestion"]
                ))

        # Adjust accessibility score based on automated checks
        contrast_score = contrast.get("score", 100)
        sizing_score = sizing.get("score", 100)
        automated_accessibility = (contrast_score + sizing_score) / 2

        # Blend with vision model's assessment (70% vision, 30% automated)
        adjusted_accessibility = (
            vision_result.scores.accessibility * 0.70 +
            automated_accessibility * 0.30
        )

        # Recalculate overall score
        adjusted_overall = (
            vision_result.scores.visual_hierarchy * 0.30 +
            vision_result.scores.color_typography * 0.25 +
            vision_result.scores.ux_interaction * 0.25 +
            adjusted_accessibility * 0.20
        )

        # Create enhanced result
        enhanced_result = CritiqueResult(
            scores=vision_result.scores.model_copy(update={
                "accessibility": round(adjusted_accessibility, 1),
                "overall": round(adjusted_overall, 1)
            }),
            issues=enhanced_issues,
            suggestions=vision_result.suggestions,
            screenshot_path=vision_result.screenshot_path,
            provider=vision_result.provider,
            context={
                **vision_result.context,
                "automated_checks": {
                    "contrast": contrast.get("pass", True),
                    "sizing": sizing.get("pass", True)
                }
            }
        )

        return enhanced_result
