"""
Anthropic Claude Vision Provider

Implements vision critique using Claude's vision capabilities.
Supports Claude 3+ models with vision understanding.
"""

import base64
import json
from pathlib import Path
from typing import Optional

import anthropic

from ..models import CritiqueResult, Issue, Scores
from .base import VisionProvider


class AnthropicProvider(VisionProvider):
    """
    Vision provider using Anthropic's Claude models.

    Uses Claude 3.5 Sonnet or later models with vision capabilities
    to analyze UI screenshots and provide structured feedback.

    Features:
    - High-quality vision understanding
    - Structured JSON output
    - Consistent scoring
    - Detailed issue identification

    Example:
        provider = AnthropicProvider(api_key="sk-ant-...")
        result = await provider.critique(screenshot_path, context)
    """

    def __init__(
        self,
        api_key: str,
        model: str = "claude-3-5-sonnet-20241022"
    ):
        """
        Initialize Anthropic provider.

        Args:
            api_key: Anthropic API key (get from https://console.anthropic.com/)
            model: Claude model to use (default: claude-3-5-sonnet-20241022)
                   Must be a vision-capable model (Claude 3+)
        """
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model
        self._api_key = api_key

    @property
    def name(self) -> str:
        """Provider name for identification"""
        return "anthropic"

    def is_available(self) -> bool:
        """
        Check if Anthropic provider is configured.

        Returns:
            True if API key is set, False otherwise
        """
        return self._api_key is not None and len(self._api_key) > 0

    async def critique(
        self,
        screenshot_path: Path,
        context: Optional[dict] = None
    ) -> CritiqueResult:
        """
        Analyze screenshot using Claude vision model.

        Sends screenshot to Claude with structured prompt requesting
        JSON output with scores, issues, and suggestions.

        Args:
            screenshot_path: Path to PNG screenshot
            context: Optional context (project name, tab, description, etc.)

        Returns:
            CritiqueResult with complete analysis

        Raises:
            RuntimeError: If API call fails or response is invalid
        """
        if not screenshot_path.exists():
            raise RuntimeError(f"Screenshot not found: {screenshot_path}")

        try:
            # Encode image as base64
            image_data = self._encode_image(screenshot_path)

            # Build prompt
            prompt = self._build_critique_prompt(context)

            # Call Claude API
            response = self.client.messages.create(
                model=self.model,
                max_tokens=2048,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": image_data
                            }
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ]
                }]
            )

            # Extract text response
            response_text = response.content[0].text

            # Parse JSON from response
            critique_data = self._parse_response(response_text)

            # Build CritiqueResult
            result = CritiqueResult(
                scores=Scores(**critique_data["scores"]),
                issues=[Issue(**issue) for issue in critique_data.get("issues", [])],
                suggestions=critique_data.get("suggestions", []),
                screenshot_path=str(screenshot_path),
                provider=self.name,
                context=context or {}
            )

            return result

        except anthropic.APIError as e:
            raise RuntimeError(f"Anthropic API error: {str(e)}") from e
        except Exception as e:
            raise RuntimeError(f"Failed to critique screenshot: {str(e)}") from e

    def _encode_image(self, image_path: Path) -> str:
        """
        Encode image file as base64 string.

        Args:
            image_path: Path to image file

        Returns:
            Base64-encoded image data
        """
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    def _parse_response(self, response_text: str) -> dict:
        """
        Parse structured JSON response from Claude.

        Handles both direct JSON and JSON embedded in markdown code blocks.

        Args:
            response_text: Raw response text from Claude

        Returns:
            Parsed critique data dictionary

        Raises:
            RuntimeError: If response cannot be parsed as valid JSON
        """
        # Try to extract JSON from markdown code blocks
        if "```json" in response_text:
            start = response_text.find("```json") + 7
            end = response_text.find("```", start)
            json_text = response_text[start:end].strip()
        elif "```" in response_text:
            start = response_text.find("```") + 3
            end = response_text.find("```", start)
            json_text = response_text[start:end].strip()
        else:
            json_text = response_text.strip()

        try:
            data = json.loads(json_text)

            # Validate required fields
            if "scores" not in data:
                raise ValueError("Response missing 'scores' field")

            # Calculate overall score if not provided
            if "overall" not in data["scores"]:
                scores = data["scores"]
                data["scores"]["overall"] = (
                    scores.get("visual_hierarchy", 0) * 0.30 +
                    scores.get("color_typography", 0) * 0.25 +
                    scores.get("ux_interaction", 0) * 0.25 +
                    scores.get("accessibility", 0) * 0.20
                )

            return data

        except json.JSONDecodeError as e:
            raise RuntimeError(
                f"Failed to parse Claude response as JSON: {str(e)}\n"
                f"Response text: {json_text[:500]}"
            ) from e
        except Exception as e:
            raise RuntimeError(f"Invalid response format: {str(e)}") from e
