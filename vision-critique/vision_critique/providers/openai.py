"""
OpenAI GPT-4V Vision Provider

Implements vision critique using OpenAI's GPT-4 with vision capabilities.
Supports GPT-4 Turbo with Vision and later models.
"""

import base64
import json
from pathlib import Path
from typing import Optional

import openai

from ..models import CritiqueResult, Issue, Scores
from .base import VisionProvider


class OpenAIProvider(VisionProvider):
    """
    Vision provider using OpenAI's GPT-4 with Vision.

    Uses GPT-4 Turbo with Vision or gpt-4o models to analyze
    UI screenshots and provide structured feedback.

    Features:
    - Strong vision understanding
    - Structured JSON output
    - Fast response times
    - Reliable API

    Example:
        provider = OpenAIProvider(api_key="sk-...")
        result = await provider.critique(screenshot_path, context)
    """

    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4o"
    ):
        """
        Initialize OpenAI provider.

        Args:
            api_key: OpenAI API key (get from https://platform.openai.com/api-keys)
            model: OpenAI model to use (default: gpt-4o)
                   Must be a vision-capable model (gpt-4-vision-preview, gpt-4o, etc.)
        """
        self.client = openai.OpenAI(api_key=api_key)
        self.model = model
        self._api_key = api_key

    @property
    def name(self) -> str:
        """Provider name for identification"""
        return "openai"

    def is_available(self) -> bool:
        """
        Check if OpenAI provider is configured.

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
        Analyze screenshot using GPT-4 with Vision.

        Sends screenshot to GPT-4V with structured prompt requesting
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

            # Call OpenAI API
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_data}",
                                "detail": "high"
                            }
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ]
                }],
                max_tokens=2048,
                temperature=0.3  # Lower temperature for more consistent output
            )

            # Extract text response
            response_text = response.choices[0].message.content

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

        except openai.APIError as e:
            raise RuntimeError(f"OpenAI API error: {str(e)}") from e
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
        Parse structured JSON response from GPT-4V.

        Handles both direct JSON and JSON embedded in markdown code blocks.

        Args:
            response_text: Raw response text from GPT-4V

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
                f"Failed to parse GPT-4V response as JSON: {str(e)}\n"
                f"Response text: {json_text[:500]}"
            ) from e
        except Exception as e:
            raise RuntimeError(f"Invalid response format: {str(e)}") from e
