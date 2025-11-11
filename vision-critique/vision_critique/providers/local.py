"""
Local LLM Vision Provider

Implements vision critique using local LLMs via Ollama.
Supports LLaVA, BakLLaVA, and other vision-capable local models.
"""

import base64
import json
import requests
from pathlib import Path
from typing import Optional

from ..models import CritiqueResult, Issue, Scores
from .base import VisionProvider


class LocalProvider(VisionProvider):
    """
    Vision provider using local LLMs through Ollama.

    Uses Ollama's API to run local vision models like LLaVA.
    Fully offline, privacy-preserving, no API costs.

    Features:
    - Runs completely offline
    - No API costs
    - Privacy-preserving (no data leaves your machine)
    - Supports multiple local vision models

    Requirements:
    - Ollama installed (https://ollama.ai/)
    - Vision model pulled (e.g., `ollama pull llava`)

    Example:
        provider = LocalProvider(
            host="http://localhost:11434",
            model="llava"
        )
        result = await provider.critique(screenshot_path, context)
    """

    def __init__(
        self,
        host: str = "http://localhost:11434",
        model: str = "llava"
    ):
        """
        Initialize local LLM provider.

        Args:
            host: Ollama server URL (default: http://localhost:11434)
            model: Vision model name (default: llava)
                   Run `ollama list` to see available models
        """
        self.host = host.rstrip("/")
        self.model = model

    @property
    def name(self) -> str:
        """Provider name for identification"""
        return "local"

    def is_available(self) -> bool:
        """
        Check if Ollama server is running and model is available.

        Returns:
            True if server is reachable, False otherwise
        """
        try:
            response = requests.get(f"{self.host}/api/tags", timeout=2)
            return response.status_code == 200
        except requests.RequestException:
            return False

    async def critique(
        self,
        screenshot_path: Path,
        context: Optional[dict] = None
    ) -> CritiqueResult:
        """
        Analyze screenshot using local LLM via Ollama.

        Sends screenshot to Ollama with structured prompt requesting
        JSON output with scores, issues, and suggestions.

        Args:
            screenshot_path: Path to PNG screenshot
            context: Optional context (project name, tab, description, etc.)

        Returns:
            CritiqueResult with complete analysis

        Raises:
            RuntimeError: If Ollama is not running or request fails

        Note:
            Local models may produce lower quality output than cloud APIs.
            Consider this a privacy-preserving alternative with trade-offs.
        """
        if not screenshot_path.exists():
            raise RuntimeError(f"Screenshot not found: {screenshot_path}")

        if not self.is_available():
            raise RuntimeError(
                f"Ollama server not reachable at {self.host}. "
                f"Make sure Ollama is running: `ollama serve`"
            )

        try:
            # Encode image as base64
            image_data = self._encode_image(screenshot_path)

            # Build prompt (simplified for local models)
            prompt = self._build_simplified_prompt(context)

            # Call Ollama API
            response = requests.post(
                f"{self.host}/api/generate",
                json={
                    "model": self.model,
                    "prompt": prompt,
                    "images": [image_data],
                    "stream": False,
                    "options": {
                        "temperature": 0.3,
                        "num_predict": 1024
                    }
                },
                timeout=120  # Local models can be slow
            )

            if response.status_code != 200:
                raise RuntimeError(f"Ollama API error: {response.text}")

            response_data = response.json()
            response_text = response_data.get("response", "")

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

        except requests.RequestException as e:
            raise RuntimeError(f"Failed to connect to Ollama: {str(e)}") from e
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

    def _build_simplified_prompt(self, context: Optional[dict] = None) -> str:
        """
        Build simplified prompt optimized for local models.

        Local models often have smaller context windows and may not
        handle complex structured prompts as well as cloud models.

        Args:
            context: Optional context dictionary

        Returns:
            Simplified prompt text
        """
        ctx = context or {}
        project_name = ctx.get("project_name", "the application")
        tab_name = ctx.get("tab_name", "this view")

        prompt = f"""Analyze this UI screenshot from {project_name} ({tab_name}).

Rate these aspects from 0-100:
1. Visual hierarchy and layout
2. Color and typography quality
3. UX and interaction design
4. Accessibility

Identify 2-3 specific issues and suggest improvements.

Respond in JSON format:
{{
  "scores": {{
    "visual_hierarchy": <score>,
    "color_typography": <score>,
    "ux_interaction": <score>,
    "accessibility": <score>,
    "overall": <average>
  }},
  "issues": [
    {{
      "dimension": "<dimension>",
      "severity": "<high|medium|low>",
      "description": "<problem>",
      "location": "<where>",
      "suggestion": "<fix>"
    }}
  ],
  "suggestions": ["<top improvement 1>", "<top improvement 2>"]
}}"""

        return prompt

    def _parse_response(self, response_text: str) -> dict:
        """
        Parse structured JSON response from local LLM.

        More lenient parsing for local models which may produce
        less perfectly formatted output.

        Args:
            response_text: Raw response text

        Returns:
            Parsed critique data dictionary

        Raises:
            RuntimeError: If response cannot be parsed
        """
        # Try to extract JSON
        if "```json" in response_text:
            start = response_text.find("```json") + 7
            end = response_text.find("```", start)
            json_text = response_text[start:end].strip()
        elif "```" in response_text:
            start = response_text.find("```") + 3
            end = response_text.find("```", start)
            json_text = response_text[start:end].strip()
        elif "{" in response_text and "}" in response_text:
            start = response_text.find("{")
            end = response_text.rfind("}") + 1
            json_text = response_text[start:end].strip()
        else:
            json_text = response_text.strip()

        try:
            data = json.loads(json_text)

            # Provide defaults for missing fields
            if "scores" not in data:
                data["scores"] = {
                    "visual_hierarchy": 50,
                    "color_typography": 50,
                    "ux_interaction": 50,
                    "accessibility": 50,
                    "overall": 50
                }

            if "overall" not in data["scores"]:
                scores = data["scores"]
                data["scores"]["overall"] = (
                    scores.get("visual_hierarchy", 50) * 0.30 +
                    scores.get("color_typography", 50) * 0.25 +
                    scores.get("ux_interaction", 50) * 0.25 +
                    scores.get("accessibility", 50) * 0.20
                )

            if "issues" not in data:
                data["issues"] = []

            if "suggestions" not in data:
                data["suggestions"] = []

            return data

        except json.JSONDecodeError as e:
            # Local models may fail to produce valid JSON
            # Return a minimal valid response
            return {
                "scores": {
                    "visual_hierarchy": 50,
                    "color_typography": 50,
                    "ux_interaction": 50,
                    "accessibility": 50,
                    "overall": 50
                },
                "issues": [{
                    "dimension": "accessibility",
                    "severity": "medium",
                    "description": f"Local model failed to analyze properly: {str(e)}",
                    "location": "General",
                    "suggestion": "Consider using a cloud provider for better results"
                }],
                "suggestions": ["Local model output was malformed - try cloud providers"]
            }
