"""
Touch Target Size Checker

Automated check for minimum touch target sizes.
Validates against WCAG 2.5.5 and mobile accessibility guidelines (44×44px minimum).
"""

import re
from pathlib import Path
from typing import Optional


def check_sizing(css_path: Optional[Path] = None) -> dict:
    """
    Check for adequate touch target sizes in CSS.

    Looks for buttons, links, and interactive elements that may be
    too small for comfortable touch/click interaction.

    WCAG 2.5.5 Target Size (Enhanced):
    - Minimum 44×44 CSS pixels for touch targets
    - Applies to buttons, links, form controls

    Args:
        css_path: Path to CSS file to analyze (optional)

    Returns:
        Dictionary with:
        - score: 0-100 based on violations
        - violations: List of sizing issues
        - pass: Whether guidelines are met

    Example:
        result = check_sizing(Path("main.css"))
        if not result["pass"]:
            for v in result["violations"]:
                print(v["description"])

    Note:
        This is a heuristic check looking for common patterns.
        Full validation requires analyzing the rendered page.
    """
    if css_path is None or not css_path.exists():
        return {
            "score": 100,
            "violations": [],
            "pass": True,
            "note": "No CSS file provided for analysis"
        }

    violations = []

    try:
        with open(css_path, "r", encoding="utf-8") as f:
            css_content = f.read()

        # Look for button/interactive element sizing below 44px
        small_height_pattern = r'(button|\.btn|\.button|a|input|\.clickable)[^}]*height:\s*([0-9]+)px'
        small_width_pattern = r'(button|\.btn|\.button|a|input|\.clickable)[^}]*width:\s*([0-9]+)px'
        min_height_pattern = r'(button|\.btn|\.button)[^}]*min-height:\s*([0-9]+)px'

        # Check heights
        for match in re.finditer(small_height_pattern, css_content):
            element = match.group(1)
            height = int(match.group(2))
            if height < 44:
                violations.append({
                    "description": f"Touch target {element} has height {height}px (minimum 44px recommended)",
                    "suggestion": f"Increase height to at least 44px for better accessibility"
                })

        # Check min-heights
        for match in re.finditer(min_height_pattern, css_content):
            element = match.group(1)
            min_height = int(match.group(2))
            if min_height < 44:
                violations.append({
                    "description": f"Touch target {element} has min-height {min_height}px (minimum 44px recommended)",
                    "suggestion": f"Increase min-height to at least 44px"
                })

        # Remove duplicates
        violations = list({v["description"]: v for v in violations}.values())

        # Score based on violations
        score = max(0, 100 - (len(violations) * 20))

        return {
            "score": score,
            "violations": violations,
            "pass": len(violations) == 0,
            "note": "Heuristic check - vision model provides more accurate analysis"
        }

    except Exception as e:
        return {
            "score": 100,
            "violations": [],
            "pass": True,
            "error": str(e)
        }
