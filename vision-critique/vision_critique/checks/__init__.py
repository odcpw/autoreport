"""
Automated Accessibility Checks

Complementary automated checks to supplement vision model analysis.
Focuses on measurable accessibility metrics.
"""

from .contrast import check_contrast
from .sizing import check_sizing

__all__ = ["check_contrast", "check_sizing"]
