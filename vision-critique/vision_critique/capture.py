"""
Screenshot Capture Module

Captures high-quality screenshots of web pages using Playwright.
Handles navigation, element waiting, and viewport configuration.
"""

import asyncio
from datetime import datetime
from pathlib import Path
from typing import Optional

from playwright.async_api import async_playwright, Browser, Page, TimeoutError as PlaywrightTimeout


class ScreenshotCapturer:
    """
    Captures screenshots of web pages using headless Playwright browser.

    Features:
    - Configurable viewport sizes
    - CSS selector navigation (for tabs, modals, etc.)
    - Wait for element rendering
    - Full-page or element-specific captures
    - Automatic output path generation

    Example:
        capturer = ScreenshotCapturer(viewport={"width": 1920, "height": 1080})
        path = await capturer.capture(
            url="file:///path/to/index.html",
            selector="[data-tab='photosorter']",
            wait_for=".photosorter__main"
        )
    """

    def __init__(
        self,
        viewport: Optional[dict] = None,
        output_dir: Optional[Path] = None
    ):
        """
        Initialize screenshot capturer.

        Args:
            viewport: Viewport dimensions {"width": int, "height": int}
                     Defaults to 1920x1080
            output_dir: Directory to save screenshots
                       Defaults to ./screenshots/
        """
        self.viewport = viewport or {"width": 1920, "height": 1080}
        self.output_dir = output_dir or Path("screenshots")
        self.output_dir.mkdir(parents=True, exist_ok=True)

    async def capture(
        self,
        url: str,
        selector: Optional[str] = None,
        wait_for: Optional[str] = None,
        output_path: Optional[Path] = None,
        full_page: bool = True,
        wait_timeout: int = 5000
    ) -> Path:
        """
        Capture screenshot of a web page.

        Workflow:
        1. Launch headless Chromium browser
        2. Navigate to URL
        3. Optionally click selector (e.g., tab button)
        4. Wait for element to render
        5. Capture screenshot
        6. Return path to saved image

        Args:
            url: Page URL to capture (file:// or http(s)://)
            selector: CSS selector to click before capture (e.g., "[data-tab='import']")
            wait_for: CSS selector to wait for before capture (e.g., ".import-panel")
            output_path: Custom path for screenshot. If None, auto-generates
            full_page: Capture full scrollable page (True) or viewport only (False)
            wait_timeout: Milliseconds to wait for elements (default: 5000)

        Returns:
            Path to saved screenshot file

        Raises:
            PlaywrightTimeout: If elements don't appear within timeout
            RuntimeError: If browser fails to launch or navigate

        Example:
            # Capture specific tab
            path = await capturer.capture(
                url="file:///home/user/project/index.html",
                selector="[data-tab='photosorter']",
                wait_for=".photosorter__main"
            )
        """
        try:
            async with async_playwright() as p:
                # Launch browser
                browser = await p.chromium.launch(headless=True)
                page = await browser.new_page(viewport=self.viewport)

                # Navigate to URL
                await page.goto(url, wait_until="networkidle", timeout=wait_timeout)

                # Click selector if provided (e.g., activate a tab)
                if selector:
                    try:
                        await page.click(selector, timeout=wait_timeout)
                    except PlaywrightTimeout:
                        raise RuntimeError(
                            f"Failed to find clickable element: {selector}"
                        )

                # Wait for element if specified (ensure content is rendered)
                if wait_for:
                    try:
                        await page.wait_for_selector(wait_for, timeout=wait_timeout)
                    except PlaywrightTimeout:
                        raise RuntimeError(
                            f"Timeout waiting for element: {wait_for}"
                        )

                # Small delay to ensure rendering completes
                await page.wait_for_timeout(500)

                # Determine output path
                screenshot_path = output_path or self._generate_path(url, selector)

                # Capture screenshot
                await page.screenshot(
                    path=str(screenshot_path),
                    full_page=full_page,
                    type="png"
                )

                await browser.close()

                return screenshot_path

        except Exception as e:
            raise RuntimeError(f"Screenshot capture failed: {str(e)}") from e

    async def capture_element(
        self,
        url: str,
        element_selector: str,
        selector: Optional[str] = None,
        wait_for: Optional[str] = None,
        output_path: Optional[Path] = None
    ) -> Path:
        """
        Capture screenshot of a specific element only.

        Useful for capturing just a component, modal, or section.

        Args:
            url: Page URL
            element_selector: CSS selector of element to capture
            selector: Optional selector to click before capture
            wait_for: Optional selector to wait for
            output_path: Custom output path

        Returns:
            Path to saved screenshot

        Example:
            # Capture only the sidebar
            path = await capturer.capture_element(
                url="file:///project/index.html",
                element_selector=".photosorter__sidebar"
            )
        """
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page(viewport=self.viewport)

            await page.goto(url, wait_until="networkidle")

            if selector:
                await page.click(selector)

            if wait_for:
                await page.wait_for_selector(wait_for)

            await page.wait_for_timeout(500)

            # Find and screenshot specific element
            element = await page.query_selector(element_selector)
            if not element:
                raise RuntimeError(f"Element not found: {element_selector}")

            screenshot_path = output_path or self._generate_path(url, element_selector)

            await element.screenshot(path=str(screenshot_path), type="png")

            await browser.close()

            return screenshot_path

    def _generate_path(self, url: str, selector: Optional[str] = None) -> Path:
        """
        Generate unique screenshot path based on URL and selector.

        Format: screenshot_{timestamp}_{url_fragment}_{selector}.png

        Args:
            url: Page URL
            selector: Optional CSS selector for naming

        Returns:
            Path to screenshot file
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Extract filename-safe part from URL
        url_part = url.split("/")[-1].split(".")[0] if "/" in url else "page"
        url_part = "".join(c for c in url_part if c.isalnum() or c in "-_")[:20]

        # Extract filename-safe part from selector
        if selector:
            selector_part = selector.replace("[", "").replace("]", "").replace("'", "")
            selector_part = "".join(c for c in selector_part if c.isalnum() or c in "-_")[:20]
            filename = f"screenshot_{timestamp}_{url_part}_{selector_part}.png"
        else:
            filename = f"screenshot_{timestamp}_{url_part}.png"

        return self.output_dir / filename

    async def capture_multiple(
        self,
        url: str,
        states: list[dict]
    ) -> list[Path]:
        """
        Capture multiple states of the same page efficiently.

        Reuses browser instance for better performance.

        Args:
            url: Base URL
            states: List of state configurations, each with:
                   - name: State identifier
                   - selector: Optional CSS selector to click
                   - wait_for: Optional selector to wait for

        Returns:
            List of screenshot paths

        Example:
            paths = await capturer.capture_multiple(
                url="file:///project/index.html",
                states=[
                    {"name": "import", "selector": "[data-tab='import']"},
                    {"name": "photosorter", "selector": "[data-tab='photosorter']"},
                ]
            )
        """
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page(viewport=self.viewport)
            await page.goto(url, wait_until="networkidle")

            paths = []

            for state in states:
                if state.get("selector"):
                    await page.click(state["selector"])

                if state.get("wait_for"):
                    await page.wait_for_selector(state["wait_for"])

                await page.wait_for_timeout(500)

                path = self._generate_path(url, state.get("name", f"state_{len(paths)}"))
                await page.screenshot(path=str(path), full_page=True, type="png")
                paths.append(path)

            await browser.close()

            return paths
