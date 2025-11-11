"""
Command-Line Interface

Beautiful CLI using rich for colored output, progress indicators,
and formatted results. Entry point for coding agents and users.
"""

import asyncio
import json
import sys
from pathlib import Path
from typing import Optional

import click
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.syntax import Syntax

from . import __version__
from .config import load_config
from .providers import get_provider
from .scorer import UIScorer


console = Console()


@click.command()
@click.option(
    '--url',
    default=None,
    help='URL to critique (file:// or http(s)://). Defaults to AutoBericht if not specified.'
)
@click.option(
    '--tab',
    default=None,
    help='Tab name to focus on (e.g., photosorter, import)'
)
@click.option(
    '--selector',
    default=None,
    help='CSS selector to click before capture (e.g., "[data-tab=\'photosorter\']")'
)
@click.option(
    '--wait-for',
    default=None,
    help='CSS selector to wait for before capture'
)
@click.option(
    '--provider',
    default=None,
    type=click.Choice(['anthropic', 'openai', 'local'], case_sensitive=False),
    help='Vision provider to use. Defaults to VISION_PROVIDER from .env'
)
@click.option(
    '--output',
    default='rich',
    type=click.Choice(['rich', 'json'], case_sensitive=False),
    help='Output format: rich (colored terminal) or json (for agents)'
)
@click.option(
    '--css',
    default=None,
    type=click.Path(exists=True),
    help='CSS file path for automated accessibility checks'
)
@click.option(
    '--env-file',
    default=None,
    type=click.Path(exists=True),
    help='Path to .env file (defaults to ./.env)'
)
@click.version_option(version=__version__)
def main(
    url: Optional[str],
    tab: Optional[str],
    selector: Optional[str],
    wait_for: Optional[str],
    provider: Optional[str],
    output: str,
    css: Optional[str],
    env_file: Optional[str]
):
    """
    Vision Critique - UI Quality Analysis Tool

    Analyze web UI screenshots using vision AI and automated checks.
    Get scores, issues, and actionable improvement suggestions.

    Examples:

      # Basic usage (uses .env config)
      vision-critique capture

      # Specific tab
      vision-critique capture --tab photosorter

      # Different provider
      vision-critique capture --provider openai

      # JSON output for coding agents
      vision-critique capture --output json

      # Full custom analysis
      vision-critique capture --url file:///path/to/index.html \\
          --selector "[data-tab='photosorter']" \\
          --wait-for ".photosorter__main" \\
          --css /path/to/main.css
    """
    try:
        # Load configuration
        config = load_config(Path(env_file) if env_file else None)

        # Determine provider
        provider_name = provider or config.vision_provider

        # Determine URL
        if url is None:
            # Default to AutoBericht in this project
            autobericht_path = Path.cwd() / "AutoBericht" / "index.html"
            if autobericht_path.exists():
                target_url = f"file://{autobericht_path.absolute()}"
            else:
                console.print("[red]❌ No URL specified and AutoBericht not found[/red]")
                console.print("\nUsage: vision-critique capture --url <url>")
                sys.exit(1)
        else:
            target_url = url

        # Auto-detect selector if tab is specified but selector isn't
        if tab and not selector:
            # Try common patterns: id="tab-{name}" or data-tab="{name}"
            selector = f"#tab-{tab}"  # AutoBericht uses id="tab-photosorter"

        # Determine CSS path
        css_path = None
        if css:
            css_path = Path(css)
        elif (Path.cwd() / "AutoBericht" / "css" / "main.css").exists():
            css_path = Path.cwd() / "AutoBericht" / "css" / "main.css"

        # Run critique
        result = asyncio.run(_run_critique(
            url=target_url,
            tab=tab,
            selector=selector,
            wait_for=wait_for,
            provider_name=provider_name,
            config=config,
            css_path=css_path
        ))

        # Output result
        if output == 'json':
            _output_json(result)
        else:
            _output_rich(result, provider_name)

    except KeyboardInterrupt:
        console.print("\n[yellow]⚠️  Interrupted by user[/yellow]")
        sys.exit(130)
    except Exception as e:
        console.print(f"[red]❌ Error: {str(e)}[/red]")
        if '--debug' in sys.argv:
            console.print_exception()
        sys.exit(1)


async def _run_critique(
    url: str,
    tab: Optional[str],
    selector: Optional[str],
    wait_for: Optional[str],
    provider_name: str,
    config,
    css_path: Optional[Path]
):
    """Run the critique workflow with progress indicators"""

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console
    ) as progress:

        # Initialize provider
        task = progress.add_task("[cyan]Initializing vision provider...", total=None)
        try:
            provider = get_provider(provider_name, config)
        except ValueError as e:
            console.print(f"[red]❌ {str(e)}[/red]")
            sys.exit(1)

        if not provider.is_available():
            console.print(f"[red]❌ Provider '{provider_name}' is not available[/red]")
            sys.exit(1)

        progress.update(task, description="[cyan]Creating scorer...")
        scorer = UIScorer(provider, config)

        progress.update(task, description="[cyan]Capturing screenshot...")

        progress.update(task, description="[cyan]Analyzing UI with vision model...")

        # Run critique
        result = await scorer.score(
            url=url,
            tab=tab,
            selector=selector,
            wait_for=wait_for,
            css_path=css_path
        )

        progress.update(task, description="[green]✓ Analysis complete", completed=True)

    return result


def _output_rich(result, provider_name: str):
    """Output result in rich formatted terminal output"""

    console.print()
    console.print(Panel.fit(
        f"[bold]UI Quality Analysis[/bold]\n"
        f"Provider: {provider_name}",
        border_style="cyan"
    ))

    # Scores table
    console.print("\n[bold]📊 Scores[/bold]")
    scores_table = Table(show_header=True, header_style="bold magenta")
    scores_table.add_column("Dimension", style="cyan")
    scores_table.add_column("Score", justify="right")
    scores_table.add_column("Grade", justify="center")

    def get_score_color(score: float) -> str:
        if score >= 90:
            return "green"
        elif score >= 75:
            return "yellow"
        elif score >= 60:
            return "orange"
        else:
            return "red"

    scores_table.add_row(
        "Visual Hierarchy",
        f"[{get_score_color(result.scores.visual_hierarchy)}]{result.scores.visual_hierarchy}/100[/]",
        _get_grade_emoji(result.scores.visual_hierarchy)
    )
    scores_table.add_row(
        "Color & Typography",
        f"[{get_score_color(result.scores.color_typography)}]{result.scores.color_typography}/100[/]",
        _get_grade_emoji(result.scores.color_typography)
    )
    scores_table.add_row(
        "UX & Interaction",
        f"[{get_score_color(result.scores.ux_interaction)}]{result.scores.ux_interaction}/100[/]",
        _get_grade_emoji(result.scores.ux_interaction)
    )
    scores_table.add_row(
        "Accessibility",
        f"[{get_score_color(result.scores.accessibility)}]{result.scores.accessibility}/100[/]",
        _get_grade_emoji(result.scores.accessibility)
    )
    scores_table.add_row(
        "[bold]Overall[/bold]",
        f"[bold][{get_score_color(result.scores.overall)}]{result.scores.overall}/100[/][/bold]",
        f"[bold]{_get_grade_emoji(result.scores.overall)}[/bold]"
    )

    console.print(scores_table)

    # Issues
    if result.issues:
        console.print(f"\n[bold]🔍 Issues Found ({len(result.issues)})[/bold]")

        high_issues = [i for i in result.issues if i.severity == "high"]
        medium_issues = [i for i in result.issues if i.severity == "medium"]
        low_issues = [i for i in result.issues if i.severity == "low"]

        if high_issues:
            console.print("\n[bold red]High Priority:[/bold red]")
            for issue in high_issues:
                console.print(f"  🔴 [{issue.dimension}] {issue.description}")
                console.print(f"     💡 {issue.suggestion}\n")

        if medium_issues:
            console.print("[bold yellow]Medium Priority:[/bold yellow]")
            for issue in medium_issues:
                console.print(f"  🟡 [{issue.dimension}] {issue.description}")
                console.print(f"     💡 {issue.suggestion}\n")

        if low_issues:
            console.print("[bold green]Low Priority:[/bold green]")
            for issue in low_issues:
                console.print(f"  🟢 [{issue.dimension}] {issue.description}")
    else:
        console.print("\n[bold green]✓ No issues found![/bold green]")

    # Suggestions
    if result.suggestions:
        console.print(f"\n[bold]💡 Top Suggestions[/bold]")
        for i, suggestion in enumerate(result.suggestions[:5], 1):
            console.print(f"  {i}. {suggestion}")

    # Screenshot path
    console.print(f"\n[dim]📸 Screenshot: {result.screenshot_path}[/dim]")
    console.print()


def _output_json(result):
    """Output result as JSON for coding agents"""
    output = {
        "scores": {
            "visual_hierarchy": result.scores.visual_hierarchy,
            "color_typography": result.scores.color_typography,
            "ux_interaction": result.scores.ux_interaction,
            "accessibility": result.scores.accessibility,
            "overall": result.scores.overall,
            "grade": result.scores.get_grade()
        },
        "issues": [
            {
                "dimension": issue.dimension,
                "severity": issue.severity,
                "description": issue.description,
                "location": issue.location,
                "suggestion": issue.suggestion
            }
            for issue in result.issues
        ],
        "suggestions": result.suggestions,
        "screenshot_path": result.screenshot_path,
        "timestamp": result.timestamp,
        "provider": result.provider
    }

    print(json.dumps(output, indent=2))


def _get_grade_emoji(score: float) -> str:
    """Get emoji for grade"""
    if score >= 90:
        return "🌟"
    elif score >= 75:
        return "✅"
    elif score >= 60:
        return "⚠️"
    else:
        return "❌"


if __name__ == "__main__":
    main()
