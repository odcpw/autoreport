# Vision Critique

> Vision-based UI quality analysis tool for coding agents

A command-line tool that uses vision AI to analyze web UI screenshots and provide actionable feedback on design quality, accessibility, and user experience.

## Features

- **Vision AI Analysis** - Uses Claude, GPT-4V, or local LLMs to analyze UI screenshots
- **Automated Checks** - Complementary WCAG accessibility validation
- **Multiple Providers** - Supports Anthropic Claude, OpenAI GPT-4V, and local Ollama models
- **Rich Output** - Beautiful terminal output or JSON for coding agents
- **Fast & Offline Capable** - Local LLM support for privacy-preserving analysis

## Quick Start

### Installation

```bash
# Install with uv (recommended)
cd vision-critique
uv tool install .

# Or install for development
uv sync
```

### Configuration

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Add your API keys:
   ```bash
   # For Claude (recommended)
   ANTHROPIC_API_KEY=sk-ant-your-key-here

   # Or for OpenAI
   OPENAI_API_KEY=sk-your-key-here

   # Or use local Ollama (no API key needed)
   OLLAMA_HOST=http://localhost:11434
   OLLAMA_MODEL=llava
   ```

3. Install Playwright browsers:
   ```bash
   playwright install chromium
   ```

### Basic Usage

```bash
# Analyze AutoBericht (auto-detected)
vision-critique capture

# Analyze specific tab
vision-critique capture --tab photosorter

# Use different provider
vision-critique capture --provider openai

# JSON output for coding agents
vision-critique capture --output json
```

## Usage Examples

### For Coding Agents

Coding agents like Claude Code can invoke this tool to get visual feedback:

```bash
# Make UI changes...
# Then validate visually:
vision-critique capture --tab photosorter --output json > critique.json

# Read the feedback and iterate
```

Example JSON output:
```json
{
  "scores": {
    "visual_hierarchy": 72.5,
    "color_typography": 68.0,
    "ux_interaction": 75.0,
    "accessibility": 65.0,
    "overall": 70.3,
    "grade": "C"
  },
  "issues": [
    {
      "dimension": "accessibility",
      "severity": "high",
      "description": "Button contrast is 2.8:1, below WCAG AA minimum of 4.5:1",
      "location": ".button-panel__grid button",
      "suggestion": "Change button background to #0066cc or text color to #000"
    }
  ],
  "suggestions": [
    "Increase button min-height to 44px for better touch targets",
    "Improve text contrast on category buttons",
    "Add visual feedback for active/hover states"
  ]
}
```

### Advanced Usage

```bash
# Full custom analysis
vision-critique capture \
  --url file:///path/to/index.html \
  --selector "[data-tab='photosorter']" \
  --wait-for ".photosorter__main" \
  --css /path/to/main.css \
  --provider anthropic \
  --output rich

# Use local LLM (privacy-preserving)
vision-critique capture --provider local

# With custom env file
vision-critique capture --env-file .env.production
```

## Configuration Options

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `ANTHROPIC_API_KEY` | Anthropic API key | (required for Claude) |
| `OPENAI_API_KEY` | OpenAI API key | (required for GPT-4V) |
| `OLLAMA_HOST` | Ollama server URL | `http://localhost:11434` |
| `OLLAMA_MODEL` | Ollama model name | `llava` |
| `VISION_PROVIDER` | Default provider | `anthropic` |
| `VIEWPORT_WIDTH` | Screenshot width | `1920` |
| `VIEWPORT_HEIGHT` | Screenshot height | `1080` |

### CLI Options

```
Options:
  --url TEXT              URL to critique (file:// or http(s)://)
  --tab TEXT              Tab name to focus on
  --selector TEXT         CSS selector to click before capture
  --wait-for TEXT         CSS selector to wait for before capture
  --provider [anthropic|openai|local]
                          Vision provider to use
  --output [rich|json]    Output format
  --css PATH              CSS file for automated checks
  --env-file PATH         Path to .env file
  --version               Show version
  --help                  Show this message
```

## Architecture

### Modular Design

```
vision-critique/
├── vision_critique/
│   ├── models.py           # Pydantic data models
│   ├── capture.py          # Playwright screenshot engine
│   ├── scorer.py           # Main orchestrator
│   ├── config.py           # Configuration management
│   │
│   ├── providers/          # Vision model providers
│   │   ├── base.py        # Abstract interface
│   │   ├── anthropic.py   # Claude implementation
│   │   ├── openai.py      # GPT-4V implementation
│   │   └── local.py       # Ollama implementation
│   │
│   └── checks/             # Automated checks
│       ├── contrast.py    # WCAG contrast validation
│       └── sizing.py      # Touch target validation
```

### Provider Pattern

Easy to extend with new vision providers:

```python
from vision_critique.providers.base import VisionProvider

class CustomProvider(VisionProvider):
    @property
    def name(self) -> str:
        return "custom"

    async def critique(self, screenshot_path, context):
        # Your implementation
        pass

    def is_available(self) -> bool:
        # Check if configured
        pass
```

## Scoring Dimensions

### Visual Hierarchy (30% weight)
- Layout clarity and flow
- Element sizing and prominence
- Spacing consistency
- Focal point clarity

### Color & Typography (25% weight)
- Color harmony and palette
- WCAG contrast ratios
- Font choices and hierarchy
- Readability

### UX & Interaction (25% weight)
- Touch target sizes (44×44px minimum)
- Information hierarchy
- Logical grouping
- Interaction patterns

### Accessibility (20% weight)
- WCAG 2.1 AA compliance
- Visual impairment considerations
- Keyboard navigation indicators
- Sufficient visual cues

**Overall Score**: Weighted average of all dimensions

## Comparison: Cloud vs Local

| Feature | Cloud (Claude/GPT-4V) | Local (Ollama) |
|---------|----------------------|----------------|
| **Quality** | Excellent | Good |
| **Speed** | 3-10 seconds | 10-60 seconds |
| **Cost** | ~$0.01-0.03/image | Free |
| **Privacy** | Data sent to API | Fully local |
| **Setup** | API key only | Ollama + model download |
| **Offline** | No | Yes |

## Integration Examples

### With Claude Code

```bash
# Claude Code workflow
# 1. Make UI improvements
# 2. Run vision critique
vision-critique capture --output json | jq '.scores.overall'

# 3. Read feedback and iterate
```

### In CI/CD

```yaml
# .github/workflows/ui-quality.yml
- name: UI Quality Check
  run: |
    vision-critique capture --output json > critique.json
    score=$(jq '.scores.overall' critique.json)
    if (( $(echo "$score < 70" | bc -l) )); then
      echo "UI quality below threshold: $score/100"
      exit 1
    fi
```

### As Development Tool

```bash
# Development workflow
while true; do
  # Make changes
  vim AutoBericht/css/main.css

  # Check quality
  vision-critique capture

  # Iterate based on feedback
done
```

## Troubleshooting

### "Provider not available"

**Claude/OpenAI:**
```bash
# Check API key is set
echo $ANTHROPIC_API_KEY

# Verify key is valid
curl https://api.anthropic.com/v1/messages \
  -H "x-api-key: $ANTHROPIC_API_KEY" \
  -H "anthropic-version: 2023-06-01"
```

**Local (Ollama):**
```bash
# Check Ollama is running
curl http://localhost:11434/api/tags

# Start Ollama if needed
ollama serve

# Pull vision model
ollama pull llava
```

### "Screenshot capture failed"

```bash
# Install Playwright browsers
playwright install chromium

# Check file path is correct
ls -la AutoBericht/index.html

# Try with explicit URL
vision-critique capture --url file://$(pwd)/AutoBericht/index.html
```

### "Failed to parse response"

This usually means the vision model returned malformed JSON. Try:

1. Use a different provider (`--provider openai`)
2. Simplify the prompt (edit `providers/base.py`)
3. Check model compatibility

## Development

### Setup

```bash
# Clone and install
git clone <repo>
cd vision-critique
uv sync

# Install pre-commit hooks
pre-commit install
```

### Running Tests

```bash
# Run tests
uv run pytest

# With coverage
uv run pytest --cov=vision_critique
```

### Adding a New Provider

1. Create `vision_critique/providers/your_provider.py`
2. Inherit from `VisionProvider`
3. Implement `critique()` and `is_available()`
4. Add to `providers/__init__.py`
5. Update .env.example

## Roadmap

- [ ] Multi-viewport testing (mobile, tablet, desktop)
- [ ] Comparison mode (before/after)
- [ ] Custom scoring weights
- [ ] HTML report generation
- [ ] Video/animation analysis
- [ ] Component-specific analysis
- [ ] Design system extraction
- [ ] Batch processing

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new features
4. Ensure all tests pass
5. Submit a pull request

## License

MIT

## Credits

Built for AutoBericht project. Inspired by ReLook's vision-based UI improvement approach.

---

**For support or questions:**
- Open an issue on GitHub
- Check the troubleshooting section
- Review example configurations in `examples/`
