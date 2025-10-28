# Agent Playbook

Practical guidelines for working as the coding agent across projects.

## Mission

- Deliver accurate answers or code changes with minimal churn.
- Follow instruction priority: system → developer → user → repo docs.
- Preserve existing work; never undo user edits without clear approval.

## Operating Principles

- **Think first**: gather context before editing; use quick searches instead of scanning whole files.
- **Plan when it helps**: craft a lightweight plan for anything non-trivial; skip it for obvious one-off edits.
- **Keep diffs tight**: prefer focused changes, avoid drive-by refactors, and match the project’s style.
- **Validate**: run the smallest meaningful check (tests, lints, scripts); if you cannot, state why and suggest how.
- **Document decisions**: explain non-obvious intent in concise comments or commit notes when needed.
- **Report clearly**: summarize outcomes, list affected files with line references, and offer next logical steps only when they add value.
- **Ask before assuming**: don't add fallbacks or fabricate data when blocked; pause and ask the user for the needed details, explaining what is missing.

## Best Practices

- Use tooling expected by the project (package managers, test runners) instead of ad-hoc commands.
- Rely on evidence from logs, stack traces, or tests before fixing bugs—no guessing.
- Avoid destructive commands unless the user explicitly requests them.
- Keep communication concise, collaborative, and actionable.
- Update this playbook when recurring patterns emerge so the next session starts from a better baseline.
- Maintain clear, centralized logs with consistent severity levels so issues can be tailed and diagnosed quickly.
- Use a temporary scratchpad (e.g., `docs/scratch.md`) for complex tasks to capture context and decisions; tidy or archive it when done.
- Break work into small, composable modules with clear interfaces so future agents (and LLMs) can reason about them quickly.
- Keep file-level headers or module docstrings current—they should explain purpose, dependencies, and how to extend the code safely.
- Replace magic numbers with named constants or configuration values to keep intent transparent.
- Source every environment-specific setting from an `.env` file and document expected keys.
- Structure `.env` files into logical sections with comments so dependencies and secrets are easy to spot.
- Maintain a running changelog that records user-visible behaviour changes for continuity between sessions.
- Write comments as definitive descriptions (no transitional phrasing like “enhanced” or “replaced”); avoid TODOs unless explicitly agreed.
- Assume the operator is LLM-only—deliver end-to-end instructions that require no extra coding knowledge.

## Error Handling

Prefer result-based error handling over exception-based flow control for business logic. This makes failure modes explicit in function signatures and improves reasoning about error paths.

### When to Use Result Types

**DO use Result[T, E] for:**
- New business logic where failures are expected (validation, resource not found, etc.)
- Service layer methods with predictable failure modes
- Operations where callers need to handle errors explicitly
- Functions where the signature should document what can go wrong

**Example (Preferred):**
```python
from sjbot.core.result import Result, Ok, Err

def get_wallet_balance(self) -> Result[Dict[str, Any], str]:
    """Fetch wallet balance from exchange API."""
    try:
        response = self.client.get("/accounts", auth=True)
        payload = response.json()
        if payload.get("result") != "success":
            return Err(f"Exchange returned error: {payload.get('error')}")
        accounts = payload.get("accounts")
        if not accounts:
            return Err("No accounts found in response")
        return Ok(accounts)
    except httpx.HTTPError as e:
        return Err(f"HTTP request failed: {e}")
```

**Calling code:**
```python
result = service.get_wallet_balance()
if result.is_ok():
    balance = result.unwrap()
    logger.info("Balance: %s", balance)
else:
    error = result.unwrap_err()
    logger.error("Failed to fetch balance: %s", error)
    return  # Handle gracefully
```

### When Exceptions Are Acceptable

**Exceptions are fine for:**
- Wrapping third-party libraries that raise exceptions (httpx, SQLAlchemy, etc.)
- Truly exceptional conditions (programming errors, system failures)
- Top-level error boundaries and retry logic
- Infrastructure code where propagating detailed errors isn't useful

**Example (Acceptable):**
```python
@retry(wait=wait_exponential(multiplier=1, min=1, max=8), stop=stop_after_attempt(5))
def get(self, path: str, params: Dict[str, Any] | None = None) -> httpx.Response:
    """Issue HTTP GET request with automatic retry (raises on failure)."""
    response = self.client.get(path, params=params)  # May raise httpx.HTTPError
    response.raise_for_status()  # Let exception propagate for retry logic
    return response
```

### Migration Strategy

**For existing code:**
1. DO NOT refactor working error handling without clear value
2. Add Result types to NEW code and NEW functions
3. Consider Result types when touching existing service methods IF:
   - The function has clear, documented failure modes
   - Callers currently catch and handle specific errors
   - Tests already cover error paths
   - The change is isolated and low-risk

**Incremental adoption:**
- Start with leaf functions (no dependencies)
- Move up the call stack as needed
- Keep HTTP/DB wrappers exception-based unless there's clear benefit
- Prioritize code clarity over dogmatic adherence

### Result Type Utilities

Available in `sjbot/core/result.py`:
- `Result[T, E]`: Union type for Ok[T] | Err[E]
- `Ok(value)`: Success variant
- `Err(error)`: Error variant
- `try_result(fn)`: Wrap exception-throwing code in Result
- Methods: `is_ok()`, `is_err()`, `unwrap()`, `unwrap_or()`, `map()`, `and_then()`

See module docstring for detailed usage examples.

## Toolbelt

These CLI tools are typically available on the host and worth using:

- `ripgrep` (`rg`) for fast code search.
- `fd-find` (`fd`) for quick file discovery.
- `jq` for JSON inspection.
- `httpie` for readable HTTP requests.
- `wget` (`apt install wget`) to mirror doc sites: `wget --mirror --convert-links --adjust-extension --no-parent https://example.com/docs/`.
- `lynx` (`apt install lynx`) for quick text dumps: `lynx -dump -nolist -nonumbers https://example.com/docs/ > docs.txt`.
- `pandoc` (`apt install pandoc`) to convert HTML into Markdown: `curl -sL https://example.com/docs/ | pandoc -f html -t gfm -o docs/tmp_scrape/docs.md`.
- `playwright` (`pipx install playwright && playwright install chromium`) to render SPA docs before conversion.
- `fzf` for fuzzy search across files, history, and git.
- `just` for lightweight task automation.
- `bat` for syntax-highlighted file previews.
- `entr` for rerunning commands on file changes.
- `btop` / `glances` for live system monitoring.
- `lnav` for interactive log exploration.
- `tmux` for managing parallel terminal sessions efficiently.

Example SPA docs scrape:

```bash
python - <<'PY'
from pathlib import Path
from playwright.sync_api import sync_playwright

url = "https://hyperliquid.gitbook.io/hyperliquid-docs/for-developers/api"
html_path = Path("docs/tmp_scrape/hyperliquid_api.html")
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto(url, wait_until="networkidle")
    html_path.write_text(page.content(), encoding="utf8")
    browser.close()
PY
pandoc docs/tmp_scrape/hyperliquid_api.html -f html -t gfm -o docs/tmp_scrape/hyperliquid_api.md
```
