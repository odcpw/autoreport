# Oracle browser mode (ChatGPT)

Fast way to run Oracle with the ChatGPT web UI on this machine.

## One-time setup (already done)
- Wrapper script `tools/oracle-chrome.sh` adds `--no-sandbox`/`--disable-dev-shm-usage` so Chrome can launch in this environment.
- Use a local state dir `.oracle/` (git‑ignored) to avoid permission issues under `~/.oracle`.

## Quick smoke test (works)
```bash
cd /home/oliver/Projects/odcpw/autoreport
ORACLE_HOME_DIR="$PWD/.oracle" \
bunx @steipete/oracle \
  --engine browser \
  --browser-chrome-path "$PWD/tools/oracle-chrome.sh" \
  --prompt "Return the word OK" \
  --file .oracle/test.txt \
  --slug "oracle-browser-smoke" \
  --verbose
```
- Expect a visible Chrome window; keep it open until the run finishes.
- Output is saved to `.oracle/sessions/<slug>/output.log` (see the smoke run at `.oracle/sessions/oracle-browser-smoke/output.log`).

## Using it for real prompts
- Swap `--prompt/--file` with your request + repo globs. Keep prompts concise; long, code‑fenced messages can trip the ChatGPT prompt-commit check.
- Keep `ORACLE_HOME_DIR="$PWD/.oracle"` and `--browser-chrome-path "$PWD/tools/oracle-chrome.sh"` to avoid sandbox and permission errors.
- If you need manual paste instead, use `--render --copy --engine browser` (no automation) or switch to API mode with `OPENAI_API_KEY` set.

## Troubleshooting
- `ECONNREFUSED 127.0.0.1:9222`: Chrome failed to start; ensure the wrapper script is executable and on the given path. You can dry-run Chrome with `tools/oracle-chrome.sh --remote-debugging-port=9222 about:blank`.
- Prompt didn’t appear / commit timeout: shorten the prompt, try `--browser-input-timeout 120s`, or fall back to `--render --copy`.
- Not logged in: when the browser opens, sign into chatgpt.com in that window, then re-run.
