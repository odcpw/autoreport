# AutoBericht Offline MVP (Scaffold)

This folder contains the standalone HTML/JS prototype for the offline AutoBericht experience. Open `index.html` in Edge/Chrome (Chromium) to run it without a web server. All assets and libraries must be bundled locally to satisfy corporate offline constraints.

> **Tip:** Chromium blocks ES module imports over `file://` by default. Launch Chrome/Edge with `--allow-file-access-from-files` or serve the folder via `python -m http.server` while developing.

## Directory Layout

- `index.html` — entry point that loads the three-panel UI.
- `css/` — global styles; `main.css` defines the current shell theme.
- `js/` — ES modules (state, schema definitions, UI helpers, utilities).
- `assets/` — logos/icons; keep high-resolution test photos outside version control.
- `libs/` — third-party libraries (markdown-it, CodeMirror, AJV, Paged.js, PptxGenJS, pica, hotkeys.js, Papa Parse, etc.).
- `fixtures/` — sample JSON payloads for local testing.

## Getting Started

1. Copy each required third-party library into `libs/` (see `libs/README.md` for suggestions). No CDN access is allowed at runtime.
2. Optionally load sample fixtures:
   - Open the app and upload `fixtures/master.sample.json` and `fixtures/self_eval.sample.json`.
   - Upload a folder of test photos (Chrome tip: use the directory picker, then allow nested folders).
3. Launch the browser with `--allow-file-access-from-files` (or a local server) before opening `index.html`.
4. Keyboard shortcuts: `Alt+1/2/3` switch tabs, `Ctrl+/` toggles the shortcut overlay, `Esc` closes it.
5. Validation must be green before exports enable; the Settings panel lists current issues and recent actions.

### PDF / PPTX Layout Options

- **Margins:** configurable in millimetres; defaults to 25mm.
- **Header/Footer:** toggle to hide logos and footer entirely (useful for draft prints).
- **Footer text:** supports `{company}`, `{date}`, `{time}` tokens; defaults to `Company — Date` with locale-aware formatting.
- **Locale:** affects date/time formatting in PDF exports.

### Export Overview

- **PDF Report:** Uses Paged.js to render cover page (logos, TOC) and chapters. Prints via browser dialog.
- **Report PPTX:** Creates one slide per included finding with Markdown bullets and chapter photos where available.
- **Seminar PPTX:** Groups photos by seminar tag; each slide lists references and shows thumbnail placeholders.
- **project.json:** Snapshot of the full project state (settings + findings + photos) for re-import/autosave.

## Development Notes

- The current code base focuses on the UI shell, tab navigation, and data-ingestion stubs.
- Schema validation is stubbed; integrate AJV bundle by replacing `js/schema/validator.js`.
- Fixture loading is a no-op; extend `js/state/fixtures.js` to hydrate the state automatically when running from a dev server.
- Keep all code modular to ease future expansion: PhotoSorter grid, Markdown editors, PPT/PDF pipelines, and FS Access toggles will plug into the existing skeleton.

---

## Related Documentation

- **[HTML Workflow Guide](../docs/guides/html-workflow.md)** - Complete user guide for web UI
- **[Getting Started](../docs/guides/getting-started.md)** - End-to-end first report
- **[System Overview](../docs/architecture/system-overview.md)** - How web UI integrates with VBA
- **[Data Model](../docs/architecture/data-model.md)** - JSON schema reference
- **[Vision Critique](../vision-critique/README.md)** - UI quality analysis tool
