# AutoBericht application files

Start the app with `start-autobericht.cmd` at the repository root. It calls this folder's PowerShell launcher, serves the application locally and opens `mini/`.

- `mini/`: AutoBericht, PhotoSorter and LibraryMaker.
- `data/`: default libraries, checklists and score weights.
- `libs/`, `shared/`: bundled dependencies and shared support.
- `project-setup.json`: folders and template files created in a project.
- `tools/`: local HTTP server.
- `tests/`: application and installation regression checks.
- `experiments/`: development prototypes, separate from the production workflow.
- `build/`: generated package metadata.

The server serves this folder at its existing URLs and maps `/templates/` to the repository's root `templates/`. It does not serve the whole repository. Existing project-local templates take precedence over bundled ones; folder setup copies missing templates without replacing existing files.

## Tests

```sh
cd program
npm test
```

The core tests use Node.js built-ins. The installation checks use Windows PowerShell on Windows, or `AUTOBERICHT_POWERSHELL` when supplied elsewhere. They exercise the real sync script and server against a local archive. The browser export check uses Firefox and the optional `ooxml` CLI; strict Word validation also uses `ooxml` when available.

The skill helpers are checked from the repository root with:

```sh
python3 -m unittest discover -s skill/source/scripts -p 'test_*.py'
```

See [technical documentation](../docs/autobericht/README.md) and [user guides](../guides/README.md).
