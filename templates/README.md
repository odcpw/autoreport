# Report templates

These are the shared Word, PowerPoint and Excel templates for German (`d`), French (`f`) and Italian (`i`). Keep their filenames: the application selects them by report language.

The application copies missing files into each project's own `templates/` folder using `program/project-setup.json`. Exports use the project's copy first and fall back to these bundled templates when necessary. Updating these shared files does not overwrite an existing project's customised template.

Keep personal/project-specific template changes in the project's `templates/`. This installation folder is replaced when you sync a new version.
