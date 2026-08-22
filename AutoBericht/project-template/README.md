# Project Template Scaffold

This folder is served by the local AutoBericht web server.

On first load of an empty project directory, the app copies the scaffold defined in `manifest.json` into the selected project folder:
- base directories (`inputs`, `outputs`, `backup`, `photos/...`, `templates`)
- report export templates in `templates/`

The scaffold is intentionally explicit here so users can see what gets created in their project directories.

PhotoSorter treats `photos/raw` as an immutable intake archive. Import creates
resized image copies in `photos/resized` and video copies in `photos/videos`;
it never moves or deletes the originals.
