# Workflow (2026)

This is the current folder-first workflow.

## Requirements

- Windows PC with Edge or Chrome
- Local run via `start-autobericht.cmd`
- File System Access API available (Edge/Chrome)

## Setup

1. Place repo locally.
2. Create an empty project folder.
3. Start `start-autobericht.cmd` from repo root.
4. In app, click **Open Project Folder** and select the empty project folder.
5. On **Project**, choose **Language** to bootstrap report content.

## Working Flow

1. Import self-assessment on **Project** (optional when available).
2. Use **PhotoSorter**:
   - `Import / Export Photos` -> `Import photos`
   - original images stay in `photos/raw`; resized images go to `photos/resized`
     and videos are moved to `photos/videos`
   - tag photos (report / observations / training)
   - `Export tagged folders` when needed
3. Use **AutoBericht** chapter editor:
   - findings/recommendations
   - include/done/priority
   - library action per row (`Off`, `Append`, `Replace`)
4. Update library on **Project**:
   - `Generate / Update Library`
   - `Export Library Excel` (optional)
5. Export outputs on **Project**:
   - `Word Export`
   - `PowerPoint Export (Report)`
   - `PowerPoint Export (Training)`

## Photo filenames and dates

Imported copies use `YYYY-MM-DD-HH-MM_abc_0001.jpg`, where `abc` is the three-character source folder. The sequence follows original filename order within that folder; it is separate from the stable `Photo 038` number shown in PhotoSorter. Original files in `photos/raw` retain their names.

For JPEG files, the date/time comes from EXIF `DateTimeOriginal` and retains the camera's recorded local time. The EXIF modification date is not a capture date. If capture metadata is absent, invalid, unreadable or in an unsupported image format, import falls back to the file's modification date and reports the number of affected photos in its completion message. A download/copy can change that fallback date. The current reader handles JPEG EXIF; it does not yet extract capture dates from HEIC or other image containers.

The capture-date fix applies to new imports. Existing resized filenames and Sidecar photo paths are not automatically changed. Reimporting raw files into an already populated resized folder can therefore fail the existing resume check if the corrected names differ. Keep the tagged project intact; repairing existing names needs a coordinated update of file paths and Sidecar references, not an independent bulk rename.

## Sidecar vs Library

- `project_sidecar.json`: project state and edits for this project.
- `library_user_*.json`: reusable knowledge base across projects.
- `Append`/`Replace` is queued in sidecar and applied to library on **Generate / Update Library**.

## Troubleshooting

- Use **Save debug log** to write diagnostics in the project folder.
