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
5. On **Project**, choose `Locale` to bootstrap seed content.

## Working Flow

1. Import self-assessment on **Project** (optional when available).
2. Use **PhotoSorter**:
   - `Import / Export Photos` -> `Import photos`
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

## Sidecar vs Library

- `project_sidecar.json`: project state and edits for this project.
- `library_user_*.json`: reusable knowledge base across projects.
- `Append`/`Replace` is queued in sidecar and applied to library on **Generate / Update Library**.

## Troubleshooting

- Use **Save debug log** to write diagnostics in the project folder.
