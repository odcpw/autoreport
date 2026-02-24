# Redesign Workflow (2026)

This guide describes the minimal, folder-first workflow.

## Requirements

- Windows PC with Edge or Chrome
- No browser flags
- File System Access API enabled (if available in your environment)

## Project Setup

1. Create a project folder on local disk (for speed).
2. Place the following inside:
   - `AutoBericht/` (the app bundle)
   - `Inputs/` (self-assessment + customer docs)
   - `Photos/` (raw photos)
   - `Outputs/` (exports land here)

## Editor Workflow

1. Start the local server (from the AutoBericht folder):
   ```
   start-autobericht.cmd
   ```
2. Open **Open Project Folder**.
3. The app loads `project_sidecar.json` if present; otherwise it loads a user library.
   If neither exists, it keeps the project blank until a report language is selected.
4. Selecting report language on the Project page bootstraps seed content for that locale.
   Sidecar is written when you leave the Project page or save manually.
5. Fill in project settings when prompted (moderator, company, locale).
6. Edit by chapter (findings and recommendations).
7. Click **Save sidecar** to persist updates (autosave runs in the background).

## Troubleshooting

- Use **Save debug log** to write a `.txt` file into the project folder.
- Attach the log file in GitHub so it can be reviewed on the other machine.
