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
3. Click **Load sidecar** (if missing, a template is loaded).
4. Edit by chapter (findings and recommendations).
5. Click **Save sidecar** to persist updates.

## Troubleshooting

- Use **Save debug log** to write a `.txt` file into the project folder.
- Attach the log file in GitHub so it can be reviewed on the other machine.
