# AutoBericht

Create safety reports with the local browser application and your own reusable library.

1. Run **start-autobericht.cmd**.
2. Choose your project folder in the application. Keep project folders separate from this installation.
3. For a new project, choose its language; the app creates the folders and copies the templates it needs.

## Find what you need

| Location | Purpose |
|---|---|
| [start-autobericht.cmd](start-autobericht.cmd) | Start the application |
| [sync-autobericht.ps1](sync-autobericht.ps1) | Download the current version beside this script |
| [skill/](skill/README.md) | Skill ZIP or portable document to attach to a chat |
| [templates/](templates/README.md) | Word, PowerPoint and Excel templates |
| [guides/](guides/README.md) | Instructions in German, French and Italian |
| [program/](program/README.md) | Application files; no manual setup needed inside |
| [docs/](docs/autobericht/README.md) | Technical documentation and development notes |

## Update

Close the application and its server window. In the **installation folder**, keep `sync-autobericht.ps1` and erase the other installation files, then run the script. Leave your separate project folders and personal libraries in place. The script installs beside itself even when run from another working directory. An explicit `-TargetFolder` still selects a different destination.

## Use the skill

Attach [skill/autobericht-skill.zip](skill/autobericht-skill.zip) to a chat that can read ZIP files. If it cannot, use [skill/AutoBericht_portable.md](skill/AutoBericht_portable.md).

For a new visit, supply your recording, project Sidecar and personal style guide. The Sidecar already contains the library. The skill also builds a Masterbericht, library and style guide from past reports. Site photos stay on the work computer.

## Project files

`project_sidecar.json` holds the current project's report, assessment and photo metadata. `library_user_*.json` carries reusable wording between projects. Existing project folders and their own `templates/` keep the same structure. Personal material belongs in your private storage, outside the program installation.

For troubleshooting, use **Save debug log** in the application. Development and test instructions are in [program/README.md](program/README.md).
