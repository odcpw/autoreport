# Ready-to-upload AutoBericht skill

After `git pull` or `sync-autobericht.ps1`, drag [autobericht-skill.zip](../../autobericht-skill.zip) from the repository root into a ChatGPT or Copilot conversation that can read ZIP attachments. The ZIP contains the complete `autobericht/` skill folder. If the chat cannot unpack ZIPs, attach **AutoBericht_portable.md**, which includes the same instructions and helper source as a single text document.

Then say:

> Follow the attached AutoBericht skill. Start with the files I provide and use my writing profile and library. If current application context would help, propose reading or cloning https://github.com/odcpw/autoreport. Keep comments about improving the software separate from the report.

For a report, supply the project Sidecar, your style guide and permitted recording or transcript. The Sidecar already contains the library; no duplicate library or examples pack is required. Site photos stay on your work computer. Attaching the package supplies instructions; it does not install a skill, enable transcription or grant repository access. Capabilities depend on the particular ChatGPT/Copilot environment.

For initial author setup, supply anonymised past reports and identify the base Word report if you have a preference. Ask for an editable consolidated Masterbericht, personal library and style guide. The skill also covers returning reviewed Sidecars to enrich the library and guide for later projects. Personal material is supplied separately and is never bundled in this shared package.

## For maintainers

Edit `skills/autobericht/`, then run from the repository root:

```sh
python3 skills/autobericht/scripts/build_distribution.py
```

The default build puts `autobericht-skill.zip` in the repository root and the portable Markdown and manifest in this folder. Commit the source changes together with the regenerated ZIP, portable Markdown, manifest and the compatibility copy at `docs/autobericht/AutoBericht_portable.md`. The package contains only skill resources, not adjacent application or client files. The manifest records artifact sizes and SHA-256 hashes. CI regenerates the package and checks that the committed copies are current.
