# AutoBericht skill

Attach **[autobericht-skill.zip](autobericht-skill.zip)** to a chat that can read ZIP files. If ZIP extraction is unavailable, attach **[AutoBericht_portable.md](AutoBericht_portable.md)** instead. Both contain the complete workflow and helpers.

For a visit, provide the recording, project Sidecar and your style guide. The library is already in the Sidecar. For initial setup, provide anonymised past reports and ask for a Masterbericht, personal library and style guide. Photos stay on the work computer. Available transcription and file tools depend on the chat environment.

The `source/` subfolder is for maintaining the skill, not an additional upload requirement. Personal guides and libraries are not bundled here.

## Maintenance

Edit [source/SKILL.md](source/SKILL.md) and its resources, then run from the repository root:

```sh
python3 skill/source/scripts/build_distribution.py
```

Commit the source changes, the ZIP and portable document in this folder, and `program/build/skill-manifest.json`. CI rebuilds them and checks for differences. There is one distributed copy of each format; temporary builds can use `--output` outside the source folder. To install directly from source, copy `skill/source/` into the destination's `autobericht` skill directory.
