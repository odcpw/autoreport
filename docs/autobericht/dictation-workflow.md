# AutoBericht: speak through the visit, review the draft

This package guides you through building your own recommendation library and writing profile, preparing reports from spoken evaluations and local photo references, and enriching your library from reviewed reports. It contains the workflow and file helpers; bring your own reports, library and profile.

## Start in ChatGPT Work or another capable assistant

Attach the ready-made [AutoBericht_portable.md](AutoBericht_portable.md) from this repository and say:

> Follow the attached AutoBericht workflow and guide me through it. Start with the files I have provided. Build my personal library and writing profile if needed. For a new visit, choose and assemble the appropriate recommendation passages yourself and return a complete sidecar for my review. Keep imported customer answers intact. Leave Done for my review. Ask only for what is missing to take the next step.

The portable document includes all instructions, examples, the profile template and helper source code. The assistant can follow it directly and materialise helpers if its session supports file execution. Attaching instructions does not install a skill or grant missing tools.

Where ChatGPT Skills are available, use Plugins → Skills to upload the skill package. **autobericht-skill.zip** contains the full `autobericht` skill directory. If the uploader does not accept the archive, use the portable document with Create with chat. Workspace availability and permissions apply; this package has not yet been installed/tested in a colleague's ChatGPT workspace. [Official skill instructions](https://help.openai.com/en/articles/20001066)

For Codex, place the extracted `autobericht` directory in your configured skills directory and invoke `$autobericht`. A local Codex installation does not automatically install in ChatGPT.

## If you have no library yet

Upload copies of your previous reports with company information and site photos removed locally. You may supply them in batches. The assistant reads them, assesses your style, extracts reusable recommendations and guides you to supply an AutoBericht library/export for the actual question structure. You receive your own library, writing profile and a short coverage/gaps review. No other consultant's personal library is included.

## For a new report

Provide the latest sidecar, preferably after importing the self-assessment, your current library/profile, your visit recording and the photo references already in the sidecar. **The photos stay on the work computer.** Look at them locally and dictate their filenames/categories and what matters. No image upload, contact sheet or screen sharing is needed. The returned sidecar reconnects to the image files in the local app. The assistant preserves yes/no answers and comments, adapts findings lightly, interprets natural assessment language, selects and combines recommendations, and links photos and applicable checklists.

You can make one long MP3. Say photo filenames when useful and distinguish interviews, your judgment, corrections and anything to leave out. Try the full recording first. If the session cannot process it reliably, split it with overlap and original time offsets; the supplied helper handles that. If no transcription tool is already available, the skill checks whether Work can install and run a multilingual engine in its task environment. This needs suitable execution, model downloads and resources. German and French are supported by Whisper; Swiss German needs a real speech test and may be captured as Standard German before drafting in the report language. If the runtime cannot transcribe, supply a complete transcript from an available tool. Chunking solves size/runtime problems, not an absence of transcription capability.

An interview theme may inform several questions. The assistant gives each destination its own supported angle and avoids repeating the full paragraph or changing every related score.

You receive a complete draft sidecar, a readable preview, coverage/validation results and only the unresolved decisions. Review the substance and mark the completed items Done. The intended result is a nearly final draft; real audio transcription and professional judgment still need this review.

## After review

Say: “Close out this report and enrich my current library from the reviewed sidecar.” The assistant generalises useful additions, removes company particulars, deduplicates them and returns the next library copy. Use that version for the next project. A new library file does not automatically refresh a sidecar already created from an older library.

Repository access is optional for editing a compatible existing sidecar. For current app checks, supply an accessible checkout/source archive or connect a repository you can access. The skill does not grant repository access or fetch it automatically.

Read [validation.md](dictation-validation.md) for checks completed and the remaining real-recording test.

## Source of truth and distribution

The shared workflow and skill source live in this repository at [skills/autobericht](../../skills/autobericht/SKILL.md). Personal libraries, writing profiles and report collections belong in each consultant's private storage, such as their Drive. Live company photos remain on the work computer. Do not commit personal material to distribute the workflow.

From the repository root, generate a colleague package with:

```sh
python3 skills/autobericht/scripts/build_distribution.py --output output/autobericht-skill
```

The output contains a skill ZIP, a complete portable Markdown document and hashes. Generated copies can be handed to colleagues; update the repository source when improving the common workflow. The skill can also be installed directly from its source directory. Neither workflow requires sending the photo files to ChatGPT.

After changing the skill, regenerate the package and copy `output/autobericht-skill/AutoBericht_portable.md` to `docs/autobericht/AutoBericht_portable.md` before committing. This checked-in copy lets a colleague start without Python or a local packaging step.
