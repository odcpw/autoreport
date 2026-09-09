# AutoBericht: speak through the visit, review the draft

Prepare a report from your MP3 recordings, project Sidecar, this shared skill and your personal style guide. The library is already loaded into the Sidecar. The skill selects and adapts that content to the visit and writes in your style.

## The four inputs for a report

| Input | What it supplies |
|---|---|
| MP3 recording(s) | Your observations, interviews, judgments and corrections |
| Project Sidecar | The question structure, loaded library and existing customer/photo metadata |
| Shared skill | The workflow and file helpers; use the ZIP or portable Markdown, or reuse the installed skill |
| Personal style guide | Your wording and report-writing conventions |

An “empty” Sidecar is not yet filled for this visit; the application has already loaded its library. You do not need to upload that library again, collect extra examples or make an author pack. Photos remain on your computer.

Keep your style guide privately and reuse it with each new project. “Author profile” elsewhere in the instructions means this same guide. If you need one created, the assistant can derive it from your existing anonymised reports and confirmed corrections. A few representative examples can live inside the guide; preparing a separate example collection is not part of the normal workflow.

## Start in ChatGPT Work or another capable assistant

After syncing the repository, attach [autobericht-skill.zip](../../autobericht-skill.zip) from the repository root to a ChatGPT or Copilot window that can read ZIPs. If ZIP extraction is unavailable, attach the complete [AutoBericht_portable.md](../../dist/autobericht/AutoBericht_portable.md) instead. Then say:

> Use this AutoBericht skill, my style guide, the project Sidecar and these MP3s to prepare my report. The library is already in the Sidecar. Choose and assemble the relevant passages and write in my style. Return the complete draft Sidecar and a clean preview, with working questions and software feedback separately. Keep imported customer answers intact. Leave Done for my review.

The portable document includes all instructions, examples, the profile template and helper source code. The assistant can follow it directly and materialise helpers if its session supports file execution. Attaching instructions does not install a skill or grant missing tools.

Where ChatGPT Skills are available, use Plugins → Skills to upload the skill package. **autobericht-skill.zip** contains the full `autobericht` skill directory. If the uploader does not accept the archive, use the portable document with Create with chat. Workspace availability and permissions apply; this package has not yet been installed/tested in a colleague's ChatGPT workspace. [Official skill instructions](https://help.openai.com/en/articles/20001066)

For Codex, place the extracted `autobericht` directory in your configured skills directory and invoke `$autobericht`. A local Codex installation does not automatically install in ChatGPT.

## If you have no library yet

Upload copies of your previous reports with company information and site photos removed locally. You may supply them in batches. The assistant reads them, assesses your style, extracts reusable recommendations and guides you to supply an AutoBericht library/export for the actual question structure. You receive your own library, writing profile and a short coverage/gaps review. No other consultant's personal library is included.

## For a new report

Provide the project Sidecar, preferably after importing the self-assessment, your style guide, the shared skill and your visit recording(s). Library text and photo references are already in the Sidecar. **The photos stay on the work computer.** Look at them locally and dictate their filenames/categories and what matters. No image upload, contact sheet or screen sharing is needed. The returned sidecar reconnects to the image files in the local app. The assistant preserves yes/no answers and comments, adapts findings lightly, interprets natural assessment language, selects and combines recommendations, and links photos and applicable checklists.

You can make one long MP3. Say photo filenames when useful and distinguish interviews, your judgment, corrections and anything to leave out. Try the full recording first. If the session cannot process it reliably, split it with overlap and original time offsets; the supplied helper handles that. If no transcription tool is already available, the skill checks whether Work can install and run a multilingual engine in its task environment. This needs suitable execution, model downloads and resources. German and French are supported by Whisper; Swiss German needs a real speech test and may be captured as Standard German before drafting in the report language. If the runtime cannot transcribe, supply a complete transcript from an available tool. Chunking solves size/runtime problems, not an absence of transcription capability.

An interview theme may inform several questions. The assistant gives each destination its own supported angle and avoids repeating the full paragraph or changing every related score.

The assistant establishes your writing conventions from your instructions, confirmed examples and library before drafting. It keeps useful existing wording and practical detail, then checks every changed report passage against those conventions. You do not need to prescribe every sentence or repair unwanted assistant narration.

You receive a complete draft sidecar and a clean preview containing the same report wording. Transcript coverage, source mappings and unresolved decisions stay in separate working material; comments about improving AutoBericht go into their own Markdown. Review the substance and mark the completed items Done. The intended result is usable report prose at the first delivery; real audio transcription and professional judgment still need this review.

## After review

Say: “Close out this report and enrich my current library from the reviewed sidecar.” The assistant generalises useful additions, removes company particulars, deduplicates them and returns the next library copy. Use that version for the next project. A new library file does not automatically refresh a sidecar already created from an older library.

Repository access is optional for editing a compatible existing sidecar. When current app behaviour or integration checks would help, the skill proposes reading or cloning https://github.com/odcpw/autoreport. It can reuse a checkout, source archive or authorized GitHub connection; when the user requests repository context and shell/network access allow it, it can clone the repo. This does not grant new credentials or authorize source changes.

Read [validation.md](dictation-validation.md) for checks completed and the remaining real-recording test.

## Source of truth and distribution

The shared workflow and skill source live in this repository at [skills/autobericht](../../skills/autobericht/SKILL.md). Personal libraries, writing profiles and report collections belong in each consultant's private storage, such as their Drive. Live company photos remain on the work computer. Do not commit personal material to distribute the workflow.

From the repository root, generate a colleague package with:

```sh
python3 skills/autobericht/scripts/build_distribution.py
```

The default build writes the ready-to-upload `autobericht-skill.zip` in the repository root, with the complete portable Markdown document and hashes in `dist/autobericht/`. Both `git pull` and the existing Windows sync script bring these files to colleagues without requiring Python on their computers. Update the repository source when improving the common workflow. The skill can also be installed directly from its source directory. Neither workflow requires sending the photo files to ChatGPT.

After changing the skill, run the builder and commit the source and generated files together. The default build also refreshes `docs/autobericht/AutoBericht_portable.md` to preserve its existing download URL. CI rejects stale generated copies. Use `--output <directory>` for a separate handoff build; that option does not update the compatibility copy.
