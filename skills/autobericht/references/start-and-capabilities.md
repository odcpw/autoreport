# Starting in ChatGPT Work, ChatGPT or Codex

## What a colleague provides

First-time setup can start with anonymised past reports (Word preferred; PDF or existing JSON/sidecars also useful). The assistant guides them to provide a base AutoBericht library or exported project when mapping requires its actual question IDs. They do not need to arrive with a writing profile already prepared.

For a new visit, provide:

- this skill package or the complete portable instruction document;
- the latest `project_sidecar.json` with imported self-assessment when available;
- their current `library_user_*.json` and author profile;
- the visit recording or full transcript;
- photo references already in the sidecar, or a permitted text-only filename manifest linking spoken identifiers to project-relative paths. The image files stay on the work computer.

A sidecar stores photo metadata, not the photo image bytes. Use the consultant’s descriptions as the evidence; do not request images or claim to inspect them. Do not substitute thumbnails, contact sheets, embedded report photos or screen sharing for uploads. Check that an export intended for Work contains no embedded image bytes. Reopen the returned sidecar on the work computer, where the app resolves its existing photo paths. A sidecar may contain some library text, but it is not always the complete latest reusable library.

## Minimal first message from the colleague

“Use the AutoBericht instructions I attached. Guide me through the process. If I have not yet built my library/profile, start from my anonymised reports. Otherwise use my library, sidecar and recording to prepare a complete sidecar for review. The photos stay on my work computer; use my descriptions and the sidecar’s filenames. Choose and combine the appropriate recommendation passages yourself. Keep customer answers and unrelated project data intact. Leave Done for my review. Tell me only what is missing to take the next step.”

## Capability check, without ceremony

Inspect the actual session for file reading/writing, ZIP extraction, audio transcription, code execution, model-download access and repository access; image access is not required. Do not ask the consultant to diagnose tools that you can inspect yourself. Use available capabilities and state only material gaps.

**Audio file accepted and transcription available:** try the original recording. Verify duration, language and processed coverage. If it fails because of size or runtime, split and resume. An MP3 attachment being accepted does not prove its entire speech has been transcribed.

**Audio transcription absent:** first check whether the session can prepare and run a multilingual transcription engine itself; see [transcription-runtime.md](transcription-runtime.md). Handle setup and processing when the tools permit it. If execution or necessary resources are absent, use the complete transcript from the user’s recorder, ChatGPT Record where available, or a transcription tool they choose. Preserve the audio for verification. A meeting summary is insufficient when details, negation or later corrections matter. Do not initiate a separately billed API service without the user’s authorization.

**Files can be read but not written:** produce the text and external edit plan, then explain that a tool-enabled session is required to return an edited sidecar. Do not claim the file was created.

**Repository context:** when current implementation or app-level validation would help, suggest using https://github.com/odcpw/autoreport. Reuse an accessible checkout or GitHub connection; offer a clone when execution and network access permit. Follow [repository-context.md](repository-context.md). A compatible existing sidecar can still be edited using this package alone. If the schema differs, inspect its source/docs or obtain an app-exported current sidecar before applying incompatible edits.

## Official product guidance checked on 6 September 2026

Capabilities and account availability may change; recheck these official pages when setup differs.

- ChatGPT supports reusable skills containing instructions, examples and code. Eligible workspaces can upload and share them through Plugins → Skills; workspace permissions apply. A normal attachment does not prove a skill was installed. If the uploader does not accept the supplied archive, use the complete portable document with “Create with chat”, or supply it as explicit task instructions. [Skills in ChatGPT](https://help.openai.com/en/articles/20001066)
- Work is intended for longer tasks and finished deliverables. Desktop Work can use granted local files; web/mobile Work requires uploads or connected sources. Voice in Work/Codex is documented for eligible desktop accounts. Local execution can still send task context to the cloud; it does not satisfy the photo restriction by itself. Keep image contents out of the task. [ChatGPT Work and Codex](https://help.openai.com/en/articles/20001275/)
- The generic upload guide lists document, spreadsheet and text formats. It does not establish universal transcription of arbitrary long MP3 attachments in every Work session. [Supported file types](https://help.openai.com/en/articles/8983675-what-types-of-files-are-supported)
- Record provides recording/transcription. The documentation checked lists supported paid workspaces, macOS availability and a four-hour session cap. This is a recording feature, not evidence that any uploaded MP3 will be handled identically. [ChatGPT Record](https://help.openai.com/en/articles/11487532-chatgpt-record)
- The separately used transcription API accepts MP3 and other audio formats, with a 25 MB file limit at the date checked. This is an API limit, not the ChatGPT attachment limit. Larger input can be split; API use needs the appropriate tools/account. [File transcription API](https://developers.openai.com/api/docs/guides/speech-to-text)
- The GitHub app retrieves repository content to which the colleague has granted access. Sharing this skill does not share repository permissions or credentials. Reading repository code is different from executing it or editing a local checkout. [Connecting GitHub](https://help.openai.com/en/articles/11145903)

## Long recordings

The consultant may make one long MP3 while thinking, reading and commenting. They need not pre-organise the report. Prefer natural spoken markers: “photo IMG_0123”, “question on instruction”, “from the interview”, “my assessment”, “correction”, “keep this out of the report”.

If splitting becomes necessary, a practical starting point is 20-minute chunks with 5 seconds of overlap. These are processing choices, not mandatory recording rules or product limits. `audio_chunks.py` provides offsets. Shorten chunks further for the actual service limit. Reconcile duplicated overlap and apply late corrections to earlier topics using original recording time.

Keep a progress file with source filename/hash, duration, chunks processed, failed intervals, current topic mappings and outstanding corrections. Resume unfinished ranges after interruptions. Distinguish silence, reading from a checklist, actual visit evidence and a proposed recommendation. Reading a negative generic finding aloud does not mean the author confirms that defect.

## Reuse and sharing

Use a dedicated ChatGPT project or reusable skill where available, but supply the current library version explicitly. Keep each company’s live case separate. Anonymised onboarding examples and generalised reusable language may be shared as authorized; the colleague package itself contains neither a personal corpus nor a client project. Installing in local Codex does not establish installation or synchronization in ChatGPT Work.
