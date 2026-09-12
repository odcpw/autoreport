# Obtain the transcript inside the available environment

The consultant supplies permitted audio; the assistant should handle transcription mechanics when the session allows it. The photos remain on the work computer throughout. Audio language and final report language are separate choices: someone may speak Swiss German and receive a French report in their established style.

## Capability and execution

Use an existing audio transcription tool first. Otherwise inspect the actual code/shell runtime, Python, ffmpeg, package/model availability, permitted network access, disk, memory and execution limits. ChatGPT Work can expose code/shell execution, with separate network controls; availability is configuration-dependent. Do not assume an unrestricted workstation, GPU or installed speech model. [Work execution and network controls](https://learn.chatgpt.com/docs/enterprise/chatgpt-work-cloud-security)

If the user has asked to process audio and task-environment installation is permitted, prepare an isolated environment, install a multilingual speech engine and download its model weights. Avoid system-wide installation or requesting administrator access for an optional route. Do not make the user run commands that the session can run itself. If the actual runtime blocks a dependency/download, report that specific limitation and choose an available alternative. Do not endlessly retry or call a separately billed API without authorization.

Open-source Whisper is one executable option. It needs Python dependencies, ffmpeg and downloaded model weights; larger models require more resources. Installing it does not require an OpenAI API key. Use a multilingual model rather than a model ending in `.en`. Adapt this illustrative recipe to the actual environment and capacity; it is not an instruction to install software just because the skill was attached. [Official Whisper setup](https://github.com/openai/whisper)

```sh
python3 -m venv .venv-transcription
.venv-transcription/bin/python -m pip install openai-whisper
.venv-transcription/bin/whisper visit.mp3 --model turbo --language German --task transcribe --output_format all --output_dir transcript
```

This is a Unix example; on Windows the environment executables are under `Scripts`. Check ffmpeg separately. For French speech use `--language French`. For mixed languages, avoid forcing the entire recording into one language; inspect natural segments and use an engine that handles the combination or transcribe segments with the correct hint. The command downloads weights if not cached. Test execution on a short excerpt before spending the runtime on the full recording; this does not require the consultant to re-record or pre-chunk anything. Then process the full original, splitting only if needed. Keep the raw transcript and structured segment timestamps.

## German, French and Swiss German

German and French are listed Whisper languages. Swiss German is not a separate language option in that list; `German` is a starting hint, not a guarantee of dialect accuracy. Do not invent a `gsw` or `de-CH` option unless the chosen engine documents it. [Whisper language definitions](https://github.com/openai/whisper/blob/main/whisper/tokenizer.py)

For Swiss German, aim to preserve meaning in a Standard German transcript before drafting in the requested report language. Do not require phonetic dialect spelling. Check a representative passage for negations, partial implementation, specialist SST terms, photo numbers and corrections. If recognition is poor, try an available stronger or dialect-suitable engine, or ask only about the uncertain passages; do not disguise guessed speech as a clean transcript. For a first dialect trial, report quality as untested until there is actual audio evidence. Chunking addresses size and runtime, not dialect recognition accuracy.

Preserve language changes, speaker attribution, reading-aloud versus assessment, and original timing. Transcription is not report drafting. The final French/German wording comes from the author's library and profile after the content has been understood. For Whisper, `--task translate` means translation into English and is inappropriate as a shortcut to a French report. [Whisper usage](https://github.com/openai/whisper)

If audio itself must also remain on the work computer, use an authorized local transcription process and provide only the permitted transcript to Work. A local file path opened by a cloud-connected assistant is not evidence that its contents remain offline. No photo access is needed for any of these transcription routes.
