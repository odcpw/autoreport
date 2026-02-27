# AI WebGPU Dev Guide (AutoBericht)

This guide documents the current AI spike setup, how to run it locally, and what needs to be present for Whisper + LiquidAI to work offline.

## Goals
- Run Whisper (ASR) and LiquidAI (vision/chat) **locally** in the browser.
- Keep all runtime assets in a self-contained `AutoBericht/AI/` bundle.
- Avoid external network calls when `Allow remote models` is unchecked.

## Folder layout (required)
```
AutoBericht/
├── AI/
│   ├── vendor/
│   │   ├── transformers.min.js
│   │   └── ort-1.23.2/
│   │       ├── ort.webgpu.min.js
│   │       └── ort-wasm-*.wasm / *.mjs
│   └── models/
│       ├── Xenova/
│       │   ├── whisper-tiny/
│       │   └── whisper-base/
│       └── LiquidAI/
│           └── LFM2.5-VL-1.6B-ONNX/
└── experiments/
    └── ai-webgpu-spike/
        ├── index.html
        ├── app.js
        └── liquid-processor.js
```

Notes:
- `AutoBericht/AI/` is **gitignored** and is intended to be copied in (or zipped) per machine.
- The spike uses **relative paths**: `../../AI/vendor/` and `../../AI/models/`.

## Local server (offline)
Use the built-in launcher from the `AutoBericht/` folder:
```
start-autobericht.cmd
```
This starts a local server bound to **127.0.0.1** and opens the UI.

Important: `AutoBericht/tools/serve-autobericht.ps1` serves `.mjs` and `.wasm` with correct MIME types. This is required for ORT 1.23.2 asyncify.

## Running the spike
1) Open: `http://127.0.0.1:<port>/AutoBericht/experiments/ai-webgpu-spike/index.html`
2) Click **Load library**.
3) Click **Check WebGPU**.
4) Click **Load ONNX (WebGPU)**.
5) Use **Analyze image** or **Ask (LiquidAI)** with a local image file.
6) Use **Run ASR** with a local audio file.

## Whisper (ASR)
Models supported in the local bundle:
- `Xenova/whisper-tiny` (fast)
- `Xenova/whisper-base` (better quality)

The spike chooses fp16 when `shader-f16` is available. If not, it falls back to fp32.

## LiquidAI (vision + chat)
Model: `LiquidAI/LFM2.5-VL-1.6B-ONNX`

Pipeline:
1) `embed_tokens_fp16.onnx`
2) `embed_images_fp16.onnx`
3) Merge image embeddings into token embeddings at `<image>` tokens
4) `decoder_q4.onnx` (greedy decode)

Preprocessing:
- Uses `liquid-processor.js` (local fallback) if AutoProcessor fails.
- Loads `chat_template.jinja` when tokenizer lacks a chat template.

## Local-only mode (privacy)
- Keep **Allow remote models** unchecked.
- Keep `transformers.min.js` and ORT files **local**.
- The spike fetches only local assets (`../AI/...`), and logs show local paths.

## Troubleshooting
### Error: `Failed to fetch dynamically imported module ... asyncify.mjs`
Fix:
- Ensure the local server serves `.mjs` and `.wasm` with correct MIME types.
- Restart the server after updating `serve-autobericht.ps1`.

### Error: `... file was not found locally`
Fix:
- Verify the file exists under `AutoBericht/AI/models/...`.
- Re-check the model id and ONNX filenames.

### WebGPU `shader-f16=false`
Fix:
- Enable WebGPU developer features in Chrome, or
- Use the WASM load button as a fallback.

## What to copy to another machine
Copy or unzip **only**:
- `AutoBericht/AI/`
- `AutoBericht/experiments/`

Then run the local server from `AutoBericht/`.
