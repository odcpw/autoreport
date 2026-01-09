# AI WebGPU Spike Notes (Whisper + LiquidAI)

Date: 2026-01-07

## Whisper (ASR)

### What was failing
- `decoder_model_merged_q4.onnx` missing; Transformers.js ASR expected this filename and failed when only `decoder_model_q4.onnx` existed.
- Passing `{ array, sampling_rate }` to the ASR pipeline triggered `e.subarray is not a function` because the pipeline expects a raw typed array (Float32Array/Float64Array).
- WebGPU fp16 failed with `The device (webgpu) does not support fp16` because the adapter does not expose `shader-f16` on this machine.

### What works now
- Audio is passed as a Float32Array directly into the ASR pipeline.
- The fp16 path is preferred, but if `shader-f16` is missing we fall back to fp32 automatically.
- fp16 ONNX files are now present for Whisper tiny + base (multilingual) in `AI/models/Xenova/`:
  - tiny fp16: `encoder_model_fp16.onnx`, `decoder_model_fp16.onnx`, `decoder_with_past_model_fp16.onnx`, `decoder_model_merged_fp16.onnx`
  - base fp16: `encoder_model_fp16.onnx`, `decoder_model_fp16.onnx`, `decoder_with_past_model_fp16.onnx`, `decoder_model_merged_fp16.onnx`

### Notes
- If `shader-f16` is missing, WebGPU runs fp32. This is slower than fp16 but avoids errors.
- For WASM performance, `crossOriginIsolated` must be true to enable multithreading; current runs are single-threaded.
- Added ASR timing logs (pipeline load, audio decode, inference, total) plus backend info to help isolate slow runs.
- Added `?auto=1&audio=URL` support for auto-running ASR after load without manual file selection.

## LiquidAI (LFM2.5‑VL‑1.6B‑ONNX)

### Required pipeline (per model card)
- Run `embed_images(_fp16).onnx` and `embed_tokens(_fp16).onnx`.
- Merge image embeddings into token embeddings at the `<image>` token positions.
- Run `decoder_q4.onnx` with `inputs_embeds`, `attention_mask`, and KV cache.

### External data
- The model uses `.onnx_data` shards, which must be provided to ORT Web via `externalData` when creating sessions.

### Image + text preprocessing
- The official preprocessing uses the LFM2‑VL processor: tiling to 512×512, patching (16×16), normalization, and chat template insertion of `<image>` tokens.

### What was failing before
- We used dummy inputs for LiquidAI sessions, which does not satisfy the required preprocessing or embedding merge.
- Decoder sessions were missing external data.

### What’s implemented now
- Added external data loading for `*.onnx_data` files during session creation.
- Added a LiquidAI path that:
  - Loads tokenizer + processor from Transformers.js.
  - Builds the chat prompt with the image token.
  - Runs processor to get `pixel_values`, `pixel_attention_mask`, and `spatial_shapes`.
  - Runs `embed_tokens` and `embed_images`.
  - Merges image embeddings into token embeddings at image token positions.
  - Runs greedy decode with KV cache updates.
- Fixed a recursion bug in `resolveOrtProviderForModel` that could cause a stack overflow when selecting providers.
- Prevented the "Load ONNX (WebGPU)" path for LiquidAI fp16 models when `shader-f16` is unavailable.

### Known constraints
- WebGPU fp16 requires `shader-f16`. If unavailable, LiquidAI runs in fp32 and will be slower.
- Model card guidance uses fp16 encoders + q4 decoder for WebGPU.

## Lunar Lake (Arch) driver + WebGPU findings

Date: 2026-01-07

### Hardware + driver stack
- GPU: Intel Lunar Lake Arc Graphics 130V/140V (PCI 8086:64a0, rev 04).
- Kernel driver: `xe` (kernel 6.18.3-arch1-1).
- Mesa: 25.3.3; Vulkan Intel driver: 25.3.3; libva: 2.22.0; linux-firmware: 20251125-2.

### Vulkan capabilities (float16 + int8)
- Vulkan reports shader float16 support (`shaderFloat16=true`) and `VK_KHR_shader_float16_int8` is present.
- Float16 atomics are partially supported (some true, some false).
 - `vulkaninfo --summary` fails with `Permission denied` on `/dev/dri/renderD128` and `ERROR_INITIALIZATION_FAILED`.
   - User is not in the `render` group (`id -nG`), which likely blocks Vulkan device access.

### WebGPU capabilities (current browser)
- Initially, the WebGPU adapter feature list did NOT include `shader-f16`.
- Result: fp16 WGSL kernels failed with errors like:
  - `f16 type used without f16 extension enabled`
  - Invalid shader modules, invalid pipelines, invalid bind groups, and command buffer submission errors.
- This matched the crashes seen in the LiquidAI WebGPU demo (and local fp16 LiquidAI attempts).

### WebGPU after enabling Chrome developer flags
- After enabling WebGPU developer features, the adapter reports `shader-f16=true`.
- GPU report includes `shader-f16` and experimental subgroup features.
- ASR performance improved (90s clip finished in ~17.6s).

### LiquidAI WebGPU failure (ORT 1.18.0)
- Creating the WebGPU session fails with:
  - `no available backend found. ERR: [webgpu] RuntimeError: function signature mismatch`
- This occurs during ORT session creation for `embed_tokens_fp16.onnx` (before inference).
- Likely causes:
  - ORT WebGPU runtime mismatch (JS vs WASM files not from same version), or
  - ORT webgpu backend bug in 1.18.0 on this browser/driver combo.
- Next steps:
  - Try ORT 1.17.3 in the UI (loads from `vendor/ort-1.17.3/`).
  - If still failing, update ORT WebGPU files to a newer release (ensure JS + WASM are from the same version).

### LiquidAI WebGPU failure (ORT 1.17.3 + 1.18.0)
- The same `function signature mismatch` happens with both 1.17.3 and 1.18.0.
- Investigation: the WebGPU bundle defaults to `ort-wasm-simd.wasm` (non‑JSEP) while WebGPU expects the JSEP wasm.
- Fix applied in `ai-webgpu-spike.js`: when `bundle=webgpu`, map:
  - `ort-wasm-simd.wasm` → `ort-wasm-simd.jsep.wasm`
  - `ort-wasm-simd-threaded.wasm` → `ort-wasm-simd-threaded.jsep.wasm`
- Next step: reload and retry LiquidAI with this mapping; if it still fails, update ORT to a newer release or test a different model variant.

### Hugging Face Space parity work
- The Space uses newer deps: `onnxruntime-web` ^1.23.2 and `@huggingface/transformers` ^3.7.1. (package.json)
- It loads ONNX into memory (`Uint8Array`) and creates sessions with `executionProviders: ['webgpu','wasm']`.
- Added ORT 1.23.2 assets in `vendor/ort-1.23.2/` and updated the spike to:
  - default to 1.23.2,
  - allow `webgpu` + `wasm` providers,
  - stream large external data via URL, and
  - create sessions from `Uint8Array` buffers.
- 1.23.2 requires extra assets not in 1.18.0: `ort-wasm-simd-threaded.asyncify.mjs/.wasm` and `ort-wasm-simd-threaded.jsep.mjs`.
- Fixed `configureOrt()` to map the missing wasm filenames to the only available 1.23.2 wasm (`ort-wasm-simd-threaded*.wasm`).
- Lowered the external-data streaming threshold to 256MB to avoid loading 850MB–1.2GB shards into JS memory (decoder was aborting in WASM).
- Downloaded model-side config files into `AI/models/LiquidAI/LFM2.5-VL-1.6B-ONNX/` and added `preprocessor_config.json` (copied from `processor_config.json`) because Transformers.js looked for `preprocessor_config.json` locally.
- Added top-level `image_processor_type` to `processor_config.json` and `preprocessor_config.json` to satisfy Transformers.js AutoProcessor detection.
- AutoProcessor still fails with `No image_processor_type or feature_extractor_type found in the config`.
 - Added a local LiquidAI image processor fallback (`liquid-processor.js`) copied from the HF Space’s `vl-processor.js`.
   - `getProcessor()` now falls back to this local processor for LiquidAI models if AutoProcessor fails.
   - `runLiquidProcessor()` now builds `pixel_values`, `pixel_attention_mask`, and `spatial_shapes` from this fallback.
 - Added chat template loading from `chat_template.jinja` when tokenizer has no `chat_template`; `buildLiquidPrompt()` now loads the template from disk for LiquidAI models.
 - Fixed token embed feeds to always include `input_ids`/`attention_mask`/`position_ids` when metadata entries are empty or missing those keys.
 - Some ORT builds report token embed input metadata as numeric keys (e.g. `0`), but `session.inputNames` still lists `input_ids`.
   - `buildLiquidTokenInputs()` now accepts `inputNames` and uses them when metadata keys are numeric/empty.
   - Numeric-only metadata entries are now ignored to avoid feeding invalid input names like `0`.
 - Applied the same fix for image embedding inputs (`pixel_values`, `pixel_attention_mask`, `spatial_shapes`) using `session.inputNames`.
 - Added a decoder input builder that uses `session.inputNames` for `inputs_embeds`, `attention_mask`, `position_ids`, and cache tensors when metadata keys are numeric.
 - Updated cache initialization to use `decoderSession.inputNames`, ensuring all required `past_conv.*` and `past_key_values.*` inputs are prefilled.
 - Chat path now expands `<image>` to N copies of `image_token_id` (with optional `<|image_start|>`/`<|image_end|>`), matching the HF Space logic, then replaces those positions 1:1 with `image_features`.

### Notes
- WebGPU fp16 requires `shader-f16` to be exposed by the adapter and explicitly enabled at device creation.
- Vulkan float16 support does not guarantee WebGPU `shader-f16` exposure in the browser.
- If `shader-f16` remains unavailable, use fp32 or WASM for LiquidAI on this machine.

### WebGPU load fix (ORT 1.23.2 asyncify)
- Issue: `Failed to fetch dynamically imported module ... ort-wasm-simd-threaded.asyncify.mjs` and `initWasm()` failed.
- Cause: local HTTP server did not serve `.mjs` and `.wasm` with proper MIME types.
- Fix: updated `AutoBericht/tools/serve-autobericht.ps1` to return:
  - `.mjs` → `text/javascript; charset=utf-8`
  - `.wasm` → `application/wasm`
- Result: WebGPU ONNX sessions load successfully after restarting the server.

## Security & privacy assessment (local-only setup)

Date: 2026-01-08

### Scope & assumptions
- You run the spike from `AutoBericht/experiments/` and keep assets in `AutoBericht/AI/`.
- `Allow remote models` stays unchecked (local-only).
- `transformers.min.js` and `onnxruntime-web` files are loaded from the local `AI/vendor/` paths.
- No additional analytics scripts or browser extensions are capturing page data.

### What the code does (network-wise)
- The spike fetches **local** assets by default:
  - `../AI/vendor/transformers.min.js`
  - `../AI/vendor/ort-1.23.2/*`
  - `../AI/models/**` (configs + ONNX + `.onnx_data`)
- There are **no hard-coded external endpoints** in the runtime JS. All `fetch()` calls are relative to the local model path and local template/config files (see `ai-webgpu-spike.js`).
- With `Allow remote models` unchecked, the app does not contact external servers for model or runtime assets.

### Positive security posture (local-only mode)
- All model files, runtime files, and configs are served from disk under `AutoBericht/AI/`.
- The inference stack runs entirely in the browser (WebGPU/WASM) and does not require any cloud service.
- Log output in the spike shows only local file paths when loading models and assets.

### Data flow (local-only configuration)
- **Audio**: loaded from the file input, decoded in-browser, passed to the local Whisper model.
- **Images**: loaded from file input, preprocessed locally, embedded + decoded locally by LiquidAI ONNX.
- **Text prompts/responses**: all local in JS memory; not posted anywhere.
- **Logs**: displayed on-page; no upload/telemetry implemented.

### Can anyone “see” when you run a localhost server?
- A server bound to **127.0.0.1/localhost** is only reachable on the same machine.
- A server bound to **0.0.0.0** can be reached by other devices on your LAN (if they know the IP/port).
- It is **not** visible to the public internet unless you explicitly set up port-forwarding or a tunnel.

### Conclusion (this setup)
- With local assets and `Allow remote models` disabled, **audio, images, prompts, and outputs remain on the machine**.
- There is **no external communication** required to run Whisper or LiquidAI in this spike.
