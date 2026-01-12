import { processImage as liquidProcessImage } from "./liquid-processor.js";

const byId = (id) => document.getElementById(id);

const asrModelEl = byId("asr-model");
const asrModeEl = byId("asr-mode");
const loadLibBtn = byId("load-lib");
const holdBtn = byId("hold-to-talk");
const statusEl = byId("status");
const transcriptEl = byId("transcript");
const cleanedEl = byId("cleaned");
const structuredEl = byId("structured");
const notesVideoEl = byId("notes-video");
const notesPreviewEl = byId("notes-preview");
const notesStartCameraBtn = byId("notes-start-camera");
const notesCaptureBtn = byId("notes-capture");
const notesStopCameraBtn = byId("notes-stop-camera");
const notesFileEl = byId("notes-file");
const notesRunBtn = byId("notes-run");
const notesTranscriptEl = byId("notes-transcript");
const notesStructuredEl = byId("notes-structured");
const notesStatusEl = byId("notes-status");
const copyLogBtn = byId("copy-log");
const logEl = byId("log");

const DEFAULTS = {
  transformersUrl: "../AI/vendor/transformers.min.js",
  localModelPath: "../AI/models/",
};

const EKAS_CHAPTERS = [
  "1. Leitbild, Sicherheitsziele und Strategien",
  "2. Organisation",
  "3. Befähigung, Schulung, Kommunikation",
  "4. Sicherheitsstandards",
  "5. Gefährdungsermittlung, Risikobeurteilung",
  "6. Massnahmen",
  "7. Notfallorganisation",
  "8. Mitwirkung",
  "9. Gesundheitsschutz",
  "10. Kontrolle",
];
const EKAS_CHAPTER_LIST = EKAS_CHAPTERS.join("\n");

const CLEANUP_MODEL = "LiquidAI/LFM2.5-VL-1.6B-ONNX";
const CLEANUP_MAX_NEW_TOKENS_CAP = 256;
const CLEANUP_MIN_NEW_TOKENS = 64;
const STRUCTURE_MAX_NEW_TOKENS_CAP = 384;
const STRUCTURE_MIN_NEW_TOKENS = 120;
const NOTES_TRANSCRIBE_PROMPT =
  "You are transcribing a handwritten interview note page. Return ONLY the words you can read. " +
  "Preserve line breaks where possible. If a word is unclear, write [unclear]. " +
  "Do not summarize, infer, or add missing words. Do not include special tokens.";
const NOTES_TRANSCRIBE_MAX_TOKENS = 256;
const ASR_TIMEOUT_MS = 45000;
const STREAM_TIMESLICE_MS = 1200;
const ASR_LANGUAGE = "auto";
const LARGE_EXTERNAL_DATA_THRESHOLD = 256 * 1024 * 1024;
const ORT_PROVIDERS = {
  webgpu: ["webgpu", "wasm"],
  wasm: ["wasm"],
};
const LIQUIDAI_MODELS = [
  {
    match: "LFM2.5-VL-1.6B-ONNX",
    embedTokens: "onnx/embed_tokens_fp16.onnx",
    embedImages: "onnx/embed_images_fp16.onnx",
    decoder: "onnx/decoder_q4.onnx",
  },
];

const state = {
  lib: null,
  pipeline: null,
  env: null,
  pipelines: new Map(),
  tokenizers: new Map(),
  processors: new Map(),
  modelConfigs: new Map(),
  ortSessions: new Map(),
  externalDataCache: new Map(),
  webgpu: {
    supportsFp16: null,
  },
  notes: {
    stream: null,
    captureBlob: null,
    previewUrl: "",
    isRunning: false,
  },
  isRecording: false,
  isTranscribing: false,
  isCleaning: false,
  isStructuring: false,
  isStreaming: false,
  streamQueue: Promise.resolve(),
  streamSegment: 0,
  streamTranscript: "",
  streamAudioCtx: null,
  streamSource: null,
  streamProcessor: null,
  streamTimer: null,
  streamSamples: [],
  streamSamplesLen: 0,
  streamSampleRate: 0,
  recorder: null,
  chunks: [],
  stream: null,
};

function log(message) {
  if (!logEl) return;
  const line = `[${new Date().toISOString()}] ${message}`;
  logEl.textContent += `${line}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function setStatus(message) {
  statusEl.textContent = message;
  log(message);
}

async function copyLogToClipboard() {
  const content = logEl?.textContent || "";
  if (!content) {
    log("Log is empty; nothing to copy.");
    return;
  }
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(content);
      log("Copied log to clipboard.");
      return;
    } catch (err) {
      log(`Clipboard copy failed: ${err.message}`);
    }
  }
  const textarea = document.createElement("textarea");
  textarea.value = content;
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  try {
    const ok = document.execCommand("copy");
    log(ok ? "Copied log to clipboard." : "Clipboard copy failed.");
  } catch (err) {
    log(`Clipboard copy failed: ${err.message}`);
  } finally {
    document.body.removeChild(textarea);
  }
}

function stripModelTokens(text) {
  if (!text || typeof text !== "string") return "";
  return text
    .replace(/<\\|[^>]+\\|>/g, "")
    .replace(/<\\/?s>/g, "")
    .replace(/\\*\\*/g, "")
    .trim();
}

function formatMs(value) {
  if (!Number.isFinite(value)) return "n/a";
  if (value < 1000) return `${value.toFixed(1)}ms`;
  return `${(value / 1000).toFixed(2)}s`;
}

function logPerf(label, ms) {
  log(`${label}: ${formatMs(ms)}`);
}

async function ensureWebGpuFeatures() {
  if (!("gpu" in navigator)) {
    state.webgpu.supportsFp16 = false;
    return;
  }
  if (state.webgpu.supportsFp16 !== null) return;
  try {
    const adapter = await navigator.gpu.requestAdapter();
    state.webgpu.supportsFp16 = adapter?.features?.has("shader-f16") ?? false;
  } catch (err) {
    state.webgpu.supportsFp16 = false;
  }
}

function applyEnv() {
  if (!state.env) return;
  state.env.allowRemoteModels = false;
  state.env.allowLocalModels = true;
  state.env.localModelPath = DEFAULTS.localModelPath;
  state.env.useBrowserCache = true;
}

function getOrtConfig() {
  const fallback = { version: "1.23.2", base: "../AI/vendor/ort-1.23.2/", bundle: "webgpu" };
  if (window.__ortConfig && window.__ortConfig.version && window.__ortConfig.base) {
    return window.__ortConfig;
  }
  return fallback;
}

async function ensureOrtLoaded() {
  if (window.__ortLoadPromise) {
    await window.__ortLoadPromise;
  }
  if (!window.ort) {
    throw new Error("onnxruntime-web not loaded.");
  }
}

function configureOrt() {
  if (!window.ort?.env?.wasm) return;
  const { version, base: basePath, bundle } = getOrtConfig();
  const base = new URL(basePath, window.location.href).toString();
  const cacheBust = `?v=${version}`;
  const wantsWebgpu = (bundle || "").toLowerCase() === "webgpu";
  const isOrt123 = /^1\\.23\\./.test(version);
  let wasmPaths = {
    "ort-wasm.wasm": `${base}ort-wasm.wasm${cacheBust}`,
    "ort-wasm-simd.wasm": `${base}ort-wasm-simd.wasm${cacheBust}`,
    "ort-wasm-simd.jsep.wasm": `${base}ort-wasm-simd.jsep.wasm${cacheBust}`,
    "ort-wasm-threaded.wasm": `${base}ort-wasm-threaded.wasm${cacheBust}`,
    "ort-wasm-simd-threaded.wasm": `${base}ort-wasm-simd-threaded.wasm${cacheBust}`,
    "ort-wasm-simd-threaded.jsep.wasm": `${base}ort-wasm-simd-threaded.jsep.wasm${cacheBust}`,
  };
  if (isOrt123) {
    const simdThreaded = `${base}ort-wasm-simd-threaded.wasm${cacheBust}`;
    const simdThreadedJsep = `${base}ort-wasm-simd-threaded.jsep.wasm${cacheBust}`;
    wasmPaths = {
      "ort-wasm.wasm": simdThreaded,
      "ort-wasm-simd.wasm": simdThreaded,
      "ort-wasm-simd.jsep.wasm": simdThreadedJsep,
      "ort-wasm-threaded.wasm": simdThreaded,
      "ort-wasm-simd-threaded.wasm": wantsWebgpu ? simdThreadedJsep : simdThreaded,
      "ort-wasm-simd-threaded.jsep.wasm": simdThreadedJsep,
    };
  }
  if (wantsWebgpu) {
    wasmPaths["ort-wasm-simd.wasm"] = `${base}ort-wasm-simd.jsep.wasm${cacheBust}`;
    wasmPaths["ort-wasm-simd-threaded.wasm"] = `${base}ort-wasm-simd-threaded.jsep.wasm${cacheBust}`;
  }
  window.ort.env.wasm.wasmPaths = wasmPaths;
  window.ort.env.wasm.simd = true;
  window.ort.env.wasm.proxy = false;
  const canThread = typeof crossOriginIsolated !== "undefined" && crossOriginIsolated;
  window.ort.env.wasm.numThreads = canThread ? Math.min(4, navigator.hardwareConcurrency || 1) : 1;
}

function resolveOrtProvider() {
  const device = getDeviceOption();
  if (device === "webgpu" && "gpu" in navigator) return ORT_PROVIDERS.webgpu;
  return ORT_PROVIDERS.wasm;
}

async function resolveExternalOrtData(modelPath) {
  if (!modelPath.endsWith(".onnx")) return null;
  const dataUrl = `${modelPath}_data`;
  if (state.externalDataCache.has(dataUrl)) {
    return state.externalDataCache.get(dataUrl);
  }
  const loader = (async () => {
    try {
      const head = await fetch(dataUrl, { method: "HEAD", cache: "no-store" });
      if (!head.ok) return null;
      const size = Number(head.headers.get("content-length") || "0");
      const path = dataUrl.split("/").pop();
      if (size && size > LARGE_EXTERNAL_DATA_THRESHOLD) {
        log(`ORT external data streaming: ${path} (${size} bytes)`);
        return [{ path, data: dataUrl }];
      }
      const response = await fetch(dataUrl, { cache: "no-store" });
      if (!response.ok) return null;
      const buffer = new Uint8Array(await response.arrayBuffer());
      log(`ORT external data loaded: ${path} (${buffer.byteLength} bytes)`);
      return [{ path, data: buffer }];
    } catch (err) {
      log(`ORT external data fetch failed: ${err.message}`);
      return null;
    }
  })();
  state.externalDataCache.set(dataUrl, loader);
  return loader;
}

async function loadOrtSession(modelPath, providers) {
  await ensureOrtLoaded();
  const cacheKey = `${modelPath}::${providers.join(",")}`;
  if (state.ortSessions.has(cacheKey)) return state.ortSessions.get(cacheKey);
  configureOrt();
  const response = await fetch(modelPath, { cache: "no-store" });
  if (!response.ok) throw new Error(`ORT fetch failed (${response.status}): ${modelPath}`);
  const onnxBuffer = new Uint8Array(await response.arrayBuffer());
  const externalData = await resolveExternalOrtData(modelPath);
  log(`ORT create session: ${modelPath} providers=${providers.join(",")}`);
  const session = await window.ort.InferenceSession.create(onnxBuffer, {
    executionProviders: providers,
    externalData: externalData || undefined,
  });
  state.ortSessions.set(cacheKey, session);
  return session;
}

async function loadLibrary() {
  setStatus(`Loading library from ${DEFAULTS.transformersUrl} ...`);
  try {
    const mod = await import(DEFAULTS.transformersUrl);
    if (!mod.pipeline || !mod.env) {
      throw new Error("Module missing pipeline/env exports.");
    }
    state.lib = mod;
    state.pipeline = mod.pipeline;
    state.env = mod.env;
    applyEnv();
    setStatus("Library loaded. Ready.");
  } catch (err) {
    setStatus(`Library load failed: ${err.message}`);
    throw err;
  }
}

async function ensureLibraryLoaded() {
  if (state.pipeline && state.env) return;
  await loadLibrary();
}

function getDeviceOption() {
  if ("gpu" in navigator) return "webgpu";
  return "wasm";
}

function getAsrDtypeOverride() {
  const option = asrModelEl?.selectedOptions?.[0];
  const dtype = option?.dataset?.dtype;
  return dtype ? dtype.toLowerCase() : null;
}

function getAsrMode() {
  return asrModeEl?.value === "stream" ? "stream" : "batch";
}

async function resolveAsrDtypeForDevice(deviceChoice) {
  await ensureWebGpuFeatures();
  const wantsWebgpu = deviceChoice === "webgpu";
  if (wantsWebgpu && state.webgpu.supportsFp16 === false) {
    return "fp32";
  }
  if (deviceChoice === "wasm") {
    return "fp32";
  }
  return "fp16";
}

async function pickAvailableAsrDtypeForDevice(modelId, deviceChoice) {
  const preferred = await resolveAsrDtypeForDevice(deviceChoice);
  const tried = [preferred];
  try {
    await ensureAsrModelFiles(modelId, preferred);
    return preferred;
  } catch (err) {
    if (preferred !== "fp16") {
      tried.push("fp16");
      try {
        await ensureAsrModelFiles(modelId, "fp16");
        log(`ASR fp32 missing; falling back to fp16.`);
        return "fp16";
      } catch (fp16Err) {
        throw err;
      }
    }
    throw err;
  }
}

function logAudioStats(audio, sampleRate) {
  if (!audio?.length) return;
  let max = 0;
  let sumSq = 0;
  for (let i = 0; i < audio.length; i += 1) {
    const v = Math.abs(audio[i]);
    if (v > max) max = v;
    sumSq += audio[i] * audio[i];
  }
  const rms = Math.sqrt(sumSq / audio.length);
  const duration = audio.length / sampleRate;
  log(`ASR audio stats: duration=${duration.toFixed(2)}s max=${max.toFixed(4)} rms=${rms.toFixed(4)}`);
  if (max < 0.01) {
    log("ASR warning: audio level is very low; mic may be muted or too quiet.");
  }
}

function getLocalModelBase(modelId) {
  const root = DEFAULTS.localModelPath;
  return root.endsWith("/") ? `${root}${modelId}/` : `${root}/${modelId}/`;
}

async function ensureAsrModelFiles(modelId, dtype) {
  const base = `${getLocalModelBase(modelId)}onnx/`;
  const lowered = dtype.toLowerCase();
  const suffix = lowered === "fp32" || lowered === "float32" ? "" : `_${lowered}`;
  const files = [
    `encoder_model${suffix}.onnx`,
    `decoder_model${suffix}.onnx`,
    `decoder_with_past_model${suffix}.onnx`,
    `decoder_model_merged${suffix}.onnx`,
  ];
  const missing = [];
  for (const file of files) {
    try {
      const response = await fetch(`${base}${file}`, { method: "HEAD", cache: "no-store" });
      if (!response.ok) missing.push(file);
    } catch (err) {
      missing.push(file);
    }
  }
  if (missing.length) {
    throw new Error(`Missing ASR ${lowered} files in ${base}: ${missing.join(", ")}`);
  }
}

async function decodeAudioBlob(blob) {
  const data = await blob.arrayBuffer();
  const AudioCtx = window.AudioContext || window.webkitAudioContext;
  if (!AudioCtx) throw new Error("AudioContext not available in this browser.");
  const audioCtx = new AudioCtx();
  const decoded = await audioCtx.decodeAudioData(data.slice(0));
  const targetRate = 16000;
  let buffer = decoded;
  if (decoded.sampleRate !== targetRate) {
    const length = Math.ceil(decoded.duration * targetRate);
    const offline = new OfflineAudioContext(1, length, targetRate);
    const source = offline.createBufferSource();
    source.buffer = decoded;
    source.connect(offline.destination);
    source.start(0);
    buffer = await offline.startRendering();
  }
  const channelData = buffer.getChannelData(0);
  const audio = channelData instanceof Float32Array ? new Float32Array(channelData) : Float32Array.from(channelData);
  audioCtx.close();
  return { audio, sampling_rate: targetRate };
}

function loadImageFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error("Image decode failed."));
      img.src = reader.result;
    };
    reader.onerror = () => reject(new Error("File read failed."));
    reader.readAsDataURL(file);
  });
}

async function getPipeline(task, modelId, extraOptions = {}) {
  const device = extraOptions.device ?? getDeviceOption();
  const key = JSON.stringify({ task, modelId, device, extraOptions });
  if (state.pipelines.has(key)) return state.pipelines.get(key);
  await ensureLibraryLoaded();
  applyEnv();
  const options = { ...extraOptions };
  if (device) options.device = device;
  const pipe = await state.pipeline(task, modelId, options);
  state.pipelines.set(key, pipe);
  return pipe;
}

function resolveLogitsOutput(result) {
  const entries = Object.entries(result || {});
  if (!entries.length) return null;
  const found = entries.find(([name]) => name.toLowerCase().includes("logits"));
  return found ? found[1] : entries[0][1];
}

function argmaxLogits(logitsTensor) {
  const { data, dims } = logitsTensor;
  const vocab = dims[dims.length - 1];
  const seq = dims.length >= 3 ? dims[dims.length - 2] : dims[0];
  const offset = (seq - 1) * vocab;
  let best = 0;
  let bestVal = -Infinity;
  for (let i = 0; i < vocab; i += 1) {
    const val = data[offset + i];
    if (val > bestVal) {
      bestVal = val;
      best = i;
    }
  }
  return best;
}

async function getTokenizer(modelId) {
  const base = getLocalModelBase(modelId);
  if (state.tokenizers.has(base)) return state.tokenizers.get(base);
  if (!state.env || !state.lib?.AutoTokenizer) {
    throw new Error("Transformers.js AutoTokenizer not available.");
  }
  const tokenizer = await state.lib.AutoTokenizer.from_pretrained(modelId, {
    local_files_only: true,
  });
  state.tokenizers.set(base, tokenizer);
  return tokenizer;
}

function resolveOrtModelConfig(modelId) {
  return LIQUIDAI_MODELS.find((entry) => modelId?.includes(entry.match)) || null;
}

function modelUsesFp16(config) {
  return Boolean(config?.embedTokens?.includes("_fp16") || config?.embedImages?.includes("_fp16"));
}

async function resolveOrtProviderForModel(modelId, config) {
  await ensureWebGpuFeatures();
  const providers = resolveOrtProvider();
  const usesFp16 = modelUsesFp16(config);
  if (providers === ORT_PROVIDERS.webgpu && usesFp16 && state.webgpu.supportsFp16 === false) {
    log("WebGPU lacks shader-f16; falling back to WASM for LiquidAI fp16 models.");
    return ORT_PROVIDERS.wasm;
  }
  return providers;
}

function isLiquidModelId(modelId) {
  if (!modelId) return false;
  if (modelId.toLowerCase().includes("liquidai")) return true;
  return Boolean(resolveOrtModelConfig(modelId));
}

function createLiquidProcessor() {
  return {
    __liquidProcessor: true,
    async processImage(img) {
      return liquidProcessImage(img);
    },
  };
}

async function getProcessor(modelId) {
  const base = getLocalModelBase(modelId);
  if (state.processors.has(base)) return state.processors.get(base);
  if (!state.lib?.AutoProcessor) {
    if (isLiquidModelId(modelId)) {
      const processor = createLiquidProcessor();
      state.processors.set(base, processor);
      return processor;
    }
    throw new Error("Transformers.js AutoProcessor not available. Update transformers.min.js.");
  }
  try {
    const processor = await state.lib.AutoProcessor.from_pretrained(modelId, {
      local_files_only: true,
    });
    state.processors.set(base, processor);
    return processor;
  } catch (err) {
    if (isLiquidModelId(modelId)) {
      log(`AutoProcessor failed for ${modelId}. Falling back to local image processor.`);
      const processor = createLiquidProcessor();
      state.processors.set(base, processor);
      return processor;
    }
    throw err;
  }
}

async function getModelConfig(modelId) {
  const base = getLocalModelBase(modelId);
  if (state.modelConfigs.has(base)) return state.modelConfigs.get(base);
  const response = await fetch(`${base}config.json`, { cache: "no-store" });
  if (!response.ok) return null;
  const config = await response.json();
  state.modelConfigs.set(base, config);
  return config;
}

function normalizeTokenIds(value) {
  if (!value) return null;
  if (Array.isArray(value)) return value;
  if (Array.isArray(value.ids)) return value.ids;
  if (Array.isArray(value.input_ids)) return value.input_ids;
  if (Array.isArray(value.inputIds)) return value.inputIds;
  if (value?.data && Array.isArray(value.data)) return value.data;
  return null;
}

async function tokenizeQuestion(tokenizer, text) {
  if (typeof tokenizer.encode === "function") {
    const encoded = await tokenizer.encode(text);
    const ids = normalizeTokenIds(encoded);
    if (ids) return ids;
  }
  if (typeof tokenizer === "function") {
    const encoded = await tokenizer(text);
    const ids = normalizeTokenIds(encoded);
    if (ids) return ids;
  }
  throw new Error("Tokenizer output format not recognized.");
}

function decodeTokens(tokenizer, ids) {
  if (typeof tokenizer.decode === "function") {
    try {
      return tokenizer.decode(ids);
    } catch (err) {
      return ids.join(" ");
    }
  }
  return ids.join(" ");
}

function toBigInt64Array(data) {
  if (data instanceof BigInt64Array) return data;
  if (Array.isArray(data)) return new BigInt64Array(data.map((value) => BigInt(value)));
  if (ArrayBuffer.isView(data)) {
    const values = Array.from(data, (value) => BigInt(value));
    return new BigInt64Array(values);
  }
  return new BigInt64Array([]);
}

function buildAttentionMaskFromLength(length) {
  const data = new BigInt64Array(length);
  data.fill(1n);
  return data;
}

function buildPositionIdsFromLength(length) {
  const data = new BigInt64Array(length);
  for (let i = 0; i < length; i += 1) {
    data[i] = BigInt(i);
  }
  return data;
}

async function ensureChatTemplate(modelId, tokenizer) {
  if (tokenizer?.chat_template) return tokenizer.chat_template;
  const base = getLocalModelBase(modelId);
  try {
    const response = await fetch(`${base}chat_template.jinja`, { cache: "no-store" });
    if (!response.ok) return null;
    const template = await response.text();
    if (template && tokenizer) {
      tokenizer.chat_template = template;
    }
    return template || null;
  } catch (err) {
    return null;
  }
}

function resolveTokenId(tokenizer, token, fallback) {
  if (!tokenizer || !token) return fallback ?? null;
  if (typeof tokenizer.convert_tokens_to_ids === "function") {
    const id = tokenizer.convert_tokens_to_ids(token);
    if (Number.isFinite(id) && id >= 0) return id;
  }
  return fallback ?? null;
}

function expandImageTokens(inputIds, tokensPerImage, imageTokenId, imageStartTokenId, imageEndTokenId) {
  if (!imageTokenId || !tokensPerImage?.length) return inputIds;
  const expanded = [];
  let imageIdx = 0;
  for (const id of inputIds) {
    if (id === imageTokenId && imageIdx < tokensPerImage.length) {
      if (imageStartTokenId != null) expanded.push(imageStartTokenId);
      const count = tokensPerImage[imageIdx];
      for (let i = 0; i < count; i += 1) {
        expanded.push(imageTokenId);
      }
      if (imageEndTokenId != null) expanded.push(imageEndTokenId);
      imageIdx += 1;
    } else {
      expanded.push(id);
    }
  }
  return expanded;
}

function normalizeProcessorTensor(value, desiredType) {
  if (!value) return null;
  if (value instanceof window.ort.Tensor) return value;
  if (value?.data && (value.dims || value.shape)) {
    const dims = value.dims || value.shape;
    const inferred = value.type || (value.data instanceof BigInt64Array ? "int64" : "float32");
    const type = desiredType || inferred;
    let data = value.data;
    if (type === "int64" && !(data instanceof BigInt64Array)) {
      data = toBigInt64Array(data);
    }
    if (type === "float32" && !(data instanceof Float32Array)) {
      data = new Float32Array(Array.from(data));
    }
    return new window.ort.Tensor(type, data, dims);
  }
  return null;
}

async function buildLiquidPrompt(tokenizer, modelId, question) {
  const promptText = question || "Describe the image.";
  const messages = [
    {
      role: "user",
      content: [{ type: "image" }, { type: "text", text: promptText }],
    },
  ];
  if (typeof tokenizer.apply_chat_template === "function") {
    const template = await ensureChatTemplate(modelId, tokenizer);
    if (template) {
      return tokenizer.apply_chat_template(messages, {
        add_generation_prompt: true,
        tokenize: false,
        chat_template: template,
      });
    }
    if (tokenizer.chat_template) {
      return tokenizer.apply_chat_template(messages, { add_generation_prompt: true, tokenize: false });
    }
  }
  return `User: ${promptText}\nAssistant:`;
}

async function runLiquidProcessor(processor, img, prompt) {
  if (processor?.__liquidProcessor && typeof processor.processImage === "function") {
    const processed = await processor.processImage(img);
    const patchCount = processed.shape?.[1] ?? 0;
    return {
      pixel_values: { data: processed.pixelValues, dims: processed.shape },
      pixel_attention_mask: { data: processed.attentionMask, dims: [processed.numTiles, patchCount] },
      spatial_shapes: { data: processed.spatialShapes, dims: [processed.numTiles, 2] },
    };
  }
  const attempts = [
    () => processor([img], { text: prompt, return_tensors: "np" }),
    () => processor(img, { text: prompt, return_tensors: "np" }),
    () => processor({ images: [img], text: prompt, return_tensors: "np" }),
    () => processor({ image: img, text: prompt, return_tensors: "np" }),
    () => processor(img, prompt),
  ];
  let lastError = null;
  for (const attempt of attempts) {
    try {
      const result = await attempt();
      if (result && (result.pixel_values || result.pixel_attention_mask || result.spatial_shapes)) {
        return result;
      }
    } catch (err) {
      lastError = err;
    }
  }
  throw lastError || new Error("LiquidAI processor call failed.");
}

function buildLiquidImageInputs(inputMeta, inputNames, processed) {
  const pixelValuesRaw = processed.pixel_values;
  const pixelAttentionRaw = processed.pixel_attention_mask || processed.pixel_mask || processed.attention_mask;
  const spatialShapesRaw = processed.spatial_shapes;
  const inputs = {};
  const { entries, namesFromSession, numericOnly, useNames } = resolveInputNames(
    inputMeta,
    inputNames,
    ["pixel_values", "pixel_attention_mask", "spatial_shapes"]
  );
  let hasPixelValues = false;
  let hasPixelMask = false;
  let hasSpatial = false;

  for (const [name, meta] of entries) {
    if (numericOnly && namesFromSession.length) {
      continue;
    }
    const lower = name.toLowerCase();
    if ((lower.includes("pixel") || lower.includes("image")) && pixelValuesRaw) {
      inputs[name] = normalizeProcessorTensor(pixelValuesRaw, meta.type || "float32");
      hasPixelValues = true;
      continue;
    }
    if (pixelAttentionRaw && lower.includes("mask")) {
      inputs[name] = normalizeProcessorTensor(pixelAttentionRaw, meta.type || "int64");
      hasPixelMask = true;
      continue;
    }
    if (spatialShapesRaw && lower.includes("spatial")) {
      inputs[name] = normalizeProcessorTensor(spatialShapesRaw, meta.type || "int64");
      hasSpatial = true;
      continue;
    }
    const dims = resolveDimensions(name, meta);
    const total = dims.reduce((acc, value) => acc * value, 1);
    const data = toOrtTypedArray(meta.type, total);
    fillDefaultValues(meta.type, data);
    inputs[name] = new window.ort.Tensor(meta.type || "float32", data, dims);
  }

  if (!hasPixelValues && useNames.includes("pixel_values") && pixelValuesRaw) {
    inputs.pixel_values = normalizeProcessorTensor(pixelValuesRaw, "float32");
  }
  if (!hasPixelMask && useNames.includes("pixel_attention_mask") && pixelAttentionRaw) {
    inputs.pixel_attention_mask = normalizeProcessorTensor(pixelAttentionRaw, "int64");
  }
  if (!hasSpatial && useNames.includes("spatial_shapes") && spatialShapesRaw) {
    inputs.spatial_shapes = normalizeProcessorTensor(spatialShapesRaw, "int64");
  }
  return inputs;
}

function mergeLiquidEmbeds(tokenEmbeds, imageEmbeds, inputIds, imageTokenId) {
  if (!tokenEmbeds?.data || !imageEmbeds?.data) return tokenEmbeds;
  const tokenDims = tokenEmbeds.dims;
  const imageDims = imageEmbeds.dims;
  const hidden = tokenDims[tokenDims.length - 1];
  const imageTokens = imageDims.length === 3 ? imageDims[1] : imageDims[0];
  const imageHidden = imageDims.length === 3 ? imageDims[2] : imageDims[1];
  if (hidden !== imageHidden) {
    log(`LiquidAI embed mismatch: token hidden=${hidden} image hidden=${imageHidden}`);
    return tokenEmbeds;
  }
  const positions = inputIds
    .map((value, index) => (value === imageTokenId ? index : -1))
    .filter((value) => value >= 0);
  const maxCopies = Math.min(positions.length, imageTokens);
  for (let i = 0; i < maxCopies; i += 1) {
    const pos = positions[i];
    const dstBase = pos * hidden;
    const srcBase = i * hidden;
    tokenEmbeds.data.set(imageEmbeds.data.subarray(srcBase, srcBase + hidden), dstBase);
  }
  return tokenEmbeds;
}

async function buildCleanupPrompt(tokenizer, modelId, text) {
  const promptText =
    `Clean up the transcript below. Remove filler words, fix punctuation and casing, ` +
    `keep the original meaning, and return a single short paragraph.\\n\\nTranscript:\\n${text}`;
  const messages = [
    {
      role: "user",
      content: promptText,
    },
  ];
  if (typeof tokenizer.apply_chat_template === "function") {
    const template = await ensureChatTemplate(modelId, tokenizer);
    if (template) {
      return tokenizer.apply_chat_template(messages, {
        add_generation_prompt: true,
        tokenize: false,
        chat_template: template,
      });
    }
    if (tokenizer.chat_template) {
      return tokenizer.apply_chat_template(messages, { add_generation_prompt: true, tokenize: false });
    }
  }
  return `User: ${promptText}\\nAssistant:`;
}

async function buildStructurePrompt(tokenizer, modelId, text) {
  const promptText =
    `Use ONLY the following EKAS chapters (exact headings and order):\\n` +
    `${EKAS_CHAPTER_LIST}\\n\\n` +
    `Organize the interview notes below into those 10 chapters using ONLY the provided text. ` +
    `Output plain text with exactly those 10 headings, one per line, in the same order. ` +
    `Under each heading, add bullet points only when the notes explicitly mention something relevant. ` +
    `Do not create a fixed number of bullets per chapter; include only what is present. ` +
    `Do not add generic or filler bullets. If a chapter has no relevant notes, leave it blank after the heading. ` +
    `Do not invent details or rephrase beyond minimal cleanup. No markdown.\\n\\nNotes:\\n${text}`;
  const messages = [
    {
      role: "user",
      content: promptText,
    },
  ];
  if (typeof tokenizer.apply_chat_template === "function") {
    const template = await ensureChatTemplate(modelId, tokenizer);
    if (template) {
      return tokenizer.apply_chat_template(messages, {
        add_generation_prompt: true,
        tokenize: false,
        chat_template: template,
      });
    }
    if (tokenizer.chat_template) {
      return tokenizer.apply_chat_template(messages, { add_generation_prompt: true, tokenize: false });
    }
  }
  return `User: ${promptText}\\nAssistant:`;
}

function resolveInputNames(inputMeta, inputNames, fallbackNames) {
  const entries = inputMeta instanceof Map ? Array.from(inputMeta.entries()) : Object.entries(inputMeta || {});
  const keyNames = entries.map(([name]) => name);
  const namesFromSession = Array.isArray(inputNames) ? inputNames : [];
  const keyNamesNumeric =
    keyNames.length > 0 && keyNames.every((name) => typeof name === "string" && /^\d+$/.test(name));
  const sessionNamesNumeric =
    namesFromSession.length > 0 &&
    namesFromSession.every((name) => typeof name === "string" && /^\d+$/.test(name));
  const numericOnly = keyNamesNumeric && (namesFromSession.length === 0 || sessionNamesNumeric);
  const useNames =
    namesFromSession.length > 0 && !sessionNamesNumeric
      ? namesFromSession
      : (!entries.length || numericOnly) && namesFromSession.length
        ? namesFromSession
        : numericOnly && fallbackNames?.length
          ? fallbackNames
          : keyNames;
  return { entries, keyNames, namesFromSession, numericOnly, useNames };
}

function toOrtTypedArray(type, length) {
  switch (type) {
    case "int64":
      return new BigInt64Array(length);
    case "int32":
      return new Int32Array(length);
    case "float16":
      return new Uint16Array(length);
    case "bool":
      return new Uint8Array(length);
    default:
      return new Float32Array(length);
  }
}

function fillDefaultValues(type, array) {
  if (type === "int64") {
    array.fill(1n);
    return;
  }
  if (type === "int32") {
    array.fill(1);
    return;
  }
  array.fill(0);
}

function resolveDimensions(name, meta) {
  const dims = (meta?.dimensions || []).map((dim) => (typeof dim === "number" && dim > 0 ? dim : 0));
  const hasDynamic = dims.some((dim) => dim === 0);
  const next = dims.length ? [...dims] : [1];
  if (hasDynamic) {
    const isTokens = name.toLowerCase().includes("input") || name.toLowerCase().includes("token");
    if (isTokens) {
      if (next[0] === 0) next[0] = 1;
      if (next.length > 1 && next[1] === 0) next[1] = 8;
    } else {
      for (let i = 0; i < next.length; i += 1) {
        if (next[i] === 0) next[i] = 1;
      }
    }
  }
  return next;
}

function buildLiquidTokenInputs(inputMeta, inputNames, inputIds, attentionMask, positionIds) {
  const inputs = {};
  const inputIdsTensor = new window.ort.Tensor("int64", toBigInt64Array(inputIds), [1, inputIds.length]);
  const resolvedAttention = attentionMask || buildAttentionMaskFromLength(inputIds.length);
  const resolvedPosition = positionIds || buildPositionIdsFromLength(inputIds.length);
  const attentionMaskTensor = new window.ort.Tensor("int64", resolvedAttention, [1, resolvedAttention.length]);
  const positionIdsTensor = new window.ort.Tensor("int64", resolvedPosition, [1, resolvedPosition.length]);

  const { entries, namesFromSession, numericOnly, useNames } = resolveInputNames(
    inputMeta,
    inputNames,
    ["input_ids", "attention_mask", "position_ids"]
  );
  const hasNonNumericNames =
    namesFromSession.some((name) => typeof name === "string" && !/^\\d+$/.test(name)) ||
    useNames.some((name) => typeof name === "string" && !/^\\d+$/.test(name));
  const allNames = Array.from(new Set([...useNames, ...namesFromSession]));
  const keyNames = entries.map(([name]) => name);
  for (const name of keyNames) {
    if (!allNames.includes(name)) allNames.push(name);
  }
  const findName = (predicate) => {
    const direct = allNames.find((name) => predicate(name.toLowerCase()));
    return direct || null;
  };
  const explicitInputIdsName = findName((lower) => lower.includes("input") && lower.includes("id"));
  const explicitAttentionName = findName((lower) => lower.includes("attention"));
  const explicitPositionName = findName((lower) => lower.includes("position"));
  const useNamesAreNumeric =
    useNames.length > 0 && useNames.every((name) => typeof name === "string" && /^\d+$/.test(name));
  if (numericOnly || useNamesAreNumeric) {
    const ordered = [inputIdsTensor, attentionMaskTensor, positionIdsTensor];
    useNames.forEach((name, index) => {
      const tensor = ordered[index];
      if (tensor) {
        inputs[name] = tensor;
      }
    });
    if (explicitInputIdsName && !inputs[explicitInputIdsName]) {
      inputs[explicitInputIdsName] = inputIdsTensor;
    }
    if (explicitAttentionName && !inputs[explicitAttentionName]) {
      inputs[explicitAttentionName] = attentionMaskTensor;
    }
    if (explicitPositionName && !inputs[explicitPositionName]) {
      inputs[explicitPositionName] = positionIdsTensor;
    }
    return inputs;
  }
  let hasInputIds = false;
  let hasAttention = false;
  let hasPosition = false;

  for (const [name, meta] of entries) {
    if (hasNonNumericNames && typeof name === "string" && /^\\d+$/.test(name)) {
      continue;
    }
    if (numericOnly && namesFromSession.length) {
      continue;
    }
    const lower = name.toLowerCase();
    if (lower.includes("input") && lower.includes("id")) {
      inputs[name] = inputIdsTensor;
      hasInputIds = true;
      continue;
    }
    if (attentionMaskTensor && lower.includes("attention")) {
      inputs[name] = attentionMaskTensor;
      hasAttention = true;
      continue;
    }
    if (positionIdsTensor && lower.includes("position")) {
      inputs[name] = positionIdsTensor;
      hasPosition = true;
      continue;
    }
    const dims = resolveDimensions(name, meta);
    const total = dims.reduce((acc, value) => acc * value, 1);
    const data = toOrtTypedArray(meta.type, total);
    fillDefaultValues(meta.type, data);
    inputs[name] = new window.ort.Tensor(meta.type || "float32", data, dims);
  }

  if (!hasInputIds && useNames.includes("input_ids")) {
    inputs.input_ids = inputIdsTensor;
  }
  if (!hasAttention && useNames.includes("attention_mask")) {
    inputs.attention_mask = attentionMaskTensor;
  }
  if (!hasPosition && useNames.includes("position_ids")) {
    inputs.position_ids = positionIdsTensor;
  }

  if (inputs.input_ids) {
    for (const key of Object.keys(inputs)) {
      if (/^\d+$/.test(key)) {
        delete inputs[key];
      }
    }
  }

  if (Array.isArray(inputNames) && inputNames.length) {
    const hasNonNumericNames = inputNames.some((name) => typeof name === "string" && !/^\d+$/.test(name));
    if (hasNonNumericNames) {
      return Object.fromEntries(Object.entries(inputs).filter(([key]) => inputNames.includes(key)));
    }
  }
  return inputs;
}

function initLiquidCache(decoderSession, config) {
  const cache = {};
  const textConfig = config?.text_config || {};
  const hiddenSize = textConfig.hidden_size || textConfig.block_dim || 1024;
  const numHeads = textConfig.num_attention_heads || textConfig.num_heads || 8;
  const numKvHeads = textConfig.num_key_value_heads || numHeads;
  const headDim = Math.floor(hiddenSize / numHeads);
  const convCacheLen = textConfig.conv_L_cache || 3;

  const inputNames = Array.isArray(decoderSession?.inputNames) ? decoderSession.inputNames : [];
  const { entries, useNames } = resolveInputNames(
    decoderSession?.inputMetadata,
    decoderSession?.inputNames,
    inputNames
  );
  const names = useNames.length ? useNames : entries.map(([name]) => name);

  for (const name of names) {
    const lower = name.toLowerCase();
    if (lower.includes("past_conv")) {
      const data = new Float32Array(hiddenSize * convCacheLen);
      cache[name] = new window.ort.Tensor("float32", data, [1, hiddenSize, convCacheLen]);
      continue;
    }
    if (lower.includes("past_key_values")) {
      const data = new Float32Array(0);
      cache[name] = new window.ort.Tensor("float32", data, [1, numKvHeads, 0, headDim]);
    }
  }
  return cache;
}

function updateLiquidCache(cache, outputs) {
  for (const [name, value] of Object.entries(outputs || {})) {
    if (name.startsWith("present_conv")) {
      cache[name.replace("present_conv", "past_conv")] = value;
      continue;
    }
    if (name.startsWith("present.")) {
      cache[name.replace("present.", "past_key_values.")] = value;
    }
  }
}

function buildLiquidDecoderInputs(decoderSession, currentEmbeds, attnTensor, posTensor, cache) {
  const inputs = {};
  const { entries, namesFromSession, numericOnly, useNames } = resolveInputNames(
    decoderSession.inputMetadata,
    decoderSession.inputNames,
    ["inputs_embeds", "attention_mask", "position_ids"]
  );
  const hasNonNumericNames =
    namesFromSession.some((name) => typeof name === "string" && !/^\d+$/.test(name)) ||
    useNames.some((name) => typeof name === "string" && !/^\d+$/.test(name));
  const useNamesAreNumeric =
    useNames.length > 0 && useNames.every((name) => typeof name === "string" && /^\d+$/.test(name));
  const filteredUseNames = numericOnly
    ? useNames
    : useNames.filter((name) => !(typeof name === "string" && /^\d+$/.test(name)));
  const effectiveNames = filteredUseNames.length ? filteredUseNames : useNames;
  const metaMap = new Map(entries.map(([name, meta]) => [name, meta]));

  if (numericOnly || useNamesAreNumeric) {
    const ordered = [currentEmbeds, attnTensor, posTensor];
    effectiveNames.forEach((name, index) => {
      const tensor = ordered[index];
      if (tensor) {
        inputs[name] = tensor;
      }
    });
    for (const [cacheName, cacheValue] of Object.entries(cache || {})) {
      if (!inputs[cacheName]) {
        inputs[cacheName] = cacheValue;
      }
    }
    return inputs;
  }

  for (const name of effectiveNames) {
    if (hasNonNumericNames && typeof name === "string" && /^\d+$/.test(name)) {
      continue;
    }
    const lower = name.toLowerCase();
    if (lower.includes("inputs_embeds")) {
      inputs[name] = currentEmbeds;
      continue;
    }
    if (attnTensor && lower.includes("attention")) {
      inputs[name] = attnTensor;
      continue;
    }
    if (posTensor && lower.includes("position")) {
      inputs[name] = posTensor;
      continue;
    }
    if (cache[name]) {
      inputs[name] = cache[name];
      continue;
    }
    if (numericOnly) {
      continue;
    }
    const meta = metaMap.get(name);
    if (!meta) continue;
    const dims = resolveDimensions(name, meta);
    const total = dims.reduce((acc, value) => acc * value, 1);
    const data = toOrtTypedArray(meta.type, total);
    fillDefaultValues(meta.type, data);
    inputs[name] = new window.ort.Tensor(meta.type || "float32", data, dims);
  }

  if (inputs.inputs_embeds) {
    for (const key of Object.keys(inputs)) {
      if (/^\d+$/.test(key)) {
        delete inputs[key];
      }
    }
  }

  if (Array.isArray(decoderSession.inputNames) && decoderSession.inputNames.length) {
    const hasNonNumericNames = decoderSession.inputNames.some(
      (name) => typeof name === "string" && !/^\d+$/.test(name)
    );
    if (hasNonNumericNames) {
      return Object.fromEntries(
        Object.entries(inputs).filter(([key]) => decoderSession.inputNames.includes(key))
      );
    }
  }
  return inputs;
}

function isMarkerOnly(text) {
  if (!text || typeof text !== "string") return true;
  return !text.replace(/>+/g, "").trim();
}

function getAsrChunkConfig(modelId) {
  if (/whisper-base/i.test(modelId)) {
    return { chunkLength: 15, strideLength: 5 };
  }
  return { chunkLength: 30, strideLength: 5 };
}

function getAsrStreamChunkConfig(modelId) {
  if (/whisper-base/i.test(modelId)) {
    return { chunkLength: 8, strideLength: 2 };
  }
  return { chunkLength: 10, strideLength: 2 };
}

function resetStreamBuffer() {
  state.streamSamples = [];
  state.streamSamplesLen = 0;
}

function appendStreamSamples(chunk, maxSamples) {
  if (!chunk?.length) return;
  state.streamSamples.push(chunk);
  state.streamSamplesLen += chunk.length;
  while (state.streamSamplesLen > maxSamples && state.streamSamples.length) {
    const head = state.streamSamples[0];
    if (state.streamSamplesLen - head.length >= maxSamples) {
      state.streamSamples.shift();
      state.streamSamplesLen -= head.length;
    } else {
      const trim = state.streamSamplesLen - maxSamples;
      state.streamSamples[0] = head.subarray(trim);
      state.streamSamplesLen -= trim;
      break;
    }
  }
}

function getStreamTail(samplesNeeded) {
  if (state.streamSamplesLen < samplesNeeded) return null;
  const output = new Float32Array(samplesNeeded);
  let offset = samplesNeeded;
  for (let i = state.streamSamples.length - 1; i >= 0 && offset > 0; i -= 1) {
    const chunk = state.streamSamples[i];
    const take = Math.min(chunk.length, offset);
    output.set(chunk.subarray(chunk.length - take), offset - take);
    offset -= take;
  }
  return output;
}

function downsampleBuffer(buffer, inputRate, targetRate) {
  if (inputRate === targetRate) return buffer;
  if (targetRate > inputRate) return buffer;
  const ratio = inputRate / targetRate;
  const newLength = Math.round(buffer.length / ratio);
  const result = new Float32Array(newLength);
  let offset = 0;
  for (let i = 0; i < newLength; i += 1) {
    const nextOffset = Math.round((i + 1) * ratio);
    let sum = 0;
    let count = 0;
    for (let j = offset; j < nextOffset && j < buffer.length; j += 1) {
      sum += buffer[j];
      count += 1;
    }
    result[i] = count ? sum / count : 0;
    offset = nextOffset;
  }
  return result;
}

async function withTimeout(promise, ms, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms);
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    clearTimeout(timer);
  }
}

async function runAsrOnce(audio, sampling_rate, modelId, { device, chunkLength, strideLength, label, dtypeOverride }) {
  let dtype = dtypeOverride;
  if (dtype) {
    await ensureAsrModelFiles(modelId, dtype);
  } else {
    dtype = await pickAvailableAsrDtypeForDevice(modelId, device);
  }
  const loadStart = performance.now();
  const asr = await getPipeline("automatic-speech-recognition", modelId, { dtype, device });
  log(
    `ASR config${label ? ` (${label})` : ""}: model=${modelId} dtype=${dtype} device=${device} ` +
      `chunk=${chunkLength}s stride=${strideLength}s`
  );
  logPerf("ASR pipeline load", performance.now() - loadStart);
  const options = { chunk_length_s: chunkLength, stride_length_s: strideLength };
  if (ASR_LANGUAGE && ASR_LANGUAGE !== "auto") {
    options.language = ASR_LANGUAGE;
  }
  options.task = "transcribe";
  const inferStart = performance.now();
  const result = await withTimeout(asr(audio, options), ASR_TIMEOUT_MS, "ASR inference");
  logPerf("ASR inference", performance.now() - inferStart);
  if (result && typeof result === "object") {
    const keys = Object.keys(result);
    log(`ASR result keys: ${keys.join(", ") || "(none)"}`);
    if (typeof result.text === "string") {
      log(`ASR result text preview: ${result.text.slice(0, 200)}`);
    }
  }
  const text = typeof result === "string" ? result : result?.text || JSON.stringify(result, null, 2);
  return { result, text };
}

async function runAsrFromBlob(blob) {
  const modelId = asrModelEl.value.trim();
  if (!modelId) {
    setStatus("Pick an ASR model.");
    return;
  }
  const totalStart = performance.now();
  setStatus(`Loading ASR pipeline (${modelId}) ...`);
  try {
    const dtypeOverride = getAsrDtypeOverride();
    setStatus("Decoding audio ...");
    const decodeStart = performance.now();
    const { audio, sampling_rate } = await decodeAudioBlob(blob);
    logPerf("ASR audio decode", performance.now() - decodeStart);
    logAudioStats(audio, sampling_rate);
    setStatus("Transcribing ...");
    const primaryDevice = getDeviceOption();
    const { chunkLength, strideLength } = getAsrChunkConfig(modelId);
    let text = "";
    try {
      const primary = await runAsrOnce(audio, sampling_rate, modelId, {
        device: primaryDevice,
        chunkLength,
        strideLength,
        label: "primary",
        dtypeOverride,
      });
      text = primary.text;
    } catch (err) {
      if (primaryDevice === "webgpu" && /timed out/i.test(err.message)) {
        log("ASR warning: WebGPU inference timed out. Retrying with WASM.");
      } else {
        throw err;
      }
    }
    if ((!text || isMarkerOnly(text)) && primaryDevice === "webgpu") {
      if (text) {
        log("ASR warning: output looks empty or contains only special markers (e.g. >>>). Retrying with WASM.");
      }
      const retry = await runAsrOnce(audio, sampling_rate, modelId, {
        device: "wasm",
        chunkLength: Math.max(10, chunkLength - 1),
        strideLength,
        label: "retry-wasm",
        dtypeOverride,
      });
      if (!isMarkerOnly(retry.text)) {
        text = retry.text;
      } else if (!text) {
        text = retry.text;
      }
    }
    if (isMarkerOnly(text)) {
      log("ASR warning: output looks empty or contains only special markers (e.g. >>>).");
    }
    transcriptEl.value = text;
    log(`ASR transcript:\\n${text}`);
    if (structuredEl) {
      structuredEl.value = "";
    }
    logPerf("ASR total", performance.now() - totalStart);
    if (isMarkerOnly(text)) {
      setStatus("Transcription empty; skipping cleanup.");
      return;
    }
    setStatus("Transcription complete. Cleaning up ...");
    await runCleanup(text);
  } catch (err) {
    setStatus(`ASR failed: ${err.message}`);
  }
}

async function runAsrStreamAudio(audio, sampling_rate) {
  const modelId = asrModelEl.value.trim();
  if (!modelId) return;
  const primaryDevice = getDeviceOption();
  const dtypeOverride = getAsrDtypeOverride();
  const { chunkLength, strideLength } = getAsrStreamChunkConfig(modelId);
  let text = "";
  try {
    const primary = await runAsrOnce(audio, sampling_rate, modelId, {
      device: primaryDevice,
      chunkLength,
      strideLength,
      label: `stream-${state.streamSegment}`,
      dtypeOverride,
    });
    text = primary.text;
  } catch (err) {
    if (primaryDevice === "webgpu" && /timed out/i.test(err.message)) {
      log("ASR warning: streaming WebGPU timed out. Retrying with WASM.");
    } else {
      log(`ASR stream error: ${err.message}`);
      return;
    }
  }
  if ((!text || isMarkerOnly(text)) && primaryDevice === "webgpu") {
    const retry = await runAsrOnce(audio, sampling_rate, modelId, {
      device: "wasm",
      chunkLength: Math.max(5, chunkLength - 1),
      strideLength,
      label: `stream-${state.streamSegment}-wasm`,
      dtypeOverride,
    });
    text = retry.text;
  }
  if (isMarkerOnly(text)) return;
  const cleaned = text.replace(/\s+/g, " ").trim();
  if (!cleaned) return;
  state.streamTranscript = `${state.streamTranscript} ${cleaned}`.trim();
  transcriptEl.value = state.streamTranscript;
}

function queueStreamAudio(audio, sampleRate) {
  state.streamSegment += 1;
  state.streamQueue = state.streamQueue
    .then(() => runAsrStreamAudio(audio, sampleRate))
    .catch((err) => log(`ASR stream queue error: ${err.message}`));
}

function computeCleanupMaxTokens(inputLength) {
  const scaled = Math.ceil(inputLength * 1.3);
  return Math.min(CLEANUP_MAX_NEW_TOKENS_CAP, Math.max(CLEANUP_MIN_NEW_TOKENS, scaled));
}

function computeStructureMaxTokens(inputLength) {
  const scaled = Math.ceil(inputLength * 1.5);
  return Math.min(STRUCTURE_MAX_NEW_TOKENS_CAP, Math.max(STRUCTURE_MIN_NEW_TOKENS, scaled));
}

async function runCleanup(text) {
  const cleanedTarget = cleanedEl;
  if (!cleanedTarget) return;
  const trimmed = (text || "").trim();
  if (!trimmed) {
    cleanedTarget.value = "";
    setStatus("Transcription complete.");
    return;
  }
  if (state.isCleaning) return;
  state.isCleaning = true;
  const start = performance.now();
  try {
    log(`Cleanup input:\\n${trimmed}`);
    const cleaned = await runLiquidCleanup(trimmed);
    cleanedTarget.value = cleaned;
    log(`Cleanup output:\\n${cleaned}`);
    logPerf("Cleanup total", performance.now() - start);
    if (structuredEl) {
      await runStructure(cleaned || trimmed);
    } else {
      setStatus("Cleanup complete.");
    }
  } catch (err) {
    setStatus(`Cleanup failed: ${err.message}`);
  } finally {
    state.isCleaning = false;
  }
}

async function runStructure(text) {
  if (!structuredEl) {
    setStatus("Cleanup complete.");
    return;
  }
  const trimmed = (text || "").trim();
  if (!trimmed) {
    structuredEl.value = "";
    setStatus("Cleanup complete.");
    return;
  }
  if (state.isStructuring) return;
  state.isStructuring = true;
  const start = performance.now();
  try {
    log(`Structure input:\n${trimmed}`);
    setStatus("Structuring into 10 chapters ...");
    const structured = await runLiquidStructure(trimmed);
    structuredEl.value = structured;
    log(`Structure output:\n${structured}`);
    logPerf("Structure total", performance.now() - start);
    setStatus("Structuring complete.");
  } catch (err) {
    setStatus(`Structuring failed: ${err.message}`);
  } finally {
    state.isStructuring = false;
  }
}

async function runLiquidCleanup(text) {
  await ensureLibraryLoaded();
  await ensureOrtLoaded();
  await ensureWebGpuFeatures();
  const modelId = CLEANUP_MODEL;
  const base = getLocalModelBase(modelId);
  const providers =
    state.webgpu.supportsFp16 === false ? ORT_PROVIDERS.wasm : resolveOrtProvider();

  const tokenizer = await getTokenizer(modelId);
  const config = await getModelConfig(modelId);
  const eosTokenId = config?.text_config?.eos_token_id ?? config?.eos_token_id ?? null;

  const promptStart = performance.now();
  const prompt = await buildCleanupPrompt(tokenizer, modelId, text);
  log(`Cleanup prompt:\n${prompt}`);
  const inputIds = (await tokenizeQuestion(tokenizer, prompt)).map((value) => Number(value));
  logPerf("Cleanup prompt encode", performance.now() - promptStart);

  const attentionMask = buildAttentionMaskFromLength(inputIds.length);
  const positionIds = buildPositionIdsFromLength(inputIds.length);

  const embedSession = await loadOrtSession(`${base}onnx/embed_tokens_fp16.onnx`, providers);
  const decoderSession = await loadOrtSession(`${base}onnx/decoder_q4.onnx`, providers);

  const embedStart = performance.now();
  const embedMetaNames =
    embedSession.inputMetadata instanceof Map
      ? Array.from(embedSession.inputMetadata.keys())
      : Object.keys(embedSession.inputMetadata || {});
  const embedInputNames = Array.isArray(embedSession.inputNames) ? embedSession.inputNames : [];
  log(`Cleanup embed meta inputs: ${embedMetaNames.join(", ") || "(none)"}`);
  if (embedInputNames.length) {
    log(`Cleanup embed inputNames: ${embedInputNames.join(", ")}`);
  }
  const tokenInputs = buildLiquidTokenInputs(
    embedSession.inputMetadata,
    embedSession.inputNames,
    inputIds,
    attentionMask,
    positionIds
  );
  log(`Cleanup embed feeds: ${Object.keys(tokenInputs).join(", ") || "(none)"}`);
  const tokenResult = await embedSession.run(tokenInputs);
  let currentEmbeds = Object.values(tokenResult)[0];
  logPerf("Cleanup token embedding", performance.now() - embedStart);

  const maxNewTokens = computeCleanupMaxTokens(inputIds.length);
  const generated = [];
  const cache = initLiquidCache(decoderSession, config);
  const generateStart = performance.now();

  for (let step = 0; step < maxNewTokens; step += 1) {
    const allIds = [...inputIds, ...generated];
    const attnMask = buildAttentionMaskFromLength(allIds.length);
    const posIds = buildPositionIdsFromLength(allIds.length);
    const attnTensor = new window.ort.Tensor("int64", attnMask, [1, attnMask.length]);
    const posTensor = new window.ort.Tensor("int64", posIds, [1, posIds.length]);

    const decoderInputs = buildLiquidDecoderInputs(decoderSession, currentEmbeds, attnTensor, posTensor, cache);
    if (step === 0) {
      const decoderMetaNames =
        decoderSession.inputMetadata instanceof Map
          ? Array.from(decoderSession.inputMetadata.keys())
          : Object.keys(decoderSession.inputMetadata || {});
      const decoderInputNames = Array.isArray(decoderSession.inputNames) ? decoderSession.inputNames : [];
      log(`Cleanup decoder meta inputs: ${decoderMetaNames.join(", ") || "(none)"}`);
      if (decoderInputNames.length) {
        log(`Cleanup decoder inputNames: ${decoderInputNames.join(", ")}`);
      }
      log(`Cleanup decoder feeds: ${Object.keys(decoderInputs).join(", ") || "(none)"}`);
    }
    const result = await decoderSession.run(decoderInputs);
    updateLiquidCache(cache, result);
    const logitsTensor = resolveLogitsOutput(result);
    if (!logitsTensor) throw new Error("Decoder did not return logits.");
    const nextId = argmaxLogits(logitsTensor);
    generated.push(nextId);
    if (eosTokenId !== null && nextId === eosTokenId) break;

    const decodedSoFar = decodeTokens(tokenizer, generated);
    if (decodedSoFar.length >= 120 && /[.!?]\\s*$/.test(decodedSoFar)) {
      break;
    }

    const nextTokenInputs = buildLiquidTokenInputs(
      embedSession.inputMetadata,
      embedSession.inputNames,
      [nextId]
    );
    const nextTokenResult = await embedSession.run(nextTokenInputs);
    currentEmbeds = Object.values(nextTokenResult)[0];
  }

  logPerf("Cleanup generate", performance.now() - generateStart);
  return decodeTokens(tokenizer, generated).trim();
}

async function runLiquidStructure(text) {
  await ensureLibraryLoaded();
  await ensureOrtLoaded();
  await ensureWebGpuFeatures();
  const modelId = CLEANUP_MODEL;
  const base = getLocalModelBase(modelId);
  const providers =
    state.webgpu.supportsFp16 === false ? ORT_PROVIDERS.wasm : resolveOrtProvider();

  const tokenizer = await getTokenizer(modelId);
  const config = await getModelConfig(modelId);
  const eosTokenId = config?.text_config?.eos_token_id ?? config?.eos_token_id ?? null;

  const promptStart = performance.now();
  const prompt = await buildStructurePrompt(tokenizer, modelId, text);
  log(`Structure prompt:\n${prompt}`);
  const inputIds = (await tokenizeQuestion(tokenizer, prompt)).map((value) => Number(value));
  logPerf("Structure prompt encode", performance.now() - promptStart);

  const attentionMask = buildAttentionMaskFromLength(inputIds.length);
  const positionIds = buildPositionIdsFromLength(inputIds.length);

  const embedSession = await loadOrtSession(`${base}onnx/embed_tokens_fp16.onnx`, providers);
  const decoderSession = await loadOrtSession(`${base}onnx/decoder_q4.onnx`, providers);

  const embedStart = performance.now();
  const embedMetaNames =
    embedSession.inputMetadata instanceof Map
      ? Array.from(embedSession.inputMetadata.keys())
      : Object.keys(embedSession.inputMetadata || {});
  const embedInputNames = Array.isArray(embedSession.inputNames) ? embedSession.inputNames : [];
  log(`Structure embed meta inputs: ${embedMetaNames.join(", ") || "(none)"}`);
  if (embedInputNames.length) {
    log(`Structure embed inputNames: ${embedInputNames.join(", ")}`);
  }
  const tokenInputs = buildLiquidTokenInputs(
    embedSession.inputMetadata,
    embedSession.inputNames,
    inputIds,
    attentionMask,
    positionIds
  );
  log(`Structure embed feeds: ${Object.keys(tokenInputs).join(", ") || "(none)"}`);
  const tokenResult = await embedSession.run(tokenInputs);
  let currentEmbeds = Object.values(tokenResult)[0];
  logPerf("Structure token embedding", performance.now() - embedStart);

  const maxNewTokens = computeStructureMaxTokens(inputIds.length);
  const generated = [];
  const cache = initLiquidCache(decoderSession, config);
  const generateStart = performance.now();

  for (let step = 0; step < maxNewTokens; step += 1) {
    const allIds = [...inputIds, ...generated];
    const attnMask = buildAttentionMaskFromLength(allIds.length);
    const posIds = buildPositionIdsFromLength(allIds.length);
    const attnTensor = new window.ort.Tensor("int64", attnMask, [1, attnMask.length]);
    const posTensor = new window.ort.Tensor("int64", posIds, [1, posIds.length]);

    const decoderInputs = buildLiquidDecoderInputs(decoderSession, currentEmbeds, attnTensor, posTensor, cache);
    if (step === 0) {
      const decoderMetaNames =
        decoderSession.inputMetadata instanceof Map
          ? Array.from(decoderSession.inputMetadata.keys())
          : Object.keys(decoderSession.inputMetadata || {});
      const decoderInputNames = Array.isArray(decoderSession.inputNames) ? decoderSession.inputNames : [];
      log(`Structure decoder meta inputs: ${decoderMetaNames.join(", ") || "(none)"}`);
      if (decoderInputNames.length) {
        log(`Structure decoder inputNames: ${decoderInputNames.join(", ")}`);
      }
      log(`Structure decoder feeds: ${Object.keys(decoderInputs).join(", ") || "(none)"}`);
    }
    const result = await decoderSession.run(decoderInputs);
    updateLiquidCache(cache, result);
    const logitsTensor = resolveLogitsOutput(result);
    if (!logitsTensor) throw new Error("Decoder did not return logits.");
    const nextId = argmaxLogits(logitsTensor);
    generated.push(nextId);
    if (eosTokenId !== null && nextId === eosTokenId) break;

    const nextTokenInputs = buildLiquidTokenInputs(
      embedSession.inputMetadata,
      embedSession.inputNames,
      [nextId]
    );
    const nextTokenResult = await embedSession.run(nextTokenInputs);
    currentEmbeds = Object.values(nextTokenResult)[0];
  }

  logPerf("Structure generate", performance.now() - generateStart);
  return decodeTokens(tokenizer, generated).trim();
}

async function runLiquidImagePrompt({
  modelId,
  file,
  img,
  question,
  maxNewTokens = 64,
  statusEl,
}) {
  await ensureOrtLoaded();
  const config = resolveOrtModelConfig(modelId);
  if (!config) {
    throw new Error(`Unknown LiquidAI model config for ${modelId}`);
  }
  if (!file && !img) {
    throw new Error("Pick an image file first.");
  }
  if (!window.ort) {
    throw new Error("onnxruntime-web not loaded.");
  }
  if (!state.lib) {
    throw new Error("Load Transformers.js first (needed for tokenizer).");
  }

  const base = getLocalModelBase(modelId);
  const providers = await resolveOrtProviderForModel(modelId, config);
  const resolvedImg = img || (await loadImageFromFile(file));

  if (statusEl) statusEl.textContent = "Loading tokenizer + processor ...";
  const tokenizer = await getTokenizer(modelId);
  const processor = await getProcessor(modelId);
  const configJson = await getModelConfig(modelId);
  const eosTokenId = configJson?.text_config?.eos_token_id ?? configJson?.eos_token_id ?? null;
  const imageTokenId =
    configJson?.image_token_id ??
    configJson?.image_token_index ??
    (typeof tokenizer.convert_tokens_to_ids === "function" ? tokenizer.convert_tokens_to_ids("<image>") : null);

  if (statusEl) statusEl.textContent = "Encoding prompt + image ...";
  const prompt = await buildLiquidPrompt(tokenizer, modelId, question);
  const processed = await runLiquidProcessor(processor, resolvedImg, prompt);

  if (statusEl) statusEl.textContent = "Loading ONNX sessions ...";
  const embedTokenSession = await loadOrtSession(`${base}${config.embedTokens}`, providers);
  const embedImageSession = await loadOrtSession(`${base}${config.embedImages}`, providers);
  const decoderSession = await loadOrtSession(`${base}${config.decoder}`, providers);

  if (statusEl) statusEl.textContent = "Embedding image + tokens ...";
  const imageInputs = buildLiquidImageInputs(
    embedImageSession.inputMetadata,
    embedImageSession.inputNames,
    processed
  );
  const imageMetaNames =
    embedImageSession.inputMetadata instanceof Map
      ? Array.from(embedImageSession.inputMetadata.keys())
      : Object.keys(embedImageSession.inputMetadata || {});
  const imageInputNames = Array.isArray(embedImageSession.inputNames) ? embedImageSession.inputNames : [];
  log(`LiquidAI image meta inputs: ${imageMetaNames.join(", ") || "(none)"}`);
  if (imageInputNames.length) {
    log(`LiquidAI image inputNames: ${imageInputNames.join(", ")}`);
  }
  log(`LiquidAI image feeds: ${Object.keys(imageInputs).join(", ") || "(none)"}`);
  const imageResult = await embedImageSession.run(imageInputs);
  const imageEmbedding = Object.values(imageResult)[0];
  const numImageTokens = imageEmbedding?.dims?.length === 2 ? imageEmbedding.dims[0] : imageEmbedding?.dims?.[1];
  const tokensPerImage = numImageTokens ? [numImageTokens] : [];

  const imageStartTokenId = resolveTokenId(tokenizer, "<|image_start|>");
  const imageEndTokenId = resolveTokenId(tokenizer, "<|image_end|>");
  const promptIdsRaw = (await tokenizeQuestion(tokenizer, prompt)).map((value) => Number(value));
  const promptIds = expandImageTokens(
    promptIdsRaw,
    tokensPerImage,
    imageTokenId,
    imageStartTokenId,
    imageEndTokenId
  );

  const attentionMask = buildAttentionMaskFromLength(promptIds.length);
  const positionIds = buildPositionIdsFromLength(promptIds.length);
  const tokenInputs = buildLiquidTokenInputs(
    embedTokenSession.inputMetadata,
    embedTokenSession.inputNames,
    promptIds,
    attentionMask,
    positionIds
  );
  const tokenMetaNames =
    embedTokenSession.inputMetadata instanceof Map
      ? Array.from(embedTokenSession.inputMetadata.keys())
      : Object.keys(embedTokenSession.inputMetadata || {});
  const tokenInputNames = Array.isArray(embedTokenSession.inputNames) ? embedTokenSession.inputNames : [];
  log(`LiquidAI token meta inputs: ${tokenMetaNames.join(", ") || "(none)"}`);
  if (tokenInputNames.length) {
    log(`LiquidAI token inputNames: ${tokenInputNames.join(", ")}`);
  }
  log(`LiquidAI token feeds: ${Object.keys(tokenInputs).join(", ") || "(none)"}`);
  const tokenResult = await embedTokenSession.run(tokenInputs);
  const tokenEmbedding = Object.values(tokenResult)[0];

  if (imageTokenId !== null) {
    mergeLiquidEmbeds(tokenEmbedding, imageEmbedding, promptIds, imageTokenId);
  } else {
    log("LiquidAI: image token id missing; image embeddings not merged.");
  }

  const generated = [];
  let cache = initLiquidCache(decoderSession, configJson);
  let currentEmbeds = tokenEmbedding;

  if (statusEl) statusEl.textContent = "Generating (greedy) ...";
  const start = performance.now();
  for (let step = 0; step < maxNewTokens; step += 1) {
    const allIds = [...promptIds, ...generated];
    const attnMask = buildAttentionMaskFromLength(allIds.length);
    const posIds = buildPositionIdsFromLength(allIds.length);
    const attnTensor = new window.ort.Tensor("int64", attnMask, [1, attnMask.length]);
    const posTensor = new window.ort.Tensor("int64", posIds, [1, posIds.length]);

    const decoderInputs = buildLiquidDecoderInputs(decoderSession, currentEmbeds, attnTensor, posTensor, cache);
    if (step === 0) {
      const decoderMetaNames =
        decoderSession.inputMetadata instanceof Map
          ? Array.from(decoderSession.inputMetadata.keys())
          : Object.keys(decoderSession.inputMetadata || {});
      const decoderInputNames = Array.isArray(decoderSession.inputNames) ? decoderSession.inputNames : [];
      log(`LiquidAI decoder meta inputs: ${decoderMetaNames.join(", ") || "(none)"}`);
      if (decoderInputNames.length) {
        log(`LiquidAI decoder inputNames: ${decoderInputNames.join(", ")}`);
      }
      log(`LiquidAI decoder feeds: ${Object.keys(decoderInputs).join(", ") || "(none)"}`);
    }

    const result = await decoderSession.run(decoderInputs);
    updateLiquidCache(cache, result);
    const logitsTensor = resolveLogitsOutput(result);
    if (!logitsTensor) throw new Error("Decoder did not return logits.");
    const nextId = argmaxLogits(logitsTensor);
    generated.push(nextId);
    if (eosTokenId !== null && nextId === eosTokenId) break;

    const nextTokenInputs = buildLiquidTokenInputs(
      embedTokenSession.inputMetadata,
      embedTokenSession.inputNames,
      [nextId]
    );
    const nextTokenResult = await embedTokenSession.run(nextTokenInputs);
    currentEmbeds = Object.values(nextTokenResult)[0];
  }
  const elapsed = (performance.now() - start).toFixed(0);
  const decoded = decodeTokens(tokenizer, generated);

  return {
    modelId,
    providers,
    question,
    prompt,
    generatedTokenIds: generated,
    answer: decoded,
    elapsedMs: elapsed,
    image: resolvedImg,
  };
}

function setNotesStatus(message) {
  if (notesStatusEl) notesStatusEl.textContent = message;
  log(`Notes: ${message}`);
}

function clearNotesPreview() {
  if (!state.notes.previewUrl) return;
  URL.revokeObjectURL(state.notes.previewUrl);
  state.notes.previewUrl = "";
}

function setNotesPreviewFromBlob(blob) {
  if (!notesPreviewEl || !blob) return;
  clearNotesPreview();
  const url = URL.createObjectURL(blob);
  state.notes.previewUrl = url;
  notesPreviewEl.src = url;
}

function setNotesPreviewFromImage(img) {
  if (!notesPreviewEl || !img?.src) return;
  clearNotesPreview();
  notesPreviewEl.src = img.src;
}

function updateNotesCameraButtons({ hasStream }) {
  if (notesStartCameraBtn) notesStartCameraBtn.disabled = hasStream;
  if (notesCaptureBtn) notesCaptureBtn.disabled = !hasStream;
  if (notesStopCameraBtn) notesStopCameraBtn.disabled = !hasStream;
}

async function startNotesCamera() {
  if (!notesVideoEl) return;
  if (!navigator.mediaDevices?.getUserMedia) {
    setNotesStatus("Camera not available in this browser.");
    return;
  }
  try {
    setNotesStatus("Requesting camera access ...");
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: "environment" },
      audio: false,
    });
    state.notes.stream = stream;
    notesVideoEl.srcObject = stream;
    await notesVideoEl.play();
    updateNotesCameraButtons({ hasStream: true });
    setNotesStatus("Camera ready. Capture a page.");
  } catch (err) {
    setNotesStatus(`Camera error: ${err.message}`);
  }
}

function stopNotesCamera() {
  if (!notesVideoEl) return;
  if (state.notes.stream) {
    state.notes.stream.getTracks().forEach((track) => track.stop());
    state.notes.stream = null;
  }
  notesVideoEl.srcObject = null;
  updateNotesCameraButtons({ hasStream: false });
  setNotesStatus("Camera stopped.");
}

async function captureNotesPhoto() {
  if (!notesVideoEl) return;
  const width = notesVideoEl.videoWidth;
  const height = notesVideoEl.videoHeight;
  if (!width || !height) {
    setNotesStatus("Camera not ready yet.");
    return;
  }
  const maxDim = 1600;
  const scale = Math.min(1, maxDim / Math.max(width, height));
  const targetWidth = Math.max(1, Math.round(width * scale));
  const targetHeight = Math.max(1, Math.round(height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = targetWidth;
  canvas.height = targetHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    setNotesStatus("Canvas not available.");
    return;
  }
  ctx.drawImage(notesVideoEl, 0, 0, targetWidth, targetHeight);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.92));
  if (!blob) {
    setNotesStatus("Capture failed.");
    return;
  }
  state.notes.captureBlob = blob;
  if (notesFileEl) {
    notesFileEl.value = "";
  }
  setNotesPreviewFromBlob(blob);
  setNotesStatus("Photo captured.");
}

function getNotesImageBlob() {
  if (state.notes.captureBlob) return state.notes.captureBlob;
  if (notesFileEl?.files?.length) return notesFileEl.files[0];
  return null;
}

async function runNotesScan() {
  if (!notesTranscriptEl || !notesStructuredEl) return;
  const modelId = CLEANUP_MODEL;
  const file = getNotesImageBlob();
  if (!file) {
    setNotesStatus("Capture or choose an image first.");
    return;
  }
  if (state.notes.isRunning) return;
  state.notes.isRunning = true;
  if (notesRunBtn) notesRunBtn.disabled = true;
  notesTranscriptEl.value = "";
  notesStructuredEl.value = "";
  try {
    await ensureLibraryLoaded();
    setNotesStatus("Preparing image ...");
    const img = await loadImageFromFile(file);
    setNotesPreviewFromImage(img);
    log(
      `Notes scan input: ${file.name || "camera-capture"} size=${file.size || 0} type=${file.type || "image/jpeg"}`
    );

    setNotesStatus("Transcribing note ...");
    log(`Notes transcription question:\n${NOTES_TRANSCRIBE_PROMPT}`);
    const transcription = await runLiquidImagePrompt({
      modelId,
      img,
      question: NOTES_TRANSCRIBE_PROMPT,
      maxNewTokens: NOTES_TRANSCRIBE_MAX_TOKENS,
      statusEl: notesStatusEl,
    });
    const rawTranscript = transcription.answer || "";
    const cleanedTranscript = stripModelTokens(rawTranscript);
    notesTranscriptEl.value = cleanedTranscript;
    log(`Notes transcription prompt:\n${transcription.prompt}`);
    log(`Notes transcription output (raw):\n${rawTranscript}`);
    log(`Notes transcription output (clean):\n${cleanedTranscript}`);

    if (!cleanedTranscript.trim()) {
      setNotesStatus("Transcription empty; skipping structuring.");
      return;
    }

    setNotesStatus("Structuring into 10 chapters ...");
    log(`Notes structure input:\n${cleanedTranscript}`);
    const structured = await runLiquidStructure(cleanedTranscript);
    const structuredClean = stripModelTokens(structured);
    notesStructuredEl.value = structuredClean;
    log(`Notes structure output (raw):\n${structured}`);
    log(`Notes structure output (clean):\n${structuredClean}`);
    setNotesStatus("Notes scan complete.");
  } catch (err) {
    log(err?.stack || String(err));
    setNotesStatus(`Notes scan failed: ${err.message}`);
  } finally {
    state.notes.isRunning = false;
    if (notesRunBtn) notesRunBtn.disabled = false;
  }
}

function pickBestMimeType() {
  const candidates = ["audio/webm;codecs=opus", "audio/webm", "audio/ogg;codecs=opus", "audio/ogg"];
  for (const type of candidates) {
    if (MediaRecorder.isTypeSupported(type)) return type;
  }
  return "";
}

async function startRecording() {
  if (state.isRecording || state.isTranscribing || state.isCleaning || state.isStructuring) return;
  state.isRecording = true;
  holdBtn.classList.add("is-recording");
  try {
    setStatus("Requesting microphone ...");
    state.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const mimeType = pickBestMimeType();
    state.chunks = [];
    state.streamSegment = 0;
    state.streamTranscript = "";
    state.streamQueue = Promise.resolve();
    state.isStreaming = getAsrMode() === "stream";
    if (state.isStreaming) {
      transcriptEl.value = "";
      cleanedEl.value = "";
      if (structuredEl) {
        structuredEl.value = "";
      }
    }
    if (state.isStreaming) {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      const streamCtx = new AudioCtx({ sampleRate: 16000 });
      const source = streamCtx.createMediaStreamSource(state.stream);
      const processor = streamCtx.createScriptProcessor(4096, 1, 1);
      const gain = streamCtx.createGain();
      gain.gain.value = 0;
      state.streamSampleRate = streamCtx.sampleRate || 16000;
      resetStreamBuffer();
      processor.onaudioprocess = (event) => {
        if (!state.isRecording) return;
        const input = event.inputBuffer.getChannelData(0);
        const chunk = new Float32Array(input);
        const { chunkLength } = getAsrStreamChunkConfig(asrModelEl.value.trim());
        const maxSamples = Math.ceil(chunkLength * state.streamSampleRate);
        appendStreamSamples(chunk, maxSamples);
      };
      source.connect(processor);
      processor.connect(gain);
      gain.connect(streamCtx.destination);
      state.streamAudioCtx = streamCtx;
      state.streamSource = source;
      state.streamProcessor = processor;
      state.streamTimer = window.setInterval(() => {
        if (!state.isRecording) return;
        const modelId = asrModelEl.value.trim();
        if (!modelId) return;
        const { chunkLength } = getAsrStreamChunkConfig(modelId);
        const samplesNeeded = Math.ceil(chunkLength * state.streamSampleRate);
        const tail = getStreamTail(samplesNeeded);
        if (!tail) return;
        const audio = downsampleBuffer(tail, state.streamSampleRate, 16000);
        queueStreamAudio(audio, 16000);
      }, STREAM_TIMESLICE_MS);
      setStatus("Recording (streaming) ...");
      return;
    }

    state.recorder = new MediaRecorder(state.stream, mimeType ? { mimeType } : undefined);
    state.recorder.ondataavailable = (event) => {
      if (event.data && event.data.size > 0) state.chunks.push(event.data);
    };
    state.recorder.onstop = async () => {
      const blob = new Blob(state.chunks, { type: state.recorder.mimeType || "audio/webm" });
      state.chunks = [];
      state.stream?.getTracks().forEach((track) => track.stop());
      state.stream = null;
      state.recorder = null;
      state.isRecording = false;
      holdBtn.classList.remove("is-recording");
      holdBtn.disabled = false;
      state.isTranscribing = true;
      try {
        await runAsrFromBlob(blob);
      } finally {
        state.isTranscribing = false;
      }
    };
    state.recorder.start();
    setStatus("Recording ...");
  } catch (err) {
    state.isRecording = false;
    holdBtn.classList.remove("is-recording");
    setStatus(`Microphone error: ${err.message}`);
  }
}

async function stopRecording() {
  if (!state.isRecording) return;
  setStatus("Stopping recording ...");
  try {
    if (state.isStreaming) {
      state.isRecording = false;
      if (state.streamTimer) {
        clearInterval(state.streamTimer);
        state.streamTimer = null;
      }
      if (state.streamProcessor) {
        state.streamProcessor.disconnect();
        state.streamProcessor = null;
      }
      if (state.streamSource) {
        state.streamSource.disconnect();
        state.streamSource = null;
      }
      if (state.streamAudioCtx) {
        await state.streamAudioCtx.close();
        state.streamAudioCtx = null;
      }
      state.stream?.getTracks().forEach((track) => track.stop());
      state.stream = null;
      holdBtn.classList.remove("is-recording");
      holdBtn.disabled = false;
      setStatus("Finishing streaming ...");
      try {
        await state.streamQueue;
      } finally {
        state.isStreaming = false;
      }
      setStatus("Streaming complete.");
      return;
    }
    if (!state.recorder) return;
    state.recorder.stop();
  } catch (err) {
    setStatus(`Stop failed: ${err.message}`);
  }
}

loadLibBtn.addEventListener("click", () => {
  loadLibrary();
});

if (copyLogBtn) {
  copyLogBtn.addEventListener("click", () => {
    copyLogToClipboard();
  });
}

if (notesStartCameraBtn) {
  notesStartCameraBtn.addEventListener("click", () => {
    startNotesCamera();
  });
}

if (notesStopCameraBtn) {
  notesStopCameraBtn.addEventListener("click", () => {
    stopNotesCamera();
  });
}

if (notesCaptureBtn) {
  notesCaptureBtn.addEventListener("click", () => {
    captureNotesPhoto();
  });
}

if (notesFileEl) {
  notesFileEl.addEventListener("change", () => {
    const file = notesFileEl.files?.[0];
    if (!file) return;
    state.notes.captureBlob = file;
    setNotesPreviewFromBlob(file);
    setNotesStatus(`Loaded ${file.name || "image"}.`);
  });
}

if (notesRunBtn) {
  notesRunBtn.addEventListener("click", () => {
    runNotesScan();
  });
}

holdBtn.addEventListener("pointerdown", (event) => {
  event.preventDefault();
  holdBtn.setPointerCapture(event.pointerId);
  startRecording();
});

holdBtn.addEventListener("pointerup", (event) => {
  event.preventDefault();
  holdBtn.releasePointerCapture(event.pointerId);
  stopRecording();
});

holdBtn.addEventListener("pointercancel", stopRecording);
holdBtn.addEventListener("pointerleave", stopRecording);

window.addEventListener("DOMContentLoaded", () => {
  setStatus("Idle.");
});
