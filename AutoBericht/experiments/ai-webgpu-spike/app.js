import { processImage as liquidProcessImage } from "./liquid-processor.js";

const byId = (id) => document.getElementById(id);

const transformersUrlEl = byId("transformers-url");
const localModelPathEl = byId("local-model-path");
const allowRemoteEl = byId("allow-remote");
const disableWasmSimdEl = byId("disable-wasm-simd");
const ortBundleEl = byId("ort-bundle");
const ortVersionEl = byId("ort-version");
const deviceSelectEl = byId("device-select");
const loadLibBtn = byId("load-lib");
const checkWebgpuBtn = byId("check-webgpu");
const probeWasmBtn = byId("probe-wasm");
const reportWebgpuBtn = byId("report-webgpu");
const envStatusEl = byId("env-status");

const asrModelEl = byId("asr-model");
const asrFileEl = byId("asr-file");
const asrTimestampsEl = byId("asr-timestamps");
const runAsrBtn = byId("run-asr");
const asrStatusEl = byId("asr-status");
const asrOutputEl = byId("asr-output");

const visionTaskEl = byId("vision-task");
const visionModelEl = byId("vision-model");
const visionQuestionEl = byId("vision-question");
const visionFileEl = byId("vision-file");
const runVisionBtn = byId("run-vision");
const loadVisionOnnxBtn = byId("load-vision-onnx");
const loadVisionOnnxWasmBtn = byId("load-vision-onnx-wasm");
const loadVisionOnnxWebgpuBtn = byId("load-vision-onnx-webgpu");
const askVisionBtn = byId("ask-vision");
const visionStatusEl = byId("vision-status");
const visionPreviewEl = byId("vision-preview");
const visionOutputEl = byId("vision-output");

const downloadLogBtn = byId("download-log");

const logEl = byId("log");

const state = {
  lib: null,
  pipeline: null,
  env: null,
  pipelines: new Map(),
  ortSessions: new Map(),
  externalDataCache: new Map(),
  tokenizers: new Map(),
  processors: new Map(),
  modelConfigs: new Map(),
  webgpu: {
    supportsFp16: null,
  },
};

const ASR_DTYPE_PREFERRED = "fp16";
const ASR_DTYPE_FALLBACK = "fp32";
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

const DEFAULTS = {
  transformersUrl: "../../AI/vendor/transformers.min.js",
  localModelPath: "../../AI/models/",
  allowRemote: false,
  asrModel: "Xenova/whisper-tiny",
  visionModel: "LiquidAI/LFM2.5-VL-1.6B-ONNX",
};

const WASM_MAGIC = ["00", "61", "73", "6d"];

function getOrtConfig() {
  const fallback = { version: "1.23.2", base: "../../AI/vendor/ort-1.23.2/", bundle: "webgpu" };
  if (window.__ortConfig && window.__ortConfig.version && window.__ortConfig.base) {
    return window.__ortConfig;
  }
  return fallback;
}

async function ensureOrtLoaded() {
  if (window.__ortLoadPromise) {
    try {
      await window.__ortLoadPromise;
    } catch (err) {
      log(`ORT load failed: ${err.message}`);
    }
  }
}

function configureOrt() {
  if (!window.ort?.env?.wasm) return;
  const { version, base: basePath, bundle } = getOrtConfig();
  const base = new URL(basePath, window.location.href).toString();
  const cacheBust = `?v=${version}`;
  const disableSimd = disableWasmSimdEl?.checked ?? false;
  const wantsWebgpu = (bundle || "").toLowerCase() === "webgpu";
  const isOrt123 = /^1\.23\./.test(version);
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
    if (disableSimd) {
      log("WebGPU requires WASM SIMD; overriding disable SIMD toggle.");
    }
  }
  window.ort.env.wasm.wasmPaths = wasmPaths;
  window.ort.env.wasm.simd = wantsWebgpu ? true : !disableSimd;
  window.ort.env.wasm.proxy = false;
  const canThread = typeof crossOriginIsolated !== "undefined" && crossOriginIsolated;
  window.ort.env.wasm.numThreads = canThread ? Math.min(4, navigator.hardwareConcurrency || 1) : 1;
  log(
    `ORT wasmPaths set. base=${base} threads=${window.ort.env.wasm.numThreads} crossOriginIsolated=${String(
      canThread
    )} cacheBust=${cacheBust} simd=${String(window.ort.env.wasm.simd)}`
  );
}

function log(message) {
  const timestamp = new Date().toISOString().replace("T", " ").replace("Z", "");
  logEl.textContent += `[${timestamp}] ${message}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function toHex(bytes) {
  return Array.from(bytes)
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join(" ");
}

async function probeWasmFile(label, url) {
  const response = await fetch(url, { cache: "no-store" });
  const contentLength = response.headers.get("content-length") || "unknown";
  if (!response.ok) {
    log(`WASM probe failed: ${label} status=${response.status}`);
    return;
  }
  const buffer = await response.arrayBuffer();
  const header = new Uint8Array(buffer.slice(0, 4));
  const magic = toHex(header);
  const okMagic = WASM_MAGIC.join(" ") === magic;
  log(
    `WASM probe: ${label} size=${buffer.byteLength} header=${magic} ok=${okMagic} content-length=${contentLength}`
  );
}

async function probeWasmFiles() {
  await ensureOrtLoaded();
  if (!window.ort) {
    log("WASM probe skipped: onnxruntime-web not loaded.");
    return;
  }
  if (!window.ort?.env?.wasm?.wasmPaths) {
    configureOrt();
  }
  if (!window.ort?.env?.wasm?.wasmPaths) {
    log("WASM probe skipped: ort.env.wasm.wasmPaths not set yet.");
    return;
  }
  const paths = window.ort.env.wasm.wasmPaths;
  if (typeof paths === "string") {
    log(`WASM probe base path: ${paths}`);
    return;
  }
  const entries = Object.entries(paths);
  for (const [label, url] of entries) {
    await probeWasmFile(label, url);
  }
}

function setStatus(el, message) {
  el.textContent = message;
  log(message);
}

function getDeviceOption() {
  const value = deviceSelectEl.value;
  if (!value || value === "auto") return undefined;
  return value;
}

async function ensureWebGpuFeatures() {
  if (state.webgpu.supportsFp16 !== null) return;
  if (!("gpu" in navigator)) {
    state.webgpu.supportsFp16 = false;
    return;
  }
  try {
    const adapter = await navigator.gpu.requestAdapter();
    if (!adapter) {
      state.webgpu.supportsFp16 = false;
      return;
    }
    state.webgpu.supportsFp16 = adapter.features?.has("shader-f16") ?? false;
  } catch (err) {
    state.webgpu.supportsFp16 = false;
  }
}

async function checkWebGpu() {
  if (!("gpu" in navigator)) {
    setStatus(envStatusEl, "WebGPU not available: navigator.gpu missing.");
    return;
  }
  try {
    const adapter = await navigator.gpu.requestAdapter();
    if (!adapter) {
      setStatus(envStatusEl, "WebGPU adapter not available.");
      return;
    }
    state.webgpu.supportsFp16 = adapter.features?.has("shader-f16") ?? false;
    log(`WebGPU features: shader-f16=${String(state.webgpu.supportsFp16)}`);
    setStatus(envStatusEl, "WebGPU adapter ready.");
  } catch (err) {
    setStatus(envStatusEl, `WebGPU check failed: ${err.message}`);
  }
}

async function reportWebGpu() {
  if (!("gpu" in navigator)) {
    log("GPU report: navigator.gpu missing.");
    return;
  }
  try {
    const adapter = await navigator.gpu.requestAdapter();
    if (!adapter) {
      log("GPU report: adapter not available.");
      return;
    }
    const features = adapter.features ? Array.from(adapter.features.values()) : [];
    const limits = adapter.limits ? { ...adapter.limits } : {};
    state.webgpu.supportsFp16 = adapter.features?.has("shader-f16") ?? false;
    log(`GPU report: features=${features.join(", ") || "none"}`);
    log(`GPU report: limits=${JSON.stringify(limits)}`);
  } catch (err) {
    log(`GPU report failed: ${err.message}`);
  }
}

function applyEnv() {
  if (!state.env) return;
  if (typeof state.env.allowRemoteModels === "boolean") {
    state.env.allowRemoteModels = allowRemoteEl.checked;
  }
  if (typeof state.env.allowLocalModels === "boolean") {
    state.env.allowLocalModels = true;
  }
  const localPath = localModelPathEl.value.trim();
  if (localPath && typeof state.env.localModelPath === "string") {
    state.env.localModelPath = localPath;
  }
  if (typeof state.env.useBrowserCache === "boolean") {
    state.env.useBrowserCache = true;
  }
}

async function loadLibrary() {
  const url = transformersUrlEl.value.trim();
  if (!url) {
    setStatus(envStatusEl, "Enter a Transformers.js module URL first.");
    return;
  }
  setStatus(envStatusEl, `Loading library from ${url} ...`);
  try {
    const mod = await import(url);
    if (!mod.pipeline || !mod.env) {
      throw new Error("Module missing pipeline/env exports.");
    }
    state.lib = mod;
    state.pipeline = mod.pipeline;
    state.env = mod.env;
    applyEnv();
    setStatus(envStatusEl, "Library loaded. Ready.");
  } catch (err) {
    setStatus(envStatusEl, `Library load failed: ${err.message}`);
  }
}

async function ensureLibraryLoaded() {
  if (state.pipeline && state.env) return;
  await loadLibrary();
  if (!state.pipeline) {
    throw new Error("Library not loaded.");
  }
}

async function getPipeline(task, modelId, extraOptions = {}) {
  const device = getDeviceOption();
  const key = JSON.stringify({ task, modelId, device: device || "auto", extraOptions });
  if (state.pipelines.has(key)) return state.pipelines.get(key);
  await ensureLibraryLoaded();
  applyEnv();
  const options = { ...extraOptions };
  if (device) options.device = device;
  const pipe = await state.pipeline(task, modelId, options);
  state.pipelines.set(key, pipe);
  return pipe;
}

function formatMs(value) {
  if (!Number.isFinite(value)) return "n/a";
  if (value < 1000) return `${value.toFixed(1)}ms`;
  return `${(value / 1000).toFixed(2)}s`;
}

function logPerf(label, ms) {
  log(`${label}: ${formatMs(ms)}`);
}

function logPipelineBackend(pipe, label) {
  const info = {
    device: pipe?.device ?? null,
    modelDevice: pipe?.model?.device ?? null,
    backend: state.env?.backends?.onnx?.backend ?? null,
  };
  const providers =
    pipe?.model?.session?.executionProviders ||
    pipe?.model?.session?._executionProviders ||
    pipe?.model?.session?.sessionOptions?.executionProviders ||
    pipe?.model?._session?.executionProviders ||
    pipe?.model?.session?._session?.executionProviders;
  if (providers) {
    info.executionProviders = providers;
  }
  log(`${label} backend: ${JSON.stringify(info)}`);
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

async function fetchAudioBlob(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Audio fetch failed (${response.status}): ${url}`);
  }
  return response.blob();
}

async function resolveAsrDtype() {
  await ensureWebGpuFeatures();
  const deviceChoice = getDeviceOption();
  const wantsWebgpu = deviceChoice === "webgpu" || (!deviceChoice && "gpu" in navigator);
  if (wantsWebgpu && state.webgpu.supportsFp16 === false) {
    log(`WebGPU adapter lacks shader-f16; using ${ASR_DTYPE_FALLBACK}.`);
    return ASR_DTYPE_FALLBACK;
  }
  if (deviceChoice === "wasm") {
    return ASR_DTYPE_FALLBACK;
  }
  return ASR_DTYPE_PREFERRED;
}

async function runAsrInternal(modelId, fileLabel, blob) {
  if (!modelId) {
    setStatus(asrStatusEl, "Enter an ASR model id.");
    return;
  }
  setStatus(asrStatusEl, `Loading ASR pipeline (${modelId}) ...`);
  const totalStart = performance.now();
  try {
    const dtype = await resolveAsrDtype();
    await ensureAsrModelFiles(modelId, dtype);
    const loadStart = performance.now();
    const asr = await getPipeline("automatic-speech-recognition", modelId, { dtype });
    logPipelineBackend(asr, "ASR");
    logPerf("ASR pipeline load", performance.now() - loadStart);
    setStatus(asrStatusEl, "Decoding audio ...");
    const decodeStart = performance.now();
    const { audio, sampling_rate } = await decodeAudioBlob(blob);
    logPerf("ASR audio decode", performance.now() - decodeStart);
    log(`ASR audio: type=${audio?.constructor?.name || "unknown"} length=${audio?.length || 0}`);
    setStatus(asrStatusEl, "Running transcription ...");
    const options = {
      chunk_length_s: 30,
      stride_length_s: 5,
    };
    if (asrTimestampsEl.checked) {
      options.return_timestamps = true;
    }
    const inferStart = performance.now();
    const result = await asr(audio, options);
    logPerf("ASR inference", performance.now() - inferStart);
    asrOutputEl.value = typeof result === "string" ? result : JSON.stringify(result, null, 2);
    setStatus(asrStatusEl, "Transcription complete.");
    logPerf("ASR total", performance.now() - totalStart);
  } catch (err) {
    log(err?.stack || String(err));
    setStatus(asrStatusEl, `ASR failed: ${err.message}`);
  }
}

async function runAsr() {
  const modelId = asrModelEl.value.trim();
  const file = asrFileEl.files[0];
  if (!file) {
    setStatus(asrStatusEl, "Pick an audio file first.");
    return;
  }
  await runAsrInternal(modelId, file.name || "audio", file);
}

async function runAsrFromUrl(modelId, url) {
  const blob = await fetchAudioBlob(url);
  const label = url.split("/").pop() || "audio";
  await runAsrInternal(modelId, label, blob);
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

async function runVision() {
  const task = visionTaskEl.value;
  const modelId = visionModelEl.value.trim();
  const file = visionFileEl.files[0];
  if (!modelId) {
    setStatus(visionStatusEl, "Enter a vision model id.");
    return;
  }
  if (!file) {
    setStatus(visionStatusEl, "Pick an image file first.");
    return;
  }
  const useOrt = LIQUIDAI_MODELS.some((entry) => modelId.includes(entry.match));
  if (useOrt) {
    await runOrtVision(modelId, file);
    return;
  }

  setStatus(visionStatusEl, `Loading vision pipeline (${modelId}) ...`);
  try {
    const pipe = await getPipeline(task, modelId);
    const img = await loadImageFromFile(file);
    visionPreviewEl.src = img.src;
    setStatus(visionStatusEl, "Running vision inference ...");
    const result = await pipe(img);
    visionOutputEl.textContent = JSON.stringify(result, null, 2);
    setStatus(visionStatusEl, "Vision inference complete.");
  } catch (err) {
    log(err?.stack || String(err));
    setStatus(visionStatusEl, `Vision failed: ${err.message}`);
  }
}

async function downloadLog() {
  const content = logEl.textContent || "";
  const filename = `ai-webgpu-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`;

  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: "Text log", accept: { "text/plain": [".txt"] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(content);
      await writable.close();
      log(`Saved log to ${handle.name}`);
      return;
    } catch (err) {
      log(`Save canceled or failed: ${err.message}`);
    }
  }

  const blob = new Blob([content], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
  log("Downloaded log via browser fallback.");
}

function resolveOrtModelConfig(modelId) {
  return LIQUIDAI_MODELS.find((entry) => modelId.includes(entry.match)) || null;
}

function modelUsesFp16(config) {
  return Boolean(config?.embedTokens?.includes("_fp16") || config?.embedImages?.includes("_fp16"));
}

function resolveOrtProvider() {
  const choice = deviceSelectEl.value;
  if (choice === "webgpu" && "gpu" in navigator) return ORT_PROVIDERS.webgpu;
  if (choice === "wasm") return ORT_PROVIDERS.wasm;
  if ("gpu" in navigator) return ORT_PROVIDERS.webgpu;
  return ORT_PROVIDERS.wasm;
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

function getLocalModelBase(modelId) {
  const root = localModelPathEl.value.trim() || "./models/";
  return root.endsWith("/") ? `${root}${modelId}/` : `${root}/${modelId}/`;
}

async function ensureAsrModelFiles(modelId, dtype) {
  if (allowRemoteEl.checked) return;
  if (!dtype) return;
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
    throw new Error(
      `Missing ASR ${lowered} files in ${base}: ${missing.join(
        ", "
      )}. Download the matching ONNX files to enable ASR.`
    );
  }
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
    const isImage = name.toLowerCase().includes("image") || name.toLowerCase().includes("pixel");
    const isTokens = name.toLowerCase().includes("input") || name.toLowerCase().includes("token");
    if (isImage) {
      while (next.length < 4) next.push(0);
      if (next[0] === 0) next[0] = 1;
      if (next[1] === 0) next[1] = 3;
      if (next[2] === 0) next[2] = 224;
      if (next[3] === 0) next[3] = 224;
    } else if (isTokens) {
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

function float32ToFloat16(value) {
  const floatView = new Float32Array(1);
  const intView = new Uint32Array(floatView.buffer);
  floatView[0] = value;
  const x = intView[0];
  const sign = (x >> 16) & 0x8000;
  const mantissa = x & 0x7fffff;
  const exp = (x >> 23) & 0xff;
  if (exp === 0) return sign;
  if (exp === 255) return sign | 0x7c00;
  const halfExp = exp - 127 + 15;
  if (halfExp >= 31) return sign | 0x7c00;
  if (halfExp <= 0) return sign;
  return sign | (halfExp << 10) | (mantissa >> 13);
}

function buildImageTensor(img, dims, type) {
  const [n, c, h, w] = dims;
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0, w, h);
  const { data } = ctx.getImageData(0, 0, w, h);
  const size = n * c * h * w;
  const floatData = new Float32Array(size);
  let idx = 0;
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      const base = (y * w + x) * 4;
      const r = data[base] / 255;
      const g = data[base + 1] / 255;
      const b = data[base + 2] / 255;
      floatData[idx] = r; // R
      floatData[idx + h * w] = g; // G
      floatData[idx + 2 * h * w] = b; // B
      idx += 1;
    }
  }

  if (type === "float16") {
    const half = new Uint16Array(size);
    for (let i = 0; i < size; i += 1) {
      half[i] = float32ToFloat16(floatData[i]);
    }
    return half;
  }
  return floatData;
}

function buildDummyInputs(inputMeta, img) {
  const inputs = {};
  for (const [name, meta] of Object.entries(inputMeta || {})) {
    const dims = resolveDimensions(name, meta);
    const total = dims.reduce((acc, value) => acc * value, 1);
    let data;
    if (img && (name.toLowerCase().includes("image") || name.toLowerCase().includes("pixel"))) {
      data = buildImageTensor(img, dims, meta.type);
    } else {
      data = toOrtTypedArray(meta.type, total);
      fillDefaultValues(meta.type, data);
    }
    inputs[name] = new window.ort.Tensor(meta.type || "float32", data, dims);
  }
  return inputs;
}

function buildInputIdsTensor(ids) {
  const data = new BigInt64Array(ids.map((value) => BigInt(value)));
  return new window.ort.Tensor("int64", data, [1, ids.length]);
}

function buildAttentionMask(length) {
  const data = new BigInt64Array(length);
  data.fill(1n);
  return new window.ort.Tensor("int64", data, [1, length]);
}

function buildPositionIds(length) {
  const data = new BigInt64Array(length);
  for (let i = 0; i < length; i += 1) {
    data[i] = BigInt(i);
  }
  return new window.ort.Tensor("int64", data, [1, length]);
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
  if (!state.lib?.AutoTokenizer) {
    throw new Error("Transformers.js AutoTokenizer not available. Load library first.");
  }
  const tokenizer = await state.lib.AutoTokenizer.from_pretrained(modelId, {
    local_files_only: !allowRemoteEl.checked,
  });
  state.tokenizers.set(base, tokenizer);
  return tokenizer;
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
      local_files_only: !allowRemoteEl.checked,
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
  const response = await fetch(`${base}config.json`);
  if (!response.ok) return null;
  const config = await response.json();
  state.modelConfigs.set(base, config);
  return config;
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

function resolveInputNames(inputMeta, inputNames, fallbackNames) {
  const entries = inputMeta instanceof Map ? Array.from(inputMeta.entries()) : Object.entries(inputMeta || {});
  const keyNames = entries.map(([name]) => name);
  const namesFromSession = Array.isArray(inputNames) ? inputNames : [];
  const numericOnly =
    keyNames.length > 0 && keyNames.every((name) => typeof name === "string" && /^\d+$/.test(name));
  const useNames =
    (!entries.length || numericOnly) && namesFromSession.length
      ? namesFromSession
      : numericOnly && fallbackNames?.length
        ? fallbackNames
        : keyNames;
  return { entries, keyNames, namesFromSession, numericOnly, useNames };
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
  let hasInputIds = false;
  let hasAttention = false;
  let hasPosition = false;

  for (const [name, meta] of entries) {
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

  return inputs;
}

function mergeLiquidEmbeds(tokenEmbeds, imageEmbeds, inputIds, imageTokenId) {
  if (!tokenEmbeds?.data || !imageEmbeds?.data) return tokenEmbeds;
  const [batch, seqLen, hidden] = tokenEmbeds.dims;
  const imageDims = imageEmbeds.dims;
  const imageTokens = imageDims.length === 3 ? imageDims[1] : imageDims[0];
  const imageHidden = imageDims.length === 3 ? imageDims[2] : imageDims[1];
  if (hidden !== imageHidden) {
    log(`LiquidAI embed mismatch: token hidden=${hidden} image hidden=${imageHidden}`);
    return tokenEmbeds;
  }
  const positions = inputIds
    .map((value, index) => (value === imageTokenId ? index : -1))
    .filter((value) => value >= 0);
  if (!positions.length) return tokenEmbeds;

  const maxCopies = Math.min(positions.length, imageTokens);
  for (let i = 0; i < maxCopies; i += 1) {
    const srcBase = i * hidden;
    const dstBase = positions[i] * hidden;
    tokenEmbeds.data.set(imageEmbeds.data.subarray(srcBase, srcBase + hidden), dstBase);
  }
  return tokenEmbeds;
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
      continue;
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
  const metaMap = new Map(entries.map(([name, meta]) => [name, meta]));

  for (const name of useNames) {
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

  return inputs;
}

async function loadOrtSession(modelPath, providers) {
  await ensureOrtLoaded();
  const cacheKey = `${modelPath}::${providers.join(",")}`;
  if (state.ortSessions.has(cacheKey)) {
    log(`ORT session cache hit: ${modelPath}`);
    return state.ortSessions.get(cacheKey);
  }
  if (!window.ort) {
    throw new Error("onnxruntime-web not loaded (AI/vendor/ort-1.23.2/ort.webgpu.min.js missing).");
  }
  configureOrt();
  const response = await fetch(modelPath, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`ORT fetch failed (${response.status}): ${modelPath}`);
  }
  const onnxBuffer = new Uint8Array(await response.arrayBuffer());
  const externalData = await resolveExternalOrtData(modelPath);
  log(`ORT create session: ${modelPath} providers=${providers.join(",")}`);
  const session = await window.ort.InferenceSession.create(onnxBuffer, {
    executionProviders: providers,
    externalData: externalData || undefined,
  });
  state.ortSessions.set(cacheKey, session);
  log(`ORT session ready: ${modelPath}`);
  return session;
}

async function warmupOrtSessions(modelId, providersOverride) {
  await ensureOrtLoaded();
  const config = resolveOrtModelConfig(modelId);
  if (!config) {
    setStatus(visionStatusEl, `Unknown LiquidAI model config for ${modelId}`);
    return;
  }
  if (!window.ort) {
    setStatus(visionStatusEl, "onnxruntime-web not loaded (AI/vendor/ort-1.23.2/ort.webgpu.min.js missing).");
    return;
  }
  const base = getLocalModelBase(modelId);
  const providers = providersOverride || (await resolveOrtProviderForModel(modelId, config));
  try {
    setStatus(visionStatusEl, `Loading ONNX sessions (${modelId}) ...`);
    await loadOrtSession(`${base}${config.embedTokens}`, providers);
    await loadOrtSession(`${base}${config.embedImages}`, providers);
    await loadOrtSession(`${base}${config.decoder}`, providers);
    setStatus(visionStatusEl, "ONNX sessions loaded.");
  } catch (err) {
    log(err?.stack || String(err));
    setStatus(visionStatusEl, `ONNX load failed: ${err.message}`);
  }
}

async function runOrtVision(modelId, file) {
  const config = resolveOrtModelConfig(modelId);
  if (!config) {
    setStatus(visionStatusEl, `Unknown LiquidAI model config for ${modelId}`);
    return;
  }
  if (!file) {
    setStatus(visionStatusEl, "Pick an image file first.");
    return;
  }

  const base = getLocalModelBase(modelId);
  const providers = await resolveOrtProviderForModel(modelId, config);
  const img = await loadImageFromFile(file);
  visionPreviewEl.src = img.src;

  try {
    setStatus(visionStatusEl, `Loading ONNX sessions (${modelId}) ...`);
    const embedTokenSession = await loadOrtSession(`${base}${config.embedTokens}`, providers);
    const embedImageSession = await loadOrtSession(`${base}${config.embedImages}`, providers);
    const outputs = {};

    setStatus(visionStatusEl, "Running token embedding test ...");
    const tokenizer = await getTokenizer(modelId);
    const processor = await getProcessor(modelId);
    const prompt = await buildLiquidPrompt(
      tokenizer,
      modelId,
      visionQuestionEl.value.trim() || "Describe the image."
    );
    const processed = await runLiquidProcessor(processor, img, prompt);
    const inputIds = await tokenizeQuestion(tokenizer, prompt);
    const attentionMask = buildAttentionMaskFromLength(inputIds.length);
    const positionIds = buildPositionIdsFromLength(inputIds.length);
    const tokenInputs = buildLiquidTokenInputs(
      embedTokenSession.inputMetadata,
      embedTokenSession.inputNames,
      inputIds,
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
    outputs.tokenEmbedding = Object.fromEntries(
      Object.entries(tokenResult).map(([key, value]) => [key, value.dims])
    );

    setStatus(visionStatusEl, "Running image embedding test ...");
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
    outputs.imageEmbedding = Object.fromEntries(
      Object.entries(imageResult).map(([key, value]) => [key, value.dims])
    );

    const summary = {
      modelId,
      providers,
      tokenInputs: embedTokenSession.inputMetadata,
      imageInputs: embedImageSession.inputMetadata,
      outputs,
      prompt,
      note: "Embeddings only. Use Ask (LiquidAI) for a simple chat test.",
    };
    visionOutputEl.textContent = JSON.stringify(summary, null, 2);
    setStatus(visionStatusEl, "ONNX embedding tests complete.");
  } catch (err) {
    log(err?.stack || String(err));
    setStatus(visionStatusEl, `ONNX vision failed: ${err.message}`);
  }
}

async function runOrtChat(modelId, file, question) {
  await ensureOrtLoaded();
  const config = resolveOrtModelConfig(modelId);
  if (!config) {
    setStatus(visionStatusEl, `Unknown LiquidAI model config for ${modelId}`);
    return;
  }
  if (!file) {
    setStatus(visionStatusEl, "Pick an image file first.");
    return;
  }
  if (!question) {
    setStatus(visionStatusEl, "Enter a question first.");
    return;
  }
  if (!window.ort) {
    setStatus(visionStatusEl, "onnxruntime-web not loaded (AI/vendor/ort-1.23.2/ort.webgpu.min.js missing).");
    return;
  }
  if (!state.lib) {
    setStatus(visionStatusEl, "Load Transformers.js first (needed for tokenizer).");
    return;
  }

  const base = getLocalModelBase(modelId);
  const providers = resolveOrtProvider();
  const img = await loadImageFromFile(file);
  visionPreviewEl.src = img.src;

  try {
    setStatus(visionStatusEl, "Loading tokenizer + processor ...");
    const tokenizer = await getTokenizer(modelId);
    const processor = await getProcessor(modelId);
    const configJson = await getModelConfig(modelId);
    const eosTokenId = configJson?.text_config?.eos_token_id ?? configJson?.eos_token_id ?? null;
    const imageTokenId =
      configJson?.image_token_id ??
      configJson?.image_token_index ??
      (typeof tokenizer.convert_tokens_to_ids === "function" ? tokenizer.convert_tokens_to_ids("<image>") : null);

    setStatus(visionStatusEl, "Encoding prompt + image ...");
    const prompt = await buildLiquidPrompt(tokenizer, modelId, question);
    const processed = await runLiquidProcessor(processor, img, prompt);

    setStatus(visionStatusEl, "Loading ONNX sessions ...");
    const embedTokenSession = await loadOrtSession(`${base}${config.embedTokens}`, providers);
    const embedImageSession = await loadOrtSession(`${base}${config.embedImages}`, providers);
    const decoderSession = await loadOrtSession(`${base}${config.decoder}`, providers);

    setStatus(visionStatusEl, "Embedding image + tokens ...");
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

    const maxNewTokens = 64;
    const minSummaryChars = 120;
    const generated = [];
    let cache = initLiquidCache(decoderSession, configJson);
    let currentEmbeds = tokenEmbedding;

    setStatus(visionStatusEl, "Generating (greedy) ...");
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
      const decodedSoFar = decodeTokens(tokenizer, generated);
      if (decodedSoFar.length >= minSummaryChars && /[.!?]\s*$/.test(decodedSoFar)) {
        break;
      }

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

    visionOutputEl.textContent = JSON.stringify(
      {
        modelId,
        providers,
        question,
        prompt,
        generatedTokenIds: generated,
        answer: decoded,
        elapsedMs: elapsed,
      },
      null,
      2
    );
    setStatus(visionStatusEl, "LiquidAI chat complete.");
  } catch (err) {
    log(err?.stack || String(err));
    setStatus(visionStatusEl, `LiquidAI chat failed: ${err.message}`);
  }
}

function applyDefaults() {
  transformersUrlEl.value = DEFAULTS.transformersUrl;
  localModelPathEl.value = DEFAULTS.localModelPath;
  allowRemoteEl.checked = DEFAULTS.allowRemote;
  asrModelEl.value = DEFAULTS.asrModel;
  visionModelEl.value = DEFAULTS.visionModel;
  if ("gpu" in navigator) {
    deviceSelectEl.value = "webgpu";
  }
  const { version } = getOrtConfig();
  if (ortVersionEl) {
    ortVersionEl.value = version;
  }
  const { bundle } = getOrtConfig();
  if (ortBundleEl) {
    ortBundleEl.value = bundle || "webgpu";
  }
}

async function autoStart() {
  applyDefaults();
  const { version, base, bundle } = getOrtConfig();
  log(`ORT config: version=${version} base=${base} bundle=${bundle}`);
  ensureOrtLoaded().then(() => {
    if (window.ort?.env?.versions?.web) {
      log(`ORT loaded: ${window.ort.env.versions.web}`);
    }
  });
  log(`User agent: ${navigator.userAgent}`);
  log(`Cross-origin isolated: ${String(crossOriginIsolated)}`);
  configureOrt();
  await checkWebGpu();
  await loadLibrary();
  const params = new URLSearchParams(window.location.search);
  if (params.get("auto") === "1") {
    const asrModelParam = params.get("asrModel");
    if (asrModelParam) {
      asrModelEl.value = asrModelParam;
    }
    const audioUrl = params.get("audio");
    if (audioUrl) {
      try {
        await runAsrFromUrl(asrModelEl.value.trim(), audioUrl);
      } catch (err) {
        log(err?.stack || String(err));
        setStatus(asrStatusEl, `Auto ASR failed: ${err.message}`);
      }
    }
  }
}

loadLibBtn.addEventListener("click", () => {
  loadLibrary();
});

checkWebgpuBtn.addEventListener("click", () => {
  checkWebGpu();
});

probeWasmBtn.addEventListener("click", () => {
  probeWasmFiles();
});

if (reportWebgpuBtn) {
  reportWebgpuBtn.addEventListener("click", () => {
    reportWebGpu();
  });
}

runAsrBtn.addEventListener("click", () => {
  runAsr();
});

runVisionBtn.addEventListener("click", () => {
  runVision();
});

if (loadVisionOnnxBtn) {
  loadVisionOnnxBtn.addEventListener("click", () => {
    const modelId = visionModelEl.value.trim();
    if (!modelId) {
      setStatus(visionStatusEl, "Enter a vision model id first.");
      return;
    }
    warmupOrtSessions(modelId);
  });
} else {
  log("Load ONNX sessions button missing in DOM.");
}

if (loadVisionOnnxWasmBtn) {
  loadVisionOnnxWasmBtn.addEventListener("click", () => {
    const modelId = visionModelEl.value.trim();
    if (!modelId) {
      setStatus(visionStatusEl, "Enter a vision model id first.");
      return;
    }
    warmupOrtSessions(modelId, ORT_PROVIDERS.wasm);
  });
}

if (loadVisionOnnxWebgpuBtn) {
  loadVisionOnnxWebgpuBtn.addEventListener("click", async () => {
    const modelId = visionModelEl.value.trim();
    if (!modelId) {
      setStatus(visionStatusEl, "Enter a vision model id first.");
      return;
    }
    await ensureWebGpuFeatures();
    const config = resolveOrtModelConfig(modelId);
    if (config && state.webgpu.supportsFp16 === false && modelUsesFp16(config)) {
      setStatus(visionStatusEl, "WebGPU shader-f16 not available; use WASM for fp16 LiquidAI models.");
      return;
    }
    warmupOrtSessions(modelId, ORT_PROVIDERS.webgpu);
  });
}

askVisionBtn.addEventListener("click", () => {
  const modelId = visionModelEl.value.trim();
  const file = visionFileEl.files[0];
  const question = visionQuestionEl.value.trim();
  if (!modelId) {
    setStatus(visionStatusEl, "Enter a vision model id first.");
    return;
  }
  const useOrt = LIQUIDAI_MODELS.some((entry) => modelId.includes(entry.match));
  if (!useOrt) {
    setStatus(visionStatusEl, "Ask is only wired for LiquidAI ONNX models in this spike.");
    return;
  }
  runOrtChat(modelId, file, question);
});

downloadLogBtn.addEventListener("click", () => {
  downloadLog();
});

allowRemoteEl.addEventListener("change", () => {
  applyEnv();
});

if (disableWasmSimdEl) {
  disableWasmSimdEl.addEventListener("change", () => {
    configureOrt();
  });
}

if (ortVersionEl) {
  ortVersionEl.addEventListener("change", () => {
    const next = ortVersionEl.value;
    const params = new URLSearchParams(window.location.search);
    if (next === "1.18.0") {
      params.delete("ort");
    } else {
      params.set("ort", next);
    }
    const query = params.toString();
    window.location.search = query ? `?${query}` : "";
  });
}

if (ortBundleEl) {
  ortBundleEl.addEventListener("change", () => {
    const next = ortBundleEl.value;
    const params = new URLSearchParams(window.location.search);
    if (next === "webgpu") {
      params.delete("ortbundle");
    } else {
      params.set("ortbundle", next);
    }
    const query = params.toString();
    window.location.search = query ? `?${query}` : "";
  });
}

localModelPathEl.addEventListener("change", () => {
  applyEnv();
});

window.addEventListener("DOMContentLoaded", () => {
  window.addEventListener("error", (event) => {
    log(`Window error: ${event.message}`);
  });
  window.addEventListener("unhandledrejection", (event) => {
    log(`Unhandled rejection: ${event.reason?.message || String(event.reason)}`);
  });
  autoStart();
});
