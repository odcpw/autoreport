const byId = (id) => document.getElementById(id);

const transformersUrlEl = byId("transformers-url");
const localModelPathEl = byId("local-model-path");
const allowRemoteEl = byId("allow-remote");
const deviceSelectEl = byId("device-select");
const loadLibBtn = byId("load-lib");
const checkWebgpuBtn = byId("check-webgpu");
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
  tokenizers: new Map(),
  modelConfigs: new Map(),
};

const ASR_DTYPE = "q4";
const ORT_PROVIDERS = {
  webgpu: ["webgpu"],
  wasm: ["wasm"],
};

const LIQUIDAI_MODELS = [
  {
    match: "LFM2.5-VL-1.6B-ONNX",
    embedTokens: "onnx/embed_tokens_fp16.onnx",
    embedImages: "onnx/embed_images_fp16.onnx",
    decoder: "onnx/decoder_q4.onnx",
  },
  {
    match: "LFM2-VL-450M-ONNX",
    embedTokens: "onnx/embed_tokens.onnx",
    embedImages: "onnx/vision_encoder.onnx",
    decoder: "onnx/decoder_model_merged.onnx",
  },
];

const DEFAULTS = {
  transformersUrl: "./vendor/transformers.min.js",
  localModelPath: "/AutoBericht/experiments/models/",
  allowRemote: false,
  asrModel: "Xenova/whisper-tiny.en",
  visionModel: "LiquidAI/LFM2.5-VL-1.6B-ONNX",
};

function configureOrt() {
  if (!window.ort?.env?.wasm) return;
  const base = new URL("./vendor/", window.location.href).toString();
  window.ort.env.wasm.wasmPaths = {
    "ort-wasm.wasm": `${base}ort-wasm.wasm`,
    "ort-wasm-simd.wasm": `${base}ort-wasm-simd.wasm`,
    "ort-wasm-threaded.wasm": `${base}ort-wasm-threaded.wasm`,
    "ort-wasm-simd-threaded.wasm": `${base}ort-wasm-simd-threaded.wasm`,
  };
  window.ort.env.wasm.simd = true;
  const canThread = typeof crossOriginIsolated !== "undefined" && crossOriginIsolated;
  window.ort.env.wasm.numThreads = canThread ? Math.min(4, navigator.hardwareConcurrency || 1) : 1;
  log(
    `ORT wasmPaths set. base=${base} threads=${window.ort.env.wasm.numThreads} crossOriginIsolated=${String(
      canThread
    )}`
  );
}

function log(message) {
  const timestamp = new Date().toISOString().replace("T", " ").replace("Z", "");
  logEl.textContent += `[${timestamp}] ${message}\n`;
  logEl.scrollTop = logEl.scrollHeight;
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
    setStatus(envStatusEl, "WebGPU adapter ready.");
  } catch (err) {
    setStatus(envStatusEl, `WebGPU check failed: ${err.message}`);
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

async function getPipeline(task, modelId, extraOptions = {}) {
  const device = getDeviceOption();
  const key = JSON.stringify({ task, modelId, device: device || "auto", extraOptions });
  if (state.pipelines.has(key)) return state.pipelines.get(key);
  if (!state.pipeline) {
    throw new Error("Library not loaded.");
  }
  applyEnv();
  const options = { ...extraOptions };
  if (device) options.device = device;
  const pipe = await state.pipeline(task, modelId, options);
  state.pipelines.set(key, pipe);
  return pipe;
}

async function decodeAudioFile(file) {
  const data = await file.arrayBuffer();
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
  const audio = buffer.getChannelData(0);
  audioCtx.close();
  return { audio, sampling_rate: targetRate };
}

async function runAsr() {
  const modelId = asrModelEl.value.trim();
  const file = asrFileEl.files[0];
  if (!modelId) {
    setStatus(asrStatusEl, "Enter an ASR model id.");
    return;
  }
  if (!file) {
    setStatus(asrStatusEl, "Pick an audio file first.");
    return;
  }
  setStatus(asrStatusEl, `Loading ASR pipeline (${modelId}) ...`);
  try {
    const asr = await getPipeline("automatic-speech-recognition", modelId, { dtype: ASR_DTYPE });
    setStatus(asrStatusEl, "Decoding audio ...");
    const { audio, sampling_rate } = await decodeAudioFile(file);
    setStatus(asrStatusEl, "Running transcription ...");
    const options = {
      chunk_length_s: 30,
      stride_length_s: 5,
    };
    if (asrTimestampsEl.checked) {
      options.return_timestamps = true;
    }
    const result = await asr({ array: audio, sampling_rate }, options);
    asrOutputEl.value = typeof result === "string" ? result : JSON.stringify(result, null, 2);
    setStatus(asrStatusEl, "Transcription complete.");
  } catch (err) {
    log(err?.stack || String(err));
    setStatus(asrStatusEl, `ASR failed: ${err.message}`);
  }
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

function resolveOrtProvider() {
  const choice = deviceSelectEl.value;
  if (choice === "webgpu" && "gpu" in navigator) return ORT_PROVIDERS.webgpu;
  if (choice === "wasm") return ORT_PROVIDERS.wasm;
  if ("gpu" in navigator) return ORT_PROVIDERS.webgpu;
  return ORT_PROVIDERS.wasm;
}

function getLocalModelBase(modelId) {
  const root = localModelPathEl.value.trim() || "/AutoBericht/experiments/models/";
  return root.endsWith("/") ? `${root}${modelId}/` : `${root}/${modelId}/`;
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
    local_files_only: true,
  });
  state.tokenizers.set(base, tokenizer);
  return tokenizer;
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

async function loadOrtSession(modelPath, providers) {
  const cacheKey = `${modelPath}::${providers.join(",")}`;
  if (state.ortSessions.has(cacheKey)) {
    log(`ORT session cache hit: ${modelPath}`);
    return state.ortSessions.get(cacheKey);
  }
  if (!window.ort) {
    throw new Error("onnxruntime-web not loaded (vendor/ort.webgpu.min.js missing).");
  }
  configureOrt();
  log(`ORT create session: ${modelPath} providers=${providers.join(",")}`);
  const session = await window.ort.InferenceSession.create(modelPath, {
    executionProviders: providers,
  });
  state.ortSessions.set(cacheKey, session);
  log(`ORT session ready: ${modelPath}`);
  return session;
}

async function warmupOrtSessions(modelId) {
  const config = resolveOrtModelConfig(modelId);
  if (!config) {
    setStatus(visionStatusEl, `Unknown LiquidAI model config for ${modelId}`);
    return;
  }
  if (!window.ort) {
    setStatus(visionStatusEl, "onnxruntime-web not loaded (vendor/ort.webgpu.min.js missing).");
    return;
  }
  const base = getLocalModelBase(modelId);
  const providers = resolveOrtProvider();
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
  const providers = resolveOrtProvider();
  const img = await loadImageFromFile(file);
  visionPreviewEl.src = img.src;

  try {
    setStatus(visionStatusEl, `Loading ONNX sessions (${modelId}) ...`);
    const embedTokenSession = await loadOrtSession(`${base}${config.embedTokens}`, providers);
    const embedImageSession = await loadOrtSession(`${base}${config.embedImages}`, providers);
    const outputs = {};

    setStatus(visionStatusEl, "Running token embedding test ...");
    const tokenInputs = buildDummyInputs(embedTokenSession.inputMetadata, null);
    const tokenResult = await embedTokenSession.run(tokenInputs);
    outputs.tokenEmbedding = Object.fromEntries(
      Object.entries(tokenResult).map(([key, value]) => [key, value.dims])
    );

    setStatus(visionStatusEl, "Running image embedding test ...");
    const imageInputs = buildDummyInputs(embedImageSession.inputMetadata, img);
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
    setStatus(visionStatusEl, "onnxruntime-web not loaded (vendor/ort.webgpu.min.js missing).");
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
    setStatus(visionStatusEl, "Loading tokenizer ...");
    const tokenizer = await getTokenizer(modelId);
    const configJson = await getModelConfig(modelId);
    const eosTokenId = configJson?.text_config?.eos_token_id ?? configJson?.eos_token_id ?? null;

    setStatus(visionStatusEl, "Encoding question ...");
    const prompt = `User: ${question}\nAssistant:`.trim();
    const promptIds = await tokenizeQuestion(tokenizer, prompt);

    setStatus(visionStatusEl, "Loading ONNX sessions ...");
    const embedTokenSession = await loadOrtSession(`${base}${config.embedTokens}`, providers);
    const embedImageSession = await loadOrtSession(`${base}${config.embedImages}`, providers);
    const decoderSession = await loadOrtSession(`${base}${config.decoder}`, providers);

    setStatus(visionStatusEl, "Embedding image ...");
    const imageInputs = buildDummyInputs(embedImageSession.inputMetadata, img);
    const imageResult = await embedImageSession.run(imageInputs);
    const imageEmbedding = Object.values(imageResult)[0];

    const maxNewTokens = 16;
    const generated = [];

    setStatus(visionStatusEl, "Generating (greedy) ...");
    const start = performance.now();
    for (let step = 0; step < maxNewTokens; step += 1) {
      const allIds = [...promptIds, ...generated];
      const inputIdsTensor = buildInputIdsTensor(allIds);
      const attentionMaskTensor = buildAttentionMask(allIds.length);
      const positionIdsTensor = buildPositionIds(allIds.length);

      const tokenInputs = {};
      for (const [name, meta] of Object.entries(embedTokenSession.inputMetadata || {})) {
        const lower = name.toLowerCase();
        if (lower.includes("input") && lower.includes("id")) {
          tokenInputs[name] = inputIdsTensor;
        } else if (lower.includes("attention")) {
          tokenInputs[name] = attentionMaskTensor;
        } else {
          const dims = resolveDimensions(name, meta);
          const total = dims.reduce((acc, value) => acc * value, 1);
          const data = toOrtTypedArray(meta.type, total);
          fillDefaultValues(meta.type, data);
          tokenInputs[name] = new window.ort.Tensor(meta.type || "float32", data, dims);
        }
      }
      const tokenResult = await embedTokenSession.run(tokenInputs);
      const tokenEmbedding = Object.values(tokenResult)[0];

      const decoderInputs = {};
      for (const [name, meta] of Object.entries(decoderSession.inputMetadata || {})) {
        const lower = name.toLowerCase();
        if (lower.includes("input") && lower.includes("id")) {
          decoderInputs[name] = inputIdsTensor;
        } else if (lower.includes("attention")) {
          decoderInputs[name] = attentionMaskTensor;
        } else if (lower.includes("position")) {
          decoderInputs[name] = positionIdsTensor;
        } else if (lower.includes("image")) {
          decoderInputs[name] = imageEmbedding;
        } else if (lower.includes("embed")) {
          decoderInputs[name] = tokenEmbedding;
        } else {
          const dims = resolveDimensions(name, meta);
          const total = dims.reduce((acc, value) => acc * value, 1);
          const data = toOrtTypedArray(meta.type, total);
          fillDefaultValues(meta.type, data);
          decoderInputs[name] = new window.ort.Tensor(meta.type || "float32", data, dims);
        }
      }

      const result = await decoderSession.run(decoderInputs);
      const logitsTensor = resolveLogitsOutput(result);
      if (!logitsTensor) throw new Error("Decoder did not return logits.");
      const nextId = argmaxLogits(logitsTensor);
      generated.push(nextId);
      if (eosTokenId !== null && nextId === eosTokenId) break;
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
}

async function autoStart() {
  applyDefaults();
  configureOrt();
  await checkWebGpu();
  await loadLibrary();
}

loadLibBtn.addEventListener("click", () => {
  loadLibrary();
});

checkWebgpuBtn.addEventListener("click", () => {
  checkWebGpu();
});

runAsrBtn.addEventListener("click", () => {
  runAsr();
});

runVisionBtn.addEventListener("click", () => {
  runVision();
});

loadVisionOnnxBtn.addEventListener("click", () => {
  const modelId = visionModelEl.value.trim();
  if (!modelId) {
    setStatus(visionStatusEl, "Enter a vision model id first.");
    return;
  }
  warmupOrtSessions(modelId);
});

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
