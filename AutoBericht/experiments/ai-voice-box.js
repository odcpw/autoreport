const byId = (id) => document.getElementById(id);

const asrModelEl = byId("asr-model");
const loadLibBtn = byId("load-lib");
const holdBtn = byId("hold-to-talk");
const statusEl = byId("status");
const transcriptEl = byId("transcript");
const cleanedEl = byId("cleaned");
const logEl = byId("log");

const DEFAULTS = {
  transformersUrl: "../AI/vendor/transformers.min.js",
  localModelPath: "../AI/models/",
};

const CLEANUP_MODEL = "LiquidAI/LFM2.5-VL-1.6B-ONNX";
const CLEANUP_MAX_NEW_TOKENS_CAP = 256;
const CLEANUP_MIN_NEW_TOKENS = 64;
const ASR_LANGUAGE = "auto";
const LARGE_EXTERNAL_DATA_THRESHOLD = 256 * 1024 * 1024;
const ORT_PROVIDERS = {
  webgpu: ["webgpu", "wasm"],
  wasm: ["wasm"],
};

const state = {
  lib: null,
  pipeline: null,
  env: null,
  pipelines: new Map(),
  tokenizers: new Map(),
  modelConfigs: new Map(),
  ortSessions: new Map(),
  externalDataCache: new Map(),
  webgpu: {
    supportsFp16: null,
  },
  isRecording: false,
  isTranscribing: false,
  isCleaning: false,
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

async function resolveAsrDtype() {
  await ensureWebGpuFeatures();
  const deviceChoice = getDeviceOption();
  const wantsWebgpu = deviceChoice === "webgpu";
  if (wantsWebgpu && state.webgpu.supportsFp16 === false) {
    return "fp32";
  }
  if (deviceChoice === "wasm") {
    return "fp32";
  }
  return "fp16";
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

async function getPipeline(task, modelId, extraOptions = {}) {
  const device = getDeviceOption();
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

  return inputs;
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
    const dtype = await resolveAsrDtype();
    await ensureAsrModelFiles(modelId, dtype);
    const loadStart = performance.now();
    const asr = await getPipeline("automatic-speech-recognition", modelId, { dtype });
    logPerf("ASR pipeline load", performance.now() - loadStart);
    setStatus("Decoding audio ...");
    const decodeStart = performance.now();
    const { audio } = await decodeAudioBlob(blob);
    logPerf("ASR audio decode", performance.now() - decodeStart);
    setStatus("Transcribing ...");
    const inferStart = performance.now();
    const options = { chunk_length_s: 30, stride_length_s: 5 };
    if (ASR_LANGUAGE && ASR_LANGUAGE !== "auto") {
      options.language = ASR_LANGUAGE;
    }
    const result = await asr(audio, options);
    logPerf("ASR inference", performance.now() - inferStart);
    const text = typeof result === "string" ? result : result?.text || JSON.stringify(result, null, 2);
    transcriptEl.value = text;
    logPerf("ASR total", performance.now() - totalStart);
    setStatus("Transcription complete. Cleaning up ...");
    await runCleanup(text);
  } catch (err) {
    setStatus(`ASR failed: ${err.message}`);
  }
}

function computeCleanupMaxTokens(inputLength) {
  const scaled = Math.ceil(inputLength * 1.3);
  return Math.min(CLEANUP_MAX_NEW_TOKENS_CAP, Math.max(CLEANUP_MIN_NEW_TOKENS, scaled));
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
    const cleaned = await runLiquidCleanup(trimmed);
    cleanedTarget.value = cleaned;
    logPerf("Cleanup total", performance.now() - start);
    setStatus("Cleanup complete.");
  } catch (err) {
    setStatus(`Cleanup failed: ${err.message}`);
  } finally {
    state.isCleaning = false;
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

function pickBestMimeType() {
  const candidates = ["audio/webm;codecs=opus", "audio/webm", "audio/ogg;codecs=opus", "audio/ogg"];
  for (const type of candidates) {
    if (MediaRecorder.isTypeSupported(type)) return type;
  }
  return "";
}

async function startRecording() {
  if (state.isRecording || state.isTranscribing || state.isCleaning) return;
  state.isRecording = true;
  holdBtn.classList.add("is-recording");
  try {
    setStatus("Requesting microphone ...");
    state.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const mimeType = pickBestMimeType();
    state.chunks = [];
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

function stopRecording() {
  if (!state.isRecording || !state.recorder) return;
  setStatus("Stopping recording ...");
  try {
    state.recorder.stop();
  } catch (err) {
    setStatus(`Stop failed: ${err.message}`);
  }
}

loadLibBtn.addEventListener("click", () => {
  loadLibrary();
});

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
