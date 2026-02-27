const byId = (id) => document.getElementById(id);

const asrModelEl = byId("asr-model");
const asrLanguageModeEl = byId("asr-language-mode");
const asrInputModeEl = byId("asr-input-mode");
const asrBackendModeEl = byId("asr-backend-mode");
const loadAiBtn = byId("load-ai");
const pickFolderBtn = byId("pick-folder");
const loadSidecarBtn = byId("load-sidecar");
const recordToggleBtn = byId("record-toggle");
const recordTimerEl = byId("record-timer");
const recordMeterEl = byId("record-meter");
const recordMeterFillEl = byId("record-meter-fill");
const recordMeterValueEl = byId("record-meter-value");
const statusEl = byId("status");
const asrLanguageEl = byId("asr-language");
const transcriptEl = byId("transcript");
const segmentsEl = byId("segments");
const parseIdsBtn = byId("parse-ids");
const runExtractBtn = byId("run-extract");
const saveDraftBtn = byId("save-draft");
const applySidecarBtn = byId("apply-sidecar");
const idSummaryEl = byId("id-summary");
const draftListEl = byId("draft-list");
const logEl = byId("log");

const DEFAULTS = {
  transformersUrl: "../../AI/vendor/transformers.min.js",
  localModelPath: "../../AI/models/",
  liquidModel: "LiquidAI/LFM2.5-VL-1.6B-ONNX",
};
const SOURCE_SIDECAR_FILENAME = "project_sidecar.json";
const WORKING_SIDECAR_FILENAME = "project_sidecar.voice-report-working.json";

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
  fs: {
    dirHandle: null,
    sidecarDoc: null,
    sidecarProject: null,
    rowsById: new Map(),
    subIdToRowId: new Map(),
    locale: "de-CH",
  },
  recording: {
    isRecording: false,
    mode: null, // "global" | "card"
    rowId: null, // when mode === "card"
    buttonEl: null, // active card button, if any
    recorder: null,
    chunks: [],
    stream: null,
    startedAtMs: 0,
    timer: null,
    meterContext: null,
    meterSource: null,
    meterAnalyser: null,
    meterData: null,
    meterRaf: 0,
    meterLevel: 0,
    captureProcessor: null,
    captureGain: null,
    captureChunks: [],
    captureSampleRate: 0,
  },
  run: {
    parsed: null,
    draft: null,
  },
  ui: {
    cards: new Map(), // rowId -> card state
  },
};

function log(message) {
  if (!logEl) return;
  const line = `[${new Date().toISOString()}] ${message}`;
  logEl.textContent += `${line}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function logBlock(title, text, { maxChars = 12000 } = {}) {
  const raw = String(text ?? "");
  const trimmed = raw.trim();
  if (!trimmed) {
    log(`${title}: (empty)`);
    return;
  }
  let body = trimmed;
  if (Number.isFinite(maxChars) && maxChars > 0 && body.length > maxChars) {
    body = `${body.slice(0, maxChars)}\n… (truncated ${body.length - maxChars} chars)`;
  }
  const indented = body.split("\n").map((line) => `  ${line}`).join("\n");
  log(`${title}:\n${indented}`);
}

function setStatus(message) {
  if (statusEl) statusEl.textContent = message;
  log(message);
}

function formatMs(value) {
  if (!Number.isFinite(value)) return "n/a";
  if (value < 1000) return `${value.toFixed(1)}ms`;
  return `${(value / 1000).toFixed(2)}s`;
}

function formatTimerSeconds(seconds) {
  const s = Math.max(0, Math.floor(seconds));
  const mm = String(Math.floor(s / 60)).padStart(2, "0");
  const ss = String(s % 60).padStart(2, "0");
  return `${mm}:${ss}`;
}

function setRecordTimerText() {
  if (!recordTimerEl) return;
  if (!state.recording.isRecording) {
    recordTimerEl.textContent = "";
    return;
  }
  const elapsed = (performance.now() - state.recording.startedAtMs) / 1000;
  recordTimerEl.textContent = `Recording: ${formatTimerSeconds(elapsed)}`;
}

function setRecordMeter(levelPct, { live = false } = {}) {
  const clamped = Math.max(0, Math.min(100, Math.round(levelPct || 0)));
  if (recordMeterFillEl) {
    recordMeterFillEl.style.width = `${clamped}%`;
  }
  if (recordMeterValueEl) {
    recordMeterValueEl.textContent = `${clamped}%`;
  }
  if (recordMeterEl) {
    const track = recordMeterEl.querySelector(".meter-track");
    if (track) track.setAttribute("aria-valuenow", String(clamped));
    if (live) recordMeterEl.classList.add("live");
    else recordMeterEl.classList.remove("live");
  }
}

function logAsrBackend(pipe) {
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
  log(`ASR backend: ${JSON.stringify(info)}`);
}

function computeAudioSignalStats(audio) {
  if (!audio?.length) return { samples: 0, rms: 0, peak: 0 };
  let sum = 0;
  let peak = 0;
  for (let i = 0; i < audio.length; i += 1) {
    const v = Math.abs(audio[i] || 0);
    sum += v * v;
    if (v > peak) peak = v;
  }
  return {
    samples: audio.length,
    rms: Math.sqrt(sum / audio.length),
    peak,
  };
}

function downmixInputBuffer(inputBuffer) {
  const channels = inputBuffer.numberOfChannels || 1;
  const frames = inputBuffer.length || 0;
  const out = new Float32Array(frames);
  if (channels <= 1) {
    out.set(inputBuffer.getChannelData(0));
    return out;
  }
  for (let c = 0; c < channels; c += 1) {
    const data = inputBuffer.getChannelData(c);
    for (let i = 0; i < frames; i += 1) {
      out[i] += data[i];
    }
  }
  for (let i = 0; i < frames; i += 1) {
    out[i] /= channels;
  }
  return out;
}

function concatFloat32Chunks(chunks) {
  const total = chunks.reduce((sum, chunk) => sum + (chunk?.length || 0), 0);
  const out = new Float32Array(total);
  let offset = 0;
  chunks.forEach((chunk) => {
    if (!chunk || !chunk.length) return;
    out.set(chunk, offset);
    offset += chunk.length;
  });
  return out;
}

function resampleLinear(input, srcRate, dstRate) {
  if (!input?.length) return new Float32Array(0);
  if (!Number.isFinite(srcRate) || !Number.isFinite(dstRate) || srcRate <= 0 || dstRate <= 0) {
    return input;
  }
  if (srcRate === dstRate) return input;
  const ratio = srcRate / dstRate;
  const outLength = Math.max(1, Math.round(input.length / ratio));
  const out = new Float32Array(outLength);
  for (let i = 0; i < outLength; i += 1) {
    const pos = i * ratio;
    const lo = Math.floor(pos);
    const hi = Math.min(lo + 1, input.length - 1);
    const frac = pos - lo;
    out[i] = input[lo] * (1 - frac) + input[hi] * frac;
  }
  return out;
}

function buildCapturedAudioPayload() {
  const chunks = state.recording.captureChunks || [];
  const srcRate = Number(state.recording.captureSampleRate || 0);
  if (!chunks.length || !srcRate) return null;
  const merged = concatFloat32Chunks(chunks);
  const resampled = resampleLinear(merged, srcRate, 16000);
  return { audio: resampled, sampling_rate: 16000, srcRate, frames: merged.length };
}

async function stopRecordMeter() {
  if (state.recording.meterRaf) {
    window.cancelAnimationFrame(state.recording.meterRaf);
  }
  state.recording.meterRaf = 0;
  if (state.recording.meterSource) {
    try {
      state.recording.meterSource.disconnect();
    } catch {
      // no-op
    }
  }
  state.recording.meterSource = null;
  if (state.recording.captureProcessor) {
    try {
      state.recording.captureProcessor.disconnect();
    } catch {
      // no-op
    }
  }
  state.recording.captureProcessor = null;
  if (state.recording.captureGain) {
    try {
      state.recording.captureGain.disconnect();
    } catch {
      // no-op
    }
  }
  state.recording.captureGain = null;
  state.recording.meterAnalyser = null;
  state.recording.meterData = null;
  state.recording.meterLevel = 0;
  if (state.recording.meterContext) {
    try {
      await state.recording.meterContext.close();
    } catch {
      // no-op
    }
  }
  state.recording.meterContext = null;
  setRecordMeter(0, { live: false });
}

async function startRecordMeter(stream) {
  await stopRecordMeter();
  const AudioCtx = window.AudioContext || window.webkitAudioContext;
  if (!AudioCtx) {
    log("Mic level meter unavailable: AudioContext not supported.");
    return;
  }

  const meterContext = new AudioCtx();
  if (meterContext.state === "suspended") {
    try {
      await meterContext.resume();
    } catch {
      // ignore
    }
  }

  const meterSource = meterContext.createMediaStreamSource(stream);
  const meterAnalyser = meterContext.createAnalyser();
  meterAnalyser.fftSize = 1024;
  meterAnalyser.smoothingTimeConstant = 0.8;
  meterSource.connect(meterAnalyser);

  let captureProcessor = null;
  let captureGain = null;
  state.recording.captureChunks = [];
  state.recording.captureSampleRate = meterContext.sampleRate || 0;
  if (typeof meterContext.createScriptProcessor === "function") {
    captureProcessor = meterContext.createScriptProcessor(4096, 1, 1);
    captureGain = meterContext.createGain();
    captureGain.gain.value = 0;
    meterSource.connect(captureProcessor);
    captureProcessor.connect(captureGain);
    captureGain.connect(meterContext.destination);
    captureProcessor.onaudioprocess = (event) => {
      if (!state.recording.isRecording) return;
      const mixed = downmixInputBuffer(event.inputBuffer);
      state.recording.captureChunks.push(mixed);
    };
  } else {
    log("Live PCM capture unavailable: createScriptProcessor not supported. Use recorder decode input mode.");
  }

  const meterData = new Uint8Array(meterAnalyser.fftSize);
  state.recording.meterContext = meterContext;
  state.recording.meterSource = meterSource;
  state.recording.captureProcessor = captureProcessor;
  state.recording.captureGain = captureGain;
  state.recording.meterAnalyser = meterAnalyser;
  state.recording.meterData = meterData;
  state.recording.meterLevel = 0;

  const tick = () => {
    if (!state.recording.isRecording || !state.recording.meterAnalyser || !state.recording.meterData) {
      setRecordMeter(0, { live: false });
      return;
    }
    state.recording.meterAnalyser.getByteTimeDomainData(state.recording.meterData);
    let sum = 0;
    for (let i = 0; i < state.recording.meterData.length; i += 1) {
      const sample = (state.recording.meterData[i] - 128) / 128;
      sum += sample * sample;
    }
    const rms = Math.sqrt(sum / state.recording.meterData.length);
    const normalized = Math.min(1, rms * 3.2);
    const smoothed = (state.recording.meterLevel * 0.7) + (normalized * 0.3);
    state.recording.meterLevel = smoothed;
    setRecordMeter(smoothed * 100, { live: true });
    state.recording.meterRaf = window.requestAnimationFrame(tick);
  };

  setRecordMeter(0, { live: true });
  state.recording.meterRaf = window.requestAnimationFrame(tick);
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
  const defaultConfig = { version: "1.23.2", base: "../../AI/vendor/ort-1.23.2/", bundle: "webgpu" };
  if (window.__ortConfig && window.__ortConfig.version && window.__ortConfig.base) {
    return window.__ortConfig;
  }
  return defaultConfig;
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

function getDeviceOption() {
  if ("gpu" in navigator) return "webgpu";
  return "wasm";
}

function resolveOrtProvider() {
  const deviceChoice = getDeviceOption();
  if (deviceChoice === "wasm") return ORT_PROVIDERS.wasm;
  return ORT_PROVIDERS.webgpu;
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
    setStatus("AI library loaded. Ready.");
  } catch (err) {
    setStatus(`AI library load failed: ${err.message}`);
    throw err;
  }
}

async function ensureLibraryLoaded() {
  if (state.pipeline && state.env) return;
  await loadLibrary();
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

async function pickAvailableAsrDtype(modelId) {
  const preferred = await resolveAsrDtype();
  await ensureAsrModelFiles(modelId, preferred);
  return preferred;
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
  const channels = Math.max(1, Number(buffer.numberOfChannels || 1));
  const frames = Number(buffer.length || 0);
  const mixed = new Float32Array(frames);
  for (let c = 0; c < channels; c += 1) {
    const channelData = buffer.getChannelData(c);
    for (let i = 0; i < frames; i += 1) {
      mixed[i] += channelData[i];
    }
  }
  if (channels > 1) {
    for (let i = 0; i < frames; i += 1) {
      mixed[i] /= channels;
    }
  }
  const audio = mixed;
  audioCtx.close();
  return { audio, sampling_rate: targetRate };
}

async function getPipeline(task, modelId, extraOptions = {}) {
  const defaultDevice = getDeviceOption();
  const options = { ...extraOptions };
  const preferAutoDevice = options.deviceAuto === true;
  delete options.deviceAuto;
  if (preferAutoDevice) {
    delete options.device;
  } else if (!options.device && defaultDevice) {
    options.device = defaultDevice;
  }
  const key = JSON.stringify({ task, modelId, options, deviceMode: preferAutoDevice ? "auto" : "forced" });
  if (state.pipelines.has(key)) return state.pipelines.get(key);
  await ensureLibraryLoaded();
  applyEnv();
  const pipe = await state.pipeline(task, modelId, options);
  state.pipelines.set(key, pipe);
  return pipe;
}

function resolveAsrBackendMode() {
  const mode = String(asrBackendModeEl?.value || "auto").trim();
  if (mode === "auto" || mode === "wasm" || mode === "webgpu") return mode;
  throw new Error(`Unsupported ASR backend mode: ${mode}`);
}

function buildAsrPipelineOptions(dtype) {
  const backendMode = resolveAsrBackendMode();
  const options = { dtype };
  if (backendMode === "auto") {
    options.deviceAuto = true;
  } else {
    options.device = backendMode;
  }
  return { backendMode, options };
}

function pickAudioMimeType() {
  const types = [
    "audio/webm;codecs=opus",
    "audio/webm",
    "audio/ogg;codecs=opus",
    "audio/ogg",
  ];
  for (const type of types) {
    try {
      if (MediaRecorder.isTypeSupported(type)) return type;
    } catch {
      // ignore
    }
  }
  return "";
}

function setButtonRecordingState(buttonEl, isRecording, { recordingText, idleText } = {}) {
  if (!buttonEl) return;
  if (isRecording) {
    buttonEl.classList.add("is-recording");
    if (recordingText) buttonEl.textContent = recordingText;
    return;
  }
  buttonEl.classList.remove("is-recording");
  if (idleText) buttonEl.textContent = idleText;
}

async function startRecording({ mode = "global", rowId = null, buttonEl = null } = {}) {
  if (state.recording.isRecording) return;
  if (!navigator.mediaDevices?.getUserMedia) {
    throw new Error("getUserMedia not available in this browser.");
  }
  const mimeType = pickAudioMimeType();
  const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  state.recording.mode = mode;
  state.recording.rowId = rowId;
  state.recording.buttonEl = buttonEl;
  state.recording.stream = stream;
  state.recording.chunks = [];
  state.recording.isRecording = true;
  try {
    await startRecordMeter(stream);
    const recorder = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
    state.recording.recorder = recorder;
    recorder.addEventListener("dataavailable", (event) => {
      if (event.data && event.data.size > 0) {
        state.recording.chunks.push(event.data);
      }
    });
    recorder.start();
    state.recording.startedAtMs = performance.now();
    state.recording.timer = window.setInterval(setRecordTimerText, 250);
    setRecordTimerText();
  } catch (err) {
    await stopRecordMeter();
    stream.getTracks().forEach((t) => t.stop());
    state.recording.isRecording = false;
    state.recording.mode = null;
    state.recording.rowId = null;
    state.recording.buttonEl = null;
    state.recording.stream = null;
    state.recording.recorder = null;
    state.recording.chunks = [];
    throw err;
  }

  if (mode === "global") {
    setButtonRecordingState(recordToggleBtn, true, { recordingText: "Stop recording" });
  } else if (mode === "card") {
    setButtonRecordingState(buttonEl, true, { recordingText: "Stop mic" });
    if (recordToggleBtn) recordToggleBtn.disabled = true;
  }

  setStatus(mode === "card" && rowId ? `Recording (${rowId})...` : "Recording...");
}

async function stopRecording() {
  if (!state.recording.isRecording) return null;
  const recorder = state.recording.recorder;
  const stream = state.recording.stream;
  const mode = state.recording.mode;
  const rowId = state.recording.rowId;
  const buttonEl = state.recording.buttonEl;
  if (!recorder) {
    await stopRecordMeter();
    if (stream) {
      stream.getTracks().forEach((t) => t.stop());
    }
    state.recording.isRecording = false;
    state.recording.mode = null;
    state.recording.rowId = null;
    state.recording.buttonEl = null;
    state.recording.stream = null;
    state.recording.chunks = [];
    state.recording.captureChunks = [];
    state.recording.captureSampleRate = 0;
    setRecordTimerText();
    ensureButtons();
    return null;
  }
  const stopped = new Promise((resolve) => {
    recorder.addEventListener("stop", () => resolve());
  });
  recorder.stop();
  await stopped;
  await stopRecordMeter();
  if (stream) {
    stream.getTracks().forEach((t) => t.stop());
  }
  const blob = new Blob(state.recording.chunks, { type: recorder.mimeType || "audio/webm" });
  const capturedAudio = buildCapturedAudioPayload();
  state.recording.isRecording = false;
  state.recording.mode = null;
  state.recording.rowId = null;
  state.recording.buttonEl = null;
  state.recording.recorder = null;
  state.recording.stream = null;
  state.recording.chunks = [];
  state.recording.captureChunks = [];
  state.recording.captureSampleRate = 0;
  if (state.recording.timer) window.clearInterval(state.recording.timer);
  state.recording.timer = null;
  setRecordTimerText();

  if (mode === "global") {
    setButtonRecordingState(recordToggleBtn, false, { idleText: "Start recording" });
    setStatus("Recording stopped. Ready to transcribe.");
  } else if (mode === "card") {
    setButtonRecordingState(buttonEl, false, { idleText: "Mic this card" });
    setStatus(rowId ? `Recording stopped (${rowId}).` : "Recording stopped.");
  } else {
    setStatus("Recording stopped.");
  }
  ensureButtons();
  return { blob, capturedAudio };
}

async function runAsrFromBlob(blob, captureAudio = null) {
  const modelId = asrModelEl.value.trim();
  if (!modelId) {
    setStatus("Pick an ASR model.");
    return "";
  }
  const languageMode = String(asrLanguageModeEl?.value || "auto").trim();
  if (!state.fs.sidecarDoc && languageMode === "from-locale") {
    throw new Error("Load sidecar before recording so ASR language can be resolved from locale.");
  }
  const totalStart = performance.now();
  setStatus(`Loading ASR pipeline (${modelId}) ...`);
  const dtype = await pickAvailableAsrDtype(modelId);
  const { backendMode, options: pipelineOptions } = buildAsrPipelineOptions(dtype);
  const asr = await getPipeline("automatic-speech-recognition", modelId, pipelineOptions);
  log(`ASR backend mode: ${backendMode}`);
  logAsrBackend(asr);
  const inputMode = String(asrInputModeEl?.value || "blob").trim();
  let audio;
  if (inputMode === "pcm") {
    if (!captureAudio?.audio?.length) {
      throw new Error("Live PCM mode selected, but no PCM capture is available.");
    }
    audio = captureAudio.audio;
    log(
      `ASR using live PCM capture: samples=${audio.length}, srcRate=${captureAudio.srcRate}, targetRate=${captureAudio.sampling_rate}`
    );
  } else {
    setStatus("Decoding audio ...");
    const decodeStart = performance.now();
    const decoded = await decodeAudioBlob(blob);
    audio = decoded.audio;
    log(`ASR audio decode: ${formatMs(performance.now() - decodeStart)}`);
  }
  const stats = computeAudioSignalStats(audio);
  log(`ASR audio stats: samples=${stats.samples}, rms=${stats.rms.toFixed(6)}, peak=${stats.peak.toFixed(6)}`);
  if (stats.rms < 0.002 && stats.peak < 0.02) {
    throw new Error("Microphone signal is extremely low. Select the correct mic/input level and re-record.");
  }
  const durationSec = audio.length / 16000;
  setStatus("Transcribing ...");
  const inferStart = performance.now();
  const language = resolveAsrLanguageSetting();
  const options = { task: "transcribe" };
  if (durationSec > 30) {
    options.chunk_length_s = 30;
    options.stride_length_s = 5;
    log(`ASR chunking enabled (duration=${durationSec.toFixed(2)}s).`);
  } else {
    log(`ASR single-pass decode (duration=${durationSec.toFixed(2)}s).`);
  }
  if (language) {
    options.language = language;
    log(`ASR language fixed to ${language} (locale=${state.fs.locale || "n/a"}).`);
  } else {
    log("ASR language mode set to auto-detect.");
  }
  const result = await asr(audio, options);
  log(`ASR inference: ${formatMs(performance.now() - inferStart)}`);
  const text = typeof result === "string" ? result : result?.text || JSON.stringify(result, null, 2);
  logBlock("ASR transcript (raw)", text, { maxChars: 4000 });
  log(`ASR total: ${formatMs(performance.now() - totalStart)}`);
  setStatus("Transcription complete.");
  return text;
}

function stripTranscriptArtifacts(text) {
  return String(text || "").replace(/>+/g, "").trim();
}

function normalizeNewlines(value) {
  return String(value || "").replace(/\r\n/g, "\n");
}

function resolveAsrLanguage(locale) {
  const base = String(locale || "").trim().toLowerCase();
  if (!base) throw new Error("Missing locale. Load sidecar before ASR.");
  if (base.startsWith("fr")) return "fr";
  if (base.startsWith("de")) return "de";
  if (base.startsWith("it")) return "it";
  if (base.startsWith("en")) return "en";
  throw new Error(`Unsupported locale for ASR language routing: ${locale}`);
}

function resolveAsrLanguageSetting() {
  const mode = String(asrLanguageModeEl?.value || "auto").trim();
  if (mode === "auto") return null;
  if (mode === "from-locale") return resolveAsrLanguage(state.fs.locale);
  if (mode === "de" || mode === "fr" || mode === "it" || mode === "en") return mode;
  throw new Error(`Unsupported ASR language mode: ${mode}`);
}

function updateAsrLanguageBadge() {
  if (!asrLanguageEl) return;
  const mode = String(asrLanguageModeEl?.value || "auto").trim();
  if (mode === "auto") {
    asrLanguageEl.textContent = "ASR language: auto-detect";
    asrLanguageEl.classList.remove("warn");
    return;
  }
  if (mode === "de" || mode === "fr" || mode === "it" || mode === "en") {
    asrLanguageEl.textContent = `ASR language: ${mode} (manual)`;
    asrLanguageEl.classList.remove("warn");
    return;
  }
  const locale = String(state.fs.locale || "").trim();
  if (!locale || !state.fs.sidecarDoc) {
    asrLanguageEl.textContent = "ASR language: load sidecar";
    asrLanguageEl.classList.remove("warn");
    return;
  }
  try {
    const lang = resolveAsrLanguage(locale);
    asrLanguageEl.textContent = `ASR language: ${lang} (from ${locale})`;
    asrLanguageEl.classList.remove("warn");
  } catch (err) {
    asrLanguageEl.textContent = `ASR language: unsupported (${locale})`;
    asrLanguageEl.classList.add("warn");
  }
}

function isRejectedFallbackText(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) return false;
  const patterns = [
    /\binsufficient information\b/,
    /\bnot enough information\b/,
    /\bnot enough context\b/,
    /\bcannot determine\b/,
    /\bcan't determine\b/,
    /\bcannot assess\b/,
    /\bunable to assess\b/,
    /\binformation insuffisante\b/,
    /\bpas suffisamment d(?:'|\u2019)information\b/,
    /\bimpossible de déterminer\b/,
    /\binformazioni insufficienti\b/,
    /\bnon abbastanza informazioni\b/,
    /\bimpossibile determinare\b/,
    /\bnicht genug information(?:en)?\b/,
    /\bnicht ausreichend(?:e)? information(?:en)?\b/,
    /\bnicht beurteilbar\b/,
    /\bkeine ausreichenden information(?:en)?\b/,
    /\bn\/a\b/,
    /\bna\b/,
  ];
  return patterns.some((pattern) => pattern.test(text));
}

function normalizeExtractField(value, { rowId, field }) {
  const text = String(value || "").trim();
  if (!text) return "";
  if (isRejectedFallbackText(text)) {
    log(`Rejected ${field} boilerplate text for ${rowId}.`);
    return "";
  }
  return text;
}

function evaluateTranscriptQuality(text) {
  const t = String(text || "").trim();
  if (!t) return { ok: false, reason: "transcript is empty", metrics: null };
  if (/^[^\p{L}\p{N}]+$/u.test(t)) {
    return { ok: false, reason: "transcript does not contain words", metrics: null };
  }
  const normalized = t
    .toLowerCase()
    .replace(/[^\p{L}\p{N}'-]+/gu, " ")
    .trim();
  const words = normalized ? normalized.split(/\s+/).filter(Boolean) : [];
  if (t.length < 6 || words.length < 2) {
    return { ok: false, reason: "transcript is too short", metrics: { words: words.length, chars: t.length } };
  }

  const wordCounts = new Map();
  for (const word of words) {
    wordCounts.set(word, (wordCounts.get(word) || 0) + 1);
  }
  const uniqueWordCount = wordCounts.size;
  const maxWordCount = Math.max(...Array.from(wordCounts.values()));
  const uniqueWordRatio = uniqueWordCount / words.length;
  const topWordShare = maxWordCount / words.length;

  let topBigramShare = 0;
  if (words.length > 1) {
    const bigramCounts = new Map();
    for (let i = 0; i < words.length - 1; i += 1) {
      const key = `${words[i]} ${words[i + 1]}`;
      bigramCounts.set(key, (bigramCounts.get(key) || 0) + 1);
    }
    const maxBigramCount = Math.max(...Array.from(bigramCounts.values()));
    topBigramShare = maxBigramCount / (words.length - 1);
  }

  const metrics = {
    words: words.length,
    chars: t.length,
    uniqueWordRatio,
    topWordShare,
    topBigramShare,
  };

  if (words.length >= 20) {
    const repetitiveWordPattern = uniqueWordRatio < 0.28 && topWordShare > 0.2;
    const repetitiveBigramPattern = topBigramShare > 0.24;
    if (repetitiveWordPattern || repetitiveBigramPattern) {
      return {
        ok: false,
        reason: "transcript is highly repetitive and likely invalid ASR output",
        metrics,
      };
    }
  }

  return { ok: true, reason: "", metrics };
}

function assertTranscriptQuality(text, { contextLabel = "ASR transcript" } = {}) {
  const quality = evaluateTranscriptQuality(text);
  if (quality.metrics) {
    const m = quality.metrics;
    const ratio = Number.isFinite(m.uniqueWordRatio) ? m.uniqueWordRatio.toFixed(3) : "n/a";
    const topWord = Number.isFinite(m.topWordShare) ? m.topWordShare.toFixed(3) : "n/a";
    const topBigram = Number.isFinite(m.topBigramShare) ? m.topBigramShare.toFixed(3) : "n/a";
    log(
      `ASR transcript quality: words=${m.words}, chars=${m.chars}, uniqueRatio=${ratio}, topWordShare=${topWord}, topBigramShare=${topBigram}`
    );
  }
  if (!quality.ok) {
    throw new Error(`${contextLabel} failed quality check: ${quality.reason}.`);
  }
}

function safeSlugTimestamp() {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

async function readFileText(fileHandle) {
  const file = await fileHandle.getFile();
  return await file.text();
}

async function writeJsonToHandle(fileHandle, payload) {
  const writable = await fileHandle.createWritable();
  await writable.write(JSON.stringify(payload, null, 2));
  await writable.close();
}

async function writeTextToHandle(fileHandle, text) {
  const writable = await fileHandle.createWritable();
  await writable.write(text);
  await writable.close();
}

async function tryGetFileHandle(dirHandle, filename) {
  try {
    return await dirHandle.getFileHandle(filename);
  } catch (err) {
    if (err?.name === "NotFoundError") return null;
    throw err;
  }
}

async function readSidecarForEditor(dirHandle) {
  const workingHandle = await tryGetFileHandle(dirHandle, WORKING_SIDECAR_FILENAME);
  if (workingHandle) {
    return {
      handle: workingHandle,
      filename: WORKING_SIDECAR_FILENAME,
      text: await readFileText(workingHandle),
      source: "working",
    };
  }
  const sourceHandle = await dirHandle.getFileHandle(SOURCE_SIDECAR_FILENAME);
  return {
    handle: sourceHandle,
    filename: SOURCE_SIDECAR_FILENAME,
    text: await readFileText(sourceHandle),
    source: "source",
  };
}

async function ensureWritableWorkingSidecar(dirHandle) {
  const existing = await tryGetFileHandle(dirHandle, WORKING_SIDECAR_FILENAME);
  if (existing) {
    return {
      handle: existing,
      filename: WORKING_SIDECAR_FILENAME,
      text: await readFileText(existing),
      seeded: false,
    };
  }
  const sourceHandle = await dirHandle.getFileHandle(SOURCE_SIDECAR_FILENAME);
  const sourceText = await readFileText(sourceHandle);
  const workingHandle = await dirHandle.getFileHandle(WORKING_SIDECAR_FILENAME, { create: true });
  await writeTextToHandle(workingHandle, sourceText);
  return {
    handle: workingHandle,
    filename: WORKING_SIDECAR_FILENAME,
    text: sourceText,
    seeded: true,
  };
}

function extractReportProject(doc) {
  if (!doc || typeof doc !== "object") return null;
  if (doc.report && doc.report.project && doc.report.project.chapters) return doc.report.project;
  if (doc.chapters) return doc;
  return null;
}

function toText(value) {
  if (Array.isArray(value)) return value.join("\n");
  if (value == null) return "";
  return String(value);
}

function getRowTitle(row) {
  return String(row?.titleOverride || row?.id || "").trim();
}

function getRowCurrentFinding(row) {
  const ws = row?.workstate || {};
  return toText(ws.findingText);
}

function getRowCurrentRecommendation(row) {
  const ws = row?.workstate || {};
  return toText(ws.recommendationText);
}

function buildRowIndex(project) {
  const rowsById = new Map();
  const subIdToRowId = new Map();
  (project?.chapters || []).forEach((chapter) => {
    (chapter?.rows || []).forEach((row) => {
      if (!row || typeof row !== "object") return;
      if (row.kind === "section") return;
      const id = String(row.id || "").trim();
      if (!id) return;
      rowsById.set(id, row);
      const items = row.customer?.items;
      if (Array.isArray(items)) {
        items.forEach((item) => {
          if (!item || typeof item !== "object") return;
          const subId = String(item.id || "").trim();
          if (subId) subIdToRowId.set(subId, id);
        });
      }
    });
  });
  return { rowsById, subIdToRowId };
}

function seedCardsFromSidecarProject(project) {
  state.ui.cards.clear();
  (project?.chapters || []).forEach((chapter) => {
    (chapter?.rows || []).forEach((row) => {
      if (!row || typeof row !== "object") return;
      if (row.kind === "section") return;
      const id = String(row.id || "").trim();
      if (!id) return;
      upsertCard(id, {
        title: getRowTitle(row),
        currentFinding: getRowCurrentFinding(row),
        currentRecommendation: getRowCurrentRecommendation(row),
      });
    });
  });
}

function resolveLocaleKey(locale) {
  const base = String(locale || "de-CH").toLowerCase();
  if (base.startsWith("fr")) return "fr";
  if (base.startsWith("it")) return "it";
  return "de";
}

function isValidIdCandidate(value) {
  return /^[0-9]{1,2}(?:\.[0-9]{1,2}){1,3}[a-z]?$/.test(String(value || "").trim());
}

function buildMentionCandidate(parts, suffixLetter = "") {
  const nums = parts.filter((p) => p != null && String(p).trim() !== "").map((p) => String(p).trim());
  if (nums.length < 2) return "";
  const id = `${nums.join(".")}${suffixLetter || ""}`;
  return isValidIdCandidate(id) ? id : "";
}

function mapDeNumberWord(token) {
  const t = String(token || "").toLowerCase();
  const table = {
    ein: 1,
    eins: 1,
    zwei: 2,
    drei: 3,
    vier: 4,
    fuenf: 5,
    fünf: 5,
    funf: 5,
    sechs: 6,
    sieben: 7,
    acht: 8,
    neun: 9,
    zehn: 10,
    elf: 11,
    zwoelf: 12,
    zwölf: 12,
    dreizehn: 13,
    vierzehn: 14,
  };
  if (Object.prototype.hasOwnProperty.call(table, t)) return table[t];
  if (/^\d{1,2}$/.test(t)) return Number(t);
  return null;
}

function mapFrNumberWord(token) {
  const t = String(token || "").toLowerCase();
  const table = {
    un: 1,
    une: 1,
    deux: 2,
    trois: 3,
    quatre: 4,
    cinq: 5,
    six: 6,
    sept: 7,
    huit: 8,
    neuf: 9,
    dix: 10,
    onze: 11,
    douze: 12,
    treize: 13,
    quatorze: 14,
  };
  if (Object.prototype.hasOwnProperty.call(table, t)) return table[t];
  if (/^\d{1,2}$/.test(t)) return Number(t);
  return null;
}

function extractIdMentions(text, { rowsById, subIdToRowId, localeKey }) {
  const src = String(text || "");
  const candidates = [];

  const addIfResolvable = (raw, start, end, candidateId) => {
    if (!candidateId) return;
    const id = String(candidateId);
    const rowId = rowsById.has(id) ? id : (subIdToRowId.get(id) || null);
    if (!rowId) return;
    const kind = rowsById.has(id) ? "row" : "sub";
    candidates.push({ raw, start, end, id, rowId, kind });
  };

  // 1) Numeric dotted IDs like 1.1.1 or 2.1.2b
  const dotted = /\b(\d{1,2})(?:\s*\.\s*(\d{1,2}))(?:\s*\.\s*(\d{1,2}))?(?:\s*\.\s*(\d{1,2}))?(?:\s*([a-z]))?\b/gi;
  for (let match; (match = dotted.exec(src)); ) {
    const raw = match[0];
    const start = match.index;
    const end = start + raw.length;
    const suffix = (match[5] || "").trim();
    const candidate = buildMentionCandidate([match[1], match[2], match[3], match[4]], suffix);
    addIfResolvable(raw, start, end, candidate);
  }

  // 2) Numeric spaced IDs like "1 1 3" (validated against known ids)
  const spaced = /\b(\d{1,2})\s+(\d{1,2})(?:\s+(\d{1,2}))?(?:\s+(\d{1,2}))?(?:\s*([a-z]))?\b/gi;
  for (let match; (match = spaced.exec(src)); ) {
    const raw = match[0];
    const start = match.index;
    const end = start + raw.length;
    const suffix = (match[5] || "").trim();
    const candidate = buildMentionCandidate([match[1], match[2], match[3], match[4]], suffix);
    addIfResolvable(raw, start, end, candidate);
  }

  // 3) DE word-based patterns like "eins eins drei" or "elf sieben"
  if (localeKey === "de") {
    const words = /\b((?:ein|eins|zwei|drei|vier|fuenf|fünf|funf|sechs|sieben|acht|neun|zehn|elf|zwoelf|zwölf|dreizehn|vierzehn|\d{1,2}))(?:\s+(?:punkt|\.)\s+|\s+)((?:ein|eins|zwei|drei|vier|fuenf|fünf|funf|sechs|sieben|acht|neun|zehn|elf|zwoelf|zwölf|dreizehn|vierzehn|\d{1,2}))(?:\s+(?:punkt|\.)\s+|\s+)?((?:ein|eins|zwei|drei|vier|fuenf|fünf|funf|sechs|sieben|acht|neun|zehn|elf|zwoelf|zwölf|dreizehn|vierzehn|\d{1,2}))?(?:\s+(?:punkt|\.)\s+|\s+)?((?:ein|eins|zwei|drei|vier|fuenf|fünf|funf|sechs|sieben|acht|neun|zehn|elf|zwoelf|zwölf|dreizehn|vierzehn|\d{1,2}))?\b/gi;
    for (let match; (match = words.exec(src)); ) {
      const raw = match[0];
      const start = match.index;
      const end = start + raw.length;
      const parts = [match[1], match[2], match[3], match[4]].map(mapDeNumberWord).filter((v) => v != null);
      if (parts.length < 2) continue;
      // Try 2..4 segments progressively (validate against known ids)
      for (let n = 2; n <= Math.min(4, parts.length); n += 1) {
        const candidate = buildMentionCandidate(parts.slice(0, n).map(String), "");
        addIfResolvable(raw, start, end, candidate);
      }
    }
  }

  // 4) FR word-based patterns like "un un trois" or "onze sept"
  if (localeKey === "fr") {
    const words = /\b((?:un|une|deux|trois|quatre|cinq|six|sept|huit|neuf|dix|onze|douze|treize|quatorze|\d{1,2}))(?:\s+(?:point|\.)\s+|\s+)((?:un|une|deux|trois|quatre|cinq|six|sept|huit|neuf|dix|onze|douze|treize|quatorze|\d{1,2}))(?:\s+(?:point|\.)\s+|\s+)?((?:un|une|deux|trois|quatre|cinq|six|sept|huit|neuf|dix|onze|douze|treize|quatorze|\d{1,2}))?(?:\s+(?:point|\.)\s+|\s+)?((?:un|une|deux|trois|quatre|cinq|six|sept|huit|neuf|dix|onze|douze|treize|quatorze|\d{1,2}))?\b/gi;
    for (let match; (match = words.exec(src)); ) {
      const raw = match[0];
      const start = match.index;
      const end = start + raw.length;
      const parts = [match[1], match[2], match[3], match[4]].map(mapFrNumberWord).filter((v) => v != null);
      if (parts.length < 2) continue;
      for (let n = 2; n <= Math.min(4, parts.length); n += 1) {
        const candidate = buildMentionCandidate(parts.slice(0, n).map(String), "");
        addIfResolvable(raw, start, end, candidate);
      }
    }
  }

  // Resolve overlaps: prefer earliest, then prefer longer match, then prefer direct row ids.
  candidates.sort((a, b) => {
    if (a.start !== b.start) return a.start - b.start;
    if (a.kind !== b.kind) return a.kind === "row" ? -1 : 1;
    return (b.end - b.start) - (a.end - a.start);
  });

  const chosen = [];
  let lastEnd = -1;
  candidates.forEach((cand) => {
    if (cand.start < lastEnd) return;
    chosen.push(cand);
    lastEnd = cand.end;
  });

  return chosen;
}

function segmentByMentions(text, mentions) {
  const src = String(text || "");
  const segments = [];
  for (let i = 0; i < mentions.length; i += 1) {
    const m = mentions[i];
    const next = mentions[i + 1];
    const start = m.end;
    const end = next ? next.start : src.length;
    const raw = src.slice(start, end);
    const cleaned = raw.replace(/^[\s,;:.-]+/, "").trim();
    segments.push({
      mentionedId: m.id,
      rowId: m.rowId,
      kind: m.kind,
      note: cleaned,
      rawNote: raw,
      start,
      end,
    });
  }
  return segments;
}

function groupSegmentsByRowId(segments) {
  const map = new Map();
  segments.forEach((seg) => {
    if (!seg.note) return;
    const entry = map.get(seg.rowId) || { rowId: seg.rowId, mentionedIds: new Set(), notes: [] };
    entry.mentionedIds.add(seg.mentionedId);
    entry.notes.push(seg.note);
    map.set(seg.rowId, entry);
  });
  return Array.from(map.values()).map((entry) => ({
    rowId: entry.rowId,
    mentionedIds: Array.from(entry.mentionedIds),
    note: entry.notes.join("\n\n"),
    notes: entry.notes,
  }));
}

function buildSortedSegmentsText(groups, { rowsById }) {
  const sorted = [...groups].sort((a, b) => a.rowId.localeCompare(b.rowId, "de", { numeric: true }));
  const parts = [];
  sorted.forEach((g) => {
    const row = rowsById.get(g.rowId);
    const title = row ? getRowTitle(row) : "";
    parts.push(`### ${g.rowId}${title ? ` — ${title}` : ""}`);
    parts.push(g.note);
    parts.push("");
  });
  return parts.join("\n");
}

function computeExtractMaxTokens(inputLength) {
  // Keep bounded and shorter for small inputs to reduce rambly outputs.
  const cap = 256;
  const min = 32;
  const scaled = Math.ceil(inputLength * 0.6);
  return Math.min(cap, Math.max(min, scaled));
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

function resolveInputNames(inputMeta, inputNames, defaultNames) {
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
        : numericOnly && defaultNames?.length
          ? defaultNames
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
    namesFromSession.some((name) => typeof name === "string" && !/^\d+$/.test(name)) ||
    useNames.some((name) => typeof name === "string" && !/^\d+$/.test(name));
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
    if (hasNonNumericNames && typeof name === "string" && /^\d+$/.test(name)) {
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

function decodeTokens(tokenizer, tokenIds) {
  if (!tokenizer || typeof tokenizer.decode !== "function") return tokenIds.join(" ");
  return tokenizer.decode(tokenIds, { skip_special_tokens: true });
}

async function ensureChatTemplate(modelId, tokenizer) {
  if (tokenizer?.chat_template) return tokenizer.chat_template;
  const base = getLocalModelBase(modelId);
  try {
    const response = await fetch(`${base}chat_template.jinja`, { cache: "no-store" });
    if (!response.ok) return null;
    const template = await response.text();
    if (template && tokenizer) tokenizer.chat_template = template;
    return template || null;
  } catch {
    return null;
  }
}

async function getTokenizer(modelId) {
  if (state.tokenizers.has(modelId)) return state.tokenizers.get(modelId);
  await ensureLibraryLoaded();
  applyEnv();
  const tokenizer = await state.lib.AutoTokenizer.from_pretrained(modelId, { local_files_only: true });
  state.tokenizers.set(modelId, tokenizer);
  return tokenizer;
}

async function getModelConfig(modelId) {
  if (state.modelConfigs.has(modelId)) return state.modelConfigs.get(modelId);
  await ensureLibraryLoaded();
  applyEnv();
  const config = await state.lib.AutoConfig.from_pretrained(modelId, { local_files_only: true });
  state.modelConfigs.set(modelId, config);
  return config;
}

async function tokenizeQuestion(tokenizer, prompt) {
  if (!tokenizer || typeof tokenizer.encode !== "function") return [];
  const encoded = await tokenizer.encode(prompt);
  if (Array.isArray(encoded)) return encoded;
  if (encoded?.input_ids) return Array.from(encoded.input_ids);
  return [];
}

async function buildExtractPrompt(tokenizer, modelId, localeKey, items) {
  const localeName = localeKey === "fr" ? "French" : localeKey === "it" ? "Italian" : "German";
  const intro =
    `You are writing an audit report. For each item below, extract two fields in ${localeName}:\\n` +
    `- FINDING: a concise negative finding statement.\\n` +
    `- RECOMMENDATION: a concise action recommendation.\\n\\n` +
    `Rules:\\n` +
    `- Use the same language as the NOTE.\\n` +
    `- Do not invent details. Use only what is in the NOTE.\\n` +
    `- Do not invent IDs. Use exactly the given IDs.\\n` +
    `- If the NOTE does not support a field, leave that field empty.\\n` +
    `- Never output placeholder phrases like 'insufficient information', 'information insuffisante', 'informazioni insufficienti', 'nicht beurteilbar', 'N/A'.\\n` +
    `- Output strictly in this repeated format:\\n` +
    `ID: <id>\\nFINDING: <text or empty>\\nRECOMMENDATION: <text or empty>\\n\\n`;

  const body = items.map((it) => (
    `ID: ${it.rowId}\\nTITLE: ${it.title || ""}\\nNOTE: ${it.note || ""}\\n`
  )).join("\\n");

  const promptText = `${intro}Items:\\n\\n${body}`;
  const messages = [{ role: "user", content: promptText }];

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

async function runLiquidExtract(items) {
  await ensureLibraryLoaded();
  await ensureOrtLoaded();
  await ensureWebGpuFeatures();
  const modelId = DEFAULTS.liquidModel;
  const base = getLocalModelBase(modelId);
  const providers = state.webgpu.supportsFp16 === false ? ORT_PROVIDERS.wasm : resolveOrtProvider();

  const tokenizer = await getTokenizer(modelId);
  const config = await getModelConfig(modelId);
  const eosTokenId = config?.text_config?.eos_token_id ?? config?.eos_token_id ?? null;
  const localeKey = resolveLocaleKey(state.fs.locale);

  const prompt = await buildExtractPrompt(tokenizer, modelId, localeKey, items);
  const inputIds = (await tokenizeQuestion(tokenizer, prompt)).map((value) => Number(value));
  const attentionMask = buildAttentionMaskFromLength(inputIds.length);
  const positionIds = buildPositionIdsFromLength(inputIds.length);

  const embedSession = await loadOrtSession(`${base}onnx/embed_tokens_fp16.onnx`, providers);
  const decoderSession = await loadOrtSession(`${base}onnx/decoder_q4.onnx`, providers);

  const tokenInputs = buildLiquidTokenInputs(
    embedSession.inputMetadata,
    embedSession.inputNames,
    inputIds,
    attentionMask,
    positionIds
  );

  const embedStart = performance.now();
  const tokenResult = await embedSession.run(tokenInputs);
  let currentEmbeds = Object.values(tokenResult)[0];
  log(`Liquid token embedding: ${formatMs(performance.now() - embedStart)}`);

  const maxNewTokens = computeExtractMaxTokens(inputIds.length);
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
    const result = await decoderSession.run(decoderInputs);
    updateLiquidCache(cache, result);
    const logitsTensor = resolveLogitsOutput(result);
    if (!logitsTensor) {
      throw new Error("Decoder did not return logits.");
    }
    const nextId = argmaxLogits(logitsTensor);
    generated.push(nextId);
    if (eosTokenId !== null && Number(nextId) === Number(eosTokenId)) break;

    const nextTokenInputs = buildLiquidTokenInputs(
      embedSession.inputMetadata,
      embedSession.inputNames,
      [nextId]
    );
    const nextTokenResult = await embedSession.run(nextTokenInputs);
    currentEmbeds = Object.values(nextTokenResult)[0];
  }

  log(`Liquid decode: ${formatMs(performance.now() - generateStart)} (tokens=${generated.length})`);
  return {
    prompt,
    generated,
    text: decodeTokens(tokenizer, generated).trim(),
  };
}

function parseLiquidExtractOutput(text, rowIds) {
  const src = normalizeNewlines(text || "");
  const blocks = src.split(/\n(?=ID:\s*)/g).map((b) => b.trim()).filter(Boolean);
  const out = new Map();
  const requested = new Set(rowIds);
  const unexpectedIds = new Set();
  blocks.forEach((block) => {
    const idMatch = block.match(/^ID:\s*([0-9]{1,2}(?:\.[0-9]{1,2}){1,3}[a-z]?)\s*$/mi);
    if (!idMatch) return;
    const id = idMatch[1].trim();
    if (!requested.has(id)) {
      unexpectedIds.add(id);
      return;
    }
    if (out.has(id)) {
      throw new Error(`Model output repeated ID ${id}. Expected exactly one block per ID.`);
    }
    const findingMatch = block.match(/FINDING:\s*([\s\S]*?)(?:\nRECOMMENDATION:|$)/i);
    const recMatch = block.match(/RECOMMENDATION:\s*([\s\S]*)$/i);
    const finding = normalizeExtractField(findingMatch ? findingMatch[1] : "", { rowId: id, field: "finding" });
    const recommendation = normalizeExtractField(recMatch ? recMatch[1] : "", { rowId: id, field: "recommendation" });
    out.set(id, { id, finding, recommendation, raw: block });
  });
  if (unexpectedIds.size) {
    throw new Error(`Model output included unexpected IDs: ${Array.from(unexpectedIds).join(", ")}.`);
  }
  return out;
}

function upsertCard(rowId, payload) {
  const existing = state.ui.cards.get(rowId) || {
    rowId,
    title: payload.title || "",
    currentFinding: payload.currentFinding || "",
    currentRecommendation: payload.currentRecommendation || "",
    proposedFinding: "",
    proposedRecommendation: "",
    actionFinding: "append",
    actionRecommendation: "append",
    takes: [],
  };
  if (payload.title && !existing.title) existing.title = payload.title;
  if (payload.currentFinding != null) existing.currentFinding = payload.currentFinding;
  if (payload.currentRecommendation != null) existing.currentRecommendation = payload.currentRecommendation;
  if (payload.take) {
    existing.takes.push(payload.take);
    if (payload.take.finding) {
      const add = payload.take.finding.trim();
      if (add) {
        existing.proposedFinding = existing.proposedFinding
          ? `${existing.proposedFinding.trim()}\n\n${add}`
          : add;
      }
    }
    if (payload.take.recommendation) {
      const add = payload.take.recommendation.trim();
      if (add) {
        existing.proposedRecommendation = existing.proposedRecommendation
          ? `${existing.proposedRecommendation.trim()}\n\n${add}`
          : add;
      }
    }
  }
  state.ui.cards.set(rowId, existing);
}

function renderCards() {
  if (!draftListEl) return;
  draftListEl.innerHTML = "";
  const cards = Array.from(state.ui.cards.values()).sort((a, b) => a.rowId.localeCompare(b.rowId, "de", { numeric: true }));
  cards.forEach((card) => {
    const row = document.createElement("div");
    row.className = "card";

    const header = document.createElement("div");
    header.className = "card-header";
    const left = document.createElement("div");
    const title = document.createElement("div");
    title.className = "card-title";
    title.textContent = `${card.rowId} ${card.title || ""}`.trim();
    const subtitle = document.createElement("div");
    subtitle.className = "card-subtitle";
    subtitle.textContent = `Takes: ${card.takes.length}`;
    left.appendChild(title);
    left.appendChild(subtitle);
    header.appendChild(left);

    const actions = document.createElement("div");
    actions.className = "card-header-actions";
    const micBtn = document.createElement("button");
    micBtn.type = "button";
    micBtn.className = "secondary card-voice";
    const isThisCardRecording =
      state.recording.isRecording &&
      state.recording.mode === "card" &&
      state.recording.rowId === card.rowId;
    micBtn.textContent = isThisCardRecording ? "Stop mic" : "Mic this card";
    if (isThisCardRecording) {
      micBtn.classList.add("is-recording");
    }
    micBtn.disabled =
      !state.pipeline ||
      !state.fs.sidecarDoc ||
      (state.recording.isRecording && !isThisCardRecording);
    micBtn.addEventListener("click", async () => {
      if (!state.pipeline) {
        setStatus("Load AI first.");
        return;
      }
      if (!state.fs.sidecarDoc || !state.fs.rowsById) {
        setStatus("Load a project sidecar first.");
        return;
      }

      const active = state.recording;
      const isThisCardRecordingNow =
        active.isRecording && active.mode === "card" && active.rowId === card.rowId;
      if (active.isRecording && !isThisCardRecordingNow) {
        const suffix =
          active.mode === "card" && active.rowId
            ? ` (currently recording ${active.rowId})`
            : " (currently recording)";
        setStatus(`Stop the current recording first${suffix}.`);
        return;
      }

      try {
        if (!active.isRecording) {
          await startRecording({ mode: "card", rowId: card.rowId, buttonEl: micBtn });
          return;
        }

        const capture = await stopRecording();
        if (!capture?.blob) return;
        micBtn.disabled = true;
        micBtn.textContent = "Transcribing...";
        setStatus(`Transcribing ${card.rowId} ...`);
        const text = await runAsrFromBlob(capture.blob, capture.capturedAudio);
        const note = stripTranscriptArtifacts(text);
        logBlock(`Card ${card.rowId} ASR (raw)`, text);
        logBlock(`Card ${card.rowId} ASR (cleaned)`, note);
        if (!note) {
          setStatus(`Transcript empty for ${card.rowId}.`);
          return;
        }
        assertTranscriptQuality(note, {
          contextLabel: `ASR transcript for ${card.rowId}`,
        });

        micBtn.textContent = "Extracting...";
        setStatus(`Extracting finding + recommendation for ${card.rowId} ...`);
        const sidecarRow = state.fs.rowsById.get(card.rowId);
        const itemTitle = sidecarRow ? getRowTitle(sidecarRow) : card.title || "";
        const result = await runLiquidExtract([{ rowId: card.rowId, title: itemTitle, note }]);
        logBlock(`Card ${card.rowId} Liquid (raw output)`, result?.text || "");
        const parsedOut = parseLiquidExtractOutput(result.text, [card.rowId]);
        const out = parsedOut.get(card.rowId);
        if (!out) {
          throw new Error(`Model output did not include required ID ${card.rowId}. No auto-filled text was used.`);
        }
        logBlock(`Card ${card.rowId} Parsed FINDING`, out.finding || "");
        logBlock(`Card ${card.rowId} Parsed RECOMMENDATION`, out.recommendation || "");

        upsertCard(card.rowId, {
          title: itemTitle,
          currentFinding: sidecarRow ? getRowCurrentFinding(sidecarRow) : card.currentFinding,
          currentRecommendation: sidecarRow ? getRowCurrentRecommendation(sidecarRow) : card.currentRecommendation,
          take: {
            createdAt: new Date().toISOString(),
            note,
            finding: out.finding,
            recommendation: out.recommendation,
            rawModelOutput: out.raw,
          },
        });
        setStatus(`Updated ${card.rowId}.`);
      } catch (err) {
        setStatus(`Card voice failed: ${err.message}`);
      } finally {
        renderCards();
        updateApplyEnabled();
      }
    });
    actions.appendChild(micBtn);
    header.appendChild(actions);
    row.appendChild(header);

    const grid = document.createElement("div");
    grid.className = "card-grid";

    const colCurrent = document.createElement("div");
    colCurrent.className = "card-col";
    colCurrent.innerHTML = `<h3>Current (sidecar)</h3>`;
    const currFinding = document.createElement("textarea");
    currFinding.value = card.currentFinding || "";
    currFinding.disabled = true;
    const currRec = document.createElement("textarea");
    currRec.value = card.currentRecommendation || "";
    currRec.disabled = true;
    colCurrent.appendChild(document.createTextNode("Finding"));
    colCurrent.appendChild(currFinding);
    colCurrent.appendChild(document.createTextNode("Recommendation"));
    colCurrent.appendChild(currRec);

    const colProposed = document.createElement("div");
    colProposed.className = "card-col";
    colProposed.innerHTML = `<h3>Proposed (voice)</h3>`;
    const propFinding = document.createElement("textarea");
    propFinding.value = card.proposedFinding || "";
    propFinding.addEventListener("input", () => {
      card.proposedFinding = propFinding.value;
      updateApplyEnabled();
    });
    const propRec = document.createElement("textarea");
    propRec.value = card.proposedRecommendation || "";
    propRec.addEventListener("input", () => {
      card.proposedRecommendation = propRec.value;
      updateApplyEnabled();
    });
    colProposed.appendChild(document.createTextNode("Finding"));
    colProposed.appendChild(propFinding);
    colProposed.appendChild(renderActionRow("Finding action", `${card.rowId}__finding`, card.actionFinding, (val) => {
      card.actionFinding = val;
      updateApplyEnabled();
    }));
    colProposed.appendChild(document.createTextNode("Recommendation"));
    colProposed.appendChild(propRec);
    colProposed.appendChild(renderActionRow("Recommendation action", `${card.rowId}__rec`, card.actionRecommendation, (val) => {
      card.actionRecommendation = val;
      updateApplyEnabled();
    }));

    grid.appendChild(colCurrent);
    grid.appendChild(colProposed);
    row.appendChild(grid);

    draftListEl.appendChild(row);
  });
}

function renderActionRow(labelText, groupName, value, onChange) {
  const wrap = document.createElement("div");
  wrap.className = "radio-row";
  const label = document.createElement("span");
  label.textContent = labelText;
  wrap.appendChild(label);

  const mk = (key, text) => {
    const l = document.createElement("label");
    const input = document.createElement("input");
    input.type = "radio";
    input.name = groupName;
    input.value = key;
    input.checked = value === key;
    input.addEventListener("change", () => {
      if (!input.checked) return;
      onChange(key);
    });
    l.appendChild(input);
    l.appendChild(document.createTextNode(text));
    return l;
  };

  wrap.appendChild(mk("skip", "Skip"));
  wrap.appendChild(mk("append", "Append"));
  wrap.appendChild(mk("replace", "Replace"));
  return wrap;
}

function computePendingApplyCount() {
  let count = 0;
  state.ui.cards.forEach((card) => {
    const finding = String(card.proposedFinding || "").trim();
    const rec = String(card.proposedRecommendation || "").trim();
    if (card.actionFinding !== "skip" && finding) count += 1;
    if (card.actionRecommendation !== "skip" && rec) count += 1;
  });
  return count;
}

function updateApplyEnabled() {
  if (applySidecarBtn) {
    applySidecarBtn.disabled = !state.fs.dirHandle || !state.fs.sidecarDoc || computePendingApplyCount() === 0;
  }
  if (saveDraftBtn) {
    saveDraftBtn.disabled = !state.fs.dirHandle || state.ui.cards.size === 0;
  }
}

async function saveDraftJson() {
  if (!state.fs.dirHandle) return;
  const timestamp = safeSlugTimestamp();
  const filename = `voice_draft_${timestamp}.json`;
  const handle = await state.fs.dirHandle.getFileHandle(filename, { create: true });
  const payload = {
    meta: {
      createdAt: new Date().toISOString(),
      locale: state.fs.locale,
      asrModel: asrModelEl?.value || "",
      liquidModel: DEFAULTS.liquidModel,
      projectFolder: state.fs.dirHandle?.name || "",
    },
    transcriptRaw: transcriptEl?.value || "",
    segmentsSorted: segmentsEl?.value || "",
    parsed: state.run.parsed || null,
    cards: Array.from(state.ui.cards.values()).map((c) => ({
      rowId: c.rowId,
      title: c.title,
      proposedFinding: c.proposedFinding,
      proposedRecommendation: c.proposedRecommendation,
      actionFinding: c.actionFinding,
      actionRecommendation: c.actionRecommendation,
      takes: c.takes,
    })),
  };
  await writeJsonToHandle(handle, payload);
  setStatus(`Saved draft: ${filename}`);
}

async function backupSidecar(sidecarText, filenamePrefix) {
  const dir = state.fs.dirHandle;
  if (!dir) return null;
  const backupDir = await dir.getDirectoryHandle("backup", { create: true });
  const safePrefix = String(filenamePrefix || "sidecar")
    .replace(/\.json$/i, "")
    .replace(/[^a-zA-Z0-9._-]/g, "_");
  const filename = `${safePrefix}_${safeSlugTimestamp()}.json`;
  const handle = await backupDir.getFileHandle(filename, { create: true });
  await writeTextToHandle(handle, sidecarText);
  return `backup/${filename}`;
}

function mergeText(existing, next, action) {
  const a = String(existing || "").trim();
  const b = String(next || "").trim();
  if (!b) return existing;
  if (action === "replace") return b;
  if (action === "append") {
    if (!a) return b;
    return `${a}\n\n${b}`.trim();
  }
  return existing;
}

async function applyToSidecar() {
  if (!state.fs.dirHandle) return;
  setStatus(`Preparing working sidecar (${WORKING_SIDECAR_FILENAME}) ...`);
  const writableSidecar = await ensureWritableWorkingSidecar(state.fs.dirHandle);
  const currentText = writableSidecar.text;
  const backupPath = await backupSidecar(currentText, writableSidecar.filename);
  if (backupPath) log(`Backup created: ${backupPath}`);

  const doc = JSON.parse(currentText);
  const project = extractReportProject(doc);
  if (!project) throw new Error("Could not find project in sidecar.");
  const { rowsById } = buildRowIndex(project);

  let applied = 0;
  state.ui.cards.forEach((card) => {
    const row = rowsById.get(card.rowId);
    if (!row) return;
    if (!row.workstate) row.workstate = {};
    if (card.actionFinding !== "skip") {
      const next = String(card.proposedFinding || "").trim();
      if (next) {
        row.workstate.findingText = mergeText(row.workstate.findingText, next, card.actionFinding);
        applied += 1;
      }
    }
    if (card.actionRecommendation !== "skip") {
      const next = String(card.proposedRecommendation || "").trim();
      if (next) {
        row.workstate.recommendationText = mergeText(row.workstate.recommendationText, next, card.actionRecommendation);
        applied += 1;
      }
    }
  });

  if (!doc.meta) doc.meta = {};
  doc.meta.updatedAt = new Date().toISOString();
  await writeJsonToHandle(writableSidecar.handle, doc);
  const seedNote = writableSidecar.seeded ? `; seeded from ${SOURCE_SIDECAR_FILENAME}` : "";
  setStatus(
    `Applied ${applied} field updates to ${WORKING_SIDECAR_FILENAME}${seedNote} (backup: ${backupPath || "none"})`
  );
}

async function pickFolder() {
  if (!window.showDirectoryPicker) {
    setStatus("File System Access API is not available. Use Chrome/Edge via http://127.0.0.1.");
    pickFolderBtn.disabled = true;
    return;
  }
  try {
    state.fs.dirHandle = await window.showDirectoryPicker();
    setStatus(`Selected folder: ${state.fs.dirHandle.name}`);
    loadSidecarBtn.disabled = false;
    updateApplyEnabled();
  } catch (err) {
    setStatus(`Folder pick canceled or failed: ${err.message}`);
  }
}

async function loadSidecar() {
  if (!state.fs.dirHandle) return;
  setStatus(`Loading sidecar for editor (${WORKING_SIDECAR_FILENAME} preferred) ...`);
  const loaded = await readSidecarForEditor(state.fs.dirHandle);
  const text = loaded.text;
  const doc = JSON.parse(text);
  const project = extractReportProject(doc);
  if (!project) throw new Error("Could not find report project in sidecar.");
  const meta = project.meta || {};
  const locale = meta.locale || "de-CH";
  const { rowsById, subIdToRowId } = buildRowIndex(project);

  state.fs.sidecarDoc = doc;
  state.fs.sidecarProject = project;
  state.fs.rowsById = rowsById;
  state.fs.subIdToRowId = subIdToRowId;
  state.fs.locale = locale;
  state.run.parsed = null;
  state.run.draft = null;
  if (transcriptEl) transcriptEl.value = "";
  if (segmentsEl) segmentsEl.value = "";
  if (idSummaryEl) idSummaryEl.textContent = "";

  seedCardsFromSidecarProject(project);
  renderCards();
  updateApplyEnabled();

  const sourceLabel = loaded.source === "working" ? WORKING_SIDECAR_FILENAME : SOURCE_SIDECAR_FILENAME;
  setStatus(`Loaded ${sourceLabel}. Locale=${locale}. Rows=${rowsById.size}. Cards=${state.ui.cards.size}.`);
  updateAsrLanguageBadge();
  recordToggleBtn.disabled = !state.pipeline; // require AI loaded
  parseIdsBtn.disabled = false;
  runExtractBtn.disabled = true;
  saveDraftBtn.disabled = state.ui.cards.size === 0;
  applySidecarBtn.disabled = true;
}

function renderIdSummary(parsed) {
  if (!idSummaryEl) return;
  if (!parsed) {
    idSummaryEl.textContent = "";
    return;
  }
  const lines = [];
  lines.push(`Mentions: ${parsed.mentions.length}`);
  lines.push(`Resolved row IDs: ${parsed.groups.length}`);
  if (parsed.warnings.length) {
    lines.push("");
    lines.push("Warnings:");
    parsed.warnings.forEach((w) => lines.push(`- ${w}`));
  }
  idSummaryEl.textContent = lines.join("\n");
}

function refreshCardsIfLoaded() {
  if (!state.fs.sidecarDoc) return;
  renderCards();
  updateApplyEnabled();
}

function parseIdsFromTranscript() {
  const raw = stripTranscriptArtifacts(transcriptEl?.value || "");
  if (!raw) {
    setStatus("Transcript is empty.");
    return null;
  }
  if (!state.fs.rowsById || state.fs.rowsById.size === 0) {
    setStatus("Load a project sidecar first.");
    return null;
  }
  const localeKey = resolveLocaleKey(state.fs.locale);
  const mentions = extractIdMentions(raw, {
    rowsById: state.fs.rowsById,
    subIdToRowId: state.fs.subIdToRowId,
    localeKey,
  });
  const segments = segmentByMentions(raw, mentions);
  const groups = groupSegmentsByRowId(segments);
  const warnings = [];
  if (mentions.length === 0) warnings.push("No IDs detected. Try saying: 'für 1.1.1 ... für 1.1.2 ...'");
  if (segments.some((s) => !s.note)) warnings.push("Some IDs have empty text between them. Those will be skipped.");
  if (mentions.some((m) => m.kind === "sub")) warnings.push("Some mentions were subquestion IDs and were mapped to their parent row.");

  const sortedText = buildSortedSegmentsText(groups, { rowsById: state.fs.rowsById });
  if (segmentsEl) segmentsEl.value = sortedText;

  const parsed = { transcript: raw, locale: state.fs.locale, localeKey, mentions, segments, groups, warnings };
  state.run.parsed = parsed;
  renderIdSummary(parsed);
  runExtractBtn.disabled = groups.length === 0;
  if (groups.length === 0) {
    setStatus(
      `Transcript captured (${raw.length} chars). No row IDs detected yet. Say IDs like 'für 1.1.1 ...' or use per-card mic.`
    );
  } else {
    setStatus(`Transcript captured (${raw.length} chars). Parsed ${mentions.length} mentions across ${groups.length} rows.`);
  }
  return parsed;
}

async function extractDrafts() {
  const parsed = state.run.parsed || parseIdsFromTranscript();
  if (!parsed) return;
  const items = parsed.groups
    .map((g) => {
      const row = state.fs.rowsById.get(g.rowId);
      return {
        rowId: g.rowId,
        title: row ? getRowTitle(row) : "",
        note: g.note,
      };
    })
    .filter((it) => it.note && it.note.trim().length > 0);

  if (!items.length) {
    setStatus("No non-empty segments to extract.");
    return;
  }
  const invalidNotes = [];
  items.forEach((it) => {
    const quality = evaluateTranscriptQuality(it.note);
    if (!quality.ok) {
      invalidNotes.push(`${it.rowId} (${quality.reason})`);
    }
  });
  if (invalidNotes.length) {
    throw new Error(
      `Refusing extraction for low-quality transcript segments: ${invalidNotes.join(", ")}. Re-record those findings.`
    );
  }
  const start = performance.now();
  const batchSize = 6;
  for (let offset = 0; offset < items.length; offset += batchSize) {
    const batch = items.slice(offset, offset + batchSize);
    setStatus(
      `Extracting ${batch.length} items (${offset + 1}-${offset + batch.length} of ${items.length}) with LiquidAI (offline) ...`
    );
    const batchStart = performance.now();
    const result = await runLiquidExtract(batch);
    log(`Liquid extract batch: ${formatMs(performance.now() - batchStart)} (items=${batch.length})`);

    const rowIds = batch.map((it) => it.rowId);
    const parsedOut = parseLiquidExtractOutput(result.text, rowIds);
    const missing = rowIds.filter((id) => !parsedOut.has(id));
    if (missing.length) {
      throw new Error(`Model output missing required IDs: ${missing.join(", ")}. No auto-filled text was used.`);
    }

    batch.forEach((it) => {
      const row = state.fs.rowsById.get(it.rowId);
      const currentFinding = row ? getRowCurrentFinding(row) : "";
      const currentRec = row ? getRowCurrentRecommendation(row) : "";
      const out = parsedOut.get(it.rowId);
      if (!out) {
        throw new Error(`Internal parse contract failure for row ${it.rowId}.`);
      }
      upsertCard(it.rowId, {
        title: row ? getRowTitle(row) : "",
        currentFinding,
        currentRecommendation: currentRec,
        take: {
          createdAt: new Date().toISOString(),
          note: it.note,
          finding: out.finding,
          recommendation: out.recommendation,
          rawModelOutput: out.raw,
        },
      });
    });

    renderCards();
    updateApplyEnabled();
  }
  log(`Liquid extract total: ${formatMs(performance.now() - start)}`);

  renderCards();
  updateApplyEnabled();
  setStatus(`Extraction complete. Draft cards: ${state.ui.cards.size}.`);
}

function ensureButtons() {
  const hasAi = !!state.pipeline;
  const isCardRecording = state.recording.isRecording && state.recording.mode === "card";
  loadSidecarBtn.disabled = !state.fs.dirHandle;
  recordToggleBtn.disabled = isCardRecording || !hasAi || !state.fs.sidecarDoc;
  parseIdsBtn.disabled = !state.fs.sidecarDoc;
  updateAsrLanguageBadge();
  updateApplyEnabled();
}

loadAiBtn?.addEventListener("click", async () => {
  try {
    await loadLibrary();
    await ensureWebGpuFeatures();
    recordToggleBtn.disabled = !state.fs.sidecarDoc;
    ensureButtons();
    refreshCardsIfLoaded();
  } catch {
    // status already set
  }
});

pickFolderBtn?.addEventListener("click", async () => {
  await pickFolder();
  ensureButtons();
});

loadSidecarBtn?.addEventListener("click", async () => {
  try {
    await loadSidecar();
    ensureButtons();
  } catch (err) {
    setStatus(`Sidecar load failed: ${err.message}`);
  }
});

recordToggleBtn?.addEventListener("click", async () => {
  try {
    if (!state.recording.isRecording) {
      await startRecording();
      refreshCardsIfLoaded();
      return;
    }
    const capture = await stopRecording();
    if (!capture?.blob) return;
    setStatus("Transcribing recording ...");
    const text = await runAsrFromBlob(capture.blob, capture.capturedAudio);
    const cleaned = stripTranscriptArtifacts(text);
    transcriptEl.value = cleaned;
    if (!cleaned) {
      throw new Error("ASR returned empty transcript. Check mic input level and selected ASR model.");
    }
    assertTranscriptQuality(cleaned, {
      contextLabel: "ASR transcript",
    });
    parseIdsBtn.disabled = false;
    runExtractBtn.disabled = true;
    saveDraftBtn.disabled = true;
    applySidecarBtn.disabled = true;
    parseIdsFromTranscript();
    refreshCardsIfLoaded();
  } catch (err) {
    setStatus(`Recording/transcribe failed: ${err.message}`);
  }
});

parseIdsBtn?.addEventListener("click", () => {
  parseIdsFromTranscript();
});

asrLanguageModeEl?.addEventListener("change", () => {
  updateAsrLanguageBadge();
});

asrInputModeEl?.addEventListener("change", () => {
  setStatus(`ASR input mode: ${asrInputModeEl.value}`);
});

asrBackendModeEl?.addEventListener("change", () => {
  setStatus(`ASR backend mode: ${asrBackendModeEl.value}`);
});

runExtractBtn?.addEventListener("click", async () => {
  try {
    await extractDrafts();
  } catch (err) {
    setStatus(`Extraction failed: ${err.message}`);
  }
});

saveDraftBtn?.addEventListener("click", async () => {
  try {
    await saveDraftJson();
  } catch (err) {
    setStatus(`Draft save failed: ${err.message}`);
  }
});

applySidecarBtn?.addEventListener("click", async () => {
  try {
    await applyToSidecar();
  } catch (err) {
    setStatus(`Apply failed: ${err.message}`);
  }
});

// Initialization
ensureButtons();
