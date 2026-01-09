const byId = (id) => document.getElementById(id);

const asrModelEl = byId("asr-model");
const loadLibBtn = byId("load-lib");
const holdBtn = byId("hold-to-talk");
const statusEl = byId("status");
const transcriptEl = byId("transcript");
const logEl = byId("log");

const DEFAULTS = {
  transformersUrl: "../AI/vendor/transformers.min.js",
  localModelPath: "../AI/models/",
};

const state = {
  pipeline: null,
  env: null,
  pipelines: new Map(),
  webgpu: {
    supportsFp16: null,
  },
  isRecording: false,
  isTranscribing: false,
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

async function loadLibrary() {
  setStatus(`Loading library from ${DEFAULTS.transformersUrl} ...`);
  try {
    const mod = await import(DEFAULTS.transformersUrl);
    if (!mod.pipeline || !mod.env) {
      throw new Error("Module missing pipeline/env exports.");
    }
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

async function runAsrFromBlob(blob) {
  const modelId = asrModelEl.value.trim();
  if (!modelId) {
    setStatus("Pick an ASR model.");
    return;
  }
  setStatus(`Loading ASR pipeline (${modelId}) ...`);
  try {
    const dtype = await resolveAsrDtype();
    await ensureAsrModelFiles(modelId, dtype);
    const asr = await getPipeline("automatic-speech-recognition", modelId, { dtype });
    setStatus("Decoding audio ...");
    const { audio } = await decodeAudioBlob(blob);
    setStatus("Transcribing ...");
    const options = { chunk_length_s: 30, stride_length_s: 5 };
    const result = await asr(audio, options);
    const text = typeof result === "string" ? result : result?.text || JSON.stringify(result, null, 2);
    transcriptEl.value = text;
    setStatus("Transcription complete.");
  } catch (err) {
    setStatus(`ASR failed: ${err.message}`);
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
  if (state.isRecording || state.isTranscribing) return;
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
