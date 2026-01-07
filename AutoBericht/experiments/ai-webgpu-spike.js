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
const visionFileEl = byId("vision-file");
const runVisionBtn = byId("run-vision");
const visionStatusEl = byId("vision-status");
const visionPreviewEl = byId("vision-preview");
const visionOutputEl = byId("vision-output");

const logEl = byId("log");

const state = {
  lib: null,
  pipeline: null,
  env: null,
  pipelines: new Map(),
};

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

async function getPipeline(task, modelId) {
  const device = getDeviceOption();
  const key = `${task}::${modelId}::${device || "auto"}`;
  if (state.pipelines.has(key)) return state.pipelines.get(key);
  if (!state.pipeline) {
    throw new Error("Library not loaded.");
  }
  applyEnv();
  const options = {};
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
    const asr = await getPipeline("automatic-speech-recognition", modelId);
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
    setStatus(visionStatusEl, `Vision failed: ${err.message}`);
  }
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

allowRemoteEl.addEventListener("change", () => {
  applyEnv();
});

localModelPathEl.addEventListener("change", () => {
  applyEnv();
});

transformersUrlEl.value = "./vendor/transformers.min.js";
localModelPathEl.value = "/AutoBericht/experiments/models/";
allowRemoteEl.checked = false;
asrModelEl.value = "Xenova/whisper-tiny.en";
visionModelEl.value = "Xenova/vit-gpt2-image-captioning";
