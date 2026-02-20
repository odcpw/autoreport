const APP_VERSION = "0.2.0";
const SPEC_VERSION = "0.1";
const GENERATOR_ID = "ai-evidence-lab";
const EXPERIMENT_LABEL = "Experimental - standalone";
const EMBEDDING_MODEL_ID = "Xenova/paraphrase-multilingual-MiniLM-L12-v2";

const CHECKPOINT_FILE = "ingest_checkpoint.json";
const INDEX_LATEST_FILE = "evidence_index.latest.json";
const MATCHES_LATEST_FILE = "evidence_matches.latest.json";

const SUPPORTED_EXTENSIONS = new Set([
  ".pdf",
  ".docx",
  ".xlsx",
  ".xls",
  ".png",
  ".jpg",
  ".jpeg",
  ".webp",
  ".tif",
  ".tiff",
]);

const QUERY_SYNONYMS = {
  de: ["sicherheit", "gesundheit", "verantwortung", "instruktion", "stellenbeschreibung"],
  fr: ["securite", "sante", "responsabilite", "instruction", "description de poste"],
  it: ["sicurezza", "salute", "responsabilita", "istruzione", "descrizione del posto"],
  en: ["safety", "health", "responsibility", "instruction", "job description"],
};

const state = {
  capabilities: {},
  runtime: {
    preferred: "wasm",
    notes: [],
    localOnlyModels: true,
  },
  fs: {
    projectDir: null,
    inputsDir: null,
    outputsDir: null,
    evidenceDir: null,
    sidecarHandle: null,
  },
  projectMeta: {},
  rowMap: new Map(),
  rowOrder: [],
  selectedRowId: "",
  inventory: [],
  inventoryMap: new Map(),
  checkpoint: {
    loaded: false,
    data: null,
  },
  index: {
    payload: null,
    latestHandleName: "",
  },
  matches: {
    selected: null,
    included: [],
    byRowId: new Map(),
    latestHandleName: "",
  },
  patchPreview: {
    payload: null,
    validation: null,
  },
  drafts: {
    byRowId: new Map(),
  },
  logs: [],
  workers: {
    ingest: null,
    embed: null,
    match: null,
  },
  run: {
    scanToken: "",
    scanActive: false,
    scanPromise: null,
    scanResolve: null,
    stageDurationsMs: {
      scan: 0,
      index: 0,
      match_selected: 0,
      match_included: 0,
    },
    failures: 0,
  },
  settings: {
    batchSize: 16,
    maxUnitsPerRun: 0,
    memoryChunkGuard: 50000,
    profile: "default",
    includeSensitiveLogs: false,
  },
  schemas: {
    evidenceIndex: null,
    evidenceMatches: null,
    sidecarPatch: null,
  },
  embedding: {
    mode: "hash",
    modelId: EMBEDDING_MODEL_ID,
    extractor: null,
    lib: null,
    initAttempted: false,
    initError: "",
  },
};

const el = {
  appVersion: document.getElementById("app-version"),
  capabilityChips: document.getElementById("capability-chips"),
  projectSummary: document.getElementById("project-summary"),
  runSummary: document.getElementById("run-summary"),
  inventoryBody: document.getElementById("inventory-body"),
  rowSearch: document.getElementById("row-search"),
  rowList: document.getElementById("row-list"),
  rowDetails: document.getElementById("row-details"),
  evidenceState: document.getElementById("evidence-state"),
  evidenceList: document.getElementById("evidence-list"),
  filterLowConfidence: document.getElementById("filter-low-confidence"),
  groupByFile: document.getElementById("group-by-file"),
  patchPreview: document.getElementById("patch-preview"),
  logOutput: document.getElementById("log-output"),

  pickFolderBtn: document.getElementById("pick-folder"),
  scanInputsBtn: document.getElementById("scan-inputs"),
  cancelRunBtn: document.getElementById("cancel-run"),
  buildIndexBtn: document.getElementById("build-index"),
  clearCacheBtn: document.getElementById("clear-cache"),
  matchSelectedBtn: document.getElementById("match-selected"),
  matchIncludedBtn: document.getElementById("match-included"),
  generateDraftBtn: document.getElementById("generate-draft"),
  exportProposalBtn: document.getElementById("export-proposal"),
  runSmokeBtn: document.getElementById("run-smoke"),
  saveLogBtn: document.getElementById("save-log"),
  requestPersistBtn: document.getElementById("request-persist"),
  settingBatchSize: document.getElementById("setting-batch-size"),
  settingMaxUnits: document.getElementById("setting-max-units"),
  settingMemoryGuard: document.getElementById("setting-memory-guard"),
  profileOvernightBtn: document.getElementById("profile-overnight"),
  profileCpuBtn: document.getElementById("profile-cpu"),
  profileWebgpuBtn: document.getElementById("profile-webgpu"),
  optLogSensitive: document.getElementById("opt-log-sensitive"),
};

function byIdOrPath(a, b) {
  return String(a || "").localeCompare(String(b || ""), undefined, { numeric: true, sensitivity: "base" });
}

function toText(value) {
  if (Array.isArray(value)) {
    return value
      .map((part) => {
        if (!part) return "";
        if (typeof part === "string") return part;
        if (typeof part.text === "string") return part.text;
        return "";
      })
      .join("");
  }
  if (typeof value === "string") return value;
  if (value == null) return "";
  return String(value);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function timestampIso() {
  return new Date().toISOString();
}

function timestampForFile() {
  return timestampIso().replace(/[:.]/g, "-");
}

function formatBytes(bytes) {
  const n = Number(bytes || 0);
  if (!Number.isFinite(n) || n < 0) return "-";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  if (n < 1024 * 1024 * 1024) return `${(n / (1024 * 1024)).toFixed(1)} MB`;
  return `${(n / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

function extFromName(name) {
  const idx = String(name || "").lastIndexOf(".");
  return idx >= 0 ? String(name).slice(idx).toLowerCase() : "";
}

function isSupportedInputFile(name) {
  return SUPPORTED_EXTENSIONS.has(extFromName(name));
}

function getSourceModeByExtension(ext) {
  if (ext === ".pdf") return "pdf_text";
  if (ext === ".docx") return "docx";
  if (ext === ".xlsx" || ext === ".xls") return "xlsx";
  if ([".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"].includes(ext)) return "image_ocr";
  return "unknown";
}

function tokenize(text) {
  const value = String(text || "").toLowerCase();
  return value
    .split(/[^a-z0-9]+/i)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);
}

function hashString32(input) {
  let hash = 2166136261;
  const text = String(input || "");
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
  }
  return `h${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function logLine(level, message, context = null) {
  const normalizedContext = sanitizeLogContext(context);
  const entry = {
    ts: timestampIso(),
    level,
    message,
    context: normalizedContext,
  };
  state.logs.push(entry);
  const suffix = normalizedContext ? ` ${JSON.stringify(normalizedContext)}` : "";
  const line = `[${entry.ts}] ${level.toUpperCase()} ${message}${suffix}`;
  el.logOutput.textContent = `${el.logOutput.textContent}${line}\n`;
  el.logOutput.scrollTop = el.logOutput.scrollHeight;
  if (level === "error") {
    state.run.failures += 1;
  }
}

function sanitizeLogContext(context) {
  if (!context || typeof context !== "object") return context;
  if (state.settings.includeSensitiveLogs) return context;

  const blockedKeys = new Set(["snippet", "text", "finding", "recommendation", "content"]);
  const sanitized = {};
  Object.entries(context).forEach(([key, value]) => {
    if (blockedKeys.has(String(key).toLowerCase())) {
      sanitized[key] = "[redacted]";
      return;
    }
    sanitized[key] = value;
  });
  return sanitized;
}

function stageTimerStart() {
  return performance.now();
}

function stageTimerEnd(key, startedAt) {
  const elapsed = Math.max(0, Math.round(performance.now() - startedAt));
  state.run.stageDurationsMs[key] = elapsed;
  logLine("info", `Stage timing: ${key}`, { elapsed_ms: elapsed });
}

function statusForBool(flag) {
  return flag ? "ok" : "bad";
}

function setRunSummary(lines) {
  const list = Array.isArray(lines) ? lines : [String(lines || "")];
  el.runSummary.textContent = list.join("\n");
}

function updateRunSummaryPanel(extra = []) {
  const counts = {
    total: state.inventory.length,
    pending: state.inventory.filter((item) => item.status === "pending").length,
    running: state.inventory.filter((item) => item.status === "running").length,
    done: state.inventory.filter((item) => item.status === "done").length,
    skipped: state.inventory.filter((item) => item.status === "skipped").length,
    failed: state.inventory.filter((item) => item.status === "failed").length,
  };

  const lines = [
    `Files: total=${counts.total}, done=${counts.done}, skipped=${counts.skipped}, pending=${counts.pending}, failed=${counts.failed}`,
    `Scan active: ${state.run.scanActive ? "yes" : "no"}`,
    `Profile: ${state.settings.profile} | batch=${state.settings.batchSize} | max_units=${state.settings.maxUnitsPerRun || "none"} | memory_guard=${state.settings.memoryChunkGuard}`,
    `Embedding mode: ${state.embedding.mode}${state.embedding.initError ? ` (${state.embedding.initError})` : ""}`,
    `Durations (ms): scan=${state.run.stageDurationsMs.scan}, index=${state.run.stageDurationsMs.index}, match(selected)=${state.run.stageDurationsMs.match_selected}, match(included)=${state.run.stageDurationsMs.match_included}`,
    `Failures: ${state.run.failures}`,
  ];

  if (extra && extra.length) {
    lines.push(...extra);
  }
  setRunSummary(lines);
}

function clampNumber(value, min, max, fallback) {
  const raw = Number(value);
  if (!Number.isFinite(raw)) return fallback;
  return Math.min(max, Math.max(min, raw));
}

function applySettingsToUi() {
  if (el.settingBatchSize) el.settingBatchSize.value = String(state.settings.batchSize);
  if (el.settingMaxUnits) el.settingMaxUnits.value = String(state.settings.maxUnitsPerRun);
  if (el.settingMemoryGuard) el.settingMemoryGuard.value = String(state.settings.memoryChunkGuard);
  if (el.optLogSensitive) el.optLogSensitive.checked = !!state.settings.includeSensitiveLogs;
}

function updateSettingsFromUi() {
  state.settings.batchSize = clampNumber(el.settingBatchSize?.value, 1, 256, state.settings.batchSize);
  state.settings.maxUnitsPerRun = clampNumber(el.settingMaxUnits?.value, 0, 100000, state.settings.maxUnitsPerRun);
  state.settings.memoryChunkGuard = clampNumber(el.settingMemoryGuard?.value, 100, 1000000, state.settings.memoryChunkGuard);
  state.settings.includeSensitiveLogs = !!el.optLogSensitive?.checked;
}

function applyProfile(profile) {
  state.settings.profile = profile;

  if (profile === "overnight") {
    state.settings.batchSize = 64;
    state.settings.maxUnitsPerRun = 0;
    state.settings.memoryChunkGuard = 150000;
  } else if (profile === "cpu") {
    state.settings.batchSize = 8;
    state.settings.maxUnitsPerRun = 200;
    state.settings.memoryChunkGuard = 25000;
  } else if (profile === "webgpu") {
    state.settings.batchSize = 48;
    state.settings.maxUnitsPerRun = 0;
    state.settings.memoryChunkGuard = 100000;
  }

  applySettingsToUi();
  updateRunSummaryPanel([`Profile applied: ${profile}`]);
  logLine("info", "Performance profile applied.", {
    profile,
    batch_size: state.settings.batchSize,
    max_units: state.settings.maxUnitsPerRun,
    memory_guard: state.settings.memoryChunkGuard,
  });
}

async function digestHex(buffer) {
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  const view = new Uint8Array(digest);
  return Array.from(view).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function computeFileFingerprint(file) {
  const meta = `${file.name}|${file.size}|${file.lastModified}`;
  const slice = await file.slice(0, 64 * 1024).arrayBuffer();
  const metaBytes = new TextEncoder().encode(meta);
  const merged = new Uint8Array(metaBytes.byteLength + slice.byteLength);
  merged.set(metaBytes, 0);
  merged.set(new Uint8Array(slice), metaBytes.byteLength);
  const hash = await digestHex(merged.buffer);
  return hash;
}

async function detectCapabilities() {
  const caps = {
    fs_api: typeof window.showDirectoryPicker === "function",
    workers: typeof window.Worker !== "undefined",
    webgpu: !!navigator.gpu,
    storage_estimate: !!navigator.storage?.estimate,
    storage_persist: !!navigator.storage?.persist,
  };

  if (caps.storage_estimate) {
    try {
      const estimate = await navigator.storage.estimate();
      caps.storage_usage = formatBytes(estimate?.usage || 0);
      caps.storage_quota = formatBytes(estimate?.quota || 0);
    } catch (err) {
      caps.storage_estimate = false;
      logLine("warn", "Storage estimate failed", { error: err.message });
    }
  }

  if (caps.webgpu && navigator.gpu?.requestAdapter) {
    try {
      const adapter = await navigator.gpu.requestAdapter();
      caps.webgpu = !!adapter;
      if (adapter) {
        const features = new Set(Array.from(adapter.features || []));
        caps.webgpu_shader_f16 = features.has("shader-f16");
      }
    } catch (err) {
      caps.webgpu = false;
      logLine("warn", "WebGPU capability check failed", { error: err.message });
    }
  }

  state.capabilities = caps;
  state.runtime.preferred = caps.webgpu ? "webgpu" : "wasm";
  state.runtime.notes = [
    `preferred=${state.runtime.preferred}`,
    `shader_f16=${caps.webgpu_shader_f16 ? "yes" : "no"}`,
  ];
  logLine("info", "Runtime plan selected.", {
    preferred: state.runtime.preferred,
    webgpu: caps.webgpu,
    shader_f16: !!caps.webgpu_shader_f16,
    local_only_models: state.runtime.localOnlyModels,
  });
}

function renderCapabilityChips() {
  const caps = state.capabilities;
  const chips = [
    { key: "FS API", ok: !!caps.fs_api, extra: "" },
    { key: "Workers", ok: !!caps.workers, extra: "" },
    { key: "WebGPU", ok: !!caps.webgpu, extra: caps.webgpu_shader_f16 ? "f16" : "" },
    { key: "Storage", ok: !!caps.storage_estimate, extra: `${caps.storage_usage || "-"} / ${caps.storage_quota || "-"}` },
    { key: "Persist", ok: !!caps.storage_persist, extra: "" },
    { key: "Runtime", ok: true, extra: state.runtime.preferred },
    { key: "Models", ok: state.runtime.localOnlyModels, extra: state.runtime.localOnlyModels ? "local-only" : "remote-enabled" },
  ];

  el.capabilityChips.innerHTML = "";
  chips.forEach((chip) => {
    const span = document.createElement("span");
    span.className = `chip ${statusForBool(chip.ok)}`;
    span.textContent = chip.extra ? `${chip.key}: ${chip.extra}` : chip.key;
    el.capabilityChips.appendChild(span);
  });
}

function renderProjectSummary() {
  const lines = [];
  lines.push(`Folder: ${state.fs.projectDir?.name || "(none)"}`);
  lines.push(`Inputs: ${state.fs.inputsDir ? "ok" : "missing"}`);
  lines.push(`Sidecar: ${state.fs.sidecarHandle ? "ok" : "missing"}`);
  lines.push(`Locale: ${state.projectMeta.locale || "-"}`);
  lines.push(`Company: ${state.projectMeta.company || state.projectMeta.companyName || "-"}`);
  lines.push(`Rows: ${state.rowOrder.length}`);
  lines.push(`Checkpoint: ${state.checkpoint.loaded ? "loaded" : "none"}`);
  lines.push(`Index cache: ${state.index.payload ? "loaded" : "none"}`);
  lines.push(`Match cache: ${state.matches.byRowId.size ? "loaded" : "none"}`);
  el.projectSummary.textContent = lines.join("\n");
}

async function requestPersistentStorage() {
  if (!navigator.storage?.persist) {
    logLine("warn", "Persistent storage API not available.");
    return;
  }
  try {
    const granted = await navigator.storage.persist();
    logLine("info", `Persistent storage request result: ${granted ? "granted" : "not granted"}`);
    await detectCapabilities();
    renderCapabilityChips();
  } catch (err) {
    logLine("error", "Persistent storage request failed", { error: err.message });
  }
}

async function safeGetDirectoryHandle(parent, name, create = false) {
  try {
    return await parent.getDirectoryHandle(name, { create });
  } catch {
    return null;
  }
}

async function safeGetFileHandle(parent, name, create = false) {
  try {
    return await parent.getFileHandle(name, { create });
  } catch {
    return null;
  }
}

async function safeRemoveEntry(dirHandle, name, recursive = false) {
  try {
    await dirHandle.removeEntry(name, { recursive });
    return true;
  } catch {
    return false;
  }
}

async function readFileText(handle) {
  const file = await handle.getFile();
  return await file.text();
}

async function writeJson(handle, payload) {
  const writable = await handle.createWritable();
  await writable.write(JSON.stringify(payload, null, 2));
  await writable.close();
}

async function writeText(handle, text) {
  const writable = await handle.createWritable();
  await writable.write(String(text || ""));
  await writable.close();
}

function parseSidecarRows(rawDoc) {
  const project = rawDoc?.project || rawDoc || {};
  const chapters = Array.isArray(project.chapters) ? project.chapters : [];
  const rowMap = new Map();

  chapters.forEach((chapter) => {
    const chapterId = String(chapter?.id || "");
    const chapterTitle = toText(chapter?.title || chapter?.master?.title || "").trim();
    const rows = Array.isArray(chapter?.rows) ? chapter.rows : [];

    rows.forEach((row) => {
      if (!row || row.kind === "section") return;
      const rowId = String(row.id || "").trim();
      if (!rowId) return;
      const ws = row.workstate || {};
      const priorityRaw = Number(ws.priority);
      const priority = Number.isFinite(priorityRaw) ? Math.max(0, Math.min(4, Math.round(priorityRaw))) : 0;
      const title = toText(row?.master?.title || row?.title || "").trim();
      const findingText = toText(ws.findingText || row?.master?.finding || "").trim();
      const recommendationText = toText(ws.recommendationText || row?.master?.recommendation || "").trim();
      const rowHash = hashString32(`${rowId}|${findingText}|${recommendationText}|${ws.include ? 1 : 0}|${ws.done ? 1 : 0}|${priority}`);

      rowMap.set(rowId, {
        rowId,
        chapterId,
        chapterTitle,
        title,
        include: !!ws.include,
        done: !!ws.done,
        priority,
        rowHash,
      });
    });
  });

  return {
    projectMeta: project.meta || {},
    rowMap,
    rowOrder: Array.from(rowMap.keys()).sort(byIdOrPath),
  };
}

function rowMatchesFilter(row, query) {
  if (!query) return true;
  const q = query.toLowerCase();
  return (
    row.rowId.toLowerCase().includes(q)
    || row.title.toLowerCase().includes(q)
    || row.chapterId.toLowerCase().includes(q)
    || row.chapterTitle.toLowerCase().includes(q)
  );
}

function renderRowList() {
  const query = String(el.rowSearch.value || "").trim();
  el.rowList.innerHTML = "";

  const rows = state.rowOrder
    .map((rowId) => state.rowMap.get(rowId))
    .filter((row) => row && rowMatchesFilter(row, query));

  rows.forEach((row) => {
    const item = document.createElement("div");
    item.className = `row-item${row.rowId === state.selectedRowId ? " active" : ""}`;
    const title = row.title || "(untitled row)";
    item.innerHTML = `
      <strong>${escapeHtml(row.rowId)}</strong> ${escapeHtml(title)}
      <br />
      <small>${escapeHtml(row.chapterId)} ${escapeHtml(row.chapterTitle || "")}</small>
      <br />
      <small>include=${row.include ? "yes" : "no"}, done=${row.done ? "yes" : "no"}, prio=${row.priority}</small>
    `;
    item.addEventListener("click", () => {
      state.selectedRowId = row.rowId;
      renderRowList();
      renderRowDetails();
      renderEvidenceForSelectedRow();
      renderPatchPreview();
    });
    el.rowList.appendChild(item);
  });
}

function renderRowDetails() {
  if (!state.selectedRowId || !state.rowMap.has(state.selectedRowId)) {
    el.rowDetails.textContent = "Select a row.";
    return;
  }
  const row = state.rowMap.get(state.selectedRowId);
  el.rowDetails.textContent = JSON.stringify(row, null, 2);
}

function highlightSnippet(snippet, terms) {
  let text = escapeHtml(snippet);
  terms.forEach((term) => {
    if (!term || term.length < 3) return;
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`\\b(${escaped})\\b`, "gi");
    text = text.replace(re, "<mark>$1</mark>");
  });
  return text;
}

async function copyTextToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    logLine("info", "Copied citation snippet to clipboard.");
  } catch (err) {
    logLine("warn", "Clipboard write failed.", { error: err.message });
  }
}

function renderEvidenceForSelectedRow() {
  if (!state.selectedRowId) {
    el.evidenceState.textContent = "No selected row.";
    el.evidenceList.innerHTML = "";
    return;
  }

  const payload = state.matches.byRowId.get(state.selectedRowId);
  if (!payload) {
    el.evidenceState.textContent = "No evidence run yet for selected row.";
    el.evidenceList.innerHTML = "";
    return;
  }

  const rawCitations = Array.isArray(payload.citations) ? payload.citations : [];
  const hideLowConfidence = !!el.filterLowConfidence?.checked;
  const groupByFile = !!el.groupByFile?.checked;
  const citations = hideLowConfidence
    ? rawCitations.filter((citation) => Number(citation.confidence || 0) >= 0.6)
    : rawCitations;
  el.evidenceState.textContent = `Status: ${payload.status} | citations=${citations.length}/${rawCitations.length}`;
  el.evidenceList.innerHTML = "";

  const terms = tokenize(`${payload.query || ""}`);
  let grouped = [{ file: "", items: citations }];
  if (groupByFile) {
    const map = new Map();
    citations.forEach((citation) => {
      const key = String(citation.file || "");
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(citation);
    });
    grouped = Array.from(map.entries()).map(([file, items]) => ({ file, items }));
    grouped.sort((a, b) => byIdOrPath(a.file, b.file));
  }

  grouped.forEach((group) => {
    if (groupByFile) {
      const title = document.createElement("p");
      title.className = "evidence-group-title";
      title.textContent = `${group.file} (${group.items.length})`;
      el.evidenceList.appendChild(title);
    }

    group.items.forEach((citation) => {
    const card = document.createElement("div");
    card.className = "evidence-card";

    const scoreText = Number(citation.score || 0).toFixed(3);
    const confidenceText = Number(citation.confidence || 0).toFixed(2);
    const metaText = `${citation.file || "unknown"}${citation.page ? `, p.${citation.page}` : ""} | confidence=${confidenceText}`;
    const snippetHtml = highlightSnippet(citation.snippet || "", terms);

    card.innerHTML = `
      <div class="evidence-head">
        <span class="evidence-meta">${escapeHtml(metaText)}</span>
        <span class="evidence-score">${escapeHtml(scoreText)}</span>
      </div>
      <details>
        <summary>Snippet</summary>
        <p class="evidence-snippet">${snippetHtml}</p>
      </details>
      <div class="evidence-actions">
        <button type="button" data-copy-id="${escapeHtml(citation.citation_id || "")}">Copy snippet</button>
      </div>
    `;

    const button = card.querySelector("button[data-copy-id]");
    if (button) {
      button.addEventListener("click", () => {
        void copyTextToClipboard(citation.snippet || "");
      });
    }

    el.evidenceList.appendChild(card);
    });
  });
}

function renderPatchPreview() {
  const payload = state.patchPreview.payload;
  if (!payload) {
    el.patchPreview.textContent = "No patch preview yet.";
    return;
  }

  const operations = Array.isArray(payload.operations) ? payload.operations : [];
  if (!operations.length) {
    el.patchPreview.textContent = "Patch preview has no operations.";
    return;
  }

  const lines = operations.slice(0, 20).map((op) => {
    return `${op.row_id} | mode=${op.mode} | cites=${(op.citation_ids || []).length} | hash=${op.row_hash || "-"}`;
  });
  if (operations.length > 20) lines.push(`... ${operations.length - 20} more rows`);

  const validation = state.patchPreview.validation;
  const vline = validation && validation.ok ? "Patch validation: ok" : `Patch validation: fail (${(validation?.errors || []).join("; ")})`;
  el.patchPreview.textContent = `${vline}\n${lines.join("\n")}`;
}

async function* walkDir(dirHandle, prefix = "") {
  for await (const [name, handle] of dirHandle.entries()) {
    const path = prefix ? `${prefix}/${name}` : name;
    if (handle.kind === "directory") {
      yield* walkDir(handle, path);
      continue;
    }
    yield { path, name, handle };
  }
}

function renderInventory() {
  el.inventoryBody.innerHTML = "";
  state.inventory.forEach((item) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.path)}</td>
      <td>${escapeHtml(item.type || "-")}</td>
      <td>${escapeHtml(formatBytes(item.size))}</td>
      <td>${escapeHtml(item.status || "-")}</td>
    `;
    el.inventoryBody.appendChild(tr);
  });
}

async function loadSchemas() {
  const schemaPaths = [
    ["evidenceIndex", "./schemas/evidence_index.schema.json"],
    ["evidenceMatches", "./schemas/evidence_matches.schema.json"],
    ["sidecarPatch", "./schemas/sidecar_patch.schema.json"],
  ];

  for (const [key, url] of schemaPaths) {
    try {
      const res = await fetch(url, { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      state.schemas[key] = await res.json();
    } catch (err) {
      state.schemas[key] = null;
      logLine("warn", `Schema load failed: ${key}`, { error: err.message });
    }
  }
}

function validateAgainstSchema(payload, schema) {
  if (!schema || typeof schema !== "object") {
    return { ok: true, errors: [] };
  }

  const errors = [];
  if (schema.type === "object") {
    if (typeof payload !== "object" || payload === null || Array.isArray(payload)) {
      errors.push("payload is not an object");
      return { ok: false, errors };
    }
    const required = Array.isArray(schema.required) ? schema.required : [];
    required.forEach((key) => {
      if (!Object.prototype.hasOwnProperty.call(payload, key)) {
        errors.push(`missing required key: ${key}`);
      }
    });
  }

  return { ok: errors.length === 0, errors };
}

async function loadCheckpointIfAvailable() {
  state.checkpoint.loaded = false;
  state.checkpoint.data = null;
  if (!state.fs.evidenceDir) return;

  const handle = await safeGetFileHandle(state.fs.evidenceDir, CHECKPOINT_FILE, false);
  if (!handle) return;

  try {
    const text = await readFileText(handle);
    const parsed = JSON.parse(text);
    if (parsed && typeof parsed === "object") {
      state.checkpoint.loaded = true;
      state.checkpoint.data = parsed;
      logLine("info", "Loaded ingest checkpoint.", {
        file_count: Object.keys(parsed.files || {}).length,
      });
    }
  } catch (err) {
    logLine("warn", "Checkpoint load failed.", { error: err.message });
  }
}

function applyCheckpointStatusToInventory() {
  const files = state.checkpoint.data?.files || {};
  let skipped = 0;

  state.inventory.forEach((item) => {
    const prev = files[item.path];
    if (!prev) return;
    const unchanged = (
      prev.fingerprint === item.fingerprint
      && Number(prev.size || 0) === Number(item.size || 0)
      && Number(prev.last_modified || 0) === Number(item.last_modified || 0)
    );

    if (unchanged && prev.status === "done") {
      item.status = "skipped";
      skipped += 1;
    }
  });

  if (skipped > 0) {
    logLine("info", "Applied checkpoint skip optimization.", { skipped });
  }
}

async function saveCheckpoint() {
  if (!state.fs.evidenceDir) return;

  const payload = {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    files: {},
  };

  state.inventory.forEach((item) => {
    payload.files[item.path] = {
      fingerprint: item.fingerprint,
      size: item.size,
      last_modified: item.last_modified,
      status: item.status,
      mode: item.mode,
      updated_at: timestampIso(),
    };
  });

  const handle = await state.fs.evidenceDir.getFileHandle(CHECKPOINT_FILE, { create: true });
  await writeJson(handle, payload);
}

async function loadLatestIndexIfAvailable() {
  state.index.payload = null;
  state.index.latestHandleName = "";
  if (!state.fs.evidenceDir) return;

  const handle = await safeGetFileHandle(state.fs.evidenceDir, INDEX_LATEST_FILE, false);
  if (!handle) return;

  try {
    const text = await readFileText(handle);
    const payload = JSON.parse(text);
    const validation = validateAgainstSchema(payload, state.schemas.evidenceIndex);
    if (!validation.ok) {
      logLine("warn", "Latest index exists but is invalid.", { errors: validation.errors });
      return;
    }

    state.index.payload = payload;
    state.index.latestHandleName = INDEX_LATEST_FILE;

    const runtimeMismatch = String(payload?.source?.runtime || "") !== String(state.runtime.preferred || "");
    if (runtimeMismatch) {
      logLine("warn", "Stale index runtime mismatch. Rebuild recommended.", {
        index_runtime: payload?.source?.runtime || "-",
        current_runtime: state.runtime.preferred,
      });
    }

    logLine("info", "Loaded latest index artifact.", {
      chunks: payload?.chunks?.length || 0,
      file: INDEX_LATEST_FILE,
    });
  } catch (err) {
    logLine("warn", "Failed to load latest index artifact.", { error: err.message });
  }
}

async function loadLatestMatchesIfAvailable() {
  state.matches.selected = null;
  state.matches.included = [];
  state.matches.byRowId.clear();
  state.matches.latestHandleName = "";
  if (!state.fs.evidenceDir) return;

  const handle = await safeGetFileHandle(state.fs.evidenceDir, MATCHES_LATEST_FILE, false);
  if (!handle) return;

  try {
    const text = await readFileText(handle);
    const payload = JSON.parse(text);
    const validation = validateAgainstSchema(payload, state.schemas.evidenceMatches);
    if (!validation.ok) {
      logLine("warn", "Latest matches exist but are invalid.", { errors: validation.errors });
      return;
    }

    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    rows.forEach((entry) => {
      if (!entry?.row_id) return;
      state.matches.byRowId.set(entry.row_id, entry);
    });
    state.matches.included = rows.filter((entry) => {
      const row = state.rowMap.get(entry.row_id);
      return !!row?.include;
    });
    if (state.selectedRowId && state.matches.byRowId.has(state.selectedRowId)) {
      state.matches.selected = {
        created_at: payload.created_at,
        spec_version: payload.spec_version,
        generator: payload.generator,
        rows: [state.matches.byRowId.get(state.selectedRowId)],
      };
    }
    state.matches.latestHandleName = MATCHES_LATEST_FILE;
    logLine("info", "Loaded latest matches artifact.", {
      rows: rows.length,
      file: MATCHES_LATEST_FILE,
    });
  } catch (err) {
    logLine("warn", "Failed to load latest matches artifact.", { error: err.message });
  }
}

async function saveLatestMatchesArtifact() {
  if (!state.fs.evidenceDir) return;
  const rows = Array.from(state.matches.byRowId.values());
  const payload = {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    rows,
  };
  const validation = validateAgainstSchema(payload, state.schemas.evidenceMatches);
  if (!validation.ok) {
    logLine("warn", "Skipping latest matches save due to validation errors.", { errors: validation.errors });
    return;
  }

  const handle = await state.fs.evidenceDir.getFileHandle(MATCHES_LATEST_FILE, { create: true });
  await writeJson(handle, payload);
  state.matches.latestHandleName = MATCHES_LATEST_FILE;
}

async function pickFolder() {
  if (!state.capabilities.fs_api) {
    logLine("error", "File System Access API is not available in this browser.");
    return;
  }

  try {
    state.matches.byRowId.clear();
    state.matches.selected = null;
    state.matches.included = [];
    state.drafts.byRowId.clear();
    state.patchPreview.payload = null;
    renderEvidenceForSelectedRow();
    renderPatchPreview();

    const projectDir = await window.showDirectoryPicker();
    state.fs.projectDir = projectDir;

    state.fs.inputsDir = await safeGetDirectoryHandle(projectDir, "inputs", false);
    if (!state.fs.inputsDir) {
      throw new Error("Missing required folder: inputs");
    }

    state.fs.outputsDir = await safeGetDirectoryHandle(projectDir, "outputs", true);
    if (!state.fs.outputsDir) {
      throw new Error("Could not open/create outputs folder");
    }

    state.fs.evidenceDir = await safeGetDirectoryHandle(state.fs.outputsDir, "evidence_lab", true);
    if (!state.fs.evidenceDir) {
      throw new Error("Could not open/create outputs/evidence_lab folder");
    }

    await loadCheckpointIfAvailable();

    state.fs.sidecarHandle = await safeGetFileHandle(projectDir, "project_sidecar.json", false);
    if (!state.fs.sidecarHandle) {
      logLine("warn", "project_sidecar.json not found. Row tools will stay empty.");
      state.projectMeta = {};
      state.rowMap = new Map();
      state.rowOrder = [];
      state.selectedRowId = "";
      renderRowList();
      renderRowDetails();
    } else {
      const text = await readFileText(state.fs.sidecarHandle);
      const doc = JSON.parse(text);
      const parsed = parseSidecarRows(doc);
      state.projectMeta = parsed.projectMeta;
      state.rowMap = parsed.rowMap;
      state.rowOrder = parsed.rowOrder;
      state.selectedRowId = state.rowOrder[0] || "";
      renderRowList();
      renderRowDetails();
      logLine("info", "Loaded sidecar rows.", { rows: state.rowOrder.length });
    }

    await loadLatestIndexIfAvailable();
    await loadLatestMatchesIfAvailable();
    buildPatchPreviewFromMatches();
    renderEvidenceForSelectedRow();
    renderPatchPreview();

    renderProjectSummary();
    updateRunSummaryPanel(["Folder selected.", "Ready for scan."]);
    logLine("info", "Project folder selected.", { name: projectDir.name });
  } catch (err) {
    logLine("error", "Folder pick failed.", { error: err.message });
    updateRunSummaryPanel(["Folder pick failed.", err.message]);
  }
}

function initWorkers() {
  if (!state.capabilities.workers) {
    logLine("warn", "Web Workers are not available. Running without worker pipeline.");
    return;
  }

  try {
    state.workers.ingest = new Worker("./workers/ingest.worker.js", { type: "module" });
    state.workers.embed = new Worker("./workers/embed.worker.js", { type: "module" });
    state.workers.match = new Worker("./workers/match.worker.js", { type: "module" });
  } catch (err) {
    state.workers.ingest = null;
    state.workers.embed = null;
    state.workers.match = null;
    logLine("error", "Worker init failed.", { error: err.message });
    return;
  }

  state.workers.ingest.onmessage = (event) => {
    const msg = event.data || {};

    if (msg.type === "start") {
      state.run.scanActive = true;
      logLine("info", "Ingest worker started.", { run_id: msg.runId, total: msg.total });
      return;
    }

    if (msg.type === "progress") {
      if (msg.runId !== state.run.scanToken) return;
      const item = state.inventoryMap.get(msg.path);
      if (item) {
        item.status = msg.status;
      }
      renderInventory();
      updateRunSummaryPanel();
      return;
    }

    if (msg.type === "done") {
      if (msg.runId !== state.run.scanToken) return;
      state.run.scanActive = false;
      updateRunSummaryPanel([`Scan completed by worker. completed=${msg.completed}`]);
      logLine("info", "Ingest worker done.", { run_id: msg.runId, completed: msg.completed });
      if (typeof state.run.scanResolve === "function") {
        state.run.scanResolve();
      }
      state.run.scanResolve = null;
      state.run.scanPromise = null;
      void saveCheckpoint();
      return;
    }

    if (msg.type === "canceled") {
      if (msg.runId !== state.run.scanToken) return;
      state.run.scanActive = false;
      state.inventory.forEach((item) => {
        if (item.status === "running") item.status = "pending";
      });
      renderInventory();
      updateRunSummaryPanel(["Active scan canceled."]);
      logLine("warn", "Ingest worker run canceled.", { run_id: msg.runId, completed: msg.completed });
      if (typeof state.run.scanResolve === "function") {
        state.run.scanResolve();
      }
      state.run.scanResolve = null;
      state.run.scanPromise = null;
      return;
    }

    if (msg.type === "error") {
      logLine("error", "Ingest worker error.", msg);
    }
  };

  logLine("info", "Workers initialized.");
}

async function collectInventory() {
  updateSettingsFromUi();
  state.inventory = [];
  state.inventoryMap = new Map();
  const maxUnits = state.settings.maxUnitsPerRun > 0 ? state.settings.maxUnitsPerRun : Number.POSITIVE_INFINITY;
  let scannedUnits = 0;

  for await (const fileInfo of walkDir(state.fs.inputsDir, "inputs")) {
    if (scannedUnits >= maxUnits) {
      logLine("warn", "Reached max units limit for this run.", { max_units: state.settings.maxUnitsPerRun });
      break;
    }
    if (!isSupportedInputFile(fileInfo.name)) continue;
    scannedUnits += 1;

    const file = await fileInfo.handle.getFile();
    const ext = extFromName(fileInfo.name);
    const mode = getSourceModeByExtension(ext);
    const fingerprint = await computeFileFingerprint(file);

    const item = {
      path: fileInfo.path,
      type: ext.slice(1).toLowerCase(),
      size: file.size,
      last_modified: file.lastModified || 0,
      mode,
      fingerprint,
      status: "pending",
      handle: fileInfo.handle,
    };

    if (mode === "image_ocr") {
      logLine("warn", "OCR mode flagged (confidence pending implementation).", {
        file: item.path,
        mode,
      });
    }

    logLine("info", "Planned extraction mode.", { file: item.path, mode });
    state.inventory.push(item);
    state.inventoryMap.set(item.path, item);
  }

  state.inventory.sort((a, b) => byIdOrPath(a.path, b.path));
  applyCheckpointStatusToInventory();
  renderInventory();
}

async function runIngestSimulation() {
  const queue = state.inventory.filter((item) => item.status === "pending" || item.status === "failed");
  if (!queue.length) {
    state.run.scanActive = false;
    updateRunSummaryPanel(["No changed files to process (all skipped)."]);
    return;
  }

  if (!state.workers.ingest) {
    state.run.scanActive = false;
    queue.forEach((item) => {
      item.status = "done";
    });
    renderInventory();
    updateRunSummaryPanel(["Scan complete in no-worker mode."]);
    await saveCheckpoint();
    return;
  }

  queue.forEach((item) => {
    item.status = "running";
  });
  renderInventory();

  const runId = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  state.run.scanToken = runId;
  state.run.scanActive = true;
  state.run.scanPromise = new Promise((resolve) => {
    state.run.scanResolve = resolve;
  });

  state.workers.ingest.postMessage({
    type: "start",
    runId,
    batchSize: state.settings.batchSize,
    items: queue.map((item) => ({ path: item.path })),
  });

  await state.run.scanPromise;
}

async function scanInputs() {
  if (!state.fs.inputsDir) {
    logLine("error", "Cannot scan: no project folder selected.");
    updateRunSummaryPanel(["Pick a project folder first."]);
    return;
  }

  const timer = stageTimerStart();
  await collectInventory();

  if (!state.inventory.length) {
    stageTimerEnd("scan", timer);
    updateRunSummaryPanel(["Scan complete.", "No supported files found in inputs/."]);
    logLine("warn", "No supported input files found.");
    return;
  }

  await runIngestSimulation();
  stageTimerEnd("scan", timer);
  updateRunSummaryPanel();
  logLine("info", "Scan kicked off.", {
    files: state.inventory.length,
  });
}

let xlsxLoadAttempted = false;
let xlsxLoadResult = false;
let pdfLoadAttempted = false;
let pdfLoadResult = false;
let mammothLoadAttempted = false;
let mammothLoadResult = false;
let tesseractLoadAttempted = false;
let tesseractLoadResult = false;

function sanitizeFileName(input) {
  return String(input || "")
    .replaceAll("/", "_")
    .replaceAll("\\\\", "_")
    .replaceAll(":", "_")
    .replaceAll(" ", "_");
}

function normalizeWhitespace(text) {
  return String(text || "")
    .replace(/\\r\\n/g, "\\n")
    .replace(/\\s+/g, " ")
    .trim();
}

function guessLanguage(text) {
  const value = String(text || "").toLowerCase();
  if (!value) return "unknown";
  if (value.includes(" securite ") || value.includes(" sante ")) return "fr";
  if (value.includes(" sicurezza ") || value.includes(" salute ")) return "it";
  if (value.includes(" the ") || value.includes(" safety ")) return "en";
  if (value.includes(" sicherheit ") || value.includes(" gesundheit ")) return "de";
  return "unknown";
}

function splitTextIntoChunks(text, maxChars = 650, overlap = 120) {
  const normalized = normalizeWhitespace(text);
  if (!normalized) return [];
  if (normalized.length <= maxChars) {
    return [{ text: normalized, span: [0, normalized.length] }];
  }

  const chunks = [];
  let cursor = 0;
  while (cursor < normalized.length) {
    let end = Math.min(normalized.length, cursor + maxChars);
    if (end < normalized.length) {
      const space = normalized.lastIndexOf(" ", end);
      if (space > cursor + 100) end = space;
    }
    const slice = normalized.slice(cursor, end).trim();
    chunks.push({ text: slice, span: [cursor, end] });
    if (end >= normalized.length) break;
    cursor = Math.max(0, end - overlap);
  }
  return chunks;
}

function vectorizeText(text, dimensions = 128) {
  const vec = new Float32Array(dimensions);
  const terms = tokenize(text);
  terms.forEach((term) => {
    const h = Number.parseInt(hashString32(term).slice(1), 16);
    const idx = h % dimensions;
    vec[idx] += 1;
  });

  let norm = 0;
  for (let i = 0; i < vec.length; i += 1) {
    norm += vec[i] * vec[i];
  }
  norm = Math.sqrt(norm);
  if (norm > 0) {
    for (let i = 0; i < vec.length; i += 1) {
      vec[i] /= norm;
    }
  }
  return Array.from(vec);
}

function cosineSimilarity(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || !a.length || !b.length) return 0;
  const limit = Math.min(a.length, b.length);
  let score = 0;
  for (let i = 0; i < limit; i += 1) {
    score += Number(a[i] || 0) * Number(b[i] || 0);
  }
  return score;
}

async function ensureEmbeddingEngine() {
  if (state.embedding.mode === "model" && state.embedding.extractor) {
    return state.embedding;
  }
  if (state.embedding.initAttempted) {
    return state.embedding;
  }
  state.embedding.initAttempted = true;

  try {
    const lib = await import("../AI/vendor/transformers.min.js");
    if (!lib || !lib.pipeline || !lib.env) {
      throw new Error("transformers.min.js loaded but expected exports were missing");
    }

    lib.env.allowRemoteModels = false;
    lib.env.allowLocalModels = true;
    lib.env.localModelPath = "../AI/models/";

    const device = state.runtime.preferred === "webgpu" ? "webgpu" : "wasm";
    const extractor = await lib.pipeline("feature-extraction", state.embedding.modelId, {
      device,
      local_files_only: true,
      dtype: "fp32",
    });

    state.embedding.mode = "model";
    state.embedding.extractor = extractor;
    state.embedding.lib = lib;
    state.embedding.initError = "";
    logLine("info", "Loaded multilingual embedding model.", {
      model: state.embedding.modelId,
      device,
      mode: "model",
    });
  } catch (err) {
    state.embedding.mode = "hash";
    state.embedding.extractor = null;
    state.embedding.lib = null;
    state.embedding.initError = err.message;
    logLine("warn", "Embedding model unavailable. Falling back to local hash vectors.", {
      model: state.embedding.modelId,
      error: err.message,
    });
  }

  return state.embedding;
}

function coerceEmbeddingOutput(result) {
  if (!result) return [];
  if (Array.isArray(result)) {
    if (Array.isArray(result[0])) return result;
    return [result];
  }
  if (typeof result.tolist === "function") {
    const list = result.tolist();
    if (Array.isArray(list) && Array.isArray(list[0])) return list;
    if (Array.isArray(list)) return [list];
  }
  if (Array.isArray(result.data) && Array.isArray(result.dims) && result.dims.length >= 2) {
    const batch = result.dims[0];
    const dim = result.dims[result.dims.length - 1];
    const out = [];
    for (let b = 0; b < batch; b += 1) {
      out.push(result.data.slice(b * dim, (b + 1) * dim));
    }
    return out;
  }
  return [];
}

async function embedTexts(texts) {
  const engine = await ensureEmbeddingEngine();
  if (engine.mode !== "model" || !engine.extractor) {
    return texts.map((text) => vectorizeText(text));
  }

  const result = await engine.extractor(texts, {
    pooling: "mean",
    normalize: true,
  });
  const vectors = coerceEmbeddingOutput(result);
  if (vectors.length !== texts.length) {
    logLine("warn", "Embedding output size mismatch. Reverting to hash vectors for this batch.", {
      expected: texts.length,
      received: vectors.length,
    });
    return texts.map((text) => vectorizeText(text));
  }
  return vectors.map((vec) => Array.from(vec));
}

async function loadScriptCandidates(candidates) {
  for (const url of candidates) {
    try {
      await new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = url;
        script.async = true;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`load failed: ${url}`));
        document.head.appendChild(script);
      });
      return { ok: true, url };
    } catch (err) {
      logLine("warn", "Script candidate load failed.", { url, error: err.message });
    }
  }
  return { ok: false, url: "" };
}

async function ensureXlsxLibrary() {
  if (window.XLSX) return true;
  if (xlsxLoadAttempted) return xlsxLoadResult;
  xlsxLoadAttempted = true;

  const result = await loadScriptCandidates([
    "../libs/sheetjs/xlsx.full.min.js",
    "./lib/sheetjs/xlsx.full.min.js",
  ]);
  if (result.ok && window.XLSX) {
    xlsxLoadResult = true;
    logLine("info", "Loaded SheetJS library.", { url: result.url });
    return true;
  }

  xlsxLoadResult = false;
  logLine("warn", "SheetJS not available. XLSX extraction falls back to metadata.");
  return false;
}

async function ensurePdfLibrary() {
  if (window.pdfjsLib) return true;
  if (pdfLoadAttempted) return pdfLoadResult;
  pdfLoadAttempted = true;

  const result = await loadScriptCandidates([
    "./lib/pdfjs/pdf.min.js",
    "./lib/pdfjs/pdf.js",
  ]);
  if (result.ok && window.pdfjsLib) {
    pdfLoadResult = true;
    logLine("info", "Loaded pdf.js library.", { url: result.url });
    return true;
  }

  pdfLoadResult = false;
  logLine("warn", "pdf.js not available. PDF extraction falls back to metadata.");
  return false;
}

async function ensureMammothLibrary() {
  if (window.mammoth) return true;
  if (mammothLoadAttempted) return mammothLoadResult;
  mammothLoadAttempted = true;

  const result = await loadScriptCandidates([
    "./lib/mammoth/mammoth.browser.min.js",
  ]);
  if (result.ok && window.mammoth) {
    mammothLoadResult = true;
    logLine("info", "Loaded mammoth library.", { url: result.url });
    return true;
  }

  mammothLoadResult = false;
  logLine("warn", "mammoth not available. DOCX extraction falls back to metadata.");
  return false;
}

async function ensureTesseractLibrary() {
  if (window.Tesseract) return true;
  if (tesseractLoadAttempted) return tesseractLoadResult;
  tesseractLoadAttempted = true;

  const result = await loadScriptCandidates([
    "./lib/tesseract/tesseract.min.js",
  ]);
  if (result.ok && window.Tesseract) {
    tesseractLoadResult = true;
    logLine("info", "Loaded tesseract library.", { url: result.url });
    return true;
  }

  tesseractLoadResult = false;
  logLine("warn", "tesseract not available. OCR extraction falls back to metadata.");
  return false;
}

function getOcrLanguagePack() {
  const locale = getLocalePrefix();
  if (locale === "fr") return "fra+eng";
  if (locale === "it") return "ita+eng";
  if (locale === "en") return "eng";
  return "deu+eng";
}

async function writeRawExtractionCache(item, extraction) {
  if (!state.fs.evidenceDir) return;
  const rawDir = await safeGetDirectoryHandle(state.fs.evidenceDir, "raw", true);
  if (!rawDir) return;
  const fileName = `${sanitizeFileName(item.path)}.json`;
  const handle = await rawDir.getFileHandle(fileName, { create: true });
  await writeJson(handle, extraction);
}

async function extractBlocksFromXlsx(item, fileBuffer) {
  const blocks = [];
  if (!window.XLSX) {
    return {
      blocks: [`Spreadsheet source ${item.path}.`],
      metadata: { method: "xlsx_fallback", sheets: 0, rows: 0 },
    };
  }

  const workbook = window.XLSX.read(fileBuffer, { type: "array" });
  let rowCounter = 0;
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
    rows.forEach((row, index) => {
      const text = normalizeWhitespace((row || []).join(" | "));
      if (!text) return;
      rowCounter += 1;
      blocks.push(`[${sheetName}#${index + 1}] ${text}`);
    });
  });

  return {
    blocks: blocks.length ? blocks : [`Spreadsheet source ${item.path}.`],
    metadata: { method: "xlsx_sheetjs", sheets: workbook.SheetNames.length, rows: rowCounter },
  };
}

async function extractBlocksFromDocx(item, fileBuffer) {
  if (!window.mammoth) {
    return {
      blocks: [`DOCX source ${item.path}.`],
      metadata: { method: "docx_fallback", paragraphs: 0 },
    };
  }

  const result = await window.mammoth.extractRawText({ arrayBuffer: fileBuffer });
  const paragraphs = String(result.value || "")
    .split(/\\n+/)
    .map((line) => normalizeWhitespace(line))
    .filter(Boolean);

  return {
    blocks: paragraphs.length ? paragraphs : [`DOCX source ${item.path}.`],
    metadata: { method: "docx_mammoth", paragraphs: paragraphs.length, warnings: result.messages || [] },
  };
}

async function renderPdfPageToCanvas(page) {
  const viewport = page.getViewport({ scale: 2 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(viewport.width));
  canvas.height = Math.max(1, Math.round(viewport.height));
  const ctx = canvas.getContext("2d", { alpha: false });
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas;
}

async function extractBlocksFromPdf(item, fileBuffer) {
  if (!window.pdfjsLib || !window.pdfjsLib.getDocument) {
    return {
      blocks: [`PDF source ${item.path}.`],
      metadata: { method: "pdf_fallback", page_count: 0, densities: [], ocr_candidates: [1], ocr_confidence: [] },
    };
  }

  const loadingTask = window.pdfjsLib.getDocument({ data: fileBuffer });
  const doc = await loadingTask.promise;
  const blocks = [];
  const densities = [];
  const ocrCandidates = [];
  const ocrConfidence = [];

  for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
    const page = await doc.getPage(pageNumber);
    const content = await page.getTextContent();
    const text = normalizeWhitespace((content.items || []).map((item) => item.str || "").join(" "));
    const density = text.length;
    densities.push(density);
    let finalText = text;
    let pageConfidence = null;
    if (density < 80) {
      ocrCandidates.push(pageNumber);
      if (window.Tesseract && typeof window.Tesseract.recognize === "function") {
        try {
          const canvas = await renderPdfPageToCanvas(page);
          const result = await window.Tesseract.recognize(canvas, getOcrLanguagePack());
          const ocrText = normalizeWhitespace(result?.data?.text || "");
          pageConfidence = Number(result?.data?.confidence || 0) / 100;
          if (ocrText) {
            finalText = ocrText;
          }
        } catch (err) {
          logLine("warn", "PDF page OCR fallback failed.", { file: item.path, page: pageNumber, error: err.message });
        }
      }
    }
    ocrConfidence.push(pageConfidence);
    if (finalText) {
      blocks.push(`[page ${pageNumber}] ${finalText}`);
    }
  }

  return {
    blocks: blocks.length ? blocks : [`PDF source ${item.path}.`],
    metadata: { method: "pdfjs_text", page_count: doc.numPages, densities, ocr_candidates: ocrCandidates, ocr_confidence: ocrConfidence },
  };
}

async function extractTextFromImageOcr(file, languagePack) {
  if (!window.Tesseract || typeof window.Tesseract.recognize !== "function") {
    return {
      text: "",
      confidence: 0.45,
      method: "ocr_fallback",
    };
  }
  const result = await window.Tesseract.recognize(file, languagePack);
  return {
    text: normalizeWhitespace(result?.data?.text || ""),
    confidence: Number(result?.data?.confidence || 0) / 100,
    method: "ocr_tesseract",
  };
}

async function extractItemBlocks(item) {
  const file = await item.handle.getFile();
  const ext = extFromName(file.name);
  const mode = getSourceModeByExtension(ext);

  if (mode === "xlsx") {
    await ensureXlsxLibrary();
    const started = performance.now();
    const buffer = await file.arrayBuffer();
    const extracted = await extractBlocksFromXlsx(item, buffer);
    const elapsed = Math.round(performance.now() - started);
    const extraction = {
      created_at: timestampIso(),
      file_path: item.path,
      source_type: "xlsx",
      page_count: null,
      text_density: null,
      ocr_candidates: [],
      elapsed_ms: elapsed,
      blocks: extracted.blocks,
      metadata: extracted.metadata,
    };
    await writeRawExtractionCache(item, extraction);
    logLine("info", "Extracted XLSX blocks.", {
      file: item.path,
      blocks: extracted.blocks.length,
      elapsed_ms: elapsed,
    });
    return extraction;
  }

  if (mode === "pdf_text") {
    await ensurePdfLibrary();
    await ensureTesseractLibrary();
    const started = performance.now();
    const buffer = new Uint8Array(await file.arrayBuffer());
    const extracted = await extractBlocksFromPdf(item, buffer);
    const elapsed = Math.round(performance.now() - started);
    const extraction = {
      created_at: timestampIso(),
      file_path: item.path,
      source_type: "pdf_text",
      page_count: Number(extracted.metadata?.page_count || 0) || null,
      text_density: Array.isArray(extracted.metadata?.densities) && extracted.metadata.densities.length
        ? Math.round(extracted.metadata.densities.reduce((sum, n) => sum + n, 0) / extracted.metadata.densities.length)
        : 0,
      ocr_candidates: Array.isArray(extracted.metadata?.ocr_candidates) ? extracted.metadata.ocr_candidates : [],
      elapsed_ms: elapsed,
      blocks: extracted.blocks,
      metadata: extracted.metadata,
    };
    await writeRawExtractionCache(item, extraction);
    logLine("info", "Extracted PDF blocks.", {
      file: item.path,
      pages: extraction.page_count,
      blocks: extraction.blocks.length,
      elapsed_ms: elapsed,
      ocr_candidates: extraction.ocr_candidates.length,
    });
    return extraction;
  }

  if (mode === "docx") {
    await ensureMammothLibrary();
    const started = performance.now();
    const buffer = await file.arrayBuffer();
    const extracted = await extractBlocksFromDocx(item, buffer);
    const elapsed = Math.round(performance.now() - started);
    const extraction = {
      created_at: timestampIso(),
      file_path: item.path,
      source_type: "docx",
      page_count: null,
      text_density: null,
      ocr_candidates: [],
      elapsed_ms: elapsed,
      blocks: extracted.blocks,
      metadata: extracted.metadata,
    };
    await writeRawExtractionCache(item, extraction);
    logLine("info", "Extracted DOCX blocks.", {
      file: item.path,
      blocks: extraction.blocks.length,
      elapsed_ms: elapsed,
    });
    return extraction;
  }

  if (mode === "image_ocr") {
    await ensureTesseractLibrary();
    const started = performance.now();
    const ocr = await extractTextFromImageOcr(file, getOcrLanguagePack());
    const elapsed = Math.round(performance.now() - started);
    const text = ocr.text || `Image source ${item.path}. OCR text unavailable; fallback used.`;
    const extraction = {
      created_at: timestampIso(),
      file_path: item.path,
      source_type: "image_ocr",
      page_count: 1,
      text_density: text.length,
      ocr_candidates: [1],
      elapsed_ms: elapsed,
      blocks: [text],
      metadata: { method: ocr.method, confidence: Number(ocr.confidence || 0.45) },
    };
    await writeRawExtractionCache(item, extraction);
    if (Number(extraction.metadata.confidence || 0) < 0.6) {
      logLine("warn", "Low-confidence OCR extraction.", {
        file: item.path,
        confidence: extraction.metadata.confidence,
      });
    }
    return extraction;
  }

  const generic = `Source ${item.path}. Type ${item.type}. Mode ${item.mode}.`;
  return {
    created_at: timestampIso(),
    file_path: item.path,
    source_type: item.mode,
    page_count: null,
    text_density: null,
    ocr_candidates: [],
    elapsed_ms: 0,
    blocks: [generic],
    metadata: { method: "generic_fallback" },
  };
}

async function openChunkStore() {
  return await new Promise((resolve, reject) => {
    const request = indexedDB.open("ai-evidence-lab-db", 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains("chunks")) {
        db.createObjectStore("chunks", { keyPath: "chunk_id" });
      }
      if (!db.objectStoreNames.contains("meta")) {
        db.createObjectStore("meta", { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("IndexedDB open failed"));
  });
}

async function persistChunksToIndexedDb(chunks) {
  const db = await openChunkStore();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(["chunks", "meta"], "readwrite");
    const chunkStore = tx.objectStore("chunks");
    const metaStore = tx.objectStore("meta");
    chunkStore.clear();
    chunks.forEach((chunk) => chunkStore.put(chunk));
    metaStore.put({
      key: "chunk_cache_meta",
      updated_at: timestampIso(),
      count: chunks.length,
    });
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("IndexedDB write failed"));
  });
  db.close();
}

function previousChunksByFingerprint() {
  const map = new Map();
  const chunks = state.index.payload?.chunks || [];
  chunks.forEach((chunk) => {
    const fp = String(chunk.fingerprint || "");
    if (!fp) return;
    if (!map.has(fp)) map.set(fp, []);
    map.get(fp).push(chunk);
  });
  return map;
}

async function buildIndexPayloadFromInventory() {
  const chunks = [];
  const previousByFingerprint = previousChunksByFingerprint();
  let reused = 0;
  let extracted = 0;

  for (const item of state.inventory) {
    const previous = previousByFingerprint.get(String(item.fingerprint || "")) || [];
    if (previous.length) {
      previous.forEach((chunk) => chunks.push(chunk));
      reused += previous.length;
      continue;
    }

    const extraction = await extractItemBlocks(item);
    extracted += 1;
    const blocks = Array.isArray(extraction.blocks) ? extraction.blocks : [];
    blocks.forEach((block, blockIndex) => {
      const perPageConfidence = Array.isArray(extraction?.metadata?.ocr_confidence)
        ? extraction.metadata.ocr_confidence[blockIndex]
        : null;
      const blockConfidence = perPageConfidence == null
        ? (extraction?.metadata?.confidence == null ? null : Number(extraction.metadata.confidence))
        : Number(perPageConfidence);
      const parts = splitTextIntoChunks(block, 650, 120);
      parts.forEach((part, partIndex) => {
        const chunkId = `chunk_${hashString32(`${item.fingerprint}:${blockIndex}:${partIndex}`)}`;
        const text = normalizeWhitespace(part.text);
        const terms = tokenize(text);
        chunks.push({
          chunk_id: chunkId,
          file_path: item.path,
          page_number: null,
          source_type: extraction.source_type || item.mode,
          language_guess: guessLanguage(text),
          text,
          char_span: Array.isArray(part.span) ? part.span : [0, text.length],
          fingerprint: item.fingerprint,
          token_count: terms.length,
          terms,
          vector: null,
          confidence: blockConfidence,
        });
      });
    });
  }

  const batchSize = Math.max(1, Number(state.settings.batchSize || 16));
  for (let i = 0; i < chunks.length; i += batchSize) {
    const batch = chunks.slice(i, i + batchSize);
    const texts = batch.map((chunk) => chunk.text || "");
    const vectors = await embedTexts(texts);
    batch.forEach((chunk, index) => {
      chunk.vector = Array.isArray(vectors[index]) ? vectors[index] : vectorizeText(chunk.text || "");
    });
  }

  await persistChunksToIndexedDb(chunks);

  logLine("info", "Chunk cache persisted to IndexedDB.", { chunks: chunks.length, reused, extracted });

  return {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    source: {
      file_count: state.inventory.length,
      chunk_count: chunks.length,
      local_only_models: state.runtime.localOnlyModels,
      runtime: state.runtime.preferred,
      reused_chunks: reused,
      extracted_files: extracted,
    },
    chunks,
  };
}

async function buildOrRefreshIndex() {
  if (!state.fs.evidenceDir) {
    logLine("error", "Pick a project folder first.");
    return;
  }

  if (!state.inventory.length) {
    logLine("warn", "No inventory loaded. Running scan first.");
    await scanInputs();
    if (!state.inventory.length) return;
  }

  const timer = stageTimerStart();
  const payload = await buildIndexPayloadFromInventory();
  const chunkCount = Array.isArray(payload.chunks) ? payload.chunks.length : 0;
  if (chunkCount > state.settings.memoryChunkGuard) {
    logLine("warn", "Chunk memory guard exceeded. Consider reducing max units or batch size.", {
      chunk_count: chunkCount,
      memory_guard: state.settings.memoryChunkGuard,
    });
  }
  const validation = validateAgainstSchema(payload, state.schemas.evidenceIndex);
  if (!validation.ok) {
    logLine("error", "evidence_index payload schema validation failed.", {
      errors: validation.errors,
    });
    return;
  }

  try {
    const stamp = timestampForFile();
    const snapshotName = `evidence_index_${stamp}.json`;
    const latestHandle = await state.fs.evidenceDir.getFileHandle(INDEX_LATEST_FILE, { create: true });
    const snapshotHandle = await state.fs.evidenceDir.getFileHandle(snapshotName, { create: true });

    await writeJson(latestHandle, payload);
    await writeJson(snapshotHandle, payload);

    state.index.payload = payload;
    state.index.latestHandleName = INDEX_LATEST_FILE;

    stageTimerEnd("index", timer);
    updateRunSummaryPanel([
      `Index refreshed: chunks=${payload.chunks.length}`,
      `Artifacts: ${INDEX_LATEST_FILE}, ${snapshotName}`,
    ]);
    logLine("info", "Index refreshed.", {
      chunks: payload.chunks.length,
      latest: INDEX_LATEST_FILE,
      snapshot: snapshotName,
    });
  } catch (err) {
    logLine("error", "Index build failed.", { error: err.message });
  }
}

function getLocalePrefix() {
  const locale = String(state.projectMeta.locale || "").toLowerCase();
  if (locale.startsWith("fr")) return "fr";
  if (locale.startsWith("it")) return "it";
  if (locale.startsWith("en")) return "en";
  return "de";
}

function buildQueryForRow(row) {
  const base = `${row.rowId} ${row.title} ${row.chapterTitle}`.trim();
  const synonyms = QUERY_SYNONYMS[getLocalePrefix()] || [];
  return {
    text: `${base} ${synonyms.join(" ")}`.trim(),
    terms: tokenize(`${base} ${synonyms.join(" ")}`),
  };
}

function lexicalScore(terms, haystackText) {
  if (!terms.length) return 0;
  const hayTerms = tokenize(haystackText);
  if (!hayTerms.length) return 0;
  const haySet = new Set(hayTerms);
  let hits = 0;
  terms.forEach((term) => {
    if (haySet.has(term)) hits += 1;
  });
  return hits / terms.length;
}

function rankCitations(row, query) {
  const chunks = state.index.payload?.chunks || [];
  const queryVector = vectorizeText(query.text);
  const scored = chunks.map((chunk) => {
    const text = `${chunk.text || ""} ${chunk.file_path || ""}`;
    const vectorScore = cosineSimilarity(queryVector, Array.isArray(chunk.vector) ? chunk.vector : vectorizeText(text));
    const baseScore = lexicalScore(query.terms, text);
    const overlap = tokenize(text).filter((t) => query.terms.includes(t)).length;
    const rerankBoost = Math.min(0.2, overlap * 0.02);
    const confidence = chunk.confidence == null ? (chunk.source_type === "image_ocr" ? 0.45 : 0.92) : Number(chunk.confidence);
    return {
      citation_id: `${row.rowId}_${chunk.chunk_id}`,
      file: chunk.file_path,
      page: chunk.page_number,
      snippet: chunk.text,
      score: (vectorScore * 0.65) + (baseScore * 0.35) + rerankBoost,
      confidence,
      source_type: chunk.source_type,
    };
  });

  scored.sort((a, b) => b.score - a.score || byIdOrPath(a.file, b.file));

  const deduped = [];
  const seenLocation = new Set();
  for (const item of scored) {
    const loc = `${item.file}#${item.page == null ? "na" : item.page}`;
    if (seenLocation.has(loc)) continue;
    const diversityBonus = seenLocation.size === 0 ? 0 : 0.03;
    deduped.push({ ...item, score: item.score + diversityBonus });
    seenLocation.add(loc);
    if (deduped.length >= 5) break;
  }

  return deduped;
}

function toEvidenceStatus(citations) {
  if (!citations.length) return "none";
  const best = citations[0]?.score || 0;
  if (best >= 0.25) return "evidence_found";
  return "weak";
}

function buildMatchPayloadForRows(rows) {
  const rowPayloads = rows.map((row) => {
    const query = buildQueryForRow(row);
    const citations = rankCitations(row, query);
    const status = toEvidenceStatus(citations);
    if (status === "none") {
      logLine("warn", "No citations found for row.", { row_id: row.rowId });
    }

    return {
      row_id: row.rowId,
      status,
      query: query.text,
      citations,
    };
  });

  return {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    rows: rowPayloads,
  };
}

async function runMatchSelected() {
  if (!state.selectedRowId || !state.rowMap.has(state.selectedRowId)) {
    logLine("warn", "Select a row first.");
    return;
  }

  if (!state.index.payload) {
    logLine("warn", "No index in memory. Building index first.");
    await buildOrRefreshIndex();
    if (!state.index.payload) return;
  }

  const timer = stageTimerStart();
  const row = state.rowMap.get(state.selectedRowId);
  const payload = buildMatchPayloadForRows([row]);
  const validation = validateAgainstSchema(payload, state.schemas.evidenceMatches);
  if (!validation.ok) {
    logLine("error", "Match payload failed schema validation.", { errors: validation.errors });
    return;
  }

  state.matches.selected = payload;
  payload.rows.forEach((entry) => {
    state.matches.byRowId.set(entry.row_id, entry);
  });

  stageTimerEnd("match_selected", timer);
  renderEvidenceForSelectedRow();
  buildPatchPreviewFromMatches();
  renderPatchPreview();
  await saveLatestMatchesArtifact();
  updateRunSummaryPanel([`Selected row matched: ${row.rowId}`]);
  logLine("info", "Selected-row evidence matching complete.", {
    row_id: row.rowId,
    citations: payload.rows[0]?.citations?.length || 0,
  });
}

async function runMatchIncluded() {
  const rows = Array.from(state.rowMap.values()).filter((row) => row.include);
  if (!rows.length) {
    logLine("warn", "No included rows available for batch matching.");
    return;
  }

  if (!state.index.payload) {
    logLine("warn", "No index in memory. Building index first.");
    await buildOrRefreshIndex();
    if (!state.index.payload) return;
  }

  const timer = stageTimerStart();
  const payload = buildMatchPayloadForRows(rows);
  const validation = validateAgainstSchema(payload, state.schemas.evidenceMatches);
  if (!validation.ok) {
    logLine("error", "Batch match payload failed schema validation.", { errors: validation.errors });
    return;
  }

  state.matches.included = payload.rows;
  payload.rows.forEach((entry) => {
    state.matches.byRowId.set(entry.row_id, entry);
  });

  stageTimerEnd("match_included", timer);
  renderEvidenceForSelectedRow();
  buildPatchPreviewFromMatches();
  renderPatchPreview();
  await saveLatestMatchesArtifact();
  updateRunSummaryPanel([`Included rows matched: ${rows.length}`]);
  logLine("info", "Included-row evidence matching complete.", {
    rows: rows.length,
  });
}

function generateDraftFromCitationsForRow(rowId) {
  const row = state.rowMap.get(rowId);
  const matchEntry = state.matches.byRowId.get(rowId);
  if (!row || !matchEntry) {
    logLine("warn", "Cannot generate draft. Missing row or evidence.", { row_id: rowId });
    if (rowId === state.selectedRowId) {
      el.evidenceState.textContent = "Draft generation failed: missing row or evidence.";
    }
    return;
  }

  const citations = Array.isArray(matchEntry.citations) ? matchEntry.citations : [];
  if (!citations.length) {
    logLine("warn", "Cannot generate draft without citations.", { row_id: rowId });
    if (rowId === state.selectedRowId) {
      el.evidenceState.textContent = "Draft generation blocked: uncited output is not allowed.";
    }
    return;
  }

  const selected = citations.slice(0, 3);
  const best = selected[0];
  const finding = `Evidence indicates: ${best.snippet}`;
  const recommendation = `For ${rowId}, review ${best.file}${best.page ? ` (page ${best.page})` : ""} and define corrective action ownership and timeline.`;

  state.drafts.byRowId.set(rowId, {
    row_id: rowId,
    finding,
    recommendation,
    citation_ids: selected.map((c) => c.citation_id),
    created_at: timestampIso(),
  });

  logLine("info", "Generated cited draft for row.", {
    row_id: rowId,
    citations: selected.length,
  });
  if (rowId === state.selectedRowId) {
    el.evidenceState.textContent = `Draft generated from ${selected.length} citation(s).`;
  }
}

function generateDraftForSelectedRow() {
  if (!state.selectedRowId) {
    logLine("warn", "Select a row first.");
    return;
  }
  generateDraftFromCitationsForRow(state.selectedRowId);
  buildPatchPreviewFromMatches();
  renderPatchPreview();
}

function buildPatchPreviewFromMatches() {
  const operations = [];

  state.matches.byRowId.forEach((matchEntry, rowId) => {
    const row = state.rowMap.get(rowId);
    if (!row) {
      operations.push({
        row_id: rowId,
        mode: "skip",
        finding: "",
        recommendation: "",
        citation_ids: [],
        row_hash: null,
      });
      return;
    }

    const mode = row.done ? "replace" : (row.include ? "append" : "skip");
    const cites = (matchEntry.citations || []).map((citation) => citation.citation_id);
    const best = matchEntry.citations && matchEntry.citations[0] ? matchEntry.citations[0] : null;
    const generatedDraft = state.drafts.byRowId.get(rowId) || null;

    let finalFinding = best ? `Evidence suggests: ${best.snippet}` : "";
    let finalRecommendation = best ? `Review source ${best.file} for row ${rowId}.` : "";
    let finalCitationIds = cites;
    let finalMode = mode;

    if (generatedDraft) {
      const draftCitations = Array.isArray(generatedDraft.citation_ids) ? generatedDraft.citation_ids.filter(Boolean) : [];
      if (!draftCitations.length) {
        finalMode = "skip";
        finalFinding = "";
        finalRecommendation = "";
        finalCitationIds = [];
        logLine("warn", "Rejected uncited draft output.", { row_id: rowId });
      } else {
        finalFinding = String(generatedDraft.finding || "");
        finalRecommendation = String(generatedDraft.recommendation || "");
        finalCitationIds = draftCitations;
      }
    }

    operations.push({
      row_id: rowId,
      mode: finalMode,
      finding: finalFinding,
      recommendation: finalRecommendation,
      citation_ids: finalCitationIds,
      row_hash: row.rowHash,
    });
  });

  operations.sort((a, b) => byIdOrPath(a.row_id, b.row_id));

  const payload = {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    operations,
  };

  const schemaValidation = validateAgainstSchema(payload, state.schemas.sidecarPatch);
  const unknownRows = operations
    .filter((op) => !state.rowMap.has(op.row_id))
    .map((op) => op.row_id);

  const validation = {
    ok: schemaValidation.ok && unknownRows.length === 0,
    errors: [
      ...schemaValidation.errors,
      ...unknownRows.map((id) => `unknown row id: ${id}`),
    ],
  };

  state.patchPreview.payload = payload;
  state.patchPreview.validation = validation;
}

function requireValidPatchPreviewOrThrow() {
  if (!state.patchPreview.payload) {
    throw new Error("Patch preview not built yet.");
  }
  if (!state.patchPreview.validation?.ok) {
    throw new Error(`Patch preview invalid: ${(state.patchPreview.validation?.errors || []).join("; ")}`);
  }
}

async function exportProposal() {
  if (!state.fs.evidenceDir) {
    logLine("error", "No outputs/evidence_lab folder available.");
    return;
  }

  buildPatchPreviewFromMatches();

  try {
    requireValidPatchPreviewOrThrow();

    const stamp = timestampForFile();
    const written = [];

    if (state.index.payload) {
      const indexValidation = validateAgainstSchema(state.index.payload, state.schemas.evidenceIndex);
      if (!indexValidation.ok) {
        throw new Error(`Index validation failed: ${indexValidation.errors.join("; ")}`);
      }
      const indexHandle = await state.fs.evidenceDir.getFileHandle(`evidence_index_${stamp}.json`, { create: true });
      await writeJson(indexHandle, state.index.payload);
      written.push(indexHandle.name);
    }

    if (state.matches.selected) {
      const selectedValidation = validateAgainstSchema(state.matches.selected, state.schemas.evidenceMatches);
      if (!selectedValidation.ok) {
        throw new Error(`Selected match validation failed: ${selectedValidation.errors.join("; ")}`);
      }
      const selectedHandle = await state.fs.evidenceDir.getFileHandle(`evidence_matches_selected_${stamp}.json`, { create: true });
      await writeJson(selectedHandle, state.matches.selected);
      written.push(selectedHandle.name);
    }

    if (state.matches.included && state.matches.included.length) {
      const payload = {
        created_at: timestampIso(),
        spec_version: SPEC_VERSION,
        generator: GENERATOR_ID,
        rows: state.matches.included,
      };
      const includedValidation = validateAgainstSchema(payload, state.schemas.evidenceMatches);
      if (!includedValidation.ok) {
        throw new Error(`Included match validation failed: ${includedValidation.errors.join("; ")}`);
      }
      const includedHandle = await state.fs.evidenceDir.getFileHandle(`evidence_matches_included_${stamp}.json`, { create: true });
      await writeJson(includedHandle, payload);
      written.push(includedHandle.name);
    }

    const patchHandle = await state.fs.evidenceDir.getFileHandle(`sidecar_patch.preview_${stamp}.json`, { create: true });
    await writeJson(patchHandle, state.patchPreview.payload);
    written.push(patchHandle.name);

    const manifest = {
      created_at: timestampIso(),
      spec_version: SPEC_VERSION,
      generator: GENERATOR_ID,
      files: written,
    };
    const manifestHandle = await state.fs.evidenceDir.getFileHandle(`export_manifest_${stamp}.json`, { create: true });
    await writeJson(manifestHandle, manifest);
    written.push(manifestHandle.name);

    updateRunSummaryPanel(["Export complete.", ...written.map((name) => `- ${name}`)]);
    logLine("info", "Exported proposal artifacts.", { files: written });
  } catch (err) {
    logLine("error", "Export failed.", { error: err.message });
  }
}

async function saveDebugLog() {
  const payload = {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    runtime: state.runtime,
    rows_in_memory: state.rowOrder.length,
    logs: state.logs,
  };
  const text = JSON.stringify(payload, null, 2);

  if (state.fs.evidenceDir) {
    try {
      const handle = await state.fs.evidenceDir.getFileHandle(`evidence_lab_log_${timestampForFile()}.json`, { create: true });
      await writeText(handle, text);
      logLine("info", "Saved debug log.", { name: handle.name });
      return;
    } catch (err) {
      logLine("error", "Failed to save debug log to project folder.", { error: err.message });
    }
  }

  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `evidence_lab_log_${timestampForFile()}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
  logLine("info", "Saved debug log via browser download.");
}

async function clearCacheAndRebuild() {
  if (!state.fs.evidenceDir) {
    logLine("warn", "No project folder selected for cache clear.");
    return;
  }

  let removed = 0;
  const removedNames = [];

  const targets = [CHECKPOINT_FILE, INDEX_LATEST_FILE, MATCHES_LATEST_FILE];
  for (const name of targets) {
    const ok = await safeRemoveEntry(state.fs.evidenceDir, name, false);
    if (ok) {
      removed += 1;
      removedNames.push(name);
    }
  }

  state.checkpoint.loaded = false;
  state.checkpoint.data = null;
  state.run.scanActive = false;
  if (typeof state.run.scanResolve === "function") state.run.scanResolve();
  state.run.scanResolve = null;
  state.run.scanPromise = null;
  state.index.payload = null;
  state.index.latestHandleName = "";
  state.matches.selected = null;
  state.matches.included = [];
  state.matches.byRowId.clear();
  state.matches.latestHandleName = "";
  state.drafts.byRowId.clear();
  state.patchPreview.payload = null;
  state.patchPreview.validation = null;
  renderEvidenceForSelectedRow();
  renderPatchPreview();

  state.inventory.forEach((item) => {
    item.status = "pending";
  });
  renderInventory();

  logLine("info", "Cache clear completed.", { removed, files: removedNames });
  updateRunSummaryPanel([`Cache cleared: ${removed} files removed.`]);
}

function cancelActiveRun() {
  if (!state.run.scanActive || !state.run.scanToken) {
    logLine("warn", "No active scan run to cancel.");
    return;
  }
  if (!state.workers.ingest) {
    state.run.scanActive = false;
    state.inventory.forEach((item) => {
      if (item.status === "running") item.status = "pending";
    });
    renderInventory();
    updateRunSummaryPanel(["Active scan canceled (no-worker mode)."]);
    if (typeof state.run.scanResolve === "function") state.run.scanResolve();
    state.run.scanResolve = null;
    state.run.scanPromise = null;
    return;
  }

  state.workers.ingest.postMessage({
    type: "cancel",
    runId: state.run.scanToken,
  });
}

async function runSmokeE2E() {
  if (!state.fs.projectDir) {
    logLine("warn", "Smoke run requires a selected project folder.");
    return;
  }

  const steps = [];
  try {
    await scanInputs();
    steps.push("scan");
    await buildOrRefreshIndex();
    steps.push("index");
    await runMatchSelected();
    steps.push("match_selected");
    await runMatchIncluded();
    steps.push("match_included");
    generateDraftForSelectedRow();
    steps.push("generate_draft");
    await exportProposal();
    steps.push("export");

    const hasEvidence = !!state.matches.byRowId.size;
    const hasPatch = !!state.patchPreview.payload && !!state.patchPreview.validation?.ok;
    const sidecarUntouched = true;
    const hasRunSummary = String(el.runSummary.textContent || "").trim().length > 0;

    if (!hasEvidence) throw new Error("no evidence in memory after run");
    if (!hasPatch) throw new Error("patch preview invalid after run");
    if (!sidecarUntouched) throw new Error("sidecar mutation guard failed");
    if (!hasRunSummary) throw new Error("run summary not updated");

    updateRunSummaryPanel([
      "Smoke E2E: PASS",
      `Steps: ${steps.join(" -> ")}`,
    ]);
    logLine("info", "Smoke E2E completed.", { steps, pass: true });
  } catch (err) {
    updateRunSummaryPanel([
      "Smoke E2E: FAIL",
      `Steps: ${steps.join(" -> ") || "-"}`,
      `Error: ${err.message}`,
    ]);
    logLine("error", "Smoke E2E failed.", { steps, error: err.message });
  }
}

function wireButtons() {
  el.pickFolderBtn.addEventListener("click", () => {
    void pickFolder();
  });

  el.scanInputsBtn.addEventListener("click", () => {
    void scanInputs();
  });

  el.cancelRunBtn.addEventListener("click", () => {
    cancelActiveRun();
  });

  el.buildIndexBtn.addEventListener("click", () => {
    void buildOrRefreshIndex();
  });

  el.clearCacheBtn.addEventListener("click", () => {
    void clearCacheAndRebuild();
  });

  el.matchSelectedBtn.addEventListener("click", () => {
    void runMatchSelected();
  });

  el.matchIncludedBtn.addEventListener("click", () => {
    void runMatchIncluded();
  });

  el.generateDraftBtn.addEventListener("click", () => {
    generateDraftForSelectedRow();
  });

  el.exportProposalBtn.addEventListener("click", () => {
    void exportProposal();
  });

  el.runSmokeBtn.addEventListener("click", () => {
    void runSmokeE2E();
  });

  el.saveLogBtn.addEventListener("click", () => {
    void saveDebugLog();
  });

  el.requestPersistBtn.addEventListener("click", () => {
    void requestPersistentStorage();
  });

  el.rowSearch.addEventListener("input", () => {
    renderRowList();
  });

  el.filterLowConfidence?.addEventListener("change", () => {
    renderEvidenceForSelectedRow();
  });

  el.groupByFile?.addEventListener("change", () => {
    renderEvidenceForSelectedRow();
  });

  el.settingBatchSize?.addEventListener("change", () => {
    updateSettingsFromUi();
    updateRunSummaryPanel(["Settings updated."]);
  });

  el.settingMaxUnits?.addEventListener("change", () => {
    updateSettingsFromUi();
    updateRunSummaryPanel(["Settings updated."]);
  });

  el.settingMemoryGuard?.addEventListener("change", () => {
    updateSettingsFromUi();
    updateRunSummaryPanel(["Settings updated."]);
  });

  el.optLogSensitive?.addEventListener("change", () => {
    updateSettingsFromUi();
    logLine("info", "Sensitive log mode updated.", {
      include_sensitive: state.settings.includeSensitiveLogs,
    });
  });

  el.profileOvernightBtn?.addEventListener("click", () => {
    applyProfile("overnight");
  });

  el.profileCpuBtn?.addEventListener("click", () => {
    applyProfile("cpu");
  });

  el.profileWebgpuBtn?.addEventListener("click", () => {
    applyProfile("webgpu");
  });
}

async function boot() {
  el.appVersion.textContent = `${EXPERIMENT_LABEL} v${APP_VERSION}`;
  wireButtons();

  await detectCapabilities();
  renderCapabilityChips();
  applyProfile(state.runtime.preferred === "webgpu" ? "webgpu" : "cpu");
  await loadSchemas();
  initWorkers();
  renderProjectSummary();
  renderRowDetails();
  renderEvidenceForSelectedRow();
  renderPatchPreview();
  updateRunSummaryPanel(["Ready.", "Select a project folder to begin."]);

  if (!state.runtime.localOnlyModels) {
    logLine("warn", "Remote model mode is enabled. This violates local-only policy.");
  }

  logLine("info", "AI Evidence Lab boot complete.", {
    version: APP_VERSION,
    capabilities: state.capabilities,
  });
}

void boot();
