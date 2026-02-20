const APP_VERSION = "0.1.0";
const SPEC_VERSION = "0.1";
const GENERATOR_ID = "ai-evidence-lab";

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

const state = {
  capabilities: {},
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
  logs: [],
  workers: {
    ingest: null,
    embed: null,
    match: null,
  },
  runToken: "",
  schemas: {
    evidenceMatches: null,
    sidecarPatch: null,
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
  logOutput: document.getElementById("log-output"),

  pickFolderBtn: document.getElementById("pick-folder"),
  scanInputsBtn: document.getElementById("scan-inputs"),
  buildIndexBtn: document.getElementById("build-index"),
  matchSelectedBtn: document.getElementById("match-selected"),
  matchIncludedBtn: document.getElementById("match-included"),
  exportProposalBtn: document.getElementById("export-proposal"),
  saveLogBtn: document.getElementById("save-log"),
  requestPersistBtn: document.getElementById("request-persist"),
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

function timestampIso() {
  return new Date().toISOString();
}

function timestampForFile() {
  return timestampIso().replace(/[:.]/g, "-");
}

function logLine(level, message, context = null) {
  const entry = {
    ts: timestampIso(),
    level,
    message,
    context,
  };
  state.logs.push(entry);
  const suffix = context ? ` ${JSON.stringify(context)}` : "";
  const line = `[${entry.ts}] ${level.toUpperCase()} ${message}${suffix}`;
  el.logOutput.textContent = `${el.logOutput.textContent}${line}\n`;
  el.logOutput.scrollTop = el.logOutput.scrollHeight;
}

function setRunSummary(lines) {
  const list = Array.isArray(lines) ? lines : [String(lines || "")];
  el.runSummary.textContent = list.join("\n");
}

function statusForBool(flag) {
  return flag ? "ok" : "bad";
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
  const i = String(name || "").lastIndexOf(".");
  return i >= 0 ? String(name).slice(i).toLowerCase() : "";
}

function isSupportedInputFile(name) {
  return SUPPORTED_EXTENSIONS.has(extFromName(name));
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
      caps.storage_quota = formatBytes(estimate?.quota || 0);
      caps.storage_usage = formatBytes(estimate?.usage || 0);
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
}

function renderCapabilityChips() {
  const chips = [];
  const caps = state.capabilities;

  chips.push({ key: "FS API", ok: !!caps.fs_api, extra: "" });
  chips.push({ key: "Workers", ok: !!caps.workers, extra: "" });
  chips.push({ key: "WebGPU", ok: !!caps.webgpu, extra: caps.webgpu_shader_f16 ? "f16" : "" });
  chips.push({ key: "Storage estimate", ok: !!caps.storage_estimate, extra: `${caps.storage_usage || "-"} / ${caps.storage_quota || "-"}` });
  chips.push({ key: "Persist", ok: !!caps.storage_persist, extra: "" });

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
  const dirName = state.fs.projectDir?.name || "(none)";
  lines.push(`Folder: ${dirName}`);
  lines.push(`Inputs: ${state.fs.inputsDir ? "ok" : "missing"}`);
  lines.push(`Sidecar: ${state.fs.sidecarHandle ? "ok" : "missing"}`);
  const locale = state.projectMeta.locale || "-";
  const company = state.projectMeta.company || state.projectMeta.companyName || "-";
  lines.push(`Locale: ${locale}`);
  lines.push(`Company: ${company}`);
  lines.push(`Rows: ${state.rowOrder.length}`);
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
  } catch (err) {
    return null;
  }
}

async function safeGetFileHandle(parent, name, create = false) {
  try {
    return await parent.getFileHandle(name, { create });
  } catch (err) {
    return null;
  }
}

async function readFileText(handle) {
  const file = await handle.getFile();
  return await file.text();
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

      rowMap.set(rowId, {
        rowId,
        chapterId,
        chapterTitle,
        title,
        include: !!ws.include,
        done: !!ws.done,
        priority,
      });
    });
  });

  const rowOrder = Array.from(rowMap.keys()).sort(byIdOrPath);
  return {
    projectMeta: project.meta || {},
    rowMap,
    rowOrder,
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
  const q = String(el.rowSearch.value || "").trim();
  el.rowList.innerHTML = "";

  const filtered = state.rowOrder
    .map((rowId) => state.rowMap.get(rowId))
    .filter((row) => row && rowMatchesFilter(row, q));

  filtered.forEach((row) => {
    const item = document.createElement("div");
    item.className = `row-item${row.rowId === state.selectedRowId ? " active" : ""}`;

    const title = row.title || "(untitled row)";
    item.innerHTML = `
      <strong>${row.rowId}</strong> ${escapeHtml(title)}
      <br />
      <small>${escapeHtml(row.chapterId)} ${escapeHtml(row.chapterTitle || "")}</small>
      <br />
      <small>include=${row.include ? "yes" : "no"}, done=${row.done ? "yes" : "no"}, prio=${row.priority}</small>
    `;

    item.addEventListener("click", () => {
      state.selectedRowId = row.rowId;
      renderRowList();
      renderRowDetails();
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

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
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
  const entries = [
    ["evidenceMatches", "./schemas/evidence_matches.schema.json"],
    ["sidecarPatch", "./schemas/sidecar_patch.schema.json"],
  ];
  for (const [key, url] of entries) {
    try {
      const res = await fetch(url, { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      state.schemas[key] = await res.json();
    } catch (err) {
      logLine("warn", `Could not load schema ${key}`, { error: err.message });
      state.schemas[key] = null;
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

async function pickFolder() {
  if (!state.capabilities.fs_api) {
    logLine("error", "File System Access API is not available in this browser.");
    return;
  }

  try {
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

    renderProjectSummary();
    setRunSummary(["Folder selected.", "Ready for scan."]);
    logLine("info", "Project folder selected.", { name: projectDir.name });
  } catch (err) {
    logLine("error", "Folder pick failed.", { error: err.message });
    setRunSummary(["Folder pick failed.", err.message]);
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
      logLine("info", "Ingest worker started.", { total: msg.total, runId: msg.runId });
      return;
    }
    if (msg.type === "progress") {
      if (msg.runId !== state.runToken) return;
      const item = state.inventory.find((entry) => entry.path === msg.path);
      if (item) {
        item.status = msg.status;
        renderInventory();
      }
      return;
    }
    if (msg.type === "done") {
      if (msg.runId !== state.runToken) return;
      setRunSummary([
        "Scan complete.",
        `Files: ${state.inventory.length}`,
        `Completed: ${msg.completed}`,
      ]);
      logLine("info", "Ingest worker done.", { runId: msg.runId, completed: msg.completed });
      return;
    }
    if (msg.type === "error") {
      logLine("error", "Ingest worker error.", msg);
    }
  };

  logLine("info", "Workers initialized.");
}

async function collectInventory() {
  const items = [];
  for await (const fileInfo of walkDir(state.fs.inputsDir, "inputs")) {
    if (!isSupportedInputFile(fileInfo.name)) continue;
    const file = await fileInfo.handle.getFile();
    items.push({
      path: fileInfo.path,
      type: extFromName(fileInfo.name).slice(1).toLowerCase(),
      size: file.size,
      status: "pending",
    });
  }
  items.sort((a, b) => byIdOrPath(a.path, b.path));
  state.inventory = items;
  renderInventory();
}

async function runIngestSimulation() {
  if (!state.workers.ingest) {
    state.inventory.forEach((item) => {
      item.status = "done";
    });
    renderInventory();
    setRunSummary(["Scan complete.", `Files: ${state.inventory.length}`, "Mode: no-worker fallback"]);
    return;
  }

  const runId = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  state.runToken = runId;
  state.workers.ingest.postMessage({
    type: "start",
    runId,
    items: state.inventory.map((item) => ({ path: item.path })),
  });
}

async function scanInputs() {
  if (!state.fs.inputsDir) {
    logLine("error", "Cannot scan: no project folder selected.");
    setRunSummary("Pick a project folder first.");
    return;
  }

  const startedAt = performance.now();
  await collectInventory();

  if (!state.inventory.length) {
    setRunSummary(["Scan complete.", "No supported files found in inputs/."]);
    logLine("warn", "No supported input files found.");
    return;
  }

  await runIngestSimulation();
  const elapsedMs = performance.now() - startedAt;
  logLine("info", "Scan kicked off.", { files: state.inventory.length, elapsed_ms: Math.round(elapsedMs) });
}

function buildEvidenceMatchPayload() {
  const row = state.selectedRowId ? state.rowMap.get(state.selectedRowId) : null;
  return {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    rows: row
      ? [{ row_id: row.rowId, status: "none", citations: [] }]
      : [],
  };
}

function buildPatchPreviewPayload() {
  const row = state.selectedRowId ? state.rowMap.get(state.selectedRowId) : null;
  return {
    created_at: timestampIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    operations: row
      ? [{ row_id: row.rowId, mode: "skip", finding: "", recommendation: "", citation_ids: [] }]
      : [],
  };
}

async function exportProposal() {
  if (!state.fs.evidenceDir) {
    logLine("error", "No outputs/evidence_lab folder available.");
    return;
  }

  const evidencePayload = buildEvidenceMatchPayload();
  const patchPayload = buildPatchPreviewPayload();

  const evResult = validateAgainstSchema(evidencePayload, state.schemas.evidenceMatches);
  const patchResult = validateAgainstSchema(patchPayload, state.schemas.sidecarPatch);
  if (!evResult.ok || !patchResult.ok) {
    logLine("error", "Schema validation failed.", {
      evidence: evResult.errors,
      patch: patchResult.errors,
    });
    return;
  }

  try {
    const stamp = timestampForFile();
    const evidenceHandle = await state.fs.evidenceDir.getFileHandle(`evidence_matches_${stamp}.json`, { create: true });
    const patchHandle = await state.fs.evidenceDir.getFileHandle(`sidecar_patch.preview_${stamp}.json`, { create: true });
    await writeJson(evidenceHandle, evidencePayload);
    await writeJson(patchHandle, patchPayload);
    logLine("info", "Exported proposal JSON artifacts.", { evidence: evidenceHandle.name, patch: patchHandle.name });
  } catch (err) {
    logLine("error", "Export failed.", { error: err.message });
  }
}

async function saveDebugLog() {
  const text = state.logs
    .map((entry) => {
      const suffix = entry.context ? ` ${JSON.stringify(entry.context)}` : "";
      return `[${entry.ts}] ${entry.level.toUpperCase()} ${entry.message}${suffix}`;
    })
    .join("\n");

  if (state.fs.evidenceDir) {
    try {
      const handle = await state.fs.evidenceDir.getFileHandle(`evidence_lab_log_${timestampForFile()}.txt`, { create: true });
      await writeText(handle, text);
      logLine("info", "Saved debug log.", { name: handle.name });
      return;
    } catch (err) {
      logLine("error", "Failed to save debug log to project folder.", { error: err.message });
    }
  }

  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `evidence_lab_log_${timestampForFile()}.txt`;
  a.click();
  URL.revokeObjectURL(url);
  logLine("info", "Saved debug log via browser download.");
}

function wireButtons() {
  el.pickFolderBtn.addEventListener("click", () => {
    void pickFolder();
  });

  el.scanInputsBtn.addEventListener("click", () => {
    void scanInputs();
  });

  el.buildIndexBtn.addEventListener("click", () => {
    logLine("info", "Build/Refresh Index: placeholder (workstream 2).", {
      files: state.inventory.length,
    });
  });

  el.matchSelectedBtn.addEventListener("click", () => {
    logLine("info", "Find Evidence (Selected Row): placeholder (workstream 3).", {
      row_id: state.selectedRowId || null,
    });
  });

  el.matchIncludedBtn.addEventListener("click", () => {
    const includedCount = Array.from(state.rowMap.values()).filter((row) => row.include).length;
    logLine("info", "Find Evidence (Included Rows): placeholder (workstream 3).", {
      included_rows: includedCount,
    });
  });

  el.exportProposalBtn.addEventListener("click", () => {
    void exportProposal();
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
}

async function boot() {
  el.appVersion.textContent = `v${APP_VERSION}`;
  wireButtons();

  await detectCapabilities();
  renderCapabilityChips();
  await loadSchemas();
  initWorkers();
  renderProjectSummary();
  renderRowDetails();
  setRunSummary(["Ready.", "Select a project folder to begin."]);

  logLine("info", "AI Evidence Lab boot complete.", {
    version: APP_VERSION,
    caps: state.capabilities,
  });
}

void boot();
