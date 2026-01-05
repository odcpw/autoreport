(() => {
  const pickProjectBtn = document.getElementById("pick-project");
  const loadSidecarBtn = document.getElementById("load-sidecar");
  const loadCategoriesBtn = document.getElementById("load-categories");
  const pickPhotosBtn = document.getElementById("pick-photos");
  const scanPhotosBtn = document.getElementById("scan-photos");
  const saveSidecarBtn = document.getElementById("save-sidecar");
  const saveLogBtn = document.getElementById("save-log");
  const statusEl = document.getElementById("status");
  const photoMetaEl = document.getElementById("photo-meta");
  const photoImageEl = document.getElementById("photo-image");
  const thumbsEl = document.getElementById("thumbs");
  const filterToggleBtn = document.getElementById("filter-toggle");
  const prevBtn = document.getElementById("prev-photo");
  const nextBtn = document.getElementById("next-photo");
  const notesEl = document.getElementById("photo-notes");

  const panels = {
    report: document.getElementById("panel-report"),
    observations: document.getElementById("panel-observations"),
    training: document.getElementById("panel-training"),
  };

  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };

  const urlParams = new URLSearchParams(window.location.search);
  const demoMode = urlParams.get("demo") === "1" || urlParams.get("demo") === "true";
  const demoWorkbookParam = urlParams.get("demoWorkbook");
  const layoutParam = urlParams.get("layout");
  const storedLayout = window.localStorage?.getItem("photosorterLayout");
  let layoutMode =
    layoutParam === "tabs" || layoutParam === "stacked"
      ? layoutParam
      : storedLayout === "tabs"
        ? "tabs"
        : "stacked";
  const panelTabButtons = Array.from(document.querySelectorAll("[data-panel-tab]"));
  const layoutToggleButtons = Array.from(document.querySelectorAll("[data-layout]"));

  const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"]);
  const DEFAULT_TAGS = {
    report: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
    observations: ["Forklifts", "PPE", "Housekeeping", "Chemicals", "Workplace"],
    training: ["Vorbildliches Verhalten", "Risikoanalyse", "Audit", "Kommunikation"],
  };

  const state = {
    projectHandle: null,
    photoHandle: null,
    projectDoc: null,
    photoRootName: "",
    tagOptions: null,
    tagFilters: { report: "", observations: "", training: "" },
    photos: [],
    filterMode: "all",
    currentIndex: 0,
    activePanel: "report",
  };


  const setStatus = (message) => {
    statusEl.textContent = message;
    debug.logLine("info", message);
  };

  const applyLayoutMode = () => {
    document.body.classList.toggle("layout-tabs", layoutMode === "tabs");
    document.body.classList.toggle("layout-stacked", layoutMode === "stacked");
  };

  const updateLayoutToggle = () => {
    layoutToggleButtons.forEach((button) => {
      const isActive = button.dataset.layout === layoutMode;
      button.classList.toggle("active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
    });
  };

  const setLayoutMode = (mode) => {
    if (mode !== "tabs" && mode !== "stacked") return;
    layoutMode = mode;
    if (window.localStorage) {
      window.localStorage.setItem("photosorterLayout", mode);
    }
    applyLayoutMode();
    updateLayoutToggle();
    renderPanels();
  };

  const setActivePanel = (group) => {
    if (!panels[group]) return;
    state.activePanel = group;
    panelTabButtons.forEach((btn) => {
      const isActive = btn.dataset.panelTab === group;
      btn.classList.toggle("active", isActive);
      btn.setAttribute("aria-pressed", String(isActive));
    });
    renderPanels();
  };

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus("File System Access API is not available in this browser.");
      pickProjectBtn.disabled = true;
      return false;
    }
    return true;
  };

  const enableActions = () => {
    const hasProject = !!state.projectHandle;
    const hasPhotoFolder = !!state.photoHandle;
    const hasVisiblePhotos = getFilteredPhotos().length > 0;
    loadSidecarBtn.disabled = !hasProject;
    loadCategoriesBtn.disabled = !hasProject && !demoMode;
    pickPhotosBtn.disabled = !hasProject;
    scanPhotosBtn.disabled = !hasPhotoFolder;
    saveSidecarBtn.disabled = !hasProject;
    filterToggleBtn.disabled = state.photos.length === 0;
    prevBtn.disabled = !hasVisiblePhotos;
    nextBtn.disabled = !hasVisiblePhotos;
  };

  panelTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      setActivePanel(button.dataset.panelTab);
    });
  });

  layoutToggleButtons.forEach((button) => {
    button.addEventListener("click", () => {
      setLayoutMode(button.dataset.layout);
    });
  });

  const normalizeTagOption = (option) => {
    if (!option) return null;
    if (typeof option === "string") {
      return { value: option, label: option };
    }
    if (typeof option === "object") {
      const value = option.value || option.label;
      if (!value) return null;
      return { value, label: option.label || value };
    }
    return null;
  };

  const isNumericTag = (value) => /^\d+(?:\.\d+)*$/.test(value);

  const compareNumericTags = (a, b) => {
    const left = String(a.value || a.label || "");
    const right = String(b.value || b.label || "");
    const leftIsNum = isNumericTag(left);
    const rightIsNum = isNumericTag(right);
    if (leftIsNum && rightIsNum) {
      const leftParts = left.split(".").map((part) => Number(part));
      const rightParts = right.split(".").map((part) => Number(part));
      const max = Math.max(leftParts.length, rightParts.length);
      for (let i = 0; i < max; i += 1) {
        const l = leftParts[i] ?? -1;
        const r = rightParts[i] ?? -1;
        if (l !== r) return l - r;
      }
      return 0;
    }
    if (leftIsNum) return -1;
    if (rightIsNum) return 1;
    return left.localeCompare(right, "de", { numeric: true });
  };

  const sortOptionsForGroup = (group, options) => {
    const list = [...options];
    if (group === "report") {
      return list.sort(compareNumericTags);
    }
    return list.sort((a, b) => String(a.label || a.value || "").localeCompare(String(b.label || b.value || ""), "de", { numeric: true }));
  };

  const normalizeIncomingOptions = (options) => {
    if (!options) return {};
    const mapped = { ...options };
    if (!mapped.report && mapped.bericht) mapped.report = mapped.bericht;
    if (!mapped.training && mapped.seminar) mapped.training = mapped.seminar;
    if (!mapped.observations && mapped.topic) mapped.observations = mapped.topic;
    return mapped;
  };

  const normalizeTagOptions = (options) => {
    const incoming = normalizeIncomingOptions(options);
    return {
      report: (incoming.report || []).map(normalizeTagOption).filter(Boolean),
      observations: (incoming.observations || []).map(normalizeTagOption).filter(Boolean),
      training: (incoming.training || []).map(normalizeTagOption).filter(Boolean),
    };
  };

  const normalizePhotoTags = (tags) => {
    const incoming = normalizeIncomingOptions(tags || {});
    return {
      report: Array.from(incoming.report || []),
      observations: Array.from(incoming.observations || []),
      training: Array.from(incoming.training || []),
    };
  };

  const ensureTagOptions = (options) => {
    const incoming = normalizeIncomingOptions(options);
    const merged = {
      report: [...(DEFAULT_TAGS.report || []), ...(incoming.report || [])],
      observations: [...(DEFAULT_TAGS.observations || []), ...(incoming.observations || [])],
      training: [...(DEFAULT_TAGS.training || []), ...(incoming.training || [])],
    };
    const normalized = normalizeTagOptions(merged);
    normalized.report = sortOptionsForGroup("report", dedupeOptions(normalized.report));
    normalized.observations = sortOptionsForGroup("observations", dedupeOptions(normalized.observations));
    normalized.training = sortOptionsForGroup("training", dedupeOptions(normalized.training));
    return normalized;
  };

  const dedupeOptions = (options) => {
    const seen = new Set();
    return options.filter((opt) => {
      const key = opt.value;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const SEED_TAG_OPTIONS = window.PS_CATEGORY_LABELS_SEED
    ? ensureTagOptions(window.PS_CATEGORY_LABELS_SEED)
    : ensureTagOptions(DEFAULT_TAGS);

  const createEmptyProjectDoc = () => ({
    meta: {
      projectId: "",
      createdAt: new Date().toISOString(),
    },
    photos: {},
    photoTagOptions: SEED_TAG_OPTIONS,
    photoRoot: "",
  });

  const isImageFile = (name) => {
    const lower = name.toLowerCase();
    const dot = lower.lastIndexOf(".");
    if (dot === -1) return false;
    return IMAGE_EXTENSIONS.has(lower.slice(dot));
  };

  const isPhotoUnsorted = (photo) => {
    if (!photo?.tags) return true;
    const { report, observations, training } = photo.tags;
    return !((report || []).length || (observations || []).length || (training || []).length);
  };

  const getFilteredPhotos = () => {
    if (state.filterMode === "unsorted") {
      return state.photos.filter(isPhotoUnsorted);
    }
    return state.photos;
  };

  const getCurrentPhoto = () => {
    const filtered = getFilteredPhotos();
    if (!filtered.length) return null;
    return filtered[state.currentIndex] || null;
  };

  const updateMeta = () => {
    const filtered = getFilteredPhotos();
    const total = state.photos.length;
    const unsorted = state.photos.filter(isPhotoUnsorted).length;
    const current = filtered.length ? state.currentIndex + 1 : 0;
    photoMetaEl.textContent = `Image ${current} of ${filtered.length} • Total ${total} • Unsorted ${unsorted}`;
  };

  const clearPhotoUrls = () => {
    state.photos.forEach((photo) => {
      if (photo.url) {
        URL.revokeObjectURL(photo.url);
      }
    });
  };

  const renderViewer = () => {
    const current = getCurrentPhoto();
    if (!current) {
      photoImageEl.removeAttribute("src");
      photoImageEl.alt = "No photo loaded";
      notesEl.value = "";
      notesEl.disabled = true;
      updateMeta();
      enableActions();
      return;
    }
    photoImageEl.src = current.url;
    photoImageEl.alt = current.path;
    notesEl.value = current.notes || "";
    notesEl.disabled = false;
    updateMeta();
  };

  const renderThumbs = () => {
    thumbsEl.innerHTML = "";
    const filtered = getFilteredPhotos();
    if (!filtered.length) return;

    filtered.forEach((photo, index) => {
      const item = document.createElement("div");
      item.className = "thumb";
      if (index === state.currentIndex) {
        item.classList.add("active");
      }
      const img = document.createElement("img");
      img.src = photo.url;
      img.alt = photo.path;
      const label = document.createElement("span");
      label.textContent = photo.path.split("/").slice(-1)[0];
      item.appendChild(img);
      item.appendChild(label);
      item.addEventListener("click", () => {
        state.currentIndex = index;
        renderAll();
      });
      thumbsEl.appendChild(item);
    });
  };

  const renderTagPanel = (group, config) => {
    const container = panels[group];
    if (!container) return;

    container.classList.toggle("is-active", state.activePanel === group);
    container.innerHTML = "";
    const title = document.createElement("h3");
    title.textContent = config.title;
    const description = document.createElement("p");
    description.textContent = config.description;

    const controls = document.createElement("div");
    controls.className = "panel__controls";
    const filterInput = document.createElement("input");
    filterInput.type = "text";
    filterInput.placeholder = "Filter tags";
    filterInput.value = config.filter || "";
    filterInput.addEventListener("input", () => {
      state.tagFilters[group] = filterInput.value;
      renderPanels();
    });

    const addInput = document.createElement("input");
    addInput.type = "text";
    addInput.placeholder = "Add new tag";
    addInput.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      const value = addInput.value.trim();
      if (!value) return;
      const existing = state.tagOptions[group] || [];
      if (!existing.some((option) => option.value === value)) {
        const next = sortOptionsForGroup(group, [...existing, { value, label: value }]);
        state.tagOptions[group] = next;
      }
      addInput.value = "";
      renderPanels();
    });

    controls.append(filterInput, addInput);

    const tagsEl = document.createElement("div");
    tagsEl.className = "panel__tags";

    const current = getCurrentPhoto();
    const selected = new Set(current?.tags?.[group] || []);
    const options = (state.tagOptions?.[group] || [])
      .filter((option) => option.label.toLowerCase().includes((config.filter || "").toLowerCase()));

    options.forEach((option) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "tag-button";
      button.textContent = option.label;
      button.title = option.label;
      if (selected.has(option.value)) {
        button.classList.add("active");
      }
      button.addEventListener("click", () => {
        toggleTag(group, option.value);
      });
      tagsEl.appendChild(button);
    });

    container.append(title, description, controls, tagsEl);
  };

  const renderPanels = () => {
    renderTagPanel("report", {
      title: "Bericht",
      description: "Kapitel & Unterkapitel (1.x / 1.2 / 4.8 etc.)",
      filter: state.tagFilters.report,
    });
    renderTagPanel("observations", {
      title: "Beobachtungen",
      description: "Themen/Begriffe aus der Begehung.",
      filter: state.tagFilters.observations,
    });
    renderTagPanel("training", {
      title: "Training",
      description: "Seminar-/Schulungskategorien.",
      filter: state.tagFilters.training,
    });
  };

  const renderAll = () => {
    const filtered = getFilteredPhotos();
    if (state.currentIndex >= filtered.length) {
      state.currentIndex = Math.max(0, filtered.length - 1);
    }
    filterToggleBtn.textContent = state.filterMode === "all" ? "Show Unsorted" : "Show All";
    renderViewer();
    renderThumbs();
    renderPanels();
    enableActions();
  };

  const toggleTag = (group, tag) => {
    const current = getCurrentPhoto();
    if (!current) return;
    const list = new Set(current.tags[group] || []);
    if (list.has(tag)) {
      list.delete(tag);
    } else {
      list.add(tag);
    }
    current.tags[group] = Array.from(list);
    renderAll();
  };

  const serializePhotos = () => {
    const output = {};
    state.photos.forEach((photo) => {
      output[photo.path] = {
        notes: photo.notes || "",
        tags: photo.tags,
      };
    });
    return output;
  };

  const loadProjectSidecar = async () => {
    if (!state.projectHandle) return;
    try {
      const handle = await state.projectHandle.getFileHandle("project_sidecar.json");
      const file = await handle.getFile();
      const text = await file.text();
      state.projectDoc = JSON.parse(text);
      state.tagOptions = ensureTagOptions(state.projectDoc.photoTagOptions);
      state.photoRootName = state.projectDoc.photoRoot || "";
      setStatus("Loaded project_sidecar.json");
      debug.logLine("info", "Loaded project_sidecar.json");
    } catch (err) {
      state.projectDoc = createEmptyProjectDoc();
      state.tagOptions = ensureTagOptions(DEFAULT_TAGS);
      state.photoRootName = "";
      setStatus("Sidecar not found; starting fresh.");
      debug.logLine("warn", `Sidecar not found: ${err.message || err}`);
    }
    renderAll();
  };

  const saveProjectSidecar = async () => {
    if (!state.projectHandle) return;
    const payload = state.projectDoc && typeof state.projectDoc === "object"
      ? { ...state.projectDoc }
      : createEmptyProjectDoc();
    payload.photos = serializePhotos();
    payload.photoTagOptions = structuredClone(state.tagOptions);
    payload.photoRoot = state.photoRootName || "";
    if (!payload.meta) payload.meta = {};
    if (!payload.meta.updatedAt) payload.meta.updatedAt = new Date().toISOString();

    try {
      const handle = await state.projectHandle.getFileHandle("project_sidecar.json", { create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(payload, null, 2));
      await writable.close();
      state.projectDoc = payload;
      setStatus("Saved photo tags to project_sidecar.json");
      debug.logLine("info", "Saved photo tags to project_sidecar.json");
    } catch (err) {
      setStatus(`Save failed: ${err.message}`);
      debug.logLine("error", `Save failed: ${err.message}`);
    }
  };

  const buildPhotoEntry = (path, file) => {
    const previous = state.projectDoc?.photos?.[path];
    const tags = normalizePhotoTags(previous?.tags);
    return {
      path,
      file,
      url: URL.createObjectURL(file),
      notes: previous?.notes || "",
      tags,
    };
  };

  const collectImages = async (handle, prefix, collection) => {
    for await (const entry of handle.values()) {
      if (entry.kind === "file") {
        if (!isImageFile(entry.name)) continue;
        const file = await entry.getFile();
        const path = `${prefix}${entry.name}`;
        collection.push(buildPhotoEntry(path, file));
      } else if (entry.kind === "directory") {
        await collectImages(entry, `${prefix}${entry.name}/`, collection);
      }
    }
  };

  const scanPhotos = async () => {
    if (!state.photoHandle) return;
    clearPhotoUrls();
    setStatus("Scanning photos...");
    const collection = [];
    await collectImages(state.photoHandle, `${state.photoRootName}/`, collection);
    collection.sort((a, b) => a.path.localeCompare(b.path));
    state.photos = collection;
    state.filterMode = "all";
    state.currentIndex = 0;
    filterToggleBtn.textContent = "Show Unsorted";
    setStatus(`Loaded ${state.photos.length} photos from ${state.photoRootName}`);
    renderAll();
  };

  pickProjectBtn.addEventListener("click", async () => {
    if (!ensureFsAccess()) return;
    try {
      state.projectHandle = await window.showDirectoryPicker();
      setStatus(`Project folder: ${state.projectHandle.name}`);
      await loadProjectSidecar();
    } catch (err) {
      setStatus(`Project pick canceled: ${err.message}`);
    }
    enableActions();
  });

  loadSidecarBtn.addEventListener("click", async () => {
    await loadProjectSidecar();
  });

  const parsePSCategoryLabels = (rows) => {
    const reportLabels = [];
    const trainingLabels = [];
    const observationLabels = [];
    rows.forEach((row) => {
      const reportA = String(row[0] || "").trim();
      const training = String(row[1] || "").trim();
      const observation = String(row[2] || "").trim();
      const reportB = String(row[3] || "").trim();
      if (reportA) reportLabels.push(reportA);
      if (reportB) reportLabels.push(reportB);
      if (training) trainingLabels.push(training);
      if (observation) observationLabels.push(observation);
    });

    const reportOptions = reportLabels.map((label) => {
      const match = label.match(/^\\d+(?:\\.\\d+)*(?:\\.)?/);
      const value = match ? match[0].replace(/\\.$/, "") : label;
      return { value, label };
    });

    return {
      report: reportOptions,
      training: trainingLabels.map((label) => ({ value: label, label })),
      observations: observationLabels.map((label) => ({ value: label, label })),
    };
  };

  loadCategoriesBtn.addEventListener("click", async () => {
    if (!state.projectHandle) return;
    if (!window.showOpenFilePicker || !window.XLSX) {
      setStatus("File picker or SheetJS not available in this browser.");
      return;
    }
    try {
      const [fileHandle] = await window.showOpenFilePicker({
        types: [
          {
            description: "Excel workbook",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
                ".xlsx",
                ".xlsm",
              ],
            },
          },
        ],
        multiple: false,
      });
      const file = await fileHandle.getFile();
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames.find(
        (name) => name.toLowerCase() === "pscategorylabels"
      );
      if (!sheetName) {
        throw new Error("Sheet PSCategoryLabels not found.");
      }
      const sheet = workbook.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
      const options = parsePSCategoryLabels(rows);
      state.tagOptions = ensureTagOptions(options);
      setStatus(`Loaded categories from ${file.name}`);
      debug.logLine("info", `Loaded categories from ${file.name}`);
      renderAll();
    } catch (err) {
      setStatus(`Category load failed: ${err.message}`);
      debug.logLine("error", `Category load failed: ${err.message}`);
    }
  });

  const loadCategoriesFromUrl = async (rawUrl) => {
    if (!window.XLSX) {
      throw new Error("SheetJS not available.");
    }
    const resolved = new URL(rawUrl, window.location.origin);
    const response = await fetch(resolved.toString());
    if (!response.ok) {
      throw new Error(`Failed to fetch ${resolved.pathname}`);
    }
    const buffer = await response.arrayBuffer();
    const workbook = window.XLSX.read(buffer, { type: "array" });
    const sheetName = workbook.SheetNames.find(
      (name) => name.toLowerCase() === "pscategorylabels"
    );
    if (!sheetName) {
      throw new Error("Sheet PSCategoryLabels not found.");
    }
    const sheet = workbook.Sheets[sheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
    const options = parsePSCategoryLabels(rows);
    state.tagOptions = ensureTagOptions(options);
    renderAll();
  };

  pickPhotosBtn.addEventListener("click", async () => {
    if (!state.projectHandle) return;
    try {
      state.photoHandle = await window.showDirectoryPicker();
      state.photoRootName = state.photoHandle.name;
      setStatus(`Photo folder: ${state.photoRootName}`);
      enableActions();
    } catch (err) {
      setStatus(`Photo pick canceled: ${err.message}`);
    }
  });

  scanPhotosBtn.addEventListener("click", async () => {
    await scanPhotos();
  });

  saveSidecarBtn.addEventListener("click", async () => {
    await saveProjectSidecar();
  });

  filterToggleBtn.addEventListener("click", () => {
    state.filterMode = state.filterMode === "all" ? "unsorted" : "all";
    filterToggleBtn.textContent = state.filterMode === "all" ? "Show Unsorted" : "Show All";
    state.currentIndex = 0;
    renderAll();
  });

  prevBtn.addEventListener("click", () => {
    const filtered = getFilteredPhotos();
    if (!filtered.length) return;
    state.currentIndex = (state.currentIndex - 1 + filtered.length) % filtered.length;
    renderAll();
  });

  nextBtn.addEventListener("click", () => {
    const filtered = getFilteredPhotos();
    if (!filtered.length) return;
    state.currentIndex = (state.currentIndex + 1) % filtered.length;
    renderAll();
  });

  document.addEventListener("keydown", (event) => {
    if (layoutMode !== "tabs") return;
    const target = event.target;
    const isInput = target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA");
    if (isInput) return;
    if (event.key === "1") setActivePanel("report");
    if (event.key === "2") setActivePanel("observations");
    if (event.key === "3") setActivePanel("training");
  });

  notesEl.addEventListener("input", () => {
    const current = getCurrentPhoto();
    if (!current) return;
    current.notes = notesEl.value;
  });

  window.addEventListener("keydown", (event) => {
    const target = event.target;
    const isInput = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement;
    if (isInput) return;
    if (event.key.toLowerCase() === "a") {
      event.preventDefault();
      prevBtn.click();
    } else if (event.key.toLowerCase() === "d") {
      event.preventDefault();
      nextBtn.click();
    }
  });

  saveLogBtn.addEventListener("click", async () => {
    try {
      const result = await debug.saveLog({
        suggestedName: `photosorter-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`,
        dirHandle: state.projectHandle || null,
      });
      setStatus(`Saved log (${result.location}): ${result.filename}`);
    } catch (err) {
      setStatus(`Log save failed: ${err.message}`);
    }
  });

  state.tagOptions = SEED_TAG_OPTIONS;
  applyLayoutMode();
  updateLayoutToggle();
  setActivePanel(state.activePanel);
  ensureFsAccess();
  enableActions();

  if (window.PS_CATEGORY_LABELS_SEED) {
    setStatus("Loaded categories from bundled seed.");
  }

  if (demoMode) {
    const workbookPath = demoWorkbookParam || "/fromWork/2025-07-10 AutoBericht v0 Berichtform report selectors.xlsx";
    loadCategoriesFromUrl(workbookPath)
      .then(() => {
        setStatus(`Demo loaded categories from ${workbookPath}`);
        debug.logLine("info", `Demo loaded categories from ${workbookPath}`);
      })
      .catch((err) => {
        setStatus(`Demo load failed: ${err.message}`);
        debug.logLine("error", `Demo load failed: ${err.message}`);
      });
  }
})();
