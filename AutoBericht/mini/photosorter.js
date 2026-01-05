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
    bericht: document.getElementById("panel-bericht"),
    topic: document.getElementById("panel-topic"),
    seminar: document.getElementById("panel-seminar"),
  };

  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };

  const urlParams = new URLSearchParams(window.location.search);
  const demoMode = urlParams.get("demo") === "1" || urlParams.get("demo") === "true";
  const demoWorkbookParam = urlParams.get("demoWorkbook");

  const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"]);
  const DEFAULT_TAGS = {
    bericht: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
    topic: ["Forklifts", "PPE", "Housekeeping", "Chemicals", "Workplace"],
    seminar: ["Vorbildliches Verhalten", "Risikoanalyse", "Audit", "Kommunikation"],
  };

  const state = {
    projectHandle: null,
    photoHandle: null,
    projectDoc: null,
    photoRootName: "",
    tagOptions: null,
    tagFilters: { bericht: "", topic: "", seminar: "" },
    photos: [],
    filterMode: "all",
    currentIndex: 0,
  };


  const setStatus = (message) => {
    statusEl.textContent = message;
    debug.logLine("info", message);
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

  const normalizeTagOptions = (options) => ({
    bericht: (options?.bericht || []).map(normalizeTagOption).filter(Boolean),
    topic: (options?.topic || []).map(normalizeTagOption).filter(Boolean),
    seminar: (options?.seminar || []).map(normalizeTagOption).filter(Boolean),
  });

  const ensureTagOptions = (options) => {
    const merged = {
      bericht: [...(DEFAULT_TAGS.bericht || []), ...(options?.bericht || [])],
      topic: [...(DEFAULT_TAGS.topic || []), ...(options?.topic || [])],
      seminar: [...(DEFAULT_TAGS.seminar || []), ...(options?.seminar || [])],
    };
    const normalized = normalizeTagOptions(merged);
    normalized.bericht = dedupeOptions(normalized.bericht);
    normalized.topic = dedupeOptions(normalized.topic);
    normalized.seminar = dedupeOptions(normalized.seminar);
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
    const { bericht, topic, seminar } = photo.tags;
    return !((bericht || []).length || (topic || []).length || (seminar || []).length);
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
        const next = [...existing, { value, label: value }];
        next.sort((a, b) => a.label.localeCompare(b.label));
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
    renderTagPanel("bericht", {
      title: "Report",
      description: "Tag by report chapter or section.",
      filter: state.tagFilters.bericht,
    });
    renderTagPanel("topic", {
      title: "Topic",
      description: "Topic or folder tags.",
      filter: state.tagFilters.topic,
    });
    renderTagPanel("seminar", {
      title: "Seminar",
      description: "Training deck categories.",
      filter: state.tagFilters.seminar,
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
    const tags = previous?.tags || { bericht: [], topic: [], seminar: [] };
    return {
      path,
      file,
      url: URL.createObjectURL(file),
      notes: previous?.notes || "",
      tags: {
        bericht: Array.from(tags.bericht || []),
        topic: Array.from(tags.topic || []),
        seminar: Array.from(tags.seminar || []),
      },
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
    const seminarLabels = [];
    const topicLabels = [];
    rows.forEach((row) => {
      const reportA = String(row[0] || "").trim();
      const seminar = String(row[1] || "").trim();
      const topic = String(row[2] || "").trim();
      const reportB = String(row[3] || "").trim();
      if (reportA) reportLabels.push(reportA);
      if (reportB) reportLabels.push(reportB);
      if (seminar) seminarLabels.push(seminar);
      if (topic) topicLabels.push(topic);
    });

    const reportOptions = reportLabels.map((label) => {
      const match = label.match(/^\\d+(?:\\.\\d+)*(?:\\.)?/);
      const value = match ? match[0].replace(/\\.$/, "") : label;
      return { value, label };
    });

    return {
      bericht: reportOptions,
      seminar: seminarLabels.map((label) => ({ value: label, label })),
      topic: topicLabels.map((label) => ({ value: label, label })),
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
