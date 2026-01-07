(() => {
  const pickProjectBtn = document.getElementById("pick-project");
  const loadSidecarBtn = document.getElementById("load-sidecar");
  const loadCategoriesBtn = document.getElementById("load-categories");
  const pickPhotosBtn = document.getElementById("pick-photos");
  const scanPhotosBtn = document.getElementById("scan-photos");
  const saveSidecarBtn = document.getElementById("save-sidecar");
  const saveLogBtn = document.getElementById("save-log");
  const openSettingsBtn = document.getElementById("open-settings");
  const settingsModal = document.getElementById("settings-modal");
  const settingsCloseBtn = document.getElementById("settings-close-btn");
  const settingsBackdrop = document.getElementById("settings-close");
  const obsTagInput = document.getElementById("obs-tag-input");
  const obsTagAddBtn = document.getElementById("obs-tag-add");
  const obsTagList = document.getElementById("obs-tag-list");
  const statusEl = document.getElementById("status");
  const statusTextEl = document.getElementById("status-text");
  const statusCloseBtn = document.getElementById("status-close");
  const photoMetaEl = document.getElementById("photo-meta");
  const photoImageEl = document.getElementById("photo-image");
  const photoFilenameEl = document.getElementById("photo-filename");
  const photoUnsortedBtn = document.getElementById("photo-unsorted");
  const filterToggleBtn = document.getElementById("filter-toggle");
  const prevBtn = document.getElementById("prev-photo");
  const nextBtn = document.getElementById("next-photo");
  const notesEl = document.getElementById("photo-notes");

  const panels = {
    report: document.getElementById("panel-report"),
    observations: document.getElementById("panel-observations"),
    training: document.getElementById("panel-training"),
  };
  const observationsSlot = document.getElementById("panel-observations-slot");
  const viewerObservations = document.querySelector(".viewer__observations");

  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };
  const { setLocale = () => {} } = window.AutoBerichtI18n || {};
  const {
    saveHandle: saveFsHandle = async () => {},
    loadHandle: loadFsHandle = async () => null,
    requestHandlePermission: requestFsHandlePermission = async () => false,
  } = window.AutoBerichtFsHandle || {};

  const urlParams = new URLSearchParams(window.location.search);
  const demoMode = urlParams.get("demo") === "1" || urlParams.get("demo") === "true";
  const layoutParam = urlParams.get("layout");
  const storedLayout = window.localStorage?.getItem("photosorterLayout");
  let layoutMode =
    layoutParam === "tabs" || layoutParam === "stacked"
      ? layoutParam
      : storedLayout === "tabs"
        ? "tabs"
        : "stacked";
  const demoPhotosMode =
    urlParams.get("demoPhotos") === "1" ||
    urlParams.get("demoPhotos") === "true" ||
    demoMode;
  const panelTabButtons = Array.from(document.querySelectorAll("[data-panel-tab]"));
  const layoutToggleButtons = Array.from(document.querySelectorAll("[data-layout]"));

  const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"]);
  const DEFAULT_TAGS = {
    report: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
    observations: [
      "Absturzsicherung",
      "Arbeiten in der Höhe",
      "Arbeitsmittelprüfung",
      "Beleuchtung",
      "Brandschutz",
      "Chemikalien",
      "Druckbehälter",
      "Elektrik",
      "Ergonomie",
      "Erste Hilfe",
      "Fluchtwege",
      "Gefährdungsbeurteilung",
      "Handwerkzeuge",
      "Kennzeichnung",
      "Lagerung",
      "Lärm",
      "Leitern",
      "Maschinenwartung",
      "Maschinensicherung",
      "Notfallorganisation",
      "PSA",
      "Rutschgefahr",
      "Sauberkeit",
      "Schulung / Unterweisung",
      "Schweißarbeiten",
      "Sicherheitsdatenblätter",
      "Staplerverkehr",
      "Verkehrswege",
      "Werkstatt / Ordnung",
    ],
    training: ["Vorbildliches Verhalten", "Risikoanalyse", "Audit", "Kommunikation"],
  };

  const DEMO_PHOTO_URLS = [
    "./demo-photos/1.jpg",
    "./demo-photos/2.jpg",
    "./demo-photos/3.jpg",
    "./demo-photos/4.jpg",
    "./demo-photos/5.jpg",
  ];

  const state = {
    projectHandle: null,
    photoHandle: null,
    projectDoc: null,
    sidecarDoc: null,
    photoRootName: "",
    tagOptions: null,
    tagFilters: { report: "", observations: "", training: "" },
    photos: [],
    filterMode: "all",
    currentIndex: 0,
    activePanel: "report",
  };

  let autosaveTimer = null;
  let saveQueue = Promise.resolve();


  const updateStatusVisibility = (isHidden) => {
    statusEl.classList.toggle("is-hidden", isHidden);
  };

  const setStatusHidden = (hidden) => {
    if (window.localStorage) {
      window.localStorage.setItem("photosorterStatusHidden", hidden ? "1" : "0");
    }
    updateStatusVisibility(hidden);
  };

  const setStatus = (message) => {
    statusTextEl.textContent = message;
    debug.logLine("info", message);
  };

  const openSettings = () => {
    if (!settingsModal) return;
    settingsModal.classList.add("is-open");
    settingsModal.setAttribute("aria-hidden", "false");
    renderObservationTagList();
  };

  const closeSettings = () => {
    if (!settingsModal) return;
    settingsModal.classList.remove("is-open");
    settingsModal.setAttribute("aria-hidden", "true");
  };

  const setDefaultPhotoHandle = async () => {
    if (!state.projectHandle) return;
    if (!state.photoHandle) {
      state.photoHandle = state.projectHandle;
    }
    if (!state.photoRootName) {
      state.photoRootName = "";
    }
    const rootName = state.projectDoc?.photoRoot || "";
    if (rootName && rootName !== state.projectHandle.name) {
      try {
        const handle = await state.projectHandle.getDirectoryHandle(rootName);
        state.photoHandle = handle;
        state.photoRootName = rootName;
      } catch (err) {
        // Keep project root as fallback.
      }
    } else if (rootName === state.projectHandle.name) {
      state.photoHandle = state.projectHandle;
      state.photoRootName = "";
    }
    if (!rootName) {
      const candidates = ["Photos", "photos"];
      for (const name of candidates) {
        try {
          const handle = await state.projectHandle.getDirectoryHandle(name);
          state.photoHandle = handle;
          state.photoRootName = name;
          break;
        } catch (err) {
          // Try next candidate.
        }
      }
    }
  };

  const applyLayoutMode = () => {
    document.body.classList.toggle("layout-tabs", layoutMode === "tabs");
    document.body.classList.toggle("layout-stacked", layoutMode === "stacked");
  };

  const placeObservationsPanel = () => {
    const panel = panels.observations;
    if (!panel) return;
    const target = layoutMode === "tabs" ? observationsSlot : viewerObservations;
    if (!target || panel.parentElement === target) return;
    target.appendChild(panel);
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
    placeObservationsPanel();
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
      setStatus("File System Access API is not available. Open via http://localhost in Edge/Chrome to enable file access.");
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
    if (loadCategoriesBtn) {
      loadCategoriesBtn.disabled = !hasProject && !demoMode;
    }
    pickPhotosBtn.disabled = !hasProject;
    scanPhotosBtn.disabled = !hasPhotoFolder;
    saveSidecarBtn.disabled = !hasProject;
    filterToggleBtn.disabled = state.photos.length === 0;
    prevBtn.disabled = !hasVisiblePhotos;
    nextBtn.disabled = !hasVisiblePhotos;
  };

  const persistHandle = async (handle) => {
    try {
      await saveFsHandle(handle);
    } catch (err) {
      debug.logLine("warn", `Failed to persist folder handle: ${err.message || err}`);
    }
  };

  const loadPersistedHandle = async () => {
    try {
      return await loadFsHandle();
    } catch (err) {
      return null;
    }
  };

  const ensureHandlePermission = async (handle) => {
    try {
      return await requestFsHandlePermission(handle);
    } catch (err) {
      return false;
    }
  };

  const restoreLastHandle = async () => {
    if (!window.showDirectoryPicker) return;
    const handle = await loadPersistedHandle();
    if (!handle) return;
    const ok = await ensureHandlePermission(handle);
    if (!ok) return;
    state.projectHandle = handle;
    enableActions();
    setStatus(`Restored project folder: ${state.projectHandle.name}`);
    debug.logLine("info", `Restored project folder: ${state.projectHandle.name}`);
    await loadProjectSidecar();
  };

  if (statusCloseBtn) {
    statusCloseBtn.addEventListener("click", () => {
      setStatusHidden(true);
    });
  }

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

  const isUnsortedLabel = (value) => String(value || "").trim().toLowerCase() === "unsorted";

  const normalizeTagOption = (option) => {
    if (!option) return null;
    if (typeof option === "string") {
      if (isUnsortedLabel(option)) return null;
      return { value: option, label: option };
    }
    if (typeof option === "object") {
      const value = option.value || option.label;
      if (!value) return null;
      if (isUnsortedLabel(value)) return null;
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
    const sanitize = (list) => Array.from(list || []).filter((value) => !isUnsortedLabel(value));
    return {
      report: sanitize(incoming.report),
      observations: sanitize(incoming.observations),
      training: sanitize(incoming.training),
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

  const buildReportTagOptionsFromProject = (project) => {
    if (!project?.chapters) return null;
    const tags = new Map();
    const addTag = (value, label) => {
      const val = String(value || "").trim();
      if (!val || isUnsortedLabel(val)) return;
      let lbl = String(label || val).trim();
      if (lbl && !(lbl === val || lbl.startsWith(`${val} `) || lbl.startsWith(`${val}.`))) {
        lbl = `${val} ${lbl}`;
      }
      if (!tags.has(val)) tags.set(val, lbl || val);
    };
    const shouldSkipSection = (sectionId) => {
      const topLevel = String(sectionId || "").split(".")[0];
      return ["11", "12", "13", "14"].includes(topLevel);
    };
    project.chapters.forEach((chapter) => {
      addTag(chapter.id, chapter.title?.de || chapter.title || chapter.id);
      (chapter.rows || []).forEach((row) => {
        if (row.kind === "section") {
          if (!shouldSkipSection(row.id)) {
            addTag(row.id, row.title || row.id);
          }
          return;
        }
        if (row.sectionId) {
          if (!shouldSkipSection(row.sectionId)) {
            addTag(row.sectionId, row.sectionLabel || row.sectionId);
          }
        }
      });
    });
    const options = Array.from(tags.entries()).map(([value, label]) => ({ value, label }));
    return sortOptionsForGroup("report", options);
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
    photoTagOptions: structuredClone(SEED_TAG_OPTIONS),
    photoRoot: "",
  });

  const normalizePhotoDoc = (doc) => {
    const base = doc && typeof doc === "object" ? structuredClone(doc) : {};
    if (!base.meta) base.meta = { projectId: "", createdAt: new Date().toISOString() };
    if (!base.photos) base.photos = {};
    if (!base.photoTagOptions) base.photoTagOptions = structuredClone(SEED_TAG_OPTIONS);
    if (!base.photoRoot) base.photoRoot = "";
    return base;
  };

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
      if (photo.url && photo.url.startsWith("blob:")) {
        URL.revokeObjectURL(photo.url);
      }
    });
  };

  if (photoUnsortedBtn) {
    photoUnsortedBtn.addEventListener("click", () => {
      setPhotoUnsorted();
    });
  }

  const renderViewer = () => {
    const current = getCurrentPhoto();
    if (!current) {
      photoImageEl.removeAttribute("src");
      photoImageEl.alt = "No photo loaded";
      if (photoFilenameEl) photoFilenameEl.textContent = "";
      notesEl.value = "";
      notesEl.disabled = true;
      if (photoUnsortedBtn) {
        photoUnsortedBtn.classList.remove("active");
        photoUnsortedBtn.disabled = true;
      }
      updateMeta();
      enableActions();
      return;
    }
    photoImageEl.src = current.url;
    photoImageEl.alt = current.path;
    if (photoFilenameEl) {
      photoFilenameEl.textContent = current.path.split("/").pop();
    }
    notesEl.value = current.notes || "";
    notesEl.disabled = false;
    if (photoUnsortedBtn) {
      const isUnsorted = isPhotoUnsorted(current);
      photoUnsortedBtn.classList.toggle("active", isUnsorted);
      photoUnsortedBtn.disabled = false;
    }
    updateMeta();
  };

  const splitChapterOptions = (options) => {
    const chapters = [];
    const rest = [];
    options.forEach((option) => {
      const value = String(option.value || "");
      if (isNumericTag(value) && !value.includes(".")) {
        chapters.push(option);
      } else {
        rest.push(option);
      }
    });
    return { chapters, rest };
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

    controls.append(filterInput);
    if (config.allowAdd) {
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
      controls.append(addInput);
    }

    const current = getCurrentPhoto();
    const selected = new Set(current?.tags?.[group] || []);
    const filteredOptions = (state.tagOptions?.[group] || [])
      .filter((option) => option.label.toLowerCase().includes((config.filter || "").toLowerCase()));

    const { chapters, rest } = config.splitChapters
      ? splitChapterOptions(filteredOptions)
      : { chapters: [], rest: filteredOptions };

    if (config.splitChapters) {
      const chapterRow = document.createElement("div");
      chapterRow.className = "panel__chapters";
      chapters.forEach((option) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "tag-button tag-button--chapter";
        button.textContent = option.label;
        button.title = option.label;
        if (selected.has(option.value)) {
          button.classList.add("active");
        }
        button.addEventListener("click", () => {
          toggleTag(group, option.value);
        });
        chapterRow.appendChild(button);
      });
      container.append(chapterRow);
    }

    const tagsEl = document.createElement("div");
    tagsEl.className = "panel__tags";

    rest.forEach((option) => {
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
      splitChapters: true,
      allowAdd: false,
    });
    renderTagPanel("observations", {
      title: "Beobachtungen",
      description: "Themen/Begriffe aus der Begehung.",
      filter: state.tagFilters.observations,
      allowAdd: true,
    });
    renderTagPanel("training", {
      title: "Training",
      description: "Seminar-/Schulungskategorien.",
      filter: state.tagFilters.training,
      allowAdd: false,
    });
  };

  const renderAll = () => {
    const filtered = getFilteredPhotos();
    if (state.currentIndex >= filtered.length) {
      state.currentIndex = Math.max(0, filtered.length - 1);
    }
    filterToggleBtn.textContent = state.filterMode === "all" ? "Show Unsorted" : "Show All";
    renderViewer();
    renderPanels();
    enableActions();
  };

  const setGroupUnsorted = (group) => {
    const current = getCurrentPhoto();
    if (!current) return;
    current.tags[group] = [];
    scheduleAutosave();
    renderAll();
  };

  const setPhotoUnsorted = () => {
    const current = getCurrentPhoto();
    if (!current) return;
    current.tags.report = [];
    current.tags.observations = [];
    current.tags.training = [];
    scheduleAutosave();
    renderAll();
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
    scheduleAutosave();
    renderAll();
  };

  const removeObservationTag = (tag) => {
    if (!tag) return;
    const options = state.tagOptions?.observations || [];
    state.tagOptions.observations = options.filter((opt) => opt.value !== tag);
    state.photos.forEach((photo) => {
      const list = new Set(photo.tags?.observations || []);
      if (list.has(tag)) {
        list.delete(tag);
        photo.tags.observations = Array.from(list);
      }
    });
    scheduleAutosave();
    renderAll();
    renderObservationTagList();
  };

  const renderObservationTagList = () => {
    if (!obsTagList) return;
    obsTagList.innerHTML = "";
    const options = state.tagOptions?.observations || [];
    const counts = new Map();
    state.photos.forEach((photo) => {
      (photo.tags?.observations || []).forEach((tag) => {
        counts.set(tag, (counts.get(tag) || 0) + 1);
      });
    });

    options.forEach((option) => {
      const row = document.createElement("div");
      row.className = "settings-tags__row";
      const label = document.createElement("span");
      label.textContent = option.label;
      const count = document.createElement("span");
      count.className = "settings-tags__count";
      count.textContent = String(counts.get(option.value) || 0);
      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "Remove";
      removeBtn.addEventListener("click", () => {
        const confirmed = window.confirm(`Remove "${option.label}" and clear it from all photos?`);
        if (!confirmed) return;
        removeObservationTag(option.value);
      });
      row.append(label, count, removeBtn);
      obsTagList.appendChild(row);
    });
  };

  const addObservationTag = () => {
    if (!obsTagInput) return;
    const value = obsTagInput.value.trim();
    if (!value) return;
    const existing = state.tagOptions.observations || [];
    if (!existing.some((option) => option.value === value)) {
      state.tagOptions.observations = sortOptionsForGroup("observations", [
        ...existing,
        { value, label: value },
      ]);
      scheduleAutosave();
      renderAll();
      renderObservationTagList();
    }
    obsTagInput.value = "";
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
      const sidecar = JSON.parse(text);
      state.sidecarDoc = sidecar;
      const photoDoc = sidecar?.photos || sidecar;
      state.projectDoc = normalizePhotoDoc(photoDoc);
      if (sidecar?.report?.project?.meta?.locale) {
        setLocale(sidecar.report.project.meta.locale);
      }
      state.tagOptions = ensureTagOptions(state.projectDoc.photoTagOptions);
      const reportOptions = buildReportTagOptionsFromProject(sidecar?.report?.project);
      if (reportOptions && reportOptions.length) {
        state.tagOptions.report = reportOptions;
        state.projectDoc.photoTagOptions.report = reportOptions;
      }
      state.photoRootName = state.projectDoc.photoRoot || "";
      setStatus("Loaded project_sidecar.json");
      debug.logLine("info", "Loaded project_sidecar.json");
    } catch (err) {
      state.sidecarDoc = null;
      state.projectDoc = createEmptyProjectDoc();
      state.tagOptions = ensureTagOptions(DEFAULT_TAGS);
      state.photoRootName = "";
      setStatus("Sidecar not found; starting fresh.");
      debug.logLine("warn", `Sidecar not found: ${err.message || err}`);
    }
    await setDefaultPhotoHandle();
    const didScan = await maybeAutoScan();
    if (!didScan) {
      renderAll();
    }
  };

  const loadDemoPhotos = async () => {
    if (!DEMO_PHOTO_URLS.length) return;
    if (!state.projectDoc) {
      state.projectDoc = createEmptyProjectDoc();
    }
    setStatus("Loading demo photos...");
    const collection = [];
    for (let i = 0; i < DEMO_PHOTO_URLS.length; i += 1) {
      const url = DEMO_PHOTO_URLS[i];
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`Failed to fetch ${url}`);
      }
      const blob = await response.blob();
      const file = new File([blob], `${i + 1}.jpg`, { type: blob.type || "image/jpeg" });
      collection.push(buildPhotoEntry(`demo-photos/${i + 1}.jpg`, file));
    }
    state.photoRootName = "demo-photos";
    state.photos = collection;
    state.filterMode = "all";
    state.currentIndex = 0;
    setStatus(`Loaded ${collection.length} demo photos`);
    renderAll();
  };

  const saveProjectSidecar = async () => {
    if (!state.projectHandle) return;
    const payload = normalizePhotoDoc(state.projectDoc);
    payload.photos = serializePhotos();
    payload.photoTagOptions = structuredClone(state.tagOptions);
    payload.photoRoot = state.photoRootName || "";
    if (!payload.meta) payload.meta = {};
    payload.meta.updatedAt = new Date().toISOString();

    saveQueue = saveQueue.then(async () => {
      let existing = state.sidecarDoc;
      try {
        const existingHandle = await state.projectHandle.getFileHandle("project_sidecar.json");
        const existingFile = await existingHandle.getFile();
        const existingText = await existingFile.text();
        existing = JSON.parse(existingText);
      } catch (err) {
        existing = state.sidecarDoc;
      }
      const sidecar = existing && typeof existing === "object" ? structuredClone(existing) : {};
      if (!sidecar.meta) sidecar.meta = {};
      sidecar.meta.updatedAt = new Date().toISOString();
      sidecar.photos = payload;
      const handle = await state.projectHandle.getFileHandle("project_sidecar.json", { create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(sidecar, null, 2));
      await writable.close();
      state.projectDoc = payload;
      state.sidecarDoc = sidecar;
      setStatus("Saved photo tags to project_sidecar.json");
      debug.logLine("info", "Saved photo tags to project_sidecar.json");
    }).catch((err) => {
      setStatus(`Save failed: ${err.message}`);
      debug.logLine("error", `Save failed: ${err.message}`);
    });
    return saveQueue;
  };

  const scheduleAutosave = () => {
    if (!state.projectHandle) return;
    if (autosaveTimer) clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(async () => {
      autosaveTimer = null;
      await saveProjectSidecar();
    }, 2000);
  };

  const flushAutosave = async () => {
    if (autosaveTimer) {
      clearTimeout(autosaveTimer);
      autosaveTimer = null;
    }
    await saveProjectSidecar();
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
    const prefix = state.photoRootName ? `${state.photoRootName}/` : "";
    await collectImages(state.photoHandle, prefix, collection);
    collection.sort((a, b) => a.path.localeCompare(b.path));
    state.photos = collection;
    state.filterMode = "all";
    state.currentIndex = 0;
    filterToggleBtn.textContent = "Show Unsorted";
    setStatus(`Loaded ${state.photos.length} photos from ${state.photoRootName}`);
    renderAll();
  };

  const maybeAutoScan = async () => {
    if (!state.photoHandle) return false;
    if (state.photos.length > 0) return false;
    const savedPhotos = state.projectDoc?.photos || {};
    const hasSaved = Object.keys(savedPhotos).length > 0;
    const hasRoot = !!(state.projectDoc?.photoRoot || state.photoRootName);
    if (!hasSaved && !hasRoot) return false;
    await scanPhotos();
    return true;
  };

  pickProjectBtn.addEventListener("click", async () => {
    if (!ensureFsAccess()) return;
    try {
      state.projectHandle = await window.showDirectoryPicker({ mode: "readwrite", id: "autobericht-project" });
      await persistHandle(state.projectHandle);
      setStatus(`Project folder: ${state.projectHandle.name}`);
      await loadProjectSidecar();
      setStatus(`Project folder: ${state.projectHandle.name}`);
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
      if (training && !isUnsortedLabel(training)) trainingLabels.push(training);
      if (observation && !isUnsortedLabel(observation)) observationLabels.push(observation);
    });

    const reportOptions = reportLabels.map((label) => {
      const match = label.match(/^\\d+(?:\\.\\d+)*(?:\\.)?/);
      if (!match) return null;
      const value = match[0].replace(/\\.$/, "");
      return { value, label };
    }).filter(Boolean);

    return {
      report: reportOptions,
      training: trainingLabels.map((label) => ({ value: label, label })),
      observations: observationLabels.map((label) => ({ value: label, label })),
    };
  };

  if (loadCategoriesBtn) {
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
  }

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

  if (openSettingsBtn) {
    openSettingsBtn.addEventListener("click", openSettings);
  }
  if (settingsCloseBtn) {
    settingsCloseBtn.addEventListener("click", closeSettings);
  }
  if (settingsBackdrop) {
    settingsBackdrop.addEventListener("click", closeSettings);
  }
  if (obsTagAddBtn) {
    obsTagAddBtn.addEventListener("click", addObservationTag);
  }
  if (obsTagInput) {
    obsTagInput.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      addObservationTag();
    });
  }

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
    scheduleAutosave();
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

  window.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      flushAutosave();
    }
  });
  window.addEventListener("pagehide", () => {
    flushAutosave();
  });

  state.tagOptions = SEED_TAG_OPTIONS;
  const statusHidden = window.localStorage?.getItem("photosorterStatusHidden") === "1";
  updateStatusVisibility(statusHidden);
  applyLayoutMode();
  placeObservationsPanel();
  updateLayoutToggle();
  setActivePanel(state.activePanel);
  ensureFsAccess();
  restoreLastHandle();
  enableActions();

  if (window.PS_CATEGORY_LABELS_SEED) {
    setStatus("Loaded categories from bundled seed.");
  }


  if (demoPhotosMode) {
    loadDemoPhotos().catch((err) => {
      setStatus(`Demo photos failed: ${err.message}`);
      debug.logLine("error", `Demo photos failed: ${err.message}`);
    });
  }
})();
