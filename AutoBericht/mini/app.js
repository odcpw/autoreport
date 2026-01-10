(() => {
  const pickFolderBtn = document.getElementById("pick-folder");
  const loadSidecarBtn = document.getElementById("load-sidecar");
  const saveSidecarBtn = document.getElementById("save-sidecar");
  const importSelfBtn = document.getElementById("import-self");
  const loadSeedsBtn = document.getElementById("load-seeds");
  const saveLogBtn = document.getElementById("save-log");
  const openSettingsBtn = document.getElementById("open-settings");
  const settingsModal = document.getElementById("settings-modal");
  const closeSettingsBtn = document.getElementById("close-settings");
  const saveSettingsBtn = document.getElementById("save-settings");
  const generateLibraryBtn = document.getElementById("generate-library");
  const settingsAuthorEl = document.getElementById("settings-author");
  const settingsInitialsEl = document.getElementById("settings-initials");
  const settingsLocaleEl = document.getElementById("settings-locale");
  const settingsLibraryHintEl = document.getElementById("settings-library-hint");
  const statusEl = document.getElementById("status");
  const chapterListEl = document.getElementById("chapter-list");
  const chapterTitleEl = document.getElementById("chapter-title");
  const rowsEl = document.getElementById("rows");
  const photoOverlayEl = document.getElementById("photo-overlay");
  const photoOverlayClose = document.getElementById("photo-overlay-close");
  const photoOverlayCloseBtn = document.getElementById("photo-overlay-close-btn");
  const photoOverlayPrevBtn = document.getElementById("photo-overlay-prev");
  const photoOverlayNextBtn = document.getElementById("photo-overlay-next");
  const photoOverlayTitle = document.getElementById("photo-overlay-title");
  const photoOverlayImage = document.getElementById("photo-overlay-image");

  const filterModeEls = Array.from(document.querySelectorAll("input[name=\"filter-mode\"]"));

  let dirHandle = null;
  let sidecarDoc = null;
  let autosaveTimer = null;
  let saveQueue = Promise.resolve();
  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };
  const {
    t = (key, fallback) => fallback || key,
    setLocale = () => {},
  } = window.AutoBerichtI18n || {};
  const {
    escapeHtml = (value) => value,
    formatInlineMarkdown = (value) => value,
    markdownToHtml = (value) => value,
  } = window.AutoBerichtMarkdown || {};
  const {
    saveHandle: saveFsHandle = async () => {},
    loadHandle: loadFsHandle = async () => null,
    requestHandlePermission: requestFsHandlePermission = async () => false,
  } = window.AutoBerichtFsHandle || {};

  const defaultProject = {
    meta: {
      projectId: "2026-TEST-001",
      company: "ACME AG",
      locale: "de-CH",
      author: "consultant@example.com",
      createdAt: new Date().toISOString(),
    },
    chapters: [
      {
        id: "1",
        title: { de: "Leitbild" },
        rows: [
          {
            id: "1.1.1",
            type: "standard",
            titleOverride: "Unternehmensleitbild",
            master: {
              finding: "Das Unternehmen verfuegt nicht ueber ein Leitbild.",
              levels: {
                "1": "Eine Sicherheitscharta als ersten Schritt etablieren.",
                "2": "Die vorhandenen Werte in die Fuhrung integrieren.",
                "3": "Die dokumentierte Charta verbreiten und leben.",
                "4": "Die Charta als aktives Fuehrungsinstrument nutzen.",
              },
            },
            customer: {
              answer: 1,
              remark: "Leitbild vorhanden",
              items: [
                {
                  id: "1.1.1",
                  question: "Gibt es ein Sicherheitsleitbild?",
                  collapsedId: "1.1.1",
                },
              ],
            },
            workstate: {
              selectedLevel: 2,
              includeFinding: true,
              includeRecommendation: true,
              done: false,
              useFindingOverride: false,
              findingOverride: "",
              useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
            },
          },
          {
            id: "1.1.2",
            type: "standard",
            titleOverride: "Strategie",
            master: {
              finding: "Es fehlt eine dokumentierte Sicherheitsstrategie.",
              levels: {
                "1": "Strategie-Grundsaetze definieren.",
                "2": "Strategie in Ziele uebersetzen.",
                "3": "Strategie regelmaessig pruefen.",
                "4": "Strategie in allen Bereichen verankern.",
              },
            },
            customer: {
              answer: 0,
              remark: "",
              items: [
                {
                  id: "1.1.2",
                  question: "Gibt es eine dokumentierte Sicherheitsstrategie?",
                  collapsedId: "1.1.2",
                },
              ],
            },
            workstate: {
              selectedLevel: 3,
              includeFinding: true,
              includeRecommendation: true,
              done: true,
              useFindingOverride: true,
              findingOverride: "Eine Strategie besteht, wird aber nicht aktiv kommuniziert.",
              useLevelOverride: { "1": false, "2": true, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "Strategie sichtbar machen.", "3": "", "4": "" },
            },
          },
        ],
      },
      {
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [
          {
            id: "4.8.1",
            type: "field_observation",
            titleOverride: "Regale",
            master: {
              finding: "Regale sind nicht gegen Kippen gesichert.",
              levels: {
                "1": "Regale sichern und Sichtkontrolle definieren.",
                "2": "Regalinspektionen regelmaessig durchfuehren.",
                "3": "Sicherheitschecks dokumentieren.",
                "4": "Regalmanagement im Sicherheitsprogramm verankern.",
              },
            },
            customer: {
              answer: 0,
              remark: "",
              items: [
                {
                  id: "4.8.1",
                  question: "Sind Regale gegen Kippen gesichert?",
                  collapsedId: "4.8.1",
                },
              ],
            },
            workstate: {
              selectedLevel: 1,
              includeFinding: true,
              includeRecommendation: true,
              done: false,
              useFindingOverride: false,
              findingOverride: "",
              useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
            },
          },
        ],
      },
    ],
  };

  const state = {
    project: structuredClone(defaultProject),
    selectedChapterId: defaultProject.chapters[0].id,
    filters: {
      mode: "all",
    },
    photoIndex: {
      report: new Map(),
      observations: new Map(),
    },
    photoRoot: "",
    photoOverlay: {
      tag: "",
      items: [],
      index: 0,
      url: "",
    },
  };

  const setStatus = (message) => {
    statusEl.textContent = message;
    debug.logLine("info", message);
  };

  const readJsonFromHttp = async (path) => {
    const response = await fetch(path);
    if (!response.ok) {
      throw new Error(`Failed to fetch ${path}`);
    }
    return response.json();
  };

  const buildPhotoIndex = () => {
    const reportIndex = new Map();
    const observationIndex = new Map();
    const photoDoc = sidecarDoc?.photos;
    const photos = photoDoc?.photos || {};
    state.photoRoot = photoDoc?.photoRoot || "";
    Object.entries(photos).forEach(([path, data]) => {
      const reportTags = data?.tags?.report || [];
      reportTags.forEach((tag) => {
        const key = String(tag || "").trim();
        if (!key) return;
        if (!reportIndex.has(key)) reportIndex.set(key, []);
        reportIndex.get(key).push({
          path,
          notes: data?.notes || "",
        });
      });
      const observationTags = data?.tags?.observations || [];
      observationTags.forEach((tag) => {
        const key = String(tag || "").trim();
        if (!key) return;
        if (!observationIndex.has(key)) observationIndex.set(key, []);
        observationIndex.get(key).push({
          path,
          notes: data?.notes || "",
        });
      });
    });
    reportIndex.forEach((items) => {
      items.sort((a, b) => a.path.localeCompare(b.path, "de", { numeric: true }));
    });
    observationIndex.forEach((items) => {
      items.sort((a, b) => a.path.localeCompare(b.path, "de", { numeric: true }));
    });
    state.photoIndex = {
      report: reportIndex,
      observations: observationIndex,
    };
  };

  const getPhotosForTag = (tag, kind = "report") => {
    if (!tag) return [];
    const index = state.photoIndex?.[kind];
    if (!index) return [];
    return index.get(tag) || [];
  };

  const resolvePhotoFile = async (path) => {
    if (!dirHandle) throw new Error("Project folder not selected.");
    const parts = String(path || "").split("/").filter(Boolean);
    let current = dirHandle;
    for (let i = 0; i < parts.length - 1; i += 1) {
      current = await current.getDirectoryHandle(parts[i]);
    }
    const fileHandle = await current.getFileHandle(parts[parts.length - 1]);
    return fileHandle.getFile();
  };

  const loadOverlayImage = async (path) => {
    if (!photoOverlayImage) return;
    if (state.photoOverlay.url) {
      URL.revokeObjectURL(state.photoOverlay.url);
      state.photoOverlay.url = "";
    }
    try {
      const file = await resolvePhotoFile(path);
      const url = URL.createObjectURL(file);
      state.photoOverlay.url = url;
      photoOverlayImage.src = url;
      photoOverlayImage.alt = path;
    } catch (err) {
      photoOverlayImage.removeAttribute("src");
      photoOverlayImage.alt = "Photo not available";
      setStatus(`Photo not found: ${path}`);
    }
  };

  const renderPhotoOverlay = () => {
    if (!photoOverlayEl) return;
    const { items, index, tag } = state.photoOverlay;
    if (!items.length) return;
    const current = items[index];
    const filename = current.path.split("/").pop();
    if (photoOverlayTitle) {
      photoOverlayTitle.textContent = `${tag} • Image ${index + 1} of ${items.length} • ${filename}`;
    }
    if (photoOverlayPrevBtn) {
      photoOverlayPrevBtn.disabled = items.length <= 1;
    }
    if (photoOverlayNextBtn) {
      photoOverlayNextBtn.disabled = items.length <= 1;
    }
    loadOverlayImage(current.path);
  };

  const openPhotoOverlay = (tag, kind = "report") => {
    const items = getPhotosForTag(tag, kind);
    if (!items.length) return;
    state.photoOverlay = {
      tag,
      items,
      index: 0,
      url: "",
      kind,
    };
    if (photoOverlayEl) {
      photoOverlayEl.classList.add("is-open");
      photoOverlayEl.setAttribute("aria-hidden", "false");
    }
    renderPhotoOverlay();
  };

  const closePhotoOverlay = () => {
    if (photoOverlayEl) {
      photoOverlayEl.classList.remove("is-open");
      photoOverlayEl.setAttribute("aria-hidden", "true");
    }
    if (state.photoOverlay.url) {
      URL.revokeObjectURL(state.photoOverlay.url);
    }
    state.photoOverlay = {
      tag: "",
      items: [],
      index: 0,
      url: "",
      kind: "report",
    };
  };

  const stepPhotoOverlay = (delta) => {
    const total = state.photoOverlay.items.length;
    if (!total) return;
    state.photoOverlay.index = (state.photoOverlay.index + delta + total) % total;
    renderPhotoOverlay();
  };

  const seedPathCandidates = (filename) => ([
    ["AutoBericht", "data", "seed", filename],
  ]);

  const readSeedFromProject = async (rootHandle, filename) => {
    if (!rootHandle) return null;
    for (const parts of seedPathCandidates(filename)) {
      const data = await readJsonIfExists(rootHandle, parts);
      if (data) return data;
    }
    return null;
  };

  const seedHttpCandidates = (filename) => {
    if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
      return [];
    }
    const urls = [
      new URL(`../data/seed/${filename}`, window.location.href).toString(),
    ];
    return Array.from(new Set(urls));
  };

  const readSeedFromHttp = async (filename) => {
    const candidates = seedHttpCandidates(filename);
    for (const url of candidates) {
      try {
        return await readJsonFromHttp(url);
      } catch (err) {
        // try next candidate
      }
    }
    return null;
  };

  const autoLoadSeeds = async () => {
    if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
      debug.logLine("info", "Seed auto-load skipped (not on http).");
      return false;
    }
    try {
      state.project = await loadSeedsForProject();
      ensureProjectMeta();
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      render();
      setStatus("Loaded seed data.");
      debug.logLine("info", "Auto-loaded seed data.");
      return true;
    } catch (err) {
      debug.logLine("warn", `Seed auto-load skipped: ${err.message}`);
      return false;
    }
  };

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus("File System Access API is not available. Open via http://localhost in Edge/Chrome to enable file access.");
      pickFolderBtn.disabled = true;
      return false;
    }
    return true;
  };

  const enableActions = () => {
    const enabled = !!dirHandle;
    loadSidecarBtn.disabled = !enabled;
    saveSidecarBtn.disabled = !enabled;
    loadSeedsBtn.disabled = !enabled;
    if (importSelfBtn) {
      importSelfBtn.disabled = !enabled;
    }
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
    dirHandle = handle;
    enableActions();
    setStatus(`Restored project folder: ${dirHandle.name}`);
    debug.logLine("info", `Restored project folder: ${dirHandle.name}`);
    await loadProjectFromFolder();
  };

  const loadProjectFromFolder = async () => {
    if (!dirHandle) return false;
    try {
      const handle = await dirHandle.getFileHandle("project_sidecar.json");
      const file = await handle.getFile();
      const text = await file.text();
      sidecarDoc = JSON.parse(text);
      const reportProject = extractReportProject(sidecarDoc) || structuredClone(defaultProject);
      state.project = ensureObservationChapter(reportProject);
      ensureManagementSummaryChapter(state.project);
      syncObservationChapterRows(state.project, sidecarDoc);
      ensureProjectMeta();
      buildPhotoIndex();
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      render();
      setStatus("Loaded project_sidecar.json");
      debug.logLine("info", "Loaded project_sidecar.json");
      return true;
    } catch (err) {
      try {
        state.project = await loadSeedsForProject();
        ensureObservationChapter(state.project);
        ensureManagementSummaryChapter(state.project);
        syncObservationChapterRows(state.project, sidecarDoc);
        ensureProjectMeta();
        state.selectedChapterId = state.project.chapters[0]?.id || "";
        sidecarDoc = null;
        buildPhotoIndex();
        render();
        setStatus("Sidecar not found; loaded seed data.");
        debug.logLine("warn", `Sidecar not found; loaded seeds (${err.message || err})`);
        return true;
      } catch (seedErr) {
        state.project = structuredClone(defaultProject);
        ensureObservationChapter(state.project);
        ensureManagementSummaryChapter(state.project);
        syncObservationChapterRows(state.project, sidecarDoc);
        state.selectedChapterId = state.project.chapters[0].id;
        sidecarDoc = null;
        buildPhotoIndex();
        render();
        setStatus("Sidecar not found; loaded default template.");
        debug.logLine("warn", `Sidecar not found: ${err.message || err}`);
        return false;
      }
    }
  };

  const getChapterTitle = (chapter) => {
    if (!chapter) return "";
    if (typeof chapter.title === "string") return chapter.title;
    if (chapter.title && chapter.title.de) return chapter.title.de;
    if (chapter.id === "4.8") return "Beobachtungen";
    return chapter.id || "";
  };

  const compareIdSegments = (a, b) => {
    const aParts = String(a || "").split(".");
    const bParts = String(b || "").split(".");
    const maxLen = Math.max(aParts.length, bParts.length);
    for (let i = 0; i < maxLen; i += 1) {
      const aPart = aParts[i];
      const bPart = bParts[i];
      if (aPart == null) return -1;
      if (bPart == null) return 1;
      const aNum = Number(aPart);
      const bNum = Number(bPart);
      const aIsNum = !Number.isNaN(aNum);
      const bIsNum = !Number.isNaN(bNum);
      if (aIsNum && bIsNum) {
        if (aNum !== bNum) return aNum - bNum;
      } else {
        const cmp = String(aPart).localeCompare(String(bPart), "de", { numeric: true });
        if (cmp !== 0) return cmp;
      }
    }
    return 0;
  };

  const formatChapterLabel = (chapter) => {
    if (!chapter) return "";
    const title = getChapterTitle(chapter);
    const id = chapter.id || "";
    if (!title) return id;
    if (!id) return title;
    const normalized = title.trim();
    if (normalized === id || normalized.startsWith(`${id} `) || normalized.startsWith(`${id}.`)) {
      return normalized;
    }
    return `${id} ${title}`.trim();
  };

  const toText = (value) => {
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  };

  const calculateScore = (row) => {
    if (row?.type === "field_observation" || row?.type === "summary") return null;
    const ws = row?.workstate || {};
    if (ws.includeFinding === false) return 100;
    const level = Number(ws.selectedLevel || 1);
    return Math.max(0, Math.min(100, (level - 1) * 25));
  };

  const sanitizeFilename = (value) => String(value || "")
    .trim()
    .replace(/[^a-zA-Z0-9._-]/g, "_");

  const getLibraryFileName = () => {
    const meta = state.project?.meta || {};
    const locale = sanitizeFilename(meta.locale || "de-CH");
    const initials = sanitizeFilename(meta.initials || "");
    if (initials) return `library_user_${initials}_${locale}.json`;
    return `library_user_${locale}.json`;
  };

  const hashText = (value) => {
    const text = String(value || "");
    let hash = 2166136261;
    for (let i = 0; i < text.length; i += 1) {
      hash ^= text.charCodeAt(i);
      hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
    }
    return String(hash >>> 0);
  };

  const ensureWorkstateDefaults = (row) => {
    if (!row.workstate) row.workstate = {};
    const ws = row.workstate;
    if (ws.selectedLevel == null) ws.selectedLevel = 1;
    if (ws.includeFinding == null) ws.includeFinding = true;
    if (ws.includeRecommendation == null) ws.includeRecommendation = true;
    if (ws.done == null) ws.done = false;
    if (!ws.useFindingOverride) ws.useFindingOverride = false;
    if (!ws.findingOverride) ws.findingOverride = "";
    ws.useLevelOverride = ws.useLevelOverride || { "1": false, "2": false, "3": false, "4": false };
    ws.levelOverrides = ws.levelOverrides || { "1": "", "2": "", "3": "", "4": "" };
    ws.libraryActions = ws.libraryActions || { "1": "off", "2": "off", "3": "off", "4": "off" };
    ws.libraryHashes = ws.libraryHashes || { "1": "", "2": "", "3": "", "4": "" };
  };

  const normalizeLibraryEntries = (libraryData) => {
    if (!libraryData) return [];
    if (Array.isArray(libraryData.entries)) return libraryData.entries;
    if (libraryData.entries && typeof libraryData.entries === "object") {
      return Object.entries(libraryData.entries).map(([id, levels]) => ({ id, levels }));
    }
    return [];
  };

  const buildLibraryMap = (masterData, userData) => {
    const masterEntries = normalizeLibraryEntries(masterData);
    const userEntries = normalizeLibraryEntries(userData);
    const map = new Map(masterEntries.map((entry) => [entry.id, entry]));
    userEntries.forEach((entry) => {
      const existing = map.get(entry.id) || { id: entry.id, levels: {}, finding: "" };
      const mergedLevels = { ...(existing.levels || {}) };
      Object.entries(entry.levels || {}).forEach(([key, value]) => {
        const text = toText(value).trim();
        if (text) mergedLevels[key] = text;
      });
      const merged = {
        ...existing,
        levels: mergedLevels,
      };
      if (entry.finding) merged.finding = entry.finding;
      if (entry.lastUsed) merged.lastUsed = entry.lastUsed;
      if (!merged.lastUsed && existing.lastUsed) merged.lastUsed = existing.lastUsed;
      map.set(entry.id, merged);
    });
    return map;
  };

  const buildProjectFromSeeds = (selbstData, masterLibrary, userLibrary) => {
    const libraryMap = buildLibraryMap(masterLibrary, userLibrary);
    const groups = new Map();

    (selbstData.items || []).forEach((item) => {
      const collapsedId = item.collapsedId || item.groupId || item.id;
      if (!collapsedId) return;
      const rawId = String(collapsedId);
      let chapterId = rawId.split(".")[0];
      if (rawId === "4.8" || rawId.startsWith("4.8.")) {
        chapterId = "4.8";
      }
      const group = groups.get(collapsedId) || {
        id: collapsedId,
        chapterId,
        items: [],
        chapterLabel: item.chapterLabel || "",
      };
      group.items.push(item);
      if (!group.chapterLabel && item.chapterLabel) {
        group.chapterLabel = item.chapterLabel;
      }
      groups.set(collapsedId, group);
    });

    const chaptersById = new Map();
    groups.forEach((group) => {
      const chapterId = group.chapterId || "0";
      const chapter = chaptersById.get(chapterId) || {
        id: chapterId,
        title: {
          de: group.chapterLabel || (chapterId === "4.8" ? "Beobachtungen" : `Kapitel ${chapterId}`),
        },
        rows: [],
      };
      const master = libraryMap.get(group.id) || null;
      const sectionLabel = group.items.find((item) => item.sectionLabel)?.sectionLabel || "";
      const sectionId = String(group.id).split(".").slice(0, 2).join(".");
      chapter.rows.push({
        id: group.id,
        type: chapterId === "4.8" ? "field_observation" : "standard",
        sectionId,
        sectionLabel,
        titleOverride: group.items[0]?.question || "",
        master,
        customer: {
          answer: null,
          remark: "",
          items: group.items,
        },
        workstate: {
          selectedLevel: 1,
          includeFinding: true,
          includeRecommendation: true,
          done: false,
          useFindingOverride: false,
          findingOverride: "",
          useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
          levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
        },
      });
      chaptersById.set(chapterId, chapter);
    });

    const chapters = Array.from(chaptersById.values()).sort((a, b) => compareIdSegments(a.id, b.id));
    chapters.forEach((chapter) => {
      chapter.rows.sort((a, b) => compareIdSegments(a.id, b.id));
      const withSections = [];
      let lastSection = "";
      chapter.rows.forEach((row) => {
        if (row.sectionLabel && row.sectionId !== lastSection) {
          withSections.push({
            kind: "section",
            id: row.sectionId,
            title: row.sectionLabel,
          });
          lastSection = row.sectionId;
        }
        withSections.push(row);
      });
      chapter.rows = withSections;
    });

    if (!chaptersById.has("4.8")) {
      chapters.push({
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [],
      });
      chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    }

    return {
      meta: {
        projectId: "seed-import",
        company: "",
        locale: "de-CH",
        author: "",
        createdAt: new Date().toISOString(),
      },
      chapters,
    };
  };

  const ensureObservationChapter = (project) => {
    if (!project?.chapters) return project;
    const hasObservation = project.chapters.some((chapter) => chapter.id === "4.8");
    if (!hasObservation) {
      project.chapters.push({
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [],
      });
    }
    project.chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    return project;
  };

  const ensureManagementSummaryChapter = (project) => {
    if (!project?.chapters) return project;
    const hasSummary = project.chapters.some((chapter) => chapter.id === "0");
    if (!hasSummary) {
      project.chapters.unshift({
        id: "0",
        title: { de: "Management Summary" },
        rows: Array.from({ length: 8 }).map((_, index) => ({
          id: `0.${index + 1}`,
          type: "summary",
          titleOverride: "",
          master: {
            finding: "",
            levels: {
              "1": "",
              "2": "",
              "3": "",
              "4": "",
            },
          },
          customer: {
            answer: null,
            remark: "",
            items: [],
          },
          workstate: {
            selectedLevel: 1,
            includeFinding: true,
            includeRecommendation: true,
            done: false,
            useFindingOverride: true,
            findingOverride: "",
            useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
            levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
          },
        })),
      });
    }
    project.chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    return project;
  };

  const slugifyTag = (value) => String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/^-+|-+$/g, "") || "tag";

  const getObservationTagsFromSidecar = (doc) => {
    const options = doc?.photos?.photoTagOptions?.observations || [];
    return options
      .map((opt) => (typeof opt === "string" ? opt : opt.value || opt.label))
      .map((val) => String(val || "").trim())
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, "de", { numeric: true }));
  };

  const buildObservationRow = (tag) => ({
    id: `4.8:${slugifyTag(tag)}`,
    type: "field_observation",
    tag,
    titleOverride: tag,
    master: {
      finding: "Es wurden folgende unsichere Situationen beobachtet.",
      levels: {
        "1": "",
        "2": "",
        "3": "",
        "4": "",
      },
    },
    customer: {
      answer: null,
      remark: "",
      items: [],
    },
    workstate: {
      selectedLevel: 1,
      includeFinding: true,
      includeRecommendation: true,
      done: false,
      useFindingOverride: true,
      findingOverride: "Es wurden folgende unsichere Situationen beobachtet.",
      useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
      levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
    },
  });

  const syncObservationChapterRows = (project, doc) => {
    if (!project?.chapters) return;
    const chapter = project.chapters.find((item) => item.id === "4.8");
    if (!chapter) return;
    const tags = getObservationTagsFromSidecar(doc);
    if (!tags.length) return;

    const existingRows = (chapter.rows || []).filter((row) => row.kind !== "section");
    const byTag = new Map();
    existingRows.forEach((row) => {
      const tag = row.tag || row.titleOverride;
      if (tag) {
        row.tag = tag;
        row.titleOverride = tag;
        byTag.set(tag, row);
      }
    });

    const nextRows = tags.map((tag) => {
      const existing = byTag.get(tag);
      if (existing) {
        existing.titleOverride = tag;
        return existing;
      }
      return buildObservationRow(tag);
    });

    if (!chapter.meta) chapter.meta = {};
    if (!Array.isArray(chapter.meta.order)) {
      chapter.meta.order = nextRows.map((row) => row.id);
    } else {
      const order = chapter.meta.order.filter((id) => nextRows.some((row) => row.id === id));
      nextRows.forEach((row) => {
        if (!order.includes(row.id)) order.push(row.id);
      });
      chapter.meta.order = order;
    }

    chapter.rows = nextRows;
  };

  const extractReportProject = (doc) => {
    if (!doc || typeof doc !== "object") return null;
    if (doc.report && doc.report.project && doc.report.project.chapters) return doc.report.project;
    if (doc.chapters) return doc;
    return null;
  };

  const mergeSidecar = (baseDoc, project) => {
    const merged = baseDoc && typeof baseDoc === "object" ? structuredClone(baseDoc) : {};
    if (!merged.meta) merged.meta = {};
    merged.meta.updatedAt = new Date().toISOString();
    merged.report = { project };
    return merged;
  };

  const readJsonFromPath = async (rootHandle, pathParts) => {
    let currentHandle = rootHandle;
    for (let i = 0; i < pathParts.length - 1; i += 1) {
      currentHandle = await currentHandle.getDirectoryHandle(pathParts[i]);
    }
    const fileHandle = await currentHandle.getFileHandle(pathParts[pathParts.length - 1]);
    const file = await fileHandle.getFile();
    const text = await file.text();
    return JSON.parse(text);
  };

  const readJsonIfExists = async (rootHandle, pathParts) => {
    try {
      return await readJsonFromPath(rootHandle, pathParts);
    } catch (err) {
      return null;
    }
  };

  const loadSeedsForProject = async () => {
    let selbst = null;
    let masterLibrary = null;
    let userLibrary = null;
    if (dirHandle) {
      selbst = await readSeedFromProject(dirHandle, "selbstbeurteilung_ids.json");
      masterLibrary = await readSeedFromProject(dirHandle, "library_master.json");
      userLibrary = await readJsonIfExists(dirHandle, [getLibraryFileName()]);
    }
    if (!selbst) {
      selbst = await readSeedFromHttp("selbstbeurteilung_ids.json");
    }
    if (!masterLibrary) {
      masterLibrary = await readSeedFromHttp("library_master.json");
    }
    if (!selbst) throw new Error("Self assessment seed not found.");
    if (!masterLibrary) throw new Error("Master library seed not found.");
    return buildProjectFromSeeds(selbst, masterLibrary, userLibrary);
  };

  const ensureProjectMeta = () => {
    if (!state.project.meta) state.project.meta = {};
    if (!state.project.meta.locale) state.project.meta.locale = "de-CH";
    setLocale(state.project.meta.locale);
  };

  const saveSidecar = async () => {
    if (!dirHandle) return;
    saveQueue = saveQueue.then(async () => {
      let existing = sidecarDoc;
      try {
        const handle = await dirHandle.getFileHandle("project_sidecar.json");
        const file = await handle.getFile();
        const text = await file.text();
        existing = JSON.parse(text);
      } catch (err) {
        existing = sidecarDoc;
      }
      const payload = mergeSidecar(existing, state.project);
      const handle = await dirHandle.getFileHandle("project_sidecar.json", { create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(payload, null, 2));
      await writable.close();
      sidecarDoc = payload;
    }).catch((err) => {
      setStatus(`Autosave failed: ${err.message}`);
      debug.logLine("error", `Autosave failed: ${err.message || err}`);
    });
    return saveQueue;
  };

  const scheduleAutosave = () => {
    if (!dirHandle) return;
    if (autosaveTimer) clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(async () => {
      autosaveTimer = null;
      try {
        await saveSidecar();
        setStatus("Autosaved.");
      } catch (err) {
        setStatus(`Autosave failed: ${err.message}`);
      }
    }, 2000);
  };

  const flushAutosave = async () => {
    if (!dirHandle) return;
    if (autosaveTimer) {
      clearTimeout(autosaveTimer);
      autosaveTimer = null;
    }
    await saveSidecar();
  };

  const loadLibraryFile = async () => {
    if (!dirHandle) return { meta: {}, entries: [] };
    const name = getLibraryFileName();
    const data = await readJsonIfExists(dirHandle, [name]);
    if (data && Array.isArray(data.entries)) return data;
    return { meta: {}, entries: [] };
  };

  const saveLibraryFile = async (library, timestampSuffix) => {
    if (!dirHandle) return;
    const name = getLibraryFileName();
    const handle = await dirHandle.getFileHandle(name, { create: true });
    const writable = await handle.createWritable();
    await writable.write(JSON.stringify(library, null, 2));
    await writable.close();
    if (timestampSuffix) {
      const stamped = name.replace(/\\.json$/i, `_${timestampSuffix}.json`);
      const stampedHandle = await dirHandle.getFileHandle(stamped, { create: true });
      const stampedWritable = await stampedHandle.createWritable();
      await stampedWritable.write(JSON.stringify(library, null, 2));
      await stampedWritable.close();
    }
  };

  const generateLibrary = async () => {
    if (!dirHandle) return;
    ensureProjectMeta();
    let masterLibrary = await readSeedFromProject(dirHandle, "library_master.json");
    if (!masterLibrary) {
      masterLibrary = await readSeedFromHttp("library_master.json");
    }
    if (!masterLibrary) {
      masterLibrary = { entries: [] };
    }
    const userLibrary = await loadLibraryFile();
    const libraryMap = buildLibraryMap(masterLibrary, userLibrary);
    const entriesMap = new Map(Array.from(libraryMap.entries()).map(([id, entry]) => [id, entry]));
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const lastUsedAt = new Date().toISOString();
    let applied = 0;

    state.project.chapters.forEach((chapter) => {
      chapter.rows.forEach((row) => {
        if (row.kind === "section") return;
        ensureWorkstate(row);
        const ws = row.workstate;
        Object.entries(ws.libraryActions || {}).forEach(([levelKey, action]) => {
          if (action === "off") return;
          const text = getRecommendationText(row, levelKey).trim();
          if (!text) return;
          const currentHash = hashText(text);
          if (ws.libraryHashes?.[levelKey] === currentHash) return;
          const entry = entriesMap.get(row.id) || { id: row.id, levels: {} };
          entry.levels = entry.levels || {};
          if (action === "replace") {
            entry.levels[levelKey] = text;
          } else if (action === "append") {
            const existing = toText(entry.levels[levelKey]).trim();
            entry.levels[levelKey] = existing ? `${existing}\n\n${text}` : text;
          }
          entry.lastUsed = lastUsedAt;
          entriesMap.set(row.id, entry);
          ws.libraryHashes[levelKey] = currentHash;
          applied += 1;
        });
      });
    });

    const meta = {
      author: state.project.meta.author || "",
      initials: state.project.meta.initials || "",
      locale: state.project.meta.locale || "",
      updatedAt: new Date().toISOString(),
    };
    const output = { meta, entries: Array.from(entriesMap.values()) };
    await saveLibraryFile(output, timestamp);
    await saveSidecar();
    setStatus(`Library updated (${applied} changes).`);
    debug.logLine("info", `Library updated (${applied} changes).`);
  };

  const ensureWorkstate = (row) => {
    ensureWorkstateDefaults(row);
  };

  const getFindingText = (row) => {
    const ws = row.workstate;
    if (ws.useFindingOverride && ws.findingOverride) return ws.findingOverride;
    return toText(row.master?.finding);
  };

  const getRecommendationText = (row, level) => {
    const ws = row.workstate;
    const levelKey = String(level);
    if (ws.useLevelOverride?.[levelKey] && ws.levelOverrides?.[levelKey]) {
      return ws.levelOverrides[levelKey];
    }
    return toText(row.master?.levels?.[levelKey]);
  };

  const getAnswerState = (row) => {
    const direct = row.customer?.answer;
    if (direct === 0 || direct === 1) return direct;
    const items = row.customer?.items || [];
    const answers = new Set();
    items.forEach((item) => {
      if (item.answer === 0 || item.answer === 1) answers.add(item.answer);
    });
    if (answers.size === 1) return Array.from(answers)[0];
    if (answers.size > 1) return "mixed";
    return null;
  };

  const getAnswerComments = (row) => {
    const comments = [];
    const items = row.customer?.items || [];
    items.forEach((item) => {
      if (item.comment) {
        comments.push(`${item.id}: ${item.comment}`);
      }
    });
    if (row.customer?.comment && !comments.length) {
      comments.push(row.customer.comment);
    }
    return comments;
  };

  const getAnswerEvidence = (row) => {
    const evidence = [];
    const items = row.customer?.items || [];
    items.forEach((item) => {
      if (item.evidence) {
        evidence.push(`${item.id}: ${item.evidence}`);
      }
    });
    if (row.customer?.evidence && !evidence.length) {
      evidence.push(row.customer.evidence);
    }
    return evidence;
  };

  // Markdown utilities moved to shared/markdown.js

  const renderChapterList = () => {
    chapterListEl.innerHTML = "";
    const orderedChapters = [...state.project.chapters].sort((a, b) =>
      compareIdSegments(a.id, b.id),
    );
    orderedChapters.forEach((chapter) => {
      const button = document.createElement("button");
      const buttonLabel = formatChapterLabel(chapter);
      button.textContent = buttonLabel;
      button.title = buttonLabel;
      button.className = chapter.id === state.selectedChapterId ? "active" : "";
      button.addEventListener("click", () => {
        state.selectedChapterId = chapter.id;
        render();
      });
      chapterListEl.appendChild(button);
    });
  };

  const createPhotoPill = (tag, kind = "report") => {
    const items = getPhotosForTag(tag, kind);
    if (!items.length) return null;
    const pill = document.createElement("button");
    pill.type = "button";
    pill.className = "photo-pill";
    pill.textContent = `Photos ${items.length}`;
    pill.title = `${items.length} photos tagged ${tag}`;
    pill.addEventListener("click", () => {
      openPhotoOverlay(tag, kind);
    });
    return pill;
  };

  const renderChapterTitle = (chapter) => {
    chapterTitleEl.innerHTML = "";
    const title = document.createElement("span");
    title.textContent = formatChapterLabel(chapter);
    const pill = createPhotoPill(chapter.id);
    if (pill) chapterTitleEl.appendChild(pill);
    chapterTitleEl.appendChild(title);
  };

  const createSectionRow = (row) => {
    const sectionRow = document.createElement("div");
    sectionRow.className = "section-row";
    const topLevel = String(row.id || "").split(".")[0];
    const hideSectionPill = ["11", "12", "13", "14"].includes(topLevel);
    if (!hideSectionPill) {
      const pill = createPhotoPill(row.id);
      if (pill) sectionRow.appendChild(pill);
    }
    const label = document.createElement("span");
    label.textContent = row.title;
    sectionRow.appendChild(label);
    return sectionRow;
  };

  const createAnswerBadge = (row) => {
    const badge = document.createElement("span");
    badge.className = "badge";
    const answerState = getAnswerState(row);
    if (answerState === 1) {
      badge.textContent = "Answer: Yes";
    } else if (answerState === 0) {
      badge.textContent = "Answer: No";
    } else if (answerState === "mixed") {
      badge.textContent = "Answer: Mixed";
    } else {
      badge.textContent = "Answer: —";
    }
    return badge;
  };

  const createRowHeader = (row, score, options = {}) => {
    const header = document.createElement("div");
    header.className = "row-header";
    const meta = document.createElement("div");
    meta.className = "row-meta";
    const displayId = options.displayId || row.id;
    meta.textContent = `${displayId} ${row.titleOverride || ""}`.trim();
    const previewBtn = document.createElement("button");
    previewBtn.className = "preview-btn";
    previewBtn.type = "button";
    previewBtn.textContent = "Preview";
    const headerRight = document.createElement("div");
    headerRight.className = "row-header__right";
    headerRight.appendChild(createAnswerBadge(row));
    if (score !== null && score !== undefined) {
      const scoreBadge = document.createElement("span");
      scoreBadge.className = "score-badge";
      scoreBadge.textContent = `${score}%`;
      headerRight.appendChild(scoreBadge);
    }
    if (options.photoPill) {
      headerRight.appendChild(options.photoPill);
    }
    if (options.reorderControls) {
      headerRight.appendChild(options.reorderControls);
    }
    const comments = getAnswerComments(row);
    if (comments.length) {
      headerRight.appendChild(createBadgeTooltip("Comment", comments.join("\n"), "comment-badge"));
    }
    const evidence = getAnswerEvidence(row);
    if (evidence.length) {
      headerRight.appendChild(createBadgeTooltip("Evidence", evidence.join("\n"), "evidence-badge"));
    }
    headerRight.appendChild(previewBtn);
    header.appendChild(meta);
    header.appendChild(headerRight);
    return { header, previewBtn };
  };

  const createSelfAssessmentDetails = (row) => {
    if (row.type === "summary") return null;
    const selfItems = row.customer?.items || [];
    if (selfItems.length <= 1) return null;
    const details = document.createElement("details");
    details.className = "self-details";
    const summary = document.createElement("summary");
    summary.textContent = `Selbstbeurteilung (${selfItems.length})`;
    details.appendChild(summary);
    const list = document.createElement("ul");
    list.className = "self-list";
    selfItems.forEach((item) => {
      const li = document.createElement("li");
      li.textContent = `${item.id} — ${item.question || ""}`.trim();
      list.appendChild(li);
    });
    details.appendChild(list);
    return details;
  };

  const orderObservationRows = (chapter) => {
    const rows = (chapter.rows || []).filter((row) => row.kind !== "section");
    const order = chapter.meta?.order;
    if (!order || !order.length) return rows;
    const byId = new Map(rows.map((row) => [row.id, row]));
    const ordered = [];
    order.forEach((id) => {
      const match = byId.get(id);
      if (match) ordered.push(match);
    });
    rows.forEach((row) => {
      if (!ordered.includes(row)) ordered.push(row);
    });
    return ordered;
  };

  const moveObservationRow = (chapter, rowId, delta) => {
    if (!chapter.meta || !Array.isArray(chapter.meta.order)) return;
    const order = [...chapter.meta.order];
    const index = order.indexOf(rowId);
    if (index === -1) return;
    const next = index + delta;
    if (next < 0 || next >= order.length) return;
    [order[index], order[next]] = [order[next], order[index]];
    chapter.meta.order = order;
    scheduleAutosave();
    renderRows();
  };

  const createReorderControls = (chapter, row) => {
    const wrapper = document.createElement("div");
    wrapper.className = "reorder-controls";
    const up = document.createElement("button");
    up.type = "button";
    up.className = "ghost";
    up.textContent = "▲";
    up.addEventListener("click", () => moveObservationRow(chapter, row.id, -1));
    const down = document.createElement("button");
    down.type = "button";
    down.className = "ghost";
    down.textContent = "▼";
    down.addEventListener("click", () => moveObservationRow(chapter, row.id, 1));
    wrapper.appendChild(up);
    wrapper.appendChild(down);
    return wrapper;
  };

  const createToggle = (labelText, checked, onChange) => {
    const label = document.createElement("label");
    label.className = "field-toggle";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = checked;
    input.addEventListener("change", () => onChange(input.checked));
    label.appendChild(input);
    label.appendChild(document.createTextNode(labelText));
    return label;
  };

  const createMarkdownHint = () => {
    const hint = document.createElement("span");
    hint.className = "hint";
    const icon = document.createElement("span");
    icon.className = "hint__icon";
    icon.textContent = "?";
    const tooltip = document.createElement("span");
    tooltip.className = "hint__tooltip";
    tooltip.innerHTML = [
      t("markdown_hint_title", "Markdown:"),
      `<strong>${t("markdown_hint_bold", "**bold**")}</strong>,`,
      `<em>${t("markdown_hint_italic", "*italic*")}</em>,`,
      `<span class="hint__link">${t("markdown_hint_link", "[text](url)")}</span>`,
      `<br>${t("markdown_hint_list", "- List item")}`,
      `<br>${t("markdown_hint_paragraph", "Blank line = new paragraph")}`,
      `<br>${t("markdown_hint_linebreak", "Line breaks kept")}`,
    ].join(" ");
    hint.appendChild(icon);
    hint.appendChild(tooltip);
    return hint;
  };

  const createBadgeTooltip = (label, tooltipText, className) => {
    const wrapper = document.createElement("span");
    wrapper.className = "badge-tooltip";
    const badge = document.createElement("span");
    badge.className = className;
    badge.textContent = label;
    const tooltip = document.createElement("span");
    tooltip.className = "hint__tooltip";
    tooltip.textContent = tooltipText;
    wrapper.appendChild(badge);
    wrapper.appendChild(tooltip);
    return wrapper;
  };

  const createFindingField = (row, ws) => {
    const findingField = document.createElement("div");
    findingField.className = "field";
    const findingHeader = document.createElement("div");
    findingHeader.className = "field-header";
    const findingLabel = document.createElement("label");
    findingLabel.textContent = "Finding";
    findingLabel.appendChild(createMarkdownHint());

    const findingToggle = createToggle("Override", ws.useFindingOverride, (checked) => {
      ws.useFindingOverride = checked;
      if (ws.useFindingOverride && !ws.findingOverride) {
        ws.findingOverride = toText(row.master?.finding);
      }
      scheduleAutosave();
      renderRows();
    });
    const includeToggle = createToggle("Include", ws.includeFinding, (checked) => {
      ws.includeFinding = checked;
      scheduleAutosave();
      renderRows();
    });
    const doneToggle = createToggle("Done", ws.done, (checked) => {
      ws.done = checked;
      scheduleAutosave();
      renderRows();
    });

    const findingControls = document.createElement("div");
    findingControls.className = "field-controls";
    findingControls.appendChild(findingToggle);
    findingControls.appendChild(includeToggle);
    findingControls.appendChild(doneToggle);

    const findingArea = document.createElement("textarea");
    findingArea.value = getFindingText(row);
    findingArea.disabled = !ws.useFindingOverride;
    findingArea.addEventListener("input", () => {
      ws.findingOverride = findingArea.value;
      scheduleAutosave();
    });

    findingHeader.appendChild(findingLabel);
    findingHeader.appendChild(findingControls);
    findingField.appendChild(findingHeader);
    findingField.appendChild(findingArea);
    return findingField;
  };

  const createRecommendationField = (row, ws) => {
    const recField = document.createElement("div");
    recField.className = "field";
    const recHeader = document.createElement("div");
    recHeader.className = "field-header";
    const recLabel = document.createElement("label");
    recLabel.textContent = "Recommendation";
    recLabel.appendChild(createMarkdownHint());

    const levelGroup = document.createElement("div");
    levelGroup.className = "level-group";
    [1, 2, 3, 4].forEach((level) => {
      const label = document.createElement("label");
      label.className = "level-radio";
      const input = document.createElement("input");
      input.type = "radio";
      input.name = `level-${row.id}`;
      input.value = String(level);
      input.checked = ws.selectedLevel === level;
      input.addEventListener("change", () => {
        ws.selectedLevel = level;
        scheduleAutosave();
        renderRows();
      });
      label.appendChild(input);
      label.appendChild(document.createTextNode(String(level)));
      levelGroup.appendChild(label);
    });

    const levelKey = String(ws.selectedLevel);
    const overrideToggle = createToggle("Override", !!ws.useLevelOverride?.[levelKey], (checked) => {
      ws.useLevelOverride[levelKey] = checked;
      if (ws.useLevelOverride[levelKey] && !ws.levelOverrides[levelKey]) {
        ws.levelOverrides[levelKey] = toText(row.master?.levels?.[levelKey]);
      }
      scheduleAutosave();
      renderRows();
    });

    const recControls = document.createElement("div");
    recControls.className = "field-controls";
    recControls.appendChild(levelGroup);
    recControls.appendChild(overrideToggle);

    const libraryGroup = document.createElement("div");
    libraryGroup.className = "library-group";
    const libraryLabel = document.createElement("span");
    libraryLabel.textContent = "Library";
    libraryGroup.appendChild(libraryLabel);
    const actionButtons = [
      { key: "off", label: "Off" },
      { key: "append", label: "Append" },
      { key: "replace", label: "Replace" },
    ];
    actionButtons.forEach((option) => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = option.label;
      if ((ws.libraryActions?.[levelKey] || "off") === option.key) {
        button.classList.add("is-active");
      }
      button.addEventListener("click", () => {
        ws.libraryActions[levelKey] = option.key;
        scheduleAutosave();
        renderRows();
      });
      libraryGroup.appendChild(button);
    });
    recControls.appendChild(libraryGroup);

    const recArea = document.createElement("textarea");
    recArea.value = getRecommendationText(row, ws.selectedLevel);
    recArea.disabled = !ws.useLevelOverride?.[levelKey];
    recArea.addEventListener("input", () => {
      ws.levelOverrides[levelKey] = recArea.value;
      ws.libraryHashes[levelKey] = "";
      scheduleAutosave();
    });

    recHeader.appendChild(recLabel);
    recHeader.appendChild(recControls);
    recField.appendChild(recHeader);
    recField.appendChild(recArea);
    return recField;
  };

  const createPreviewPanel = (row, ws, previewBtn) => {
    const preview = document.createElement("div");
    preview.className = "row-preview";
    const previewFinding = document.createElement("div");
    previewFinding.className = "row-preview__col";
    const previewRec = document.createElement("div");
    previewRec.className = "row-preview__col";
    preview.appendChild(previewFinding);
    preview.appendChild(previewRec);

    const showPreview = () => {
      previewFinding.innerHTML = markdownToHtml(getFindingText(row));
      previewRec.innerHTML = markdownToHtml(getRecommendationText(row, ws.selectedLevel));
      preview.classList.add("is-visible");
    };
    const hidePreview = () => {
      preview.classList.remove("is-visible");
    };
    previewBtn.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      showPreview();
    });
    previewBtn.addEventListener("pointerup", hidePreview);
    previewBtn.addEventListener("pointerleave", hidePreview);
    previewBtn.addEventListener("pointercancel", hidePreview);
    return preview;
  };

  const shouldFilterRow = (row, ws) => {
    const answerState = getAnswerState(row);
    if (state.filters.mode === "hide-yes" && answerState === 1) return true;
    if (state.filters.mode === "include-only" && !ws.includeFinding) return true;
    if (state.filters.mode === "done-only" && !ws.done) return true;
    if (state.filters.mode === "hide-done" && ws.done) return true;
    return false;
  };

  const createRowCard = (row, options = {}) => {
    ensureWorkstate(row);
    const ws = row.workstate;
    if (shouldFilterRow(row, ws)) return null;

    const score = calculateScore(row);
    const card = document.createElement("div");
    card.className = "row-card";
    const isObservation = options.chapter?.id === "4.8";
    const observationPill = isObservation && row.tag ? createPhotoPill(row.tag, "observations") : null;
    const reorderControls = isObservation ? createReorderControls(options.chapter, row) : null;
    const { header, previewBtn } = createRowHeader(row, score, {
      displayId: options.displayId,
      photoPill: observationPill,
      reorderControls,
    });
    card.appendChild(header);

    const details = createSelfAssessmentDetails(row);
    if (details) card.appendChild(details);

    const rowBody = document.createElement("div");
    rowBody.className = "row-body";
    rowBody.appendChild(createFindingField(row, ws));
    rowBody.appendChild(createRecommendationField(row, ws));
    card.appendChild(rowBody);

    card.appendChild(createPreviewPanel(row, ws, previewBtn));
    return card;
  };

  const renderRows = () => {
    rowsEl.innerHTML = "";
    const chapter = state.project.chapters.find((c) => c.id === state.selectedChapterId);
    if (!chapter) return;
    renderChapterTitle(chapter);
    let rows = chapter.rows || [];
    let displayIds = new Map();
    if (chapter.id === "4.8") {
      rows = orderObservationRows(chapter);
      rows.forEach((row, index) => {
        displayIds.set(row.id, `4.8.${index + 1}`);
      });
    }

    rows.forEach((row) => {
      if (row.kind === "section") {
        rowsEl.appendChild(createSectionRow(row));
        return;
      }
      const displayId = displayIds.get(row.id);
      const card = createRowCard(row, { chapter, displayId });
      if (card) rowsEl.appendChild(card);
    });
  };

  const render = () => {
    renderChapterList();
    renderRows();
  };

  pickFolderBtn.addEventListener("click", async () => {
    if (!ensureFsAccess()) return;
    try {
      dirHandle = await window.showDirectoryPicker({ mode: "readwrite", id: "autobericht-project" });
      await persistHandle(dirHandle);
      enableActions();
      setStatus(`Selected folder: ${dirHandle.name}`);
      debug.logLine("info", `Selected folder: ${dirHandle.name}`);
      await loadProjectFromFolder();
    } catch (err) {
      setStatus(`Folder pick canceled or failed: ${err.message}`);
      debug.logLine("warn", `Folder pick canceled or failed: ${err.message}`);
    }
  });

  loadSidecarBtn.addEventListener("click", async () => {
    if (!dirHandle) return;
    await loadProjectFromFolder();
  });

  saveSidecarBtn.addEventListener("click", async () => {
    if (!dirHandle) return;
    try {
      await saveSidecar();
      setStatus("Saved project_sidecar.json");
      debug.logLine("info", "Saved project_sidecar.json");
    } catch (err) {
      setStatus(`Save failed: ${err.message}`);
      debug.logLine("error", `Save failed: ${err.message}`);
    }
  });

  loadSeedsBtn.addEventListener("click", async () => {
    if (!dirHandle) return;
    try {
      state.project = await loadSeedsForProject();
      ensureProjectMeta();
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      render();
      setStatus("Loaded seed data.");
      debug.logLine("info", "Loaded seed data.");
    } catch (err) {
      setStatus(`Seed load failed: ${err.message}`);
      debug.logLine("error", `Seed load failed: ${err.message}`);
    }
  });

  saveLogBtn.addEventListener("click", async () => {
    try {
      const result = await debug.saveLog({
        suggestedName: `mini-editor-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`,
        dirHandle,
      });
      setStatus(`Saved log (${result.location}): ${result.filename}`);
    } catch (err) {
      setStatus(`Log save failed: ${err.message}`);
    }
  });

  filterModeEls.forEach((input) => {
    input.addEventListener("change", () => {
      if (!input.checked) return;
      state.filters.mode = input.value;
      renderRows();
    });
  });

  const openSettings = () => {
    if (!settingsModal) return;
    settingsAuthorEl.value = state.project.meta?.author || "";
    settingsInitialsEl.value = state.project.meta?.initials || "";
    settingsLocaleEl.value = state.project.meta?.locale || "de-CH";
    const libraryName = getLibraryFileName();
    if (settingsLibraryHintEl) {
      settingsLibraryHintEl.textContent = `Library file: ${libraryName} (timestamped backup on generate).`;
    }
    settingsModal.classList.add("is-open");
    settingsModal.setAttribute("aria-hidden", "false");
  };

  const closeSettings = () => {
    if (!settingsModal) return;
    settingsModal.classList.remove("is-open");
    settingsModal.setAttribute("aria-hidden", "true");
  };

  const saveSettings = () => {
    ensureProjectMeta();
    state.project.meta.author = settingsAuthorEl.value.trim();
    state.project.meta.initials = settingsInitialsEl.value.trim();
    state.project.meta.locale = settingsLocaleEl.value || "de-CH";
    if (settingsLibraryHintEl) {
      settingsLibraryHintEl.textContent = `Library file: ${getLibraryFileName()} (timestamped backup on generate).`;
    }
    setStatus("Settings saved (remember to save sidecar).");
  };

  if (openSettingsBtn) {
    openSettingsBtn.addEventListener("click", openSettings);
  }
  if (closeSettingsBtn) {
    closeSettingsBtn.addEventListener("click", closeSettings);
  }
  if (settingsModal) {
    settingsModal.addEventListener("click", (event) => {
      if (event.target === settingsModal) closeSettings();
    });
  }
  if (saveSettingsBtn) {
    saveSettingsBtn.addEventListener("click", saveSettings);
  }
  if (generateLibraryBtn) {
    generateLibraryBtn.addEventListener("click", async () => {
      try {
        await generateLibrary();
      } catch (err) {
        setStatus(`Library update failed: ${err.message}`);
        debug.logLine("error", `Library update failed: ${err.message || err}`);
      }
    });
  }

  if (photoOverlayClose) {
    photoOverlayClose.addEventListener("click", () => {
      closePhotoOverlay();
    });
  }
  if (photoOverlayCloseBtn) {
    photoOverlayCloseBtn.addEventListener("click", () => {
      closePhotoOverlay();
    });
  }
  if (photoOverlayPrevBtn) {
    photoOverlayPrevBtn.addEventListener("click", () => {
      stepPhotoOverlay(-1);
    });
  }
  if (photoOverlayNextBtn) {
    photoOverlayNextBtn.addEventListener("click", () => {
      stepPhotoOverlay(1);
    });
  }
  document.addEventListener("keydown", (event) => {
    if (!photoOverlayEl || !photoOverlayEl.classList.contains("is-open")) return;
    if (event.target && ["INPUT", "TEXTAREA"].includes(event.target.tagName)) return;
    if (event.key === "Escape") closePhotoOverlay();
    if (event.key === "a" || event.key === "A") stepPhotoOverlay(-1);
    if (event.key === "d" || event.key === "D") stepPhotoOverlay(1);
  });

  window.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      flushAutosave();
    }
  });
  window.addEventListener("pagehide", () => {
    flushAutosave();
  });

  if (importSelfBtn) {
    importSelfBtn.addEventListener("click", async () => {
      if (!dirHandle) return;
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
        const sheetName = workbook.SheetNames.find((name) => {
          const lower = name.toLowerCase();
          return lower.includes("selbstbeurteilung")
            || lower.includes("autoévaluation")
            || lower.includes("autoevaluation")
            || lower.includes("autovalutazione");
        }) || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
        const headerRowIndex = rows.findIndex((row) => String(row[0] || "").toLowerCase().includes("nr"));
        const headerRow = headerRowIndex >= 0 ? rows[headerRowIndex] : [];
        const findCol = (label) => headerRow.findIndex((cell) => String(cell || "").toLowerCase().includes(label));
        const colId = headerRowIndex >= 0 ? findCol("nr") : 0;
        const colYes = headerRowIndex >= 0 ? findCol("ja") : 3;
        const colNo = headerRowIndex >= 0 ? findCol("nein") : 4;
        const colComment = headerRowIndex >= 0 ? findCol("bemerk") : 5;
        const pickCol = (...labels) => {
          for (const label of labels) {
            const idx = findCol(label);
            if (idx >= 0) return idx;
          }
          return -1;
        };
        const colEvidence = headerRowIndex >= 0
          ? pickCol("nachweis", "beleg", "evidence", "document")
          : -1;
        const dataRows = headerRowIndex >= 0 ? rows.slice(headerRowIndex + 1) : rows;
        const answerMap = new Map();

        dataRows.forEach((row) => {
          const rawId = String(row[colId] || "").trim();
          if (!rawId) return;
          const normalized = rawId.replace(/[.\s]+$/g, "").toLowerCase();
          if (!/^[0-9]/.test(normalized)) return;
          const yesVal = String(row[colYes] || "").trim().toLowerCase();
          const noVal = String(row[colNo] || "").trim().toLowerCase();
          const comment = String(row[colComment] || "").trim();
          const evidence = colEvidence >= 0 ? String(row[colEvidence] || "").trim() : "";
          let answer = null;
          if (yesVal === "x" || yesVal === "ja") answer = 1;
          if (noVal === "x" || noVal === "nein") answer = 0;
          answerMap.set(normalized, { answer, comment, evidence });
        });

        const idMap = new Map();
        state.project.chapters.forEach((chapter) => {
          chapter.rows.forEach((row) => {
            if (row.kind === "section") return;
            row.customer = row.customer || { items: [] };
            row.customer.items = row.customer.items || [];
            row.customer.items.forEach((item) => {
              const key = String(item.id || "").trim().toLowerCase();
              if (key) idMap.set(key, item);
            });
          });
        });

        let applied = 0;
        answerMap.forEach((payload, key) => {
          const item = idMap.get(key);
          if (!item) return;
          if (payload.answer === 0 || payload.answer === 1) {
            item.answer = payload.answer;
          }
          if (payload.comment) {
            item.comment = payload.comment;
          }
          if (payload.evidence) {
            item.evidence = payload.evidence;
          }
          applied += 1;
        });

        setStatus(`Imported self-assessment answers (${applied}).`);
        debug.logLine("info", `Imported self-assessment answers (${applied}).`);
        renderRows();
        await saveSidecar();
      } catch (err) {
        setStatus(`Import failed: ${err.message}`);
        debug.logLine("error", `Import failed: ${err.message || err}`);
      }
    });
  }

  const init = async () => {
    ensureFsAccess();
    await restoreLastHandle();
    render();
    if (!dirHandle) {
      await autoLoadSeeds();
    }
  };

  init();
})();
