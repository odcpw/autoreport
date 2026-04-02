(() => {
  const debug = window.AutoReportDebug || { logLine: () => {} };
  const i18n = window.AutoBerichtI18n || {};
  const fsHandles = window.AutoBerichtFsHandle || {};
  const stateHelpers = window.AutoBerichtState || {};
  const normalizeHelpers = window.AutoBerichtNormalize || {};
  const seeds = window.AutoBerichtSeeds || {};

  const elements = {
    statusEl: document.getElementById("status"),
    openProjectBtn: document.getElementById("lm-open-project"),
    loadSourceBtn: document.getElementById("lm-load-source"),
    undoBtn: document.getElementById("lm-undo"),
    redoBtn: document.getElementById("lm-redo"),
    saveBtn: document.getElementById("lm-save"),
    chapterListEl: document.getElementById("lm-chapter-list"),
    titleEl: document.getElementById("lm-title"),
    subtitleEl: document.getElementById("lm-subtitle"),
    saveStateEl: document.getElementById("lm-save-state"),
    targetNameEl: document.getElementById("lm-target-name"),
    targetMetaEl: document.getElementById("lm-target-meta"),
    sourceNameEl: document.getElementById("lm-source-name"),
    sourceMetaEl: document.getElementById("lm-source-meta"),
    emptyEl: document.getElementById("lm-empty"),
    rowsEl: document.getElementById("lm-rows"),
    firstRunModal: document.getElementById("first-run-modal"),
    firstRunPickBtn: document.getElementById("first-run-pick"),
    libraryChoiceModal: document.getElementById("library-choice-modal"),
    libraryListEl: document.getElementById("lm-library-list"),
  };

  const state = {
    projectHandle: null,
    projectMeta: null,
    targetLibrary: null,
    targetFileName: "",
    targetBackupName: "",
    targetRawText: "",
    sourceLibrary: null,
    sourceFileName: "",
    selectedChapterId: "",
    historyPast: [],
    historyFuture: [],
    touchedKeys: new Set(),
    dirty: false,
  };

  const runtime = {
    autosaveTimer: null,
    saveQueue: Promise.resolve(),
  };

  const t = i18n.t || ((key, fallback) => fallback || key);
  const setLocale = i18n.setLocale || (() => {});
  const resolveSpellcheckLang = i18n.resolveSpellcheckLang || ((locale) => String(locale || "en").toLowerCase().split("-")[0] || "en");
  const compareIds = stateHelpers.compareIdSegments
    || ((a, b) => String(a || "").localeCompare(String(b || ""), "de", { numeric: true }));
  const getLibraryFileName = stateHelpers.getLibraryFileName
    || ((meta = {}) => `library_user_${String(meta.locale || "de-CH")}.json`);
  const toText = stateHelpers.toText || ((value) => {
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  });

  const isPlainObject = (value) => !!value && typeof value === "object" && !Array.isArray(value);

  const getLocaleBase = (locale) => {
    const base = String(locale || "de-CH").toLowerCase().split("-")[0];
    return ["de", "fr", "it"].includes(base) ? base : "de";
  };

  const setFirstRunVisible = (visible) => {
    if (!elements.firstRunModal) return;
    if (visible) {
      elements.firstRunModal.classList.add("is-open");
      elements.firstRunModal.setAttribute("aria-hidden", "false");
    } else {
      elements.firstRunModal.classList.remove("is-open");
      elements.firstRunModal.setAttribute("aria-hidden", "true");
    }
  };

  const setLibraryChoiceVisible = (visible) => {
    if (!elements.libraryChoiceModal) return;
    if (visible) {
      elements.libraryChoiceModal.classList.add("is-open");
      elements.libraryChoiceModal.setAttribute("aria-hidden", "false");
    } else {
      elements.libraryChoiceModal.classList.remove("is-open");
      elements.libraryChoiceModal.setAttribute("aria-hidden", "true");
    }
  };

  const setStatus = (message) => {
    if (elements.statusEl) elements.statusEl.textContent = message || "";
    if (message) debug.logLine("info", message);
  };

  const setSaveState = (message) => {
    if (elements.saveStateEl) elements.saveStateEl.textContent = message || "";
  };

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus("File System Access API is not available. Open via http://localhost in Edge/Chrome to enable file access.");
      if (elements.openProjectBtn) elements.openProjectBtn.disabled = true;
      if (elements.firstRunPickBtn) elements.firstRunPickBtn.disabled = true;
      return false;
    }
    return true;
  };

  const normalizeTagOption = (option) => {
    if (!option) return null;
    if (typeof option === "string") {
      const value = String(option || "").trim();
      if (!value) return null;
      return { value, label: value };
    }
    const value = String(option.value || option.label || "").trim();
    if (!value) return null;
    const label = String(option.label || value).trim() || value;
    return { value, label };
  };

  const normalizeObservationEntry = (entry) => {
    const option = normalizeTagOption(entry);
    if (!option) return null;
    return {
      ...(isPlainObject(entry) ? structuredClone(entry) : {}),
      value: option.value,
      label: option.label,
      finding: toText(entry?.finding).replace(/\r\n/g, "\n"),
      recommendation: toText(entry?.recommendation).replace(/\r\n/g, "\n"),
    };
  };

  const normalizeObservationIdentity = (entry) => {
    const option = normalizeTagOption(entry);
    if (!option) return "";
    return String(option.label || option.value || "").trim();
  };

  const sortObservationTagOptions = (options, locale = "de") => (
    (options || [])
      .map(normalizeTagOption)
      .filter(Boolean)
      .sort((a, b) => String(a.label || a.value || "")
        .localeCompare(String(b.label || b.value || ""), locale, { numeric: true }))
  );

  const migrateLegacyObservations = (library, localeBase) => {
    if (!library) return;
    const legacyEntries = (library.entries || [])
      .filter((entry) => /^4\.8(?:\.|$)/.test(String(entry?.id || "").trim()))
      .sort((a, b) => compareIds(String(a?.id || ""), String(b?.id || "")));
    const options = (library?.tags?.observations || [])
      .map(normalizeTagOption)
      .filter(Boolean);
    const existing = (library.observations || []).map(normalizeObservationEntry).filter(Boolean);
    if (!legacyEntries.length) {
      library.entries = (library.entries || []).filter((entry) => !/^4\.8(?:\.|$)/.test(String(entry?.id || "").trim()));
      return;
    }
    if (!options.length) {
      throw new Error("Legacy 4.8 library entries exist, but observation tags are missing. Regenerate the library in AutoBericht first.");
    }
    if (legacyEntries.length > options.length) {
      throw new Error("Legacy 4.8 library entries cannot be migrated because there are more 4.8 texts than observation tags.");
    }
    const byIdentity = new Map(existing.map((entry) => [normalizeObservationIdentity(entry), entry]));
    options.forEach((option, index) => {
      const identity = normalizeObservationIdentity(option);
      const current = byIdentity.get(identity);
      const legacy = legacyEntries[index] || null;
      if (current) {
        if (legacy && !toText(current.finding).trim()) current.finding = toText(legacy.finding).replace(/\r\n/g, "\n");
        if (legacy && !toText(current.recommendation).trim()) current.recommendation = toText(legacy.recommendation).replace(/\r\n/g, "\n");
        return;
      }
      byIdentity.set(identity, {
        value: option.value,
        label: option.label,
        finding: toText(legacy?.finding).replace(/\r\n/g, "\n"),
        recommendation: toText(legacy?.recommendation).replace(/\r\n/g, "\n"),
      });
    });
    library.observations = Array.from(byIdentity.values())
      .map(normalizeObservationEntry)
      .filter(Boolean)
      .sort((a, b) => String(a.label || a.value || "")
        .localeCompare(String(b.label || b.value || ""), localeBase, { numeric: true }));
    library.entries = (library.entries || []).filter((entry) => !/^4\.8(?:\.|$)/.test(String(entry?.id || "").trim()));
  };

  const normalizeKnowledgeBaseForMaker = (knowledgeBase) => {
    seeds.validateKnowledgeBase(knowledgeBase);
    const clone = structuredClone(knowledgeBase);
    clone.meta = isPlainObject(clone.meta) ? clone.meta : {};
    clone.structure = isPlainObject(clone.structure) ? clone.structure : { items: [] };
    clone.structure.items = Array.isArray(clone.structure.items) ? clone.structure.items : [];
    clone.library = isPlainObject(clone.library) ? clone.library : { entries: [], observations: [], chapterPositives: {} };
    clone.library.entries = (clone.library.entries || [])
      .filter((entry) => isPlainObject(entry) && String(entry.id || "").trim())
      .map((entry) => ({
        ...structuredClone(entry),
        id: String(entry.id || "").trim(),
        finding: toText(entry.finding).replace(/\r\n/g, "\n"),
        recommendation: toText(entry.recommendation).replace(/\r\n/g, "\n"),
      }))
      .sort((a, b) => compareIds(a.id, b.id));
    clone.library.observations = (clone.library.observations || [])
      .map(normalizeObservationEntry)
      .filter(Boolean)
      .sort((a, b) => String(a.label || a.value || "")
        .localeCompare(String(b.label || b.value || ""), getLocaleBase(clone.meta.locale), { numeric: true }));
    clone.library.chapterPositives = isPlainObject(clone.library.chapterPositives)
      ? structuredClone(clone.library.chapterPositives)
      : {};
    clone.tags = typeof seeds.normalizeTagGroups === "function"
      ? seeds.normalizeTagGroups(clone.tags)
      : { report: [], observations: [], training: [] };
    migrateLegacyObservations(clone.library, getLocaleBase(clone.meta.locale));

    const observationMap = new Map();
    sortObservationTagOptions(clone.tags.observations || [], getLocaleBase(clone.meta.locale)).forEach((option) => {
      if (!observationMap.has(option.value)) observationMap.set(option.value, option);
    });
    clone.library.observations.forEach((entry) => {
      if (!observationMap.has(entry.value)) {
        observationMap.set(entry.value, { value: entry.value, label: entry.label || entry.value });
      }
    });
    clone.tags.observations = Array.from(observationMap.values()).sort((a, b) => String(a.label || a.value || "")
      .localeCompare(String(b.label || b.value || ""), getLocaleBase(clone.meta.locale), { numeric: true }));
    return clone;
  };

  const readJsonFileHandle = async (handle) => {
    const file = await handle.getFile();
    return JSON.parse(await file.text());
  };

  const readTextFileHandle = async (handle) => {
    const file = await handle.getFile();
    return file.text();
  };

  const getProjectFileHandle = async (dirHandle, name) => {
    try {
      return await dirHandle.getFileHandle(name);
    } catch (err) {
      if (err?.name === "NotFoundError") return null;
      throw err;
    }
  };

  const isCurrentLibraryFileName = (name) => (
    typeof name === "string"
    && name.startsWith("library_user_")
    && name.endsWith(".json")
    && !/_\d{4}-\d{2}-\d{2}/.test(name)
  );

  const listCurrentLibraryFiles = async (dirHandle) => {
    const files = [];
    if (!dirHandle?.entries) return files;
    for await (const [name, handle] of dirHandle.entries()) {
      if (handle.kind !== "file") continue;
      if (!isCurrentLibraryFileName(name)) continue;
      files.push({ name, handle });
    }
    files.sort((a, b) => a.name.localeCompare(b.name, "de", { numeric: true }));
    return files;
  };

  const pickTargetLibraryFile = async (options) => {
    if (!options?.length) return null;
    if (options.length === 1) return options[0];
    if (!elements.libraryListEl) return null;
    return new Promise((resolve) => {
      elements.libraryListEl.innerHTML = "";
      options.forEach((option) => {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = option.name;
        button.addEventListener("click", () => {
          setLibraryChoiceVisible(false);
          resolve(option);
        });
        elements.libraryListEl.appendChild(button);
      });
      setLibraryChoiceVisible(true);
    });
  };

  const extractReportProject = (doc) => {
    if (!isPlainObject(doc)) return null;
    if (isPlainObject(doc.report?.project) && Array.isArray(doc.report.project.chapters)) return doc.report.project;
    if (Array.isArray(doc.chapters)) return doc;
    return null;
  };

  const getLibraryCode = (fileName, library) => {
    const direct = String(library?.meta?.moderatorInitials || library?.meta?.initials || "").trim();
    if (direct) return direct;
    const match = /^library_user_([^_]+)_/i.exec(String(fileName || ""));
    return match ? match[1] : "";
  };

  const describeLibraryMeta = (fileName, library) => {
    if (!library) return "";
    const parts = [];
    const code = getLibraryCode(fileName, library);
    if (code) parts.push(code);
    if (library.meta?.locale) parts.push(library.meta.locale);
    if (fileName) parts.push(fileName);
    return parts.join(" • ");
  };

  const updateActionState = () => {
    const hasTarget = !!state.targetLibrary;
    const hasProject = !!state.projectHandle;
    if (elements.loadSourceBtn) elements.loadSourceBtn.disabled = !hasTarget || !hasProject;
    if (elements.saveBtn) elements.saveBtn.disabled = !hasTarget;
    if (elements.undoBtn) elements.undoBtn.disabled = state.historyPast.length === 0;
    if (elements.redoBtn) elements.redoBtn.disabled = state.historyFuture.length === 0;
  };

  const pushHistory = () => {
    if (!state.targetLibrary) return;
    state.historyPast.push(structuredClone(state.targetLibrary));
    if (state.historyPast.length > 20) state.historyPast.shift();
    state.historyFuture = [];
    updateActionState();
  };

  const markDirty = () => {
    state.dirty = true;
    setSaveState("Unsaved changes");
  };

  const scheduleAutosave = () => {
    if (!state.projectHandle || !state.targetLibrary || !state.targetFileName) return;
    if (runtime.autosaveTimer) clearTimeout(runtime.autosaveTimer);
    runtime.autosaveTimer = setTimeout(async () => {
      runtime.autosaveTimer = null;
      try {
        await saveTargetLibrary();
      } catch (err) {
        setStatus(`Autosave failed: ${err.message || err}`);
      }
    }, 1200);
  };

  const flushAutosave = async () => {
    if (runtime.autosaveTimer) {
      clearTimeout(runtime.autosaveTimer);
      runtime.autosaveTimer = null;
    }
    if (state.targetLibrary) await saveTargetLibrary();
  };

  const writeProjectJson = async (fileName, payload) => {
    const handle = await state.projectHandle.getFileHandle(fileName, { create: true });
    const writable = await handle.createWritable();
    await writable.write(JSON.stringify(payload, null, 2));
    await writable.close();
  };

  const backupTargetLibraryOnLoad = async () => {
    if (!state.projectHandle || !state.targetFileName || !state.targetRawText) return;
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const archiveName = state.targetFileName.replace(/\.json$/i, "");
    state.targetBackupName = `${archiveName}_librarymaker_${stamp}.json`;
    const handle = await state.projectHandle.getFileHandle(state.targetBackupName, { create: true });
    const writable = await handle.createWritable();
    await writable.write(state.targetRawText);
    await writable.close();
    debug.logLine("info", `LibraryMaker backup saved: ${state.targetBackupName}`);
  };

  const saveTargetLibrary = async () => {
    if (!state.projectHandle || !state.targetLibrary || !state.targetFileName) return;
    const output = structuredClone(state.targetLibrary);
    output.schemaVersion = "1.1";
    output.meta = {
      ...(output.meta || {}),
      locale: state.projectMeta?.locale || output.meta?.locale || "de-CH",
      moderator: state.projectMeta?.moderator || output.meta?.moderator || output.meta?.author || "",
      moderatorInitials: state.projectMeta?.moderatorInitials || output.meta?.moderatorInitials || output.meta?.initials || "",
      coModerator: state.projectMeta?.coModerator || output.meta?.coModerator || "",
      coModeratorInitials: state.projectMeta?.coModeratorInitials || output.meta?.coModeratorInitials || "",
      updatedAt: new Date().toISOString(),
    };
    output.meta.author = output.meta.moderator;
    output.meta.initials = output.meta.moderatorInitials;
    runtime.saveQueue = runtime.saveQueue.then(async () => {
      await writeProjectJson(state.targetFileName, output);
      state.targetLibrary = normalizeKnowledgeBaseForMaker(output);
      state.dirty = false;
      setSaveState(`Saved ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`);
      debug.logLine("info", `LibraryMaker saved: ${state.targetFileName}`);
    });
    return runtime.saveQueue;
  };

  const resolveTargetLibraryFile = async (projectHandle, projectMeta) => {
    const expectedName = getLibraryFileName(projectMeta || {});
    const expectedHandle = await getProjectFileHandle(projectHandle, expectedName);
    if (expectedHandle) {
      const rawText = await readTextFileHandle(expectedHandle);
      return { fileName: expectedName, library: JSON.parse(rawText), rawText };
    }
    const candidates = await listCurrentLibraryFiles(projectHandle);
    if (candidates.length === 1) {
      const rawText = await readTextFileHandle(candidates[0].handle);
      return { fileName: candidates[0].name, library: JSON.parse(rawText), rawText };
    }
    if (candidates.length > 1) {
      const picked = await pickTargetLibraryFile(candidates);
      if (!picked) {
        throw new Error(`Could not resolve current library file. Expected ${expectedName}, but found multiple library files in the project folder.`);
      }
      const rawText = await readTextFileHandle(picked.handle);
      return { fileName: picked.name, library: JSON.parse(rawText), rawText };
    }
    throw new Error(`Library file missing: ${expectedName}. Generate the library first in AutoBericht.`);
  };

  const loadTargetLibraryFromProject = async () => {
    if (!state.projectHandle) return;
    const sidecarHandle = await getProjectFileHandle(state.projectHandle, "project_sidecar.json");
    if (!sidecarHandle) {
      throw new Error("project_sidecar.json is missing. Open a valid AutoBericht project first.");
    }
    const sidecarDoc = await readJsonFileHandle(sidecarHandle);
    const reportProject = extractReportProject(sidecarDoc);
    if (!reportProject) {
      throw new Error("project_sidecar.json does not contain a report project.");
    }
    const normalizedProject = typeof normalizeHelpers.normalizeProject === "function"
      ? normalizeHelpers.normalizeProject(structuredClone(reportProject), setLocale)
      : structuredClone(reportProject);
    state.projectMeta = structuredClone(normalizedProject.meta || {});
    setLocale(state.projectMeta.locale || "de-CH");

    const resolved = await resolveTargetLibraryFile(state.projectHandle, state.projectMeta);
    const targetLibrary = normalizeKnowledgeBaseForMaker(resolved.library);
    const localeBase = getLocaleBase(targetLibrary.meta?.locale || state.projectMeta.locale);
    if (getLocaleBase(state.projectMeta.locale) !== localeBase) {
      throw new Error(`Current library locale (${targetLibrary.meta?.locale || "unknown"}) does not match project locale (${state.projectMeta.locale || "unknown"}).`);
    }

    state.targetLibrary = targetLibrary;
    state.targetFileName = resolved.fileName;
    state.targetRawText = resolved.rawText;
    state.sourceLibrary = null;
    state.sourceFileName = "";
    state.historyPast = [];
    state.historyFuture = [];
    state.touchedKeys = new Set();
    state.dirty = false;
    await backupTargetLibraryOnLoad();
    const chapters = getChapterObjects();
    state.selectedChapterId = chapters[0]?.id || "";
    setSaveState("Loaded");
    setStatus(`Loaded My Library: ${resolved.fileName}`);
    updateActionState();
    render();
  };

  const normalizeLibraryEntryMap = (library) => {
    const map = new Map();
    (library?.library?.entries || []).forEach((entry) => {
      const id = String(entry?.id || "").trim();
      if (!id) return;
      map.set(id, {
        ...structuredClone(entry),
        id,
        finding: toText(entry.finding).replace(/\r\n/g, "\n"),
        recommendation: toText(entry.recommendation).replace(/\r\n/g, "\n"),
      });
    });
    return map;
  };

  const normalizeObservationEntryMap = (library) => {
    const map = new Map();
    (library?.library?.observations || []).map(normalizeObservationEntry).filter(Boolean).forEach((entry) => {
      const identity = normalizeObservationIdentity(entry);
      if (!identity || map.has(identity)) return;
      map.set(identity, entry);
    });
    return map;
  };

  const buildObservationOptions = (library) => {
    const map = new Map();
    sortObservationTagOptions(library?.tags?.observations || [], getLocaleBase(library?.meta?.locale)).forEach((option) => {
      const identity = normalizeObservationIdentity(option);
      if (!identity || map.has(identity)) return;
      map.set(identity, option);
    });
    (library?.library?.observations || []).map(normalizeObservationEntry).filter(Boolean).forEach((entry) => {
      const identity = normalizeObservationIdentity(entry);
      if (!identity || map.has(identity)) return;
      map.set(identity, { value: entry.value, label: entry.label || entry.value });
    });
    return Array.from(map.values()).sort((a, b) => String(a.label || a.value || "")
      .localeCompare(String(b.label || b.value || ""), getLocaleBase(library?.meta?.locale), { numeric: true }));
  };

  const buildStructureGroups = (library) => {
    const groups = new Map();
    (library?.structure?.items || []).forEach((item) => {
      const entryId = String(item?.collapsedId || item?.groupId || item?.id || "").trim();
      if (!entryId || entryId.startsWith("4.8")) return;
      const chapterId = entryId.split(".")[0] || String(item?.chapter || "").trim();
      if (!chapterId) return;
      const existing = groups.get(entryId) || {
        entryId,
        chapterId,
        chapterLabel: String(item?.chapterLabel || "").trim(),
        sectionLabel: String(item?.sectionLabel || "").trim(),
        title: String(item?.question || "").trim(),
      };
      if (!existing.chapterLabel && item?.chapterLabel) existing.chapterLabel = String(item.chapterLabel).trim();
      if (!existing.sectionLabel && item?.sectionLabel) existing.sectionLabel = String(item.sectionLabel).trim();
      if (!existing.title && item?.question) existing.title = String(item.question).trim();
      groups.set(entryId, existing);
    });
    return groups;
  };

  const getChapterObjects = () => {
    const groups = buildStructureGroups(state.targetLibrary);
    const chapterMap = new Map();
    groups.forEach((group) => {
      if (!chapterMap.has(group.chapterId)) {
        const localeBase = getLocaleBase(state.targetLibrary?.meta?.locale || state.projectMeta?.locale);
        chapterMap.set(group.chapterId, {
          id: group.chapterId,
          title: group.chapterLabel ? { [localeBase]: group.chapterLabel } : undefined,
        });
      }
    });
    const hasObservations = buildObservationOptions(state.targetLibrary).length || buildObservationOptions(state.sourceLibrary).length;
    if (hasObservations && !chapterMap.has("4.8")) {
      chapterMap.set("4.8", { id: "4.8" });
    }
    return Array.from(chapterMap.values()).sort((a, b) => compareIds(a.id, b.id));
  };

  const buildStructuredRows = (chapterId) => {
    const groups = buildStructureGroups(state.targetLibrary);
    const targetEntries = normalizeLibraryEntryMap(state.targetLibrary);
    const sourceEntries = normalizeLibraryEntryMap(state.sourceLibrary);
    return Array.from(groups.values())
      .filter((group) => String(group.chapterId) === String(chapterId))
      .sort((a, b) => compareIds(a.entryId, b.entryId))
      .map((group) => ({
        kind: "entry",
        key: `entry:${group.entryId}`,
        entryId: group.entryId,
        displayId: group.entryId,
        title: group.title || group.entryId,
        context: group.sectionLabel || group.chapterLabel || "",
        targetEntry: targetEntries.get(group.entryId) || null,
        sourceEntry: sourceEntries.get(group.entryId) || null,
      }));
  };

  const buildObservationRows = () => {
    const union = new Map();
    const addOption = (option) => {
      const normalized = normalizeTagOption(option);
      if (!normalized) return;
      const identity = normalizeObservationIdentity(normalized);
      if (!identity) return;
      const existing = union.get(identity);
      if (!existing) {
        union.set(identity, { identity, value: normalized.value, label: normalized.label });
        return;
      }
      if (!existing.label && normalized.label) existing.label = normalized.label;
      if (!existing.value && normalized.value) existing.value = normalized.value;
    };
    buildObservationOptions(state.targetLibrary).forEach(addOption);
    buildObservationOptions(state.sourceLibrary).forEach(addOption);
    const targetEntries = normalizeObservationEntryMap(state.targetLibrary);
    const sourceEntries = normalizeObservationEntryMap(state.sourceLibrary);
    targetEntries.forEach((entry) => addOption({ value: entry.value, label: entry.label }));
    sourceEntries.forEach((entry) => addOption({ value: entry.value, label: entry.label }));

    const targetIdentities = new Set(buildObservationOptions(state.targetLibrary).map((item) => normalizeObservationIdentity(item)).filter(Boolean));
    const sourceIdentities = new Set(buildObservationOptions(state.sourceLibrary).map((item) => normalizeObservationIdentity(item)).filter(Boolean));
    return Array.from(union.values())
      .sort((a, b) => String(a.label || a.value || "")
        .localeCompare(String(b.label || b.value || ""), getLocaleBase(state.targetLibrary?.meta?.locale), { numeric: true }))
      .map((option) => ({
        kind: "observation",
        key: `obs:${option.identity}`,
        identity: option.identity,
        value: option.value,
        label: option.label || option.value,
        displayId: "4.8",
        title: option.label || option.value,
        context: "Observation tag",
        targetEntry: targetEntries.get(option.identity) || null,
        sourceEntry: sourceEntries.get(option.identity) || null,
        targetHasTag: targetIdentities.has(option.identity),
        sourceHasTag: sourceIdentities.has(option.identity),
      }));
  };

  const buildRowsForSelection = () => {
    if (!state.targetLibrary || !state.selectedChapterId) return [];
    if (state.selectedChapterId === "4.8") return buildObservationRows();
    return buildStructuredRows(state.selectedChapterId);
  };

  const getTargetFieldValue = (row, field) => toText(row?.targetEntry?.[field] || "");
  const getSourceFieldValue = (row, field) => toText(row?.sourceEntry?.[field] || "");

  const ensureTargetStructuredEntry = (entryId) => {
    let entry = (state.targetLibrary.library.entries || []).find((item) => String(item?.id || "") === String(entryId || ""));
    if (entry) return entry;
    entry = { id: String(entryId || "").trim(), finding: "", recommendation: "" };
    state.targetLibrary.library.entries.push(entry);
    return entry;
  };

  const ensureTargetObservationTag = (value, label) => {
    state.targetLibrary.tags = typeof seeds.normalizeTagGroups === "function"
      ? seeds.normalizeTagGroups(state.targetLibrary.tags)
      : state.targetLibrary.tags || { report: [], observations: [], training: [] };
    const identity = normalizeObservationIdentity({ value, label });
    const existing = (state.targetLibrary.tags.observations || []).map(normalizeTagOption)
      .find((option) => normalizeObservationIdentity(option) === identity);
    if (existing) {
      existing.label = label || existing.label;
      state.targetLibrary.tags.observations = sortObservationTagOptions(
        (state.targetLibrary.tags.observations || []).map((option) => {
          const normalized = normalizeTagOption(option);
          if (!normalized) return null;
          if (normalizeObservationIdentity(normalized) === identity) {
            normalized.value = normalized.value || value;
            normalized.label = label || normalized.label;
          }
          return normalized;
        }).filter(Boolean),
        getLocaleBase(state.targetLibrary.meta?.locale),
      );
      return;
    }
    state.targetLibrary.tags.observations = sortObservationTagOptions(
      [...(state.targetLibrary.tags.observations || []), { value, label: label || value }],
      getLocaleBase(state.targetLibrary.meta?.locale),
    );
  };

  const ensureTargetObservationEntry = (value, label) => {
    state.targetLibrary.library.observations = Array.isArray(state.targetLibrary.library.observations)
      ? state.targetLibrary.library.observations
      : [];
    const identity = normalizeObservationIdentity({ value, label });
    let entry = state.targetLibrary.library.observations
      .map(normalizeObservationEntry)
      .find((item) => normalizeObservationIdentity(item) === identity);
    if (entry) {
      entry.label = label || entry.label;
      state.targetLibrary.library.observations = state.targetLibrary.library.observations.map((item) => {
        const normalized = normalizeObservationEntry(item);
        if (!normalized) return item;
        if (normalizeObservationIdentity(normalized) === identity) {
          return {
            ...structuredClone(item),
            value: normalized.value,
            label: label || normalized.label,
            finding: normalized.finding,
            recommendation: normalized.recommendation,
          };
        }
        return normalizeObservationEntry(item);
      }).filter(Boolean);
      return state.targetLibrary.library.observations
        .map(normalizeObservationEntry)
        .find((item) => normalizeObservationIdentity(item) === identity);
    }
    entry = { value, label: label || value, finding: "", recommendation: "" };
    state.targetLibrary.library.observations.push(entry);
    return entry;
  };

  const normalizeTargetLibraryAfterMutation = () => {
    state.targetLibrary.library.entries = (state.targetLibrary.library.entries || [])
      .filter((entry) => isPlainObject(entry) && String(entry.id || "").trim())
      .map((entry) => ({
        ...structuredClone(entry),
        id: String(entry.id || "").trim(),
        finding: toText(entry.finding).replace(/\r\n/g, "\n"),
        recommendation: toText(entry.recommendation).replace(/\r\n/g, "\n"),
      }))
      .sort((a, b) => compareIds(a.id, b.id));
    state.targetLibrary.library.observations = (state.targetLibrary.library.observations || [])
      .map(normalizeObservationEntry)
      .filter(Boolean)
      .sort((a, b) => String(a.label || a.value || "")
        .localeCompare(String(b.label || b.value || ""), getLocaleBase(state.targetLibrary.meta?.locale), { numeric: true }));
    state.targetLibrary.tags.observations = sortObservationTagOptions(
      state.targetLibrary.tags?.observations || [],
      getLocaleBase(state.targetLibrary.meta?.locale),
    );
  };

  const setTargetFieldValue = (row, field, value) => {
    const normalizedValue = toText(value).replace(/\r\n/g, "\n");
    if (row.kind === "observation") {
      ensureTargetObservationTag(row.value, row.label);
      const entry = ensureTargetObservationEntry(row.value, row.label);
      entry[field] = normalizedValue;
    } else {
      const entry = ensureTargetStructuredEntry(row.entryId);
      entry[field] = normalizedValue;
    }
    normalizeTargetLibraryAfterMutation();
  };

  const appendText = (existing, addition) => {
    const left = toText(existing).trim();
    const right = toText(addition).trim();
    if (!right) return left;
    if (!left) return right;
    if (left === right) return left;
    return `${left}\n\n${right}`;
  };

  const markTouched = (key) => {
    if (!key) return;
    state.touchedKeys.add(String(key));
  };

  const applyButtonAction = (row, field, mode) => {
    const sourceText = getSourceFieldValue(row, field).trim();
    if (!sourceText) return;
    pushHistory();
    const currentText = getTargetFieldValue(row, field);
    const nextText = mode === "replace" ? sourceText : appendText(currentText, sourceText);
    setTargetFieldValue(row, field, nextText);
    markTouched(row.key);
    markDirty();
    scheduleAutosave();
    render();
  };

  const adoptObservationRow = (row) => {
    if (row.kind !== "observation") return;
    pushHistory();
    ensureTargetObservationTag(row.value, row.label);
    ensureTargetObservationEntry(row.value, row.label);
    normalizeTargetLibraryAfterMutation();
    markTouched(row.key);
    markDirty();
    scheduleAutosave();
    render();
  };

  const getChapterLabel = (chapter) => stateHelpers.formatChapterLabel
    ? stateHelpers.formatChapterLabel(chapter, state.targetLibrary?.meta?.locale || state.projectMeta?.locale || "de-CH")
    : `${chapter.id}`;

  const renderChapterList = () => {
    if (!elements.chapterListEl) return;
    elements.chapterListEl.innerHTML = "";
    const chapters = getChapterObjects();
    chapters.forEach((chapter) => {
      const button = document.createElement("button");
      button.type = "button";
      const label = getChapterLabel(chapter);
      button.textContent = label;
      button.title = label;
      if (chapter.id === state.selectedChapterId) button.classList.add("active");
      button.addEventListener("click", () => {
        state.selectedChapterId = chapter.id;
        render();
      });
      elements.chapterListEl.appendChild(button);
    });
  };

  const renderHeads = () => {
    if (elements.targetNameEl) elements.targetNameEl.textContent = state.targetFileName || "Not loaded";
    if (elements.targetMetaEl) elements.targetMetaEl.textContent = describeLibraryMeta(state.targetFileName, state.targetLibrary);
    if (elements.sourceNameEl) elements.sourceNameEl.textContent = state.sourceFileName || "Not loaded";
    if (elements.sourceMetaEl) elements.sourceMetaEl.textContent = describeLibraryMeta(state.sourceFileName, state.sourceLibrary);
  };

  const renderEmpty = (message, visible = true) => {
    if (!elements.emptyEl) return;
    elements.emptyEl.textContent = message || "";
    elements.emptyEl.classList.toggle("is-visible", visible);
  };

  const createTextarea = ({ value, readOnly = false, row, field, panelLabel }) => {
    const panel = document.createElement("div");
    panel.className = "lm-field__panel";
    const label = document.createElement("label");
    label.textContent = panelLabel;
    const textarea = document.createElement("textarea");
    const textId = `lm-${String(row?.key || "row").replace(/[^a-zA-Z0-9_-]/g, "-")}-${field}-${readOnly ? "source" : "target"}`;
    label.htmlFor = textId;
    textarea.id = textId;
    textarea.name = textId;
    textarea.value = value;
    textarea.readOnly = readOnly;
    if (!readOnly) {
      const locale = state.targetLibrary?.meta?.locale || state.projectMeta?.locale || "de-CH";
      const spellLang = resolveSpellcheckLang(locale);
      textarea.setAttribute("lang", spellLang);
      textarea.setAttribute("spellcheck", "true");
      textarea.spellcheck = true;
      textarea.addEventListener("focus", () => {
        textarea.dataset.historyArmed = "1";
      });
      textarea.addEventListener("input", () => {
        if (textarea.dataset.historyArmed === "1") {
          pushHistory();
          textarea.dataset.historyArmed = "0";
        }
        setTargetFieldValue(row, field, textarea.value);
        markTouched(row.key);
        markDirty();
        scheduleAutosave();
        const dot = textarea.closest(".lm-row")?.querySelector(".lm-row__dot");
        if (dot) dot.classList.add("is-touched");
      });
      textarea.addEventListener("blur", () => {
        delete textarea.dataset.historyArmed;
      });
    }
    panel.append(label, textarea);
    return panel;
  };

  const createFieldBlock = (row, field, title) => {
    const block = document.createElement("div");
    block.className = "lm-field";
    block.appendChild(createTextarea({
      value: getTargetFieldValue(row, field),
      readOnly: false,
      row,
      field,
      panelLabel: `My ${title}`,
    }));

    const actions = document.createElement("div");
    actions.className = "lm-field__actions";
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.textContent = "Add";
    addBtn.disabled = !getSourceFieldValue(row, field).trim();
    addBtn.addEventListener("click", () => applyButtonAction(row, field, "append"));
    const replaceBtn = document.createElement("button");
    replaceBtn.type = "button";
    replaceBtn.textContent = "Replace";
    replaceBtn.disabled = !getSourceFieldValue(row, field).trim();
    replaceBtn.addEventListener("click", () => applyButtonAction(row, field, "replace"));
    actions.append(addBtn, replaceBtn);
    block.appendChild(actions);

    block.appendChild(createTextarea({
      value: getSourceFieldValue(row, field),
      readOnly: true,
      row,
      field,
      panelLabel: `Other ${title}`,
    }));
    return block;
  };

  const renderRows = () => {
    if (!elements.rowsEl) return;
    elements.rowsEl.innerHTML = "";
    if (!state.targetLibrary) {
      renderEmpty("Open a project folder to load your library.", true);
      return;
    }
    const chapter = getChapterObjects().find((item) => item.id === state.selectedChapterId) || getChapterObjects()[0];
    if (!chapter) {
      renderEmpty("No chapters available in this library.", true);
      return;
    }
    if (state.selectedChapterId !== chapter.id) state.selectedChapterId = chapter.id;
    if (elements.subtitleEl) {
      elements.subtitleEl.textContent = `Editing ${getChapterLabel(chapter)}`;
    }

    const rows = buildRowsForSelection();
    if (!rows.length) {
      renderEmpty("No reusable texts available in this chapter yet.", true);
      return;
    }
    renderEmpty("", false);

    const fragment = document.createDocumentFragment();
    rows.forEach((row) => {
      const card = document.createElement("section");
      card.className = "lm-row";
      card.dataset.rowKey = row.key;

      const header = document.createElement("div");
      header.className = "lm-row__header";
      const titleWrap = document.createElement("div");
      titleWrap.className = "lm-row__title";
      const dot = document.createElement("span");
      dot.className = "lm-row__dot";
      if (state.touchedKeys.has(row.key)) dot.classList.add("is-touched");
      const titleText = document.createElement("div");
      titleText.className = "lm-row__title-text";
      titleText.textContent = `${row.displayId} ${row.title}`.trim();
      titleWrap.append(dot, titleText);
      header.appendChild(titleWrap);

      const meta = document.createElement("div");
      meta.className = "lm-row__context";
      const contextBits = [];
      if (row.context) contextBits.push(row.context);
      if (row.kind === "observation" && !row.targetHasTag && (row.sourceEntry || row.sourceHasTag)) {
        const badge = document.createElement("span");
        badge.className = "lm-row__badge is-new";
        badge.textContent = "New in Other Library";
        meta.appendChild(badge);
      }
      if (row.kind === "observation" && !row.targetHasTag && (row.sourceEntry || row.sourceHasTag)) {
        const adoptBtn = document.createElement("button");
        adoptBtn.type = "button";
        adoptBtn.className = "ghost";
        adoptBtn.textContent = "Add to My Library";
        adoptBtn.addEventListener("click", () => adoptObservationRow(row));
        meta.appendChild(adoptBtn);
      }
      if (contextBits.length) {
        const text = document.createElement("span");
        text.textContent = contextBits.join(" • ");
        meta.appendChild(text);
      }
      header.appendChild(meta);

      card.appendChild(header);
      card.appendChild(createFieldBlock(row, "finding", "Finding"));
      card.appendChild(createFieldBlock(row, "recommendation", "Recommendation"));
      fragment.appendChild(card);
    });
    elements.rowsEl.appendChild(fragment);
  };

  const render = () => {
    renderChapterList();
    renderHeads();
    updateActionState();
    renderRows();
  };

  const loadSourceLibrary = async () => {
    if (!state.targetLibrary) {
      setStatus("Load your project library first.");
      return;
    }
    if (!window.showOpenFilePicker) {
      setStatus("File picker is not available in this browser.");
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        id: "autobericht-librarymaker-source",
        multiple: false,
        types: [{
          description: "Library JSON",
          accept: { "application/json": [".json"] },
        }],
      });
      const raw = await readJsonFileHandle(handle);
      const library = normalizeKnowledgeBaseForMaker(raw);
      const targetBase = getLocaleBase(state.targetLibrary.meta?.locale || state.projectMeta?.locale);
      const sourceBase = getLocaleBase(library.meta?.locale);
      if (targetBase !== sourceBase) {
        throw new Error(`Other Library locale (${library.meta?.locale || "unknown"}) does not match My Library locale (${state.targetLibrary.meta?.locale || state.projectMeta?.locale || "unknown"}).`);
      }
      state.sourceLibrary = library;
      state.sourceFileName = handle.name;
      setStatus(`Loaded Other Library: ${handle.name}`);
      render();
    } catch (err) {
      if (err?.name === "AbortError") return;
      setStatus(`Loading Other Library failed: ${err.message || err}`);
    }
  };

  const pickProjectFolder = async () => {
    if (!ensureFsAccess()) return;
    try {
      const handle = await window.showDirectoryPicker({ mode: "readwrite", id: "autobericht-project" });
      state.projectHandle = handle;
      if (fsHandles.saveHandle) await fsHandles.saveHandle(handle);
      setFirstRunVisible(false);
      await loadTargetLibraryFromProject();
    } catch (err) {
      if (err?.name === "AbortError") return;
      setStatus(`Project pick failed: ${err.message || err}`);
    }
  };

  const restoreProjectHandle = async () => {
    if (!fsHandles.loadHandle || !fsHandles.requestHandlePermission) return false;
    const saved = await fsHandles.loadHandle();
    if (!saved) return false;
    const granted = await fsHandles.requestHandlePermission(saved);
    if (!granted) return false;
    state.projectHandle = saved;
    try {
      await loadTargetLibraryFromProject();
    } catch (err) {
      setStatus(`Loading My Library failed: ${err.message || err}`);
      return false;
    }
    return true;
  };

  const undo = async () => {
    if (!state.historyPast.length || !state.targetLibrary) return;
    state.historyFuture.push(structuredClone(state.targetLibrary));
    state.targetLibrary = state.historyPast.pop();
    markDirty();
    scheduleAutosave();
    render();
  };

  const redo = async () => {
    if (!state.historyFuture.length || !state.targetLibrary) return;
    state.historyPast.push(structuredClone(state.targetLibrary));
    state.targetLibrary = state.historyFuture.pop();
    markDirty();
    scheduleAutosave();
    render();
  };

  if (elements.openProjectBtn) elements.openProjectBtn.addEventListener("click", pickProjectFolder);
  if (elements.firstRunPickBtn) elements.firstRunPickBtn.addEventListener("click", pickProjectFolder);
  if (elements.loadSourceBtn) elements.loadSourceBtn.addEventListener("click", loadSourceLibrary);
  if (elements.saveBtn) elements.saveBtn.addEventListener("click", async () => {
    try {
      await flushAutosave();
      setStatus(`Saved My Library: ${state.targetFileName}`);
    } catch (err) {
      setStatus(`Save failed: ${err.message || err}`);
    }
  });
  if (elements.undoBtn) elements.undoBtn.addEventListener("click", undo);
  if (elements.redoBtn) elements.redoBtn.addEventListener("click", redo);

  window.addEventListener("beforeunload", () => {
    flushAutosave().catch((err) => debug.logLine("error", `LibraryMaker flush failed: ${err.message || err}`));
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      flushAutosave().catch((err) => debug.logLine("error", `LibraryMaker flush failed: ${err.message || err}`));
    }
  });
  window.addEventListener("pagehide", () => {
    flushAutosave().catch((err) => debug.logLine("error", `LibraryMaker flush failed: ${err.message || err}`));
  });

  const init = async () => {
    ensureFsAccess();
    render();
    const restored = await restoreProjectHandle();
    if (!restored) {
      setFirstRunVisible(true);
      setStatus("Select a project folder to load your library.");
    }
  };

  init();
})();
