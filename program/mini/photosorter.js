(() => {
  const dependencyTools = window.AutoBerichtDependencies;
  if (typeof dependencyTools?.requireModules !== "function") {
    const message = "PhotoSorter cannot start: required module AutoBerichtDependencies did not load.";
    const statusEl = document.getElementById("status-text");
    if (statusEl) statusEl.textContent = message;
    throw new Error(message);
  }
  const modules = dependencyTools.requireModules({
    AutoBerichtPhotoSorterElements: ["getElements"],
    AutoReportDebug: ["logLine", "saveLog"],
    AutoBerichtI18n: ["t", "tf", "tHint", "setLocale", "resolveSpellcheckLang"],
    AutoBerichtFsHandle: ["saveHandle", "loadHandle", "requestHandlePermission"],
    AutoBerichtPhotoImport: ["importRawPhotos", "exportTaggedPhotos"],
    AutoBerichtPhotoSorterState: ["getLayoutConfig", "createState", "createRuntime"],
    AutoBerichtPhotoSorterTags: ["createEmptyTagOptions"],
    AutoBerichtPhotoSorterPhotos: ["init"],
    AutoBerichtPhotoSorterSidecar: ["init"],
    AutoBerichtPhotoSorterRender: ["init"],
    AutoBerichtPhotoSorterBindEvents: ["bind"],
  }, { appName: "PhotoSorter", statusElementId: "status-text" });

  const elements = modules.AutoBerichtPhotoSorterElements.getElements();
  const debug = modules.AutoReportDebug;
  const i18n = modules.AutoBerichtI18n;
  const fs = modules.AutoBerichtFsHandle;
  const photoImport = modules.AutoBerichtPhotoImport;
  const stateHelpers = modules.AutoBerichtPhotoSorterState;
  const tagsApi = modules.AutoBerichtPhotoSorterTags;
  const photosModule = modules.AutoBerichtPhotoSorterPhotos;
  const ioModule = modules.AutoBerichtPhotoSorterSidecar;
  const renderModule = modules.AutoBerichtPhotoSorterRender;
  const bindModule = modules.AutoBerichtPhotoSorterBindEvents;

  const config = stateHelpers.getLayoutConfig();
  const state = stateHelpers.createState();
  const runtime = stateHelpers.createRuntime();
  try {
    if (window.localStorage) {
      state.showTagCounts = window.localStorage.getItem("photosorterShowCounts") === "1";
    }
  } catch (err) {
    state.showTagCounts = false;
  }

  const setStatus = (message) => {
    if (elements.statusTextEl) {
      elements.statusTextEl.textContent = message;
    }
    debug.logLine("info", message);
  };

  const ctx = {
    elements,
    state,
    runtime,
    debug,
    setStatus,
    i18n: {
      t: i18n.t,
      tf: i18n.tf,
      tHint: i18n.tHint,
      setLocale: i18n.setLocale,
      resolveSpellcheckLang: i18n.resolveSpellcheckLang,
    },
    fs: {
      saveHandle: fs.saveHandle,
      loadHandle: fs.loadHandle,
      requestHandlePermission: fs.requestHandlePermission,
    },
    photoImport: {
      importRawPhotos: photoImport.importRawPhotos,
      exportTaggedPhotos: photoImport.exportTaggedPhotos,
    },
    constants: {
      RESIZE_MAX: stateHelpers.RESIZE_MAX,
      RESIZE_QUALITY: stateHelpers.RESIZE_QUALITY,
    },
  };

  let renderApi = null;
  const notifyChange = () => {
    if (renderApi) renderApi.renderAll();
    // Scanning also assigns persistent numbers, even before the first tag edit.
    ioApi.scheduleAutosave();
  };

  const photosApi = photosModule.init(ctx, { tagsApi, notifyChange });

  const actions = {
    toggleTag: () => {},
    toggleFilterTag: () => {},
    clearTagFilters: () => {},
    removeObservationTag: () => {},
    renameObservationTag: () => {},
    setPhotoUnsorted: () => {},
    persistTagOptions: () => {},
    enableActions: () => {},
  };

  renderApi = renderModule.init(ctx, { elements, tagsApi, photosApi, actions });

  const ioApi = ioModule.init(ctx, {
    tagsApi,
    photosApi,
    renderApi,
    i18n: ctx.i18n,
  });
  actions.persistTagOptions = () => {
    ioApi.scheduleAutosave?.();
  };

  const syncCurrentIndex = (preferredPath = "") => {
    const filtered = photosApi.getFilteredPhotos?.() || [];
    if (!filtered.length) {
      state.currentIndex = 0;
      return;
    }
    const wanted = String(preferredPath || "").trim();
    if (wanted) {
      const idx = filtered.findIndex((photo) => photo?.path === wanted);
      if (idx >= 0) {
        state.currentIndex = idx;
        return;
      }
    }
    if (state.currentIndex >= filtered.length) {
      state.currentIndex = filtered.length - 1;
    }
    if (state.currentIndex < 0) {
      state.currentIndex = 0;
    }
  };

  actions.toggleTag = (group, tag) => {
    const current = photosApi.getCurrentPhoto?.();
    if (!current) return;
    const currentPath = current.path;
    const list = new Set(current.tags[group] || []);
    if (list.has(tag)) {
      list.delete(tag);
    } else {
      list.add(tag);
    }
    current.tags[group] = Array.from(list);
    // In "Show Unsorted" mode the photo would vanish the moment it gets a tag;
    // keep it on screen until the user moves on so the tag click is visible.
    state.keepPath = currentPath;
    syncCurrentIndex(currentPath);
    ioApi.scheduleAutosave?.();
    renderApi.renderAll();
  };

  actions.toggleFilterTag = (group, tag) => {
    if (!group || !tag) return;
    state.keepPath = "";
    const currentPath = photosApi.getCurrentPhoto?.()?.path || "";
    const active = new Set(state.activeTagFilters?.[group] || []);
    if (active.has(tag)) {
      active.delete(tag);
    } else {
      active.add(tag);
    }
    state.activeTagFilters[group] = Array.from(active);
    syncCurrentIndex(currentPath);
    renderApi.renderAll();
  };

  actions.clearTagFilters = () => {
    state.activeTagFilters = { report: [], observations: [], training: [] };
    state.filterMode = "all";
    state.keepPath = "";
    syncCurrentIndex();
    renderApi.renderAll();
  };

  // Returns true when the tag was renamed; false when nothing changed or the
  // name is already taken (reported in the status line).
  actions.renameObservationTag = (tag, nextLabel) => {
    const label = String(nextLabel || "").trim();
    if (!tag || !label) return false;
    const options = state.tagOptions?.observations || [];
    const option = options.find((opt) => opt.value === tag);
    if (!option || option.label === label) return false;
    if (options.some((opt) => opt !== option && opt.label === label)) {
      setStatus(`A tag named "${label}" already exists.`);
      return false;
    }
    // Only the display name changes. The stored value stays, so photos keep
    // their tags and the Chapter 4.8 row keeps its text; its title follows.
    option.label = label;
    state.tagOptions.observations = tagsApi.sortOptionsForGroup("observations", options);
    ioApi.scheduleAutosave?.();
    renderApi.renderAll();
    renderApi.renderObservationTagList?.();
    setStatus(`Renamed tag to "${label}".`);
    return true;
  };

  actions.removeObservationTag = (tag) => {
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
    ioApi.scheduleAutosave?.();
    renderApi.renderAll();
    renderApi.renderObservationTagList?.();
  };

  actions.setPhotoUnsorted = () => {
    const current = photosApi.getCurrentPhoto?.();
    if (!current) return;
    current.tags.report = [];
    current.tags.observations = [];
    current.tags.training = [];
    ioApi.scheduleAutosave?.();
    renderApi.renderAll();
  };

  const bindApi = bindModule.bind(ctx, {
    elements,
    tagsApi,
    photosApi,
    ioApi,
    renderApi,
    actions,
  });

  actions.enableActions = bindApi.enableActions;

  const init = async () => {
    state.tagOptions = tagsApi.createEmptyTagOptions();
    let statusHidden = false;
    try {
      statusHidden = window.localStorage?.getItem("photosorterStatusHidden") === "1";
    } catch (err) {
      statusHidden = false;
    }
    renderApi.updateStatusVisibility?.(statusHidden);
    bindApi.ensureFsAccess?.();
    actions.enableActions();
    if (!state.projectHandle && !config.demoPhotosMode) {
      const restored = await (async () => {
        if (!ctx.fs?.loadHandle || !ctx.fs?.requestHandlePermission) return false;
        try {
          const saved = await ctx.fs.loadHandle();
          if (!saved) return false;
          const granted = await ctx.fs.requestHandlePermission(saved);
          if (!granted) return false;
          state.projectHandle = saved;
          actions.enableActions();
          await ioApi.loadProjectSidecar();
          return true;
        } catch (err) {
          state.projectHandle = null;
          runtime.restoreErrorMessage = i18n.tf("status_restore_failed", "Saved project folder could not be restored: {error}", { error: err.message || err });
          actions.enableActions();
          setStatus(runtime.restoreErrorMessage);
          debug.logLine("error", runtime.restoreErrorMessage);
          return false;
        }
      })();
      if (!restored) {
        bindApi.setFirstRunVisible?.(true);
        if (!runtime.restoreErrorMessage) setStatus(i18n.t("status_select_project_folder"));
      }
    }

    if (config.demoPhotosMode) {
      photosApi.loadDemoPhotos?.().catch((err) => {
        setStatus(i18n.tf("status_demo_photos_failed", "Demo photos failed: {error}", { error: err.message || err }));
        debug.logLine("error", `Demo photos failed: ${err.message}`);
      });
    }
  };

  init().catch((err) => {
    bindApi.setFirstRunVisible?.(true);
    setStatus(i18n.tf("status_photosorter_startup_failed", "PhotoSorter startup failed: {error}", { error: err.message || err }));
    debug.logLine("error", `PhotoSorter startup failed: ${err.message || err}`);
  });
})();
