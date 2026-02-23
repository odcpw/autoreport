(() => {
  const elements = window.AutoBerichtPhotoSorterElements?.getElements?.() || {};
  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };
  const i18n = window.AutoBerichtI18n || {};
  const fs = window.AutoBerichtFsHandle || {};
  const photoImport = window.AutoBerichtPhotoImport || {};
  const stateHelpers = window.AutoBerichtPhotoSorterState || {};
  const tagsApi = window.AutoBerichtPhotoSorterTags || {};
  const photosModule = window.AutoBerichtPhotoSorterPhotos || {};
  const ioModule = window.AutoBerichtPhotoSorterSidecar || {};
  const renderModule = window.AutoBerichtPhotoSorterRender || {};
  const bindModule = window.AutoBerichtPhotoSorterBindEvents || {};

  const config = stateHelpers.getLayoutConfig ? stateHelpers.getLayoutConfig() : {
    demoMode: false,
    demoPhotosMode: false,
  };

  const state = stateHelpers.createState
    ? stateHelpers.createState()
    : {
      photos: [],
      tagFilters: { report: "", observations: "", training: "" },
      activeTagFilters: { report: [], observations: [], training: [] },
    };
  const runtime = stateHelpers.createRuntime ? stateHelpers.createRuntime() : { autosaveTimer: null, saveQueue: Promise.resolve(), renderTimer: null };
  if (window.localStorage) {
    state.showTagCounts = window.localStorage.getItem("photosorterShowCounts") === "1";
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
      setLocale: i18n.setLocale || (() => {}),
      resolveSpellcheckLang: i18n.resolveSpellcheckLang || ((locale) => String(locale || "en").toLowerCase().split("-")[0] || "en"),
    },
    fs: {
      saveHandle: fs.saveHandle || (async () => {}),
      loadHandle: fs.loadHandle || (async () => null),
      requestHandlePermission: fs.requestHandlePermission || (async () => false),
    },
    photoImport: {
      importRawPhotos: photoImport.importRawPhotos,
      exportTaggedPhotos: photoImport.exportTaggedPhotos,
    },
    constants: {
      RESIZE_MAX: stateHelpers.RESIZE_MAX || 1200,
      RESIZE_QUALITY: stateHelpers.RESIZE_QUALITY || 0.85,
    },
  };

  let renderApi = null;
  const notifyChange = () => {
    if (renderApi) renderApi.renderAll();
  };

  const photosApi = photosModule.init
    ? photosModule.init(ctx, { tagsApi, notifyChange })
    : {};

  const actions = {
    toggleTag: () => {},
    toggleFilterTag: () => {},
    clearTagFilters: () => {},
    removeObservationTag: () => {},
    setPhotoUnsorted: () => {},
    persistTagOptions: () => {},
    enableActions: () => {},
  };

  renderApi = renderModule.init
    ? renderModule.init(ctx, { elements, tagsApi, photosApi, actions })
    : { renderAll: () => {} };

  const ioApi = ioModule.init
    ? ioModule.init(ctx, {
      tagsApi,
      photosApi,
      renderApi,
      i18n: ctx.i18n,
    })
    : {};
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
    syncCurrentIndex(currentPath);
    ioApi.scheduleAutosave?.();
    renderApi.renderAll();
  };

  actions.toggleFilterTag = (group, tag) => {
    if (!group || !tag) return;
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
    syncCurrentIndex();
    renderApi.renderAll();
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

  const bindApi = bindModule.bind
    ? bindModule.bind(ctx, {
      elements,
      tagsApi,
      photosApi,
      ioApi,
      renderApi,
      actions,
    })
    : { ensureFsAccess: () => {}, enableActions: () => {} };

  actions.enableActions = bindApi.enableActions || (() => {});

  const init = async () => {
    state.tagOptions = tagsApi.createEmptyTagOptions
      ? tagsApi.createEmptyTagOptions()
      : structuredClone(tagsApi.EMPTY_TAG_OPTIONS || {
        report: [],
        observations: [],
        training: [],
      });
    const statusHidden = window.localStorage?.getItem("photosorterStatusHidden") === "1";
    renderApi.updateStatusVisibility?.(statusHidden);
    bindApi.ensureFsAccess?.();
    actions.enableActions();
    if (!state.projectHandle && !config.demoPhotosMode) {
      const restored = await (async () => {
        if (!ctx.fs?.loadHandle || !ctx.fs?.requestHandlePermission) return false;
        const saved = await ctx.fs.loadHandle();
        if (!saved) return false;
        const granted = await ctx.fs.requestHandlePermission(saved);
        if (!granted) return false;
        state.projectHandle = saved;
        actions.enableActions();
        await ioApi.loadProjectSidecar();
        return true;
      })();
      if (!restored) {
        bindApi.setFirstRunVisible?.(true);
        setStatus("Select a project folder to start.");
      }
    }

    if (config.demoPhotosMode) {
      photosApi.loadDemoPhotos?.().catch((err) => {
        setStatus(`Demo photos failed: ${err.message}`);
        debug.logLine("error", `Demo photos failed: ${err.message}`);
      });
    }
  };

  init();
})();
