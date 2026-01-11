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
    layoutMode: "stacked",
  };

  const state = stateHelpers.createState
    ? stateHelpers.createState(config.layoutMode)
    : { layoutMode: "stacked", photos: [], tagFilters: { report: "", observations: "", training: "" } };
  const runtime = stateHelpers.createRuntime ? stateHelpers.createRuntime() : { autosaveTimer: null, saveQueue: Promise.resolve(), renderTimer: null };

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
    removeObservationTag: () => {},
    setPhotoUnsorted: () => {},
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

  actions.toggleTag = (group, tag) => {
    const current = photosApi.getCurrentPhoto?.();
    if (!current) return;
    const list = new Set(current.tags[group] || []);
    if (list.has(tag)) {
      list.delete(tag);
    } else {
      list.add(tag);
    }
    current.tags[group] = Array.from(list);
    ioApi.scheduleAutosave?.();
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
    state.tagOptions = tagsApi.DEFAULT_TAG_OPTIONS;
    const statusHidden = window.localStorage?.getItem("photosorterStatusHidden") === "1";
    renderApi.updateStatusVisibility?.(statusHidden);
    renderApi.applyLayoutMode?.();
    renderApi.placeObservationsPanel?.();
    renderApi.updateLayoutToggle?.();
    renderApi.setActivePanel?.(state.activePanel);
    bindApi.ensureFsAccess?.();
    await ioApi.restoreLastHandle?.();
    actions.enableActions();

    if (config.demoPhotosMode) {
      photosApi.loadDemoPhotos?.().catch((err) => {
        setStatus(`Demo photos failed: ${err.message}`);
        debug.logLine("error", `Demo photos failed: ${err.message}`);
      });
    }
  };

  init();
})();
