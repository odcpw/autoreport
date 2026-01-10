(() => {
  const elements = window.AutoBerichtElements?.getElements?.() || {};
  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };
  const i18n = window.AutoBerichtI18n || {};
  const markdown = window.AutoBerichtMarkdown || {};
  const fsHandles = window.AutoBerichtFsHandle || {};
  const stateHelpers = window.AutoBerichtState || {};
  const normalizeHelpers = window.AutoBerichtNormalize || {};
  const seeds = window.AutoBerichtSeeds || {};
  const ioModule = window.AutoBerichtSidecar || {};
  const renderModule = window.AutoBerichtRender || {};
  const importModule = window.AutoBerichtImportSelf || {};
  const bindModule = window.AutoBerichtBindEvents || {};

  const t = i18n.t || ((key, fallback) => fallback || key);
  const setLocale = i18n.setLocale || (() => {});

  const state = stateHelpers.createState
    ? stateHelpers.createState(stateHelpers.defaultProject)
    : { project: { chapters: [] }, filters: { mode: "all" } };

  const runtime = {
    dirHandle: null,
    sidecarDoc: null,
    autosaveTimer: null,
    saveQueue: Promise.resolve(),
  };

  const setStatus = (message) => {
    if (elements.statusEl) {
      elements.statusEl.textContent = message;
    }
    debug.logLine("info", message);
  };

  const ctx = {
    elements,
    state,
    runtime,
    debug,
    setStatus,
    i18n: { t, setLocale },
    markdown,
    fs: {
      saveHandle: fsHandles.saveHandle || (async () => {}),
      loadHandle: fsHandles.loadHandle || (async () => null),
      requestHandlePermission: fsHandles.requestHandlePermission || (async () => false),
    },
  };

  const renderApi = renderModule.init
    ? renderModule.init(ctx, { stateHelpers, normalizeHelpers })
    : { render: () => {}, renderRows: () => {}, buildPhotoIndex: () => {} };

  const ioApi = ioModule.init
    ? ioModule.init(ctx, {
      stateHelpers,
      normalizeHelpers,
      seeds,
      renderApi,
    })
    : {
      loadProjectFromFolder: async () => false,
      saveSidecar: async () => {},
      generateLibrary: async () => {},
    };

  const importSelfHandler = importModule.createHandler
    ? importModule.createHandler(ctx, { renderRows: renderApi.renderRows, saveSidecar: ioApi.saveSidecar })
    : async () => {};

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus("File System Access API is not available. Open via http://localhost in Edge/Chrome to enable file access.");
      if (elements.pickFolderBtn) elements.pickFolderBtn.disabled = true;
      return false;
    }
    return true;
  };

  const enableActions = () => {
    const enabled = !!runtime.dirHandle;
    if (elements.loadSidecarBtn) elements.loadSidecarBtn.disabled = !enabled;
    if (elements.saveSidecarBtn) elements.saveSidecarBtn.disabled = !enabled;
    if (elements.loadSeedsBtn) elements.loadSeedsBtn.disabled = !enabled;
    if (elements.importSelfBtn) elements.importSelfBtn.disabled = !enabled;
  };

  const persistHandle = async (handle) => {
    try {
      await ctx.fs.saveHandle(handle);
    } catch (err) {
      debug.logLine("warn", `Failed to persist folder handle: ${err.message || err}`);
    }
  };

  const loadPersistedHandle = async () => {
    try {
      return await ctx.fs.loadHandle();
    } catch (err) {
      return null;
    }
  };

  const ensureHandlePermission = async (handle) => {
    try {
      return await ctx.fs.requestHandlePermission(handle);
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
    runtime.dirHandle = handle;
    enableActions();
    setStatus(`Restored project folder: ${runtime.dirHandle.name}`);
    debug.logLine("info", `Restored project folder: ${runtime.dirHandle.name}`);
    await ioApi.loadProjectFromFolder();
  };

  const autoLoadSeeds = async () => {
    if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
      debug.logLine("info", "Seed auto-load skipped (not on http)." );
      return false;
    }
    try {
      state.project = await seeds.loadSeedsForProject({
        dirHandle: runtime.dirHandle,
        getLibraryFileName: () => stateHelpers.getLibraryFileName(state.project.meta || {}),
      });
      normalizeHelpers.normalizeProject(state.project, setLocale);
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      renderApi.render();
      setStatus("Loaded seed data.");
      debug.logLine("info", "Auto-loaded seed data.");
      return true;
    } catch (err) {
      debug.logLine("warn", `Seed auto-load skipped: ${err.message}`);
      return false;
    }
  };

  const scheduleAutosave = () => {
    if (!runtime.dirHandle) return;
    if (runtime.autosaveTimer) clearTimeout(runtime.autosaveTimer);
    runtime.autosaveTimer = setTimeout(async () => {
      runtime.autosaveTimer = null;
      try {
        await ioApi.saveSidecar();
        setStatus("Autosaved.");
      } catch (err) {
        setStatus(`Autosave failed: ${err.message}`);
      }
    }, 2000);
  };

  const flushAutosave = async () => {
    if (!runtime.dirHandle) return;
    if (runtime.autosaveTimer) {
      clearTimeout(runtime.autosaveTimer);
      runtime.autosaveTimer = null;
    }
    await ioApi.saveSidecar();
  };

  if (renderApi.setScheduleAutosave) {
    renderApi.setScheduleAutosave(scheduleAutosave);
  }

  if (bindModule.bind) {
    bindModule.bind(ctx, {
      renderApi,
      ioApi,
      seeds,
      stateHelpers,
      normalizeHelpers,
      importSelfHandler,
      ensureFsAccess,
      enableActions,
      persistHandle,
      flushAutosave,
    });
  }

  const init = async () => {
    ensureFsAccess();
    await restoreLastHandle();
    renderApi.render();
    if (!runtime.dirHandle) {
      await autoLoadSeeds();
    }
  };

  init();
})();
