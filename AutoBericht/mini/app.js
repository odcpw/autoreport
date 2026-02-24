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
  const spiderModule = window.AutoBerichtSpider || {};
  const spiderUiModule = window.AutoBerichtSpiderUi || {};
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
    backupTimer: null,
    pendingBootstrapWrite: false,
    awaitingLocaleBootstrap: false,
  };

  const setStatus = (message) => {
    if (elements.statusEl) {
      elements.statusEl.textContent = message;
    }
    debug.logLine("info", message);
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

  const ctx = {
    elements,
    state,
    runtime,
    debug,
    setStatus,
    i18n: {
      t,
      setLocale,
      resolveSpellcheckLang: i18n.resolveSpellcheckLang || ((locale) => String(locale || "en").toLowerCase().split("-")[0] || "en"),
    },
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
      spiderModule,
    })
    : {
      loadProjectFromFolder: async () => ({ ok: false, source: "none" }),
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
      if (elements.firstRunPickBtn) elements.firstRunPickBtn.disabled = true;
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
    if (elements.openSpiderBtn) elements.openSpiderBtn.disabled = !enabled;
  };

  const setAutoBackupMinutes = (minutes) => {
    const value = Number(minutes);
    const intervalMinutes = Number.isFinite(value) ? Math.max(0, Math.floor(value)) : 30;
    if (runtime.backupTimer) {
      clearInterval(runtime.backupTimer);
      runtime.backupTimer = null;
    }
    if (!runtime.dirHandle || intervalMinutes <= 0 || !ioApi.backupSidecar) return;
    runtime.backupTimer = setInterval(async () => {
      try {
        await ioApi.backupSidecar();
      } catch (err) {
        debug.logLine("error", `Auto-backup failed: ${err.message || err}`);
      }
    }, intervalMinutes * 60 * 1000);
  };

  const applyAutoBackup = () => {
    const minutes = state.project?.meta?.autobackupMinutes;
    setAutoBackupMinutes(Number.isFinite(Number(minutes)) ? Number(minutes) : 30);
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
  if (renderApi.setSeedBootstrapHandler) {
    renderApi.setSeedBootstrapHandler(ioApi.bootstrapProjectFromSeed);
  }
  if (renderApi.setLibraryExcelExportHandler) {
    renderApi.setLibraryExcelExportHandler(ioApi.exportLibraryExcel);
  }
  if (renderApi.setSidecarMigrationHandler) {
    renderApi.setSidecarMigrationHandler(ioApi.migrateLegacySidecar);
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
      flushAutosave,
      setFirstRunVisible,
      applyAutoBackup,
    });
  }

  const restoreHandle = async () => {
    if (!ctx.fs?.loadHandle || !ctx.fs?.requestHandlePermission) return false;
    const saved = await ctx.fs.loadHandle();
    if (!saved) return false;
    const granted = await ctx.fs.requestHandlePermission(saved);
    if (!granted) return false;
    runtime.dirHandle = saved;
    enableActions();
    await ioApi.loadProjectFromFolder();
    applyAutoBackup();
    return true;
  };

  const init = async () => {
    ensureFsAccess();
    renderApi.render();
    if (!runtime.dirHandle) {
      const restored = await restoreHandle();
      if (!restored) {
        setFirstRunVisible(true);
        setStatus("Select a project folder to start.");
      }
    }
    if (spiderUiModule.init) {
      spiderUiModule.init(ctx, { spider: spiderModule, ioApi });
    }
  };

  init();
})();
