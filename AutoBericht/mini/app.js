(() => {
  const dependencyTools = window.AutoBerichtDependencies;
  if (typeof dependencyTools?.requireModules !== "function") {
    const message = "AutoBericht cannot start: required module AutoBerichtDependencies did not load.";
    const statusEl = document.getElementById("status");
    if (statusEl) statusEl.textContent = message;
    throw new Error(message);
  }
  const modules = dependencyTools.requireModules({
    AutoBerichtElements: ["getElements"],
    AutoReportDebug: ["logLine", "saveLog"],
    AutoBerichtI18n: ["t", "tf", "tHint", "tReport", "setLocale", "resolveSpellcheckLang"],
    AutoBerichtMarkdown: ["escapeHtml", "formatInlineMarkdown", "parseInlineMarkdownSegments"],
    AutoBerichtFsHandle: ["saveHandle", "loadHandle", "requestHandlePermission"],
    AutoBerichtState: ["createState"],
    AutoBerichtNormalize: ["normalizeProject"],
    AutoBerichtSeeds: ["validateKnowledgeBase", "buildProjectFromKnowledgeBase"],
    AutoBerichtSidecar: ["init"],
    AutoBerichtImportSelf: ["createHandler"],
    AutoBerichtRender: ["init"],
    AutoBerichtSpider: ["computeSpider"],
    AutoBerichtSpiderUi: ["init"],
    AutoBerichtBindEvents: ["bind"],
  }, { appName: "AutoBericht", statusElementId: "status" });

  const elements = modules.AutoBerichtElements.getElements();
  const debug = modules.AutoReportDebug;
  const i18n = modules.AutoBerichtI18n;
  const markdown = modules.AutoBerichtMarkdown;
  const fsHandles = modules.AutoBerichtFsHandle;
  const stateHelpers = modules.AutoBerichtState;
  const normalizeHelpers = modules.AutoBerichtNormalize;
  const seeds = modules.AutoBerichtSeeds;
  const ioModule = modules.AutoBerichtSidecar;
  const renderModule = modules.AutoBerichtRender;
  const importModule = modules.AutoBerichtImportSelf;
  const spiderModule = modules.AutoBerichtSpider;
  const spiderUiModule = modules.AutoBerichtSpiderUi;
  const bindModule = modules.AutoBerichtBindEvents;

  const { t, tf, setLocale } = i18n;
  const state = stateHelpers.createState(stateHelpers.defaultProject);

  const runtime = {
    dirHandle: null,
    sidecarDoc: null,
    autosaveTimer: null,
    saveQueue: Promise.resolve(),
    backupTimer: null,
    pendingBootstrapWrite: false,
    awaitingLocaleBootstrap: false,
    hasUnsavedChanges: false,
    restoreErrorMessage: "",
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
      tf,
      tHint: i18n.tHint,
      tReport: i18n.tReport,
      setLocale,
      resolveSpellcheckLang: i18n.resolveSpellcheckLang,
    },
    markdown,
    fs: {
      saveHandle: fsHandles.saveHandle,
      loadHandle: fsHandles.loadHandle,
      requestHandlePermission: fsHandles.requestHandlePermission,
    },
  };

  const renderApi = renderModule.init(ctx, { stateHelpers, normalizeHelpers });
  const ioApi = ioModule.init(ctx, {
    stateHelpers,
    normalizeHelpers,
    seeds,
    renderApi,
    spiderModule,
  });
  const importSelfHandler = importModule.createHandler(ctx, {
    renderRows: renderApi.renderRows,
    saveSidecar: ioApi.saveSidecar,
  });

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus(t("status_fs_api_unavailable"));
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
    runtime.hasUnsavedChanges = true;
    if (runtime.autosaveTimer) clearTimeout(runtime.autosaveTimer);
    runtime.autosaveTimer = setTimeout(async () => {
      runtime.autosaveTimer = null;
      try {
        await ioApi.saveSidecar();
        setStatus(t("status_autosaved"));
      } catch (err) {
        setStatus(tf("status_autosave_failed", "Autosave failed: {error}", { error: err.message || err }));
      }
    }, 2000);
  };

  const flushAutosave = async () => {
    if (!runtime.dirHandle) return;
    const shouldSave = runtime.hasUnsavedChanges || !!runtime.autosaveTimer;
    if (runtime.autosaveTimer) {
      clearTimeout(runtime.autosaveTimer);
      runtime.autosaveTimer = null;
    }
    if (!shouldSave) return;
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
  if (renderApi.setActionPlanExportHandler) {
    renderApi.setActionPlanExportHandler(ioApi.exportActionPlanExcel);
  }

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

  const restoreHandle = async () => {
    if (!ctx.fs?.loadHandle || !ctx.fs?.requestHandlePermission) return false;
    try {
      const saved = await ctx.fs.loadHandle();
      if (!saved) return false;
      const granted = await ctx.fs.requestHandlePermission(saved);
      if (!granted) return false;
      runtime.dirHandle = saved;
      enableActions();
      const result = await ioApi.loadProjectFromFolder();
      if (!result?.ok) {
        throw result?.error || new Error("The saved project folder could not be loaded.");
      }
      applyAutoBackup();
      return true;
    } catch (err) {
      runtime.dirHandle = null;
      runtime.restoreErrorMessage = tf("status_restore_failed", "Saved project folder could not be restored: {error}", { error: err.message || err });
      enableActions();
      setStatus(runtime.restoreErrorMessage);
      debug.logLine("error", `Saved project folder restore failed: ${err.message || err}`);
      return false;
    }
  };

  const init = async () => {
    ensureFsAccess();
    renderApi.render();
    if (!runtime.dirHandle) {
      const restored = await restoreHandle();
      if (!restored) {
        setFirstRunVisible(true);
        if (!runtime.restoreErrorMessage) setStatus(t("status_select_project_folder"));
      }
    }
    if (spiderUiModule.init) {
      spiderUiModule.init(ctx, { spider: spiderModule, ioApi });
    }
  };

  init().catch((err) => {
    setFirstRunVisible(true);
    setStatus(tf("status_app_startup_failed", "AutoBericht startup failed: {error}", { error: err.message || err }));
    debug.logLine("error", `AutoBericht startup failed: ${err.message || err}`);
  });
})();
