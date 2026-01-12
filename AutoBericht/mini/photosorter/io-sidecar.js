(() => {
  const init = (ctx, deps) => {
    const { state, runtime, setStatus, debug } = ctx;
    const { tagsApi, photosApi, i18n } = deps;

    const createEmptyProjectDoc = () => ({
      meta: {
        projectId: "",
        createdAt: new Date().toISOString(),
      },
      photos: {},
      photoTagOptions: structuredClone(tagsApi.EMPTY_TAG_OPTIONS),
      photoRoot: "",
    });

    const normalizePhotoDoc = (doc) => {
      const base = doc && typeof doc === "object" ? structuredClone(doc) : {};
      if (!base.meta) base.meta = { projectId: "", createdAt: new Date().toISOString() };
      if (!base.photos) base.photos = {};
      if (!base.photoTagOptions) base.photoTagOptions = structuredClone(tagsApi.EMPTY_TAG_OPTIONS);
      if (!base.photoRoot) base.photoRoot = "";
      return base;
    };

    const getNestedDirectory = async (rootHandle, parts, options = {}) => {
      let current = rootHandle;
      for (const part of parts) {
        current = await current.getDirectoryHandle(part, options);
      }
      return current;
    };

    const getDirectoryFromPath = async (rootHandle, path, options = {}) => {
      const parts = String(path || "")
        .split("/")
        .map((part) => part.trim())
        .filter(Boolean);
      if (!parts.length) return rootHandle;
      return getNestedDirectory(rootHandle, parts, options);
    };

    const readJsonFromPath = async (rootHandle, parts) => {
      let current = rootHandle;
      for (let i = 0; i < parts.length - 1; i += 1) {
        current = await current.getDirectoryHandle(parts[i]);
      }
      const fileHandle = await current.getFileHandle(parts[parts.length - 1]);
      const file = await fileHandle.getFile();
      return JSON.parse(await file.text());
    };

    const findLibraryFileHandle = async () => {
      if (!state.projectHandle?.entries) return null;
      for await (const [name, handle] of state.projectHandle.entries()) {
        if (handle.kind !== "file") continue;
        if (!name.startsWith("library_user_") || !name.endsWith(".json")) continue;
        if (/_\d{4}-\d{2}-\d{2}/.test(name)) continue;
        return handle;
      }
      return null;
    };

    const isValidKnowledgeBase = (data) => {
      if (!data || typeof data !== "object") return false;
      if (!data.schemaVersion) return false;
      if (!data.tags || typeof data.tags !== "object") return false;
      return true;
    };

    const readKnowledgeBase = async () => {
      // 1) User library in project root
      if (state.projectHandle) {
        try {
          const libraryHandle = await findLibraryFileHandle();
          if (libraryHandle) {
            const file = await libraryHandle.getFile();
            const parsed = JSON.parse(await file.text());
            if (isValidKnowledgeBase(parsed)) return parsed;
            debug.logLine("error", "User library is missing knowledge base schema.");
          }
        } catch (err) {
          // ignore and try seed
        }
      }

      // 2) Bundled seed next to the app (served by the local server)
      try {
        const baseUrl = new URL(window.location.href);
        const tryPaths = [
          "../data/seed/knowledge_base_de.json",
          "../data/seed/knowledge_base_fr.json",
          "../data/seed/knowledge_base_it.json",
        ];
        for (const rel of tryPaths) {
          const url = new URL(rel, baseUrl);
          const res = await fetch(url.toString());
          if (!res.ok) continue;
          const parsed = await res.json();
          if (isValidKnowledgeBase(parsed)) return parsed;
        }
      } catch (err) {
        debug.logLine("error", `Failed to load bundled seed via fetch: ${err.message || err}`);
      }

      return null;
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
          const handle = await getDirectoryFromPath(state.projectHandle, rootName);
          state.photoHandle = handle;
          state.photoRootName = rootName;
          return;
        } catch (err) {
          // Keep project root as fallback.
        }
      } else if (rootName === state.projectHandle.name) {
        state.photoHandle = state.projectHandle;
        state.photoRootName = "";
        return;
      }
      if (!rootName) {
        const candidates = [
          ["Photos", "resized"],
          ["photos", "resized"],
          ["Photos", "Resized"],
          ["photos", "Resized"],
          ["Photos"],
          ["photos"],
        ];
        for (const parts of candidates) {
          try {
            const handle = await getNestedDirectory(state.projectHandle, parts);
            state.photoHandle = handle;
            state.photoRootName = parts.join("/");
            break;
          } catch (err) {
            // Try next candidate.
          }
        }
      }
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
      state.projectHandle = handle;
      setStatus(`Restored project folder: ${state.projectHandle.name}`);
      debug.logLine("info", `Restored project folder: ${state.projectHandle.name}`);
      await loadProjectSidecar();
    };

    const fillMissingTagsFromSeed = async () => {
      if (state.tagOptions?.report?.length && state.tagOptions?.observations?.length && state.tagOptions?.training?.length) return;
      const knowledgeBase = await readKnowledgeBase();
      if (!knowledgeBase?.tags) return;
      const seedOptions = tagsApi.ensureTagOptions(knowledgeBase.tags);
      state.tagOptions = {
        report: state.tagOptions?.report?.length ? state.tagOptions.report : seedOptions.report,
        observations: state.tagOptions?.observations?.length ? state.tagOptions.observations : seedOptions.observations,
        training: state.tagOptions?.training?.length ? state.tagOptions.training : seedOptions.training,
      };
      if (state.projectDoc) {
        state.projectDoc.photoTagOptions = structuredClone(state.tagOptions);
      }
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
          i18n.setLocale(sidecar.report.project.meta.locale);
        }
        state.tagOptions = tagsApi.ensureTagOptions(state.projectDoc.photoTagOptions);
        const reportOptions = tagsApi.buildReportTagOptionsFromProject(sidecar?.report?.project);
        if (reportOptions && reportOptions.length) {
          state.tagOptions.report = reportOptions;
          state.projectDoc.photoTagOptions.report = reportOptions;
        }
        await fillMissingTagsFromSeed();
        state.photoRootName = state.projectDoc.photoRoot || "";
        setStatus("Loaded project_sidecar.json");
        debug.logLine("info", "Loaded project_sidecar.json");
      } catch (err) {
        state.sidecarDoc = null;
        state.projectDoc = createEmptyProjectDoc();
        state.tagOptions = tagsApi.EMPTY_TAG_OPTIONS;
        const knowledgeBase = await readKnowledgeBase();
        let statusMessage = "Sidecar not found; starting fresh.";
        if (knowledgeBase) {
          const nextOptions = tagsApi.ensureTagOptions(knowledgeBase.tags || {});
          const reportOptions = tagsApi.buildReportTagOptionsFromStructure
            ? tagsApi.buildReportTagOptionsFromStructure(knowledgeBase.structure?.items || [])
            : null;
          if (reportOptions && reportOptions.length) {
            nextOptions.report = reportOptions;
          }
          state.tagOptions = nextOptions;
          state.projectDoc.photoTagOptions = structuredClone(nextOptions);
          await fillMissingTagsFromSeed();
        } else {
          statusMessage = "Sidecar not found and knowledge base missing.";
          debug.logLine("error", "Knowledge base not found. Tags unavailable.");
        }
        state.photoRootName = "";
        setStatus(statusMessage);
        debug.logLine("warn", `Sidecar not found: ${err.message || err}`);
      }
      await setDefaultPhotoHandle();
      const didScan = await photosApi.maybeAutoScan();
      if (!didScan) {
        deps.renderApi.renderAll();
      }
    };

    const saveProjectSidecar = async () => {
      if (!state.projectHandle) return;
      const payload = normalizePhotoDoc(state.projectDoc);
      payload.photos = photosApi.serializePhotos();
      payload.photoTagOptions = structuredClone(state.tagOptions);
      payload.photoRoot = state.photoRootName || "";
      if (!payload.meta) payload.meta = {};
      payload.meta.updatedAt = new Date().toISOString();

      runtime.saveQueue = runtime.saveQueue.then(async () => {
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
      return runtime.saveQueue;
    };

    const scheduleAutosave = () => {
      if (!state.projectHandle) return;
      if (runtime.autosaveTimer) clearTimeout(runtime.autosaveTimer);
      runtime.autosaveTimer = setTimeout(async () => {
        runtime.autosaveTimer = null;
        await saveProjectSidecar();
      }, 2000);
    };

    const flushAutosave = async () => {
      if (runtime.autosaveTimer) {
        clearTimeout(runtime.autosaveTimer);
        runtime.autosaveTimer = null;
      }
      await saveProjectSidecar();
    };

    return {
      createEmptyProjectDoc,
      normalizePhotoDoc,
      getNestedDirectory,
      getDirectoryFromPath,
      setDefaultPhotoHandle,
      persistHandle,
      loadPersistedHandle,
      ensureHandlePermission,
      restoreLastHandle,
      loadProjectSidecar,
      saveProjectSidecar,
      scheduleAutosave,
      flushAutosave,
    };
  };

  window.AutoBerichtPhotoSorterSidecar = { init };
})();
