(() => {
  const init = (ctx, deps) => {
    const { state, runtime, setStatus, debug, elements } = ctx;
    const { tagsApi, photosApi, i18n } = deps;
    const seeds = window.AutoBerichtSeeds || {};
    const normalizeHelpers = window.AutoBerichtNormalize || {};
    const createEmptyTagOptions = () => (
      typeof tagsApi.createEmptyTagOptions === "function"
        ? tagsApi.createEmptyTagOptions()
        : structuredClone(tagsApi.EMPTY_TAG_OPTIONS || {
          report: [],
          observations: [],
          training: [],
        })
    );
    const isPlainObject = (value) => (
      !!value
      && typeof value === "object"
      && !Array.isArray(value)
    );
    const getSortLocale = () => document.documentElement?.getAttribute("lang") || "de";

    const createEmptyProjectDoc = () => ({
      meta: {
        projectId: "",
        createdAt: new Date().toISOString(),
      },
      photos: {},
      photoTagOptions: createEmptyTagOptions(),
      photoRoot: "",
    });

    const normalizePhotoDoc = (doc) => {
      const source = isPlainObject(doc) ? doc : {};
      const meta = isPlainObject(source.meta) ? structuredClone(source.meta) : {};
      if (!Object.prototype.hasOwnProperty.call(meta, "projectId")) meta.projectId = "";
      if (!meta.createdAt) meta.createdAt = new Date().toISOString();
      const photos = isPlainObject(source.photos) ? structuredClone(source.photos) : {};
      const photoTagOptions = tagsApi.ensureTagOptions(source.photoTagOptions || createEmptyTagOptions());
      const photoRoot = typeof source.photoRoot === "string" ? source.photoRoot : "";
      return {
        meta,
        photos,
        photoTagOptions,
        photoRoot,
      };
    };

    const isLegacyPhotoDoc = (doc) => (
      isPlainObject(doc)
      && !isPlainObject(doc.report)
      && (
        Object.prototype.hasOwnProperty.call(doc, "photoRoot")
        || Object.prototype.hasOwnProperty.call(doc, "photoTagOptions")
        || isPlainObject(doc.photos)
      )
    );

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

    const isLibraryFileName = (name) => (
      typeof name === "string"
      && name.startsWith("library_user_")
      && name.endsWith(".json")
      && !/_\d{4}-\d{2}-\d{2}/.test(name)
    );

    const listLibraryFiles = async () => {
      if (!state.projectHandle?.entries) return [];
      const matches = [];
      for await (const [name, handle] of state.projectHandle.entries()) {
        if (handle.kind !== "file") continue;
        if (!isLibraryFileName(name)) continue;
        matches.push({ name, handle });
      }
      const locale = getSortLocale();
      matches.sort((a, b) => a.name.localeCompare(b.name, locale, { numeric: true }));
      return matches;
    };

    const setLibraryModalVisible = (visible) => {
      if (!elements?.libraryModal) return;
      if (visible) {
        elements.libraryModal.classList.add("is-open");
        elements.libraryModal.setAttribute("aria-hidden", "false");
      } else {
        elements.libraryModal.classList.remove("is-open");
        elements.libraryModal.setAttribute("aria-hidden", "true");
      }
    };

    const pickLibraryFile = async (options) => {
      if (!options?.length) return null;
      if (options.length === 1) return options[0];
      if (!elements?.libraryListEl) return options[0];
      return new Promise((resolve) => {
        elements.libraryListEl.innerHTML = "";
        options.forEach((option) => {
          const button = document.createElement("button");
          button.type = "button";
          button.textContent = option.name;
          button.addEventListener("click", () => {
            setLibraryModalVisible(false);
            resolve(option);
          });
          elements.libraryListEl.appendChild(button);
        });
        setLibraryModalVisible(true);
      });
    };

    const isValidKnowledgeBase = (data, source = "knowledge base") => {
      if (!data || typeof data !== "object") {
        debug.logLine("error", `${source} missing or invalid.`);
        return false;
      }
      if (typeof seeds.validateKnowledgeBase === "function") {
        try {
          seeds.validateKnowledgeBase(data);
          return true;
        } catch (err) {
          debug.logLine("error", `${source} validation failed: ${err?.message || err}`);
          return false;
        }
      }
      if (!data.schemaVersion || !data.tags || typeof data.tags !== "object") {
        debug.logLine("error", `${source} missing required schema fields.`);
        return false;
      }
      return true;
    };

    const readKnowledgeBase = async () => {
      if (state.projectHandle) {
        try {
          const options = await listLibraryFiles();
          if (options.length) {
            const picked = await pickLibraryFile(options);
            if (picked) {
              const file = await picked.handle.getFile();
              const parsed = JSON.parse(await file.text());
              if (isValidKnowledgeBase(parsed, `User library (${picked.name})`)) return parsed;
            }
          }
        } catch (err) {
          debug.logLine("warn", `User library load failed: ${err?.message || err}`);
        }
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
          ["photos", "resized"],
          ["Photos", "resized"],
          ["photos", "Resized"],
          ["Photos", "Resized"],
          ["photos"],
          ["Photos"],
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

    const fillMissingTagsFromLibrary = async () => {
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
      state.activeTagFilters = { report: [], observations: [], training: [] };
      state.filterMode = "all";
      try {
        const handle = await state.projectHandle.getFileHandle("project_sidecar.json");
        const file = await handle.getFile();
        const text = await file.text();
        const sidecar = JSON.parse(text);
        state.sidecarDoc = sidecar;
        let photoDoc = null;
        if (isPlainObject(sidecar?.photos)) {
          photoDoc = sidecar.photos;
        } else if (isLegacyPhotoDoc(sidecar)) {
          photoDoc = sidecar;
        }
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
        await fillMissingTagsFromLibrary();
        state.photoRootName = state.projectDoc.photoRoot || "";
        setStatus("Loaded project_sidecar.json");
        debug.logLine("info", "Loaded project_sidecar.json");
      } catch (err) {
        state.sidecarDoc = null;
        state.projectDoc = createEmptyProjectDoc();
        state.tagOptions = createEmptyTagOptions();
        const knowledgeBase = await readKnowledgeBase();
        let reportProject = null;
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
          await fillMissingTagsFromLibrary();
          if (seeds.buildProjectFromKnowledgeBase && normalizeHelpers.normalizeProject) {
            try {
              reportProject = seeds.buildProjectFromKnowledgeBase(knowledgeBase);
              reportProject = normalizeHelpers.normalizeProject(reportProject, i18n.setLocale);
            } catch (projectErr) {
              debug.logLine("error", `Failed to build report project: ${projectErr.message || projectErr}`);
            }
          }
          statusMessage = "Sidecar not found; loaded tags from library.";
        } else {
          statusMessage = "Sidecar not found and no library found. Starting fresh.";
          debug.logLine("warn", "Library not found. Tag options are empty.");
        }
        if (reportProject) {
          state.sidecarDoc = { report: { project: reportProject } };
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
          if (err && err.name === "SyntaxError") {
            setStatus("project_sidecar.json is corrupted. Fix it before saving.");
            debug.logLine("error", `Sidecar parse failed: ${err.message || err}`);
            return;
          }
          existing = state.sidecarDoc;
        }
        const sidecar = existing && typeof existing === "object" ? structuredClone(existing) : {};
        if (!sidecar.meta) sidecar.meta = {};
        sidecar.meta.updatedAt = new Date().toISOString();
        sidecar.photos = payload;
        delete sidecar.photoRoot;
        delete sidecar.photoTagOptions;
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
      loadProjectSidecar,
      saveProjectSidecar,
      scheduleAutosave,
      flushAutosave,
    };
  };

  window.AutoBerichtPhotoSorterSidecar = { init };
})();
