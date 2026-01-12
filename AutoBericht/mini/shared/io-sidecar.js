(() => {
  const init = (ctx, deps) => {
    const {
      stateHelpers,
      normalizeHelpers,
      seeds,
      renderApi,
      spiderModule,
    } = deps;
    const { state, runtime, debug, setStatus } = ctx;

    const extractReportProject = (doc) => {
      if (!doc || typeof doc !== "object") return null;
      if (doc.report && doc.report.project && doc.report.project.chapters) return doc.report.project;
      if (doc.chapters) return doc;
      return null;
    };

    const mergeSidecar = (baseDoc, project, spiderData) => {
      const merged = baseDoc && typeof baseDoc === "object" ? structuredClone(baseDoc) : {};
      if (!merged.meta) merged.meta = {};
      merged.meta.updatedAt = new Date().toISOString();
      merged.report = { project };
      if (spiderData) merged.spider = spiderData;
      return merged;
    };

    const loadProjectFromFolder = async () => {
      if (!runtime.dirHandle) return false;
      try {
        const handle = await runtime.dirHandle.getFileHandle("project_sidecar.json");
        const file = await handle.getFile();
        const text = await file.text();
        runtime.sidecarDoc = JSON.parse(text);
        const reportProject = extractReportProject(runtime.sidecarDoc)
          || structuredClone(stateHelpers.defaultProject);
        state.project = normalizeHelpers.normalizeProject(reportProject, ctx.i18n.setLocale);
        normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
        state.spiderOverrides = runtime.sidecarDoc?.spider?.overrides || {};
        renderApi.buildPhotoIndex();
        state.selectedChapterId = state.project.chapters[0]?.id || "";
        renderApi.render();
        setStatus("Loaded project_sidecar.json");
        debug.logLine("info", "Loaded project_sidecar.json");
        return true;
      } catch (err) {
        try {
          state.project = await seeds.loadSeedsForProject({
            dirHandle: runtime.dirHandle,
            getLibraryFileName: () => stateHelpers.getLibraryFileName(state.project.meta || {}),
            locale: state.project?.meta?.locale,
          });
          normalizeHelpers.normalizeProject(state.project, ctx.i18n.setLocale);
          normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
          state.spiderOverrides = {};
          state.selectedChapterId = state.project.chapters[0]?.id || "";
          runtime.sidecarDoc = null;
          renderApi.buildPhotoIndex();
          renderApi.render();
          setStatus("Sidecar not found; loaded seed data.");
          debug.logLine("warn", `Sidecar not found; loaded seeds (${err.message || err})`);
          return true;
        } catch (seedErr) {
          const locale = state.project?.meta?.locale || "de-CH";
          state.project = normalizeHelpers.normalizeProject(
            {
              meta: { locale, createdAt: new Date().toISOString() },
              chapters: [],
            },
            ctx.i18n.setLocale,
          );
          normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
          state.spiderOverrides = {};
          state.selectedChapterId = "";
          runtime.sidecarDoc = null;
          renderApi.buildPhotoIndex();
          renderApi.render();
          setStatus(`Seed load failed: ${seedErr.message || seedErr}`);
          debug.logLine("error", `Seed load failed: ${seedErr.message || seedErr}`);
          return false;
        }
      }
    };

    const saveSidecar = async () => {
      if (!runtime.dirHandle) return;
      runtime.saveQueue = runtime.saveQueue.then(async () => {
        let existing = runtime.sidecarDoc;
        try {
          const handle = await runtime.dirHandle.getFileHandle("project_sidecar.json");
          const file = await handle.getFile();
          const text = await file.text();
          existing = JSON.parse(text);
        } catch (err) {
          existing = runtime.sidecarDoc;
        }
        let spiderData = null;
        if (spiderModule?.computeSpider) {
          try {
            spiderData = await spiderModule.computeSpider({
              project: state.project,
              overrides: state.spiderOverrides || {},
              dirHandle: runtime.dirHandle,
            });
          } catch (err) {
            debug.logLine("error", `Spider compute failed: ${err.message || err}`);
          }
        }
        const payload = mergeSidecar(existing, state.project, spiderData);
        const handle = await runtime.dirHandle.getFileHandle("project_sidecar.json", { create: true });
        const writable = await handle.createWritable();
        await writable.write(JSON.stringify(payload, null, 2));
        await writable.close();
        runtime.sidecarDoc = payload;
      }).catch((err) => {
        setStatus(`Autosave failed: ${err.message}`);
        debug.logLine("error", `Autosave failed: ${err.message || err}`);
      });
      return runtime.saveQueue;
    };

    const loadLibraryFile = async () => {
      if (!runtime.dirHandle) return null;
      const filename = stateHelpers.getLibraryFileName(state.project.meta || {});
      return seeds.readJsonIfExists(runtime.dirHandle, [filename]);
    };

    const saveLibraryFile = async (library, timestampSuffix) => {
      if (!runtime.dirHandle) return;
      const filename = stateHelpers.getLibraryFileName(state.project.meta || {});
      const handle = await runtime.dirHandle.getFileHandle(filename, { create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(library, null, 2));
      await writable.close();
      if (!timestampSuffix) return;
      const archiveName = filename.replace(/\.json$/i, "");
      const archiveHandle = await runtime.dirHandle.getFileHandle(
        `${archiveName}_${timestampSuffix}.json`,
        { create: true },
      );
      const archiveWritable = await archiveHandle.createWritable();
      await archiveWritable.write(JSON.stringify(library, null, 2));
      await archiveWritable.close();
    };

    const generateLibrary = async () => {
      if (!runtime.dirHandle) return;
      normalizeHelpers.ensureProjectMeta(state.project, ctx.i18n.setLocale);
      const locale = state.project.meta?.locale || "de-CH";
      const seedFilename = seeds.getKnowledgeBaseFilename(locale);
      let knowledgeBase = await loadLibraryFile();
      if (!knowledgeBase) {
        knowledgeBase = await seeds.readSeedFromProject(runtime.dirHandle, seedFilename);
      }
      if (!knowledgeBase) {
        knowledgeBase = await seeds.readSeedFromHttp(seedFilename);
      }
      if (!knowledgeBase) {
        setStatus("Knowledge base seed not found.");
        debug.logLine("error", "Knowledge base seed not found.");
        return;
      }
      try {
        seeds.validateKnowledgeBase(knowledgeBase);
      } catch (err) {
        setStatus(`Invalid knowledge base: ${err.message}`);
        debug.logLine("error", `Invalid knowledge base: ${err.message || err}`);
        return;
      }

      const existingEntries = knowledgeBase.library?.entries || [];
      const entriesMap = new Map(existingEntries.map((entry) => [entry.id, entry]));
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      const lastUsedAt = new Date().toISOString();
      let applied = 0;

      state.project.chapters.forEach((chapter) => {
        chapter.rows.forEach((row) => {
          if (row.kind === "section") return;
          normalizeHelpers.ensureWorkstateDefaults(row);
          const ws = row.workstate;
          Object.entries(ws.libraryActions || {}).forEach(([levelKey, action]) => {
            if (action === "off") return;
            const text = stateHelpers.getRecommendationText(row, levelKey).trim();
            if (!text) return;
            const currentHash = stateHelpers.hashText(text);
            if (ws.libraryHashes?.[levelKey] === currentHash) return;
            const entry = entriesMap.get(row.id) || { id: row.id, levels: {} };
            entry.levels = entry.levels || {};
            if (action === "replace") {
              entry.levels[levelKey] = text;
            } else if (action === "append") {
              const existing = stateHelpers.toText(entry.levels[levelKey]).trim();
              entry.levels[levelKey] = existing ? `${existing}\n\n${text}` : text;
            }
            entry.lastUsed = lastUsedAt;
            entriesMap.set(row.id, entry);
            ws.libraryHashes[levelKey] = currentHash;
            applied += 1;
          });
        });
      });

      const output = structuredClone(knowledgeBase);
      output.schemaVersion = output.schemaVersion || "1.0";
      output.meta = {
        ...(output.meta || {}),
        author: state.project.meta.author || "",
        initials: state.project.meta.initials || "",
        locale: state.project.meta.locale || "",
        updatedAt: new Date().toISOString(),
      };
      output.library = output.library || { entries: [] };
      output.library.entries = Array.from(entriesMap.values());
      output.tags = output.tags || { report: [], observations: [], training: [] };

      let latestSidecar = runtime.sidecarDoc;
      try {
        const existingHandle = await runtime.dirHandle.getFileHandle("project_sidecar.json");
        const existingFile = await existingHandle.getFile();
        latestSidecar = JSON.parse(await existingFile.text());
      } catch (err) {
        // Keep last loaded sidecar snapshot.
      }
      const photoTags = latestSidecar?.photos?.photoTagOptions || runtime.sidecarDoc?.photos?.photoTagOptions;
      if (photoTags) {
        const normalizedTags = seeds.normalizeTagGroups(photoTags);
        output.tags.report = normalizedTags.report || output.tags.report || [];
        output.tags.observations = normalizedTags.observations || output.tags.observations || [];
        output.tags.training = normalizedTags.training || output.tags.training || [];
      }

      await saveLibraryFile(output, timestamp);
      await saveSidecar();
      setStatus(`Library updated (${applied} changes).`);
      debug.logLine("info", `Library updated (${applied} changes).`);
    };

    return {
      extractReportProject,
      mergeSidecar,
      loadProjectFromFolder,
      saveSidecar,
      loadLibraryFile,
      saveLibraryFile,
      generateLibrary,
    };
  };

  window.AutoBerichtSidecar = { init };
})();
