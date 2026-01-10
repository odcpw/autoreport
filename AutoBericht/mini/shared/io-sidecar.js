(() => {
  const init = (ctx, deps) => {
    const {
      stateHelpers,
      normalizeHelpers,
      seeds,
      renderApi,
    } = deps;
    const { state, runtime, debug, setStatus } = ctx;

    const extractReportProject = (doc) => {
      if (!doc || typeof doc !== "object") return null;
      if (doc.report && doc.report.project && doc.report.project.chapters) return doc.report.project;
      if (doc.chapters) return doc;
      return null;
    };

    const mergeSidecar = (baseDoc, project) => {
      const merged = baseDoc && typeof baseDoc === "object" ? structuredClone(baseDoc) : {};
      if (!merged.meta) merged.meta = {};
      merged.meta.updatedAt = new Date().toISOString();
      merged.report = { project };
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
          });
          normalizeHelpers.normalizeProject(state.project, ctx.i18n.setLocale);
          normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
          state.selectedChapterId = state.project.chapters[0]?.id || "";
          runtime.sidecarDoc = null;
          renderApi.buildPhotoIndex();
          renderApi.render();
          setStatus("Sidecar not found; loaded seed data.");
          debug.logLine("warn", `Sidecar not found; loaded seeds (${err.message || err})`);
          return true;
        } catch (seedErr) {
          state.project = normalizeHelpers.normalizeProject(
            structuredClone(stateHelpers.defaultProject),
            ctx.i18n.setLocale,
          );
          normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
          state.selectedChapterId = state.project.chapters[0]?.id || "";
          runtime.sidecarDoc = null;
          renderApi.buildPhotoIndex();
          renderApi.render();
          setStatus("Sidecar not found; loaded default template.");
          debug.logLine("warn", `Sidecar not found: ${err.message || err}`);
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
        const payload = mergeSidecar(existing, state.project);
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
      let masterLibrary = await seeds.readSeedFromProject(runtime.dirHandle, "library_master.json");
      if (!masterLibrary) {
        masterLibrary = await seeds.readSeedFromHttp("library_master.json");
      }
      if (!masterLibrary) {
        masterLibrary = { entries: [] };
      }
      const userLibrary = await loadLibraryFile();
      const libraryMap = seeds.buildLibraryMap(masterLibrary, userLibrary);
      const entriesMap = new Map(Array.from(libraryMap.entries()).map(([id, entry]) => [id, entry]));
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

      const meta = {
        author: state.project.meta.author || "",
        initials: state.project.meta.initials || "",
        locale: state.project.meta.locale || "",
        updatedAt: new Date().toISOString(),
      };
      const output = { meta, entries: Array.from(entriesMap.values()) };
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
