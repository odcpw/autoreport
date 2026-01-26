(() => {
  const init = (ctx, deps) => {
    const {
      stateHelpers,
      normalizeHelpers,
      seeds,
      renderApi,
      spiderModule,
    } = deps;
    const { state, runtime, debug, setStatus, elements } = ctx;

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

    const isLibraryFileName = (name) => (
      typeof name === "string"
      && name.startsWith("library_user_")
      && name.endsWith(".json")
      && !/_\d{4}-\d{2}-\d{2}/.test(name)
    );

    const listLibraryFiles = async (dirHandle) => {
      if (!dirHandle?.entries) return [];
      const matches = [];
      for await (const [name, handle] of dirHandle.entries()) {
        if (handle.kind !== "file") continue;
        if (!isLibraryFileName(name)) continue;
        matches.push({ name, handle });
      }
      matches.sort((a, b) => a.name.localeCompare(b.name, "de", { numeric: true }));
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

    const readLibraryFromHandle = async (fileHandle) => {
      const file = await fileHandle.getFile();
      const text = await file.text();
      const data = JSON.parse(text);
      seeds.validateKnowledgeBase(data);
      return data;
    };

    const loadLibraryFromProject = async () => {
      if (!runtime.dirHandle) return null;
      const options = await listLibraryFiles(runtime.dirHandle);
      if (!options.length) return null;
      const picked = await pickLibraryFile(options);
      if (!picked) return null;
      try {
        return await readLibraryFromHandle(picked.handle);
      } catch (err) {
        setStatus(`Invalid library: ${err.message || err}`);
        debug.logLine("error", `Invalid library: ${err.message || err}`);
        return null;
      }
    };

    const loadSeedKnowledgeBase = async (locale) => {
      const seedLocale = locale || "de-CH";
      const seedFilename = seeds.getKnowledgeBaseFilename(seedLocale);
      let knowledgeBase = await seeds.readSeedFromProject(runtime.dirHandle, seedFilename);
      if (!knowledgeBase) {
        knowledgeBase = await seeds.readSeedFromHttp(seedFilename);
      }
      if (!knowledgeBase) return null;
      try {
        seeds.validateKnowledgeBase(knowledgeBase);
      } catch (err) {
        setStatus(`Invalid seed: ${err.message || err}`);
        debug.logLine("error", `Invalid seed: ${err.message || err}`);
        return null;
      }
      return knowledgeBase;
    };

    const loadProjectFromFolder = async () => {
      if (!runtime.dirHandle) return { ok: false, source: "none" };
      let sidecarDoc = null;
      try {
        const handle = await runtime.dirHandle.getFileHandle("project_sidecar.json");
        const file = await handle.getFile();
        const text = await file.text();
        sidecarDoc = JSON.parse(text);
      } catch (err) {
        sidecarDoc = null;
      }

      runtime.sidecarDoc = sidecarDoc;
      const reportProject = extractReportProject(sidecarDoc);
      if (reportProject) {
        state.project = normalizeHelpers.normalizeProject(reportProject, ctx.i18n.setLocale);
        normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
        state.spiderOverrides = runtime.sidecarDoc?.spider?.overrides || {};
        renderApi.buildPhotoIndex();
        state.selectedChapterId = state.project.chapters[0]?.id || "";
        renderApi.render();
        setStatus("Loaded project_sidecar.json");
        debug.logLine("info", "Loaded project_sidecar.json");
        return { ok: true, source: "sidecar" };
      }

      const knowledgeBase = await loadLibraryFromProject();
      let source = "seed";
      let project = null;
      if (knowledgeBase) {
        source = "library";
        project = seeds.buildProjectFromKnowledgeBase(knowledgeBase);
      } else {
        const seedBase = await loadSeedKnowledgeBase(state.project?.meta?.locale);
        if (seedBase) {
          project = seeds.buildProjectFromKnowledgeBase(seedBase);
          source = "seed";
        }
      }

      if (!project) {
        const locale = state.project?.meta?.locale || "de-CH";
        project = {
          meta: { locale, createdAt: new Date().toISOString() },
          chapters: [],
        };
        source = "empty";
      }

      state.project = normalizeHelpers.normalizeProject(project, ctx.i18n.setLocale);
      normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
      state.spiderOverrides = runtime.sidecarDoc?.spider?.overrides || {};
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      renderApi.buildPhotoIndex();
      renderApi.render();

      if (source === "library") {
        setStatus("Library loaded; initialized project.");
        debug.logLine("info", "Library loaded; initialized project.");
      } else if (source === "seed") {
        setStatus("Seed loaded; initialized project.");
        debug.logLine("info", "Seed loaded; initialized project.");
      } else {
        setStatus("Initialized empty project.");
        debug.logLine("warn", "Initialized empty project.");
      }

      await saveSidecar();
      return { ok: true, source };
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
          if (err && err.name === "SyntaxError") {
            setStatus("project_sidecar.json is corrupted. Fix it before saving.");
            debug.logLine("error", `Sidecar parse failed: ${err.message || err}`);
            return;
          }
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
          const action = ws.libraryAction || "off";
          if (action === "off") return;
          const text = stateHelpers.getRecommendationText(row).trim();
          if (!text) return;
          const currentHash = stateHelpers.hashText(text);
          if (ws.libraryHash === currentHash) return;
          const entry = entriesMap.get(row.id) || { id: row.id };
          if (action === "replace") {
            entry.recommendation = text;
          } else if (action === "append") {
            const existing = stateHelpers.toText(entry.recommendation).trim();
            entry.recommendation = existing ? `${existing}\n\n${text}` : text;
          }
          if (entry.finding == null && row.master?.finding) {
            entry.finding = row.master.finding;
          }
          entry.lastUsed = lastUsedAt;
          entriesMap.set(row.id, entry);
          ws.libraryHash = currentHash;
          applied += 1;
        });
      });

      const output = structuredClone(knowledgeBase);
      output.schemaVersion = output.schemaVersion || "1.0";
      output.meta = {
        ...(output.meta || {}),
        author: state.project.meta.moderator || state.project.meta.author || "",
        initials: state.project.meta.moderatorInitials || state.project.meta.initials || "",
        moderator: state.project.meta.moderator || "",
        moderatorInitials: state.project.meta.moderatorInitials || "",
        coModerator: state.project.meta.coModerator || "",
        coModeratorInitials: state.project.meta.coModeratorInitials || "",
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
