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

    const isPlainObject = (value) => (
      !!value
      && typeof value === "object"
      && !Array.isArray(value)
    );

    const sanitizePhotoTagOptions = (value) => {
      const normalized = typeof seeds.normalizeTagGroups === "function"
        ? seeds.normalizeTagGroups(value)
        : {
          report: value?.report || [],
          observations: value?.observations || [],
          training: value?.training || [],
        };
      return {
        report: Array.isArray(normalized.report) ? structuredClone(normalized.report) : [],
        observations: Array.isArray(normalized.observations) ? structuredClone(normalized.observations) : [],
        training: Array.isArray(normalized.training) ? structuredClone(normalized.training) : [],
      };
    };

    const sanitizePhotosBranch = (photosBranch) => {
      const source = isPlainObject(photosBranch) ? photosBranch : {};
      const meta = isPlainObject(source.meta) ? structuredClone(source.meta) : {};
      if (!meta.createdAt) meta.createdAt = new Date().toISOString();
      if (meta.updatedAt == null) meta.updatedAt = "";
      if (!Object.prototype.hasOwnProperty.call(meta, "projectId")) meta.projectId = "";
      const photoRoot = typeof source.photoRoot === "string" ? source.photoRoot : "";
      const photos = isPlainObject(source.photos) ? structuredClone(source.photos) : {};
      return {
        meta,
        photoRoot,
        photoTagOptions: sanitizePhotoTagOptions(source.photoTagOptions),
        photos,
      };
    };

    const reorderChapterRows = (project, chapterId) => {
      if (!project?.chapters) return;
      const chapter = project.chapters.find((item) => item.id === chapterId);
      if (!chapter || !Array.isArray(chapter.meta?.order)) return;
      const rows = (chapter.rows || []).filter((row) => row.kind !== "section");
      const byId = new Map(rows.map((row) => [row.id, row]));
      const ordered = [];
      chapter.meta.order.forEach((id) => {
        const match = byId.get(id);
        if (match) ordered.push(match);
      });
      rows.forEach((row) => {
        if (!ordered.includes(row)) ordered.push(row);
      });
      chapter.rows = ordered;
    };

    const mergeSidecar = (baseDoc, project, spiderData) => {
      const merged = baseDoc && typeof baseDoc === "object" ? structuredClone(baseDoc) : {};
      if (!merged.meta) merged.meta = {};
      merged.meta.updatedAt = new Date().toISOString();
      const projectCopy = structuredClone(project);
      reorderChapterRows(projectCopy, "0");
      reorderChapterRows(projectCopy, "4.8");
      merged.report = { project: projectCopy };
      if (Object.prototype.hasOwnProperty.call(merged, "photos")) {
        merged.photos = sanitizePhotosBranch(merged.photos);
      }
      delete merged.photoRoot;
      delete merged.photoTagOptions;
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

    const findDirectoryCaseInsensitive = async (parentHandle, name) => {
      if (!parentHandle) return null;
      try {
        return await parentHandle.getDirectoryHandle(name);
      } catch (err) {
        // probe entries case-insensitively
      }
      if (!parentHandle.entries) return null;
      for await (const [entryName, handle] of parentHandle.entries()) {
        if (handle.kind !== "directory") continue;
        if (String(entryName).toLowerCase() === String(name).toLowerCase()) {
          return handle;
        }
      }
      return null;
    };

    const ensureProjectScaffold = async (projectHandle) => {
      if (!projectHandle) return;
      const ensureDir = async (parentHandle, name) => {
        const existing = await findDirectoryCaseInsensitive(parentHandle, name);
        if (existing) return existing;
        return parentHandle.getDirectoryHandle(name, { create: true });
      };

      await ensureDir(projectHandle, "inputs");
      await ensureDir(projectHandle, "outputs");
      await ensureDir(projectHandle, "backup");
      const photosDir = await ensureDir(projectHandle, "photos");
      await ensureDir(photosDir, "raw");
      await ensureDir(photosDir, "resized");
      await ensureDir(photosDir, "export");
    };

    const loadProjectFromFolder = async () => {
      if (!runtime.dirHandle) return { ok: false, source: "none" };
      try {
        await ensureProjectScaffold(runtime.dirHandle);
      } catch (err) {
        setStatus(`Project scaffold failed: ${err.message || err}`);
        debug.logLine("error", `Project scaffold failed: ${err.message || err}`);
        return { ok: false, source: "none" };
      }
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
        runtime.awaitingLocaleBootstrap = false;
        runtime.pendingBootstrapWrite = false;
        renderApi.buildPhotoIndex();
        state.selectedChapterId = state.project.chapters[0]?.id || "";
        renderApi.render();
        setStatus("Loaded project_sidecar.json");
        debug.logLine("info", "Loaded project_sidecar.json");
        return { ok: true, source: "sidecar" };
      }

      const knowledgeBase = await loadLibraryFromProject();
      let source = "empty";
      let project = null;
      if (knowledgeBase) {
        source = "library";
        project = seeds.buildProjectFromKnowledgeBase(knowledgeBase);
      }

      if (!project) {
        project = {
          meta: {
            locale: "",
            moderator: "",
            moderatorInitials: "",
            coModerator: "",
            coModeratorInitials: "",
            company: "",
            companyId: "",
            address: "",
            plz: "",
            city: "",
            createdAt: new Date().toISOString(),
          },
          chapters: [],
        };
        source = "empty";
      }

      if (source === "empty") {
        state.project = project;
      } else {
        state.project = normalizeHelpers.normalizeProject(project, ctx.i18n.setLocale);
        normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
      }
      state.spiderOverrides = runtime.sidecarDoc?.spider?.overrides || {};
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      if (source === "library" || source === "empty") {
        state.selectedChapterId = "__project__";
      }
      renderApi.buildPhotoIndex();
      renderApi.render();
      runtime.awaitingLocaleBootstrap = source === "empty";
      runtime.pendingBootstrapWrite = false;

      if (source === "library") {
        setStatus("Library loaded; initialized project.");
        debug.logLine("info", "Library loaded; initialized project.");
      } else {
        setStatus("Initialized empty project. Choose report language to bootstrap seed content.");
        debug.logLine("warn", "Initialized empty project; waiting for explicit language bootstrap.");
      }

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
        runtime.pendingBootstrapWrite = false;
        runtime.awaitingLocaleBootstrap = false;
      }).catch((err) => {
        setStatus(`Autosave failed: ${err.message}`);
        debug.logLine("error", `Autosave failed: ${err.message || err}`);
      });
      return runtime.saveQueue;
    };

    const backupSidecar = async () => {
      if (!runtime.dirHandle) return;
      await saveSidecar();
      const payload = runtime.sidecarDoc;
      if (!payload) return;
      const backupDir = await runtime.dirHandle.getDirectoryHandle("backup", { create: true });
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      const filename = `project_sidecar_${timestamp}.json`;
      const handle = await backupDir.getFileHandle(filename, { create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(payload, null, 2));
      await writable.close();
      debug.logLine("info", `Auto-backup saved: backup/${filename}`);
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

    const normalizeTagOption = (tag) => {
      if (typeof tag === "string") {
        const value = String(tag || "").trim();
        if (!value) return null;
        return { value, label: value };
      }
      if (!isPlainObject(tag)) return null;
      const value = String(tag.value || tag.label || "").trim();
      if (!value) return null;
      const label = String(tag.label || value).trim() || value;
      return { value, label };
    };

    const isObservationReportTag = (value) => /^4\.8(?:\.|$)/.test(String(value || "").trim());

    const sortTagOptionsAlpha = (tags, locale = "de") => (
      tags
        .map(normalizeTagOption)
        .filter(Boolean)
        .sort((a, b) => String(a.label || a.value || "")
          .localeCompare(String(b.label || b.value || ""), locale, { numeric: true }))
    );

    const toFlatText = (value) => stateHelpers.toText(value).replace(/\r\n/g, "\n");

    const exportLibraryExcel = async () => {
      if (!runtime.dirHandle) {
        setStatus("Open project folder first.");
        return;
      }
      if (!window.XLSX?.utils?.book_new) {
        throw new Error("XLSX export library is not available.");
      }
      let library = await loadLibraryFile();
      if (!library) {
        library = await loadLibraryFromProject();
      }
      if (!library) {
        throw new Error("Library file not found. Generate or update the library first.");
      }
      seeds.validateKnowledgeBase(library);
      const locale = state.project?.meta?.locale || "de-CH";
      const localeBase = String(locale).toLowerCase().split("-")[0] || "de";
      const workbook = window.XLSX.utils.book_new();
      const appendSheet = (name, rows) => {
        const safeName = String(name || "Sheet").slice(0, 31);
        const data = Array.isArray(rows) && rows.length ? rows : [{}];
        const sheet = window.XLSX.utils.json_to_sheet(data);
        window.XLSX.utils.book_append_sheet(workbook, sheet, safeName);
      };

      const structureRows = (library.structure?.items || []).map((item, index) => {
        const itemId = String(item?.id || "");
        const entryId = String(item?.collapsedId || item?.groupId || itemId);
        return {
          row: index + 1,
          id: itemId,
          entryId,
          collapsedId: item?.collapsedId || "",
          groupId: item?.groupId || "",
          originalId: item?.originalId || "",
          chapter: item?.chapter || "",
          chapterLabel: item?.chapterLabel || "",
          sectionLabel: item?.sectionLabel || "",
          question: toFlatText(item?.question || ""),
        };
      });

      const compareIds = typeof stateHelpers.compareIdSegments === "function"
        ? stateHelpers.compareIdSegments
        : ((a, b) => String(a || "").localeCompare(String(b || ""), "de", { numeric: true }));
      const libraryRows = (library.library?.entries || [])
        .map((entry) => ({
          entryId: entry?.id || "",
          finding: toFlatText(entry?.finding || ""),
          recommendation: toFlatText(entry?.recommendation || ""),
          lastUsed: entry?.lastUsed || "",
        }))
        .sort((a, b) => compareIds(a.entryId, b.entryId));

      const chapterPositiveRows = Object.entries(library.library?.chapterPositives || {})
        .map(([chapterId, value]) => {
          if (typeof value === "string") {
            return {
              chapterId,
              text: toFlatText(value),
              lastUsed: "",
            };
          }
          return {
            chapterId,
            text: toFlatText(value?.text || ""),
            lastUsed: value?.lastUsed || "",
          };
        })
        .sort((a, b) => compareIds(a.chapterId, b.chapterId));

      const toTagRows = (tags, group) => {
        const sorted = group === "observations"
          ? sortTagOptionsAlpha(tags, localeBase)
          : (tags || []).map(normalizeTagOption).filter(Boolean);
        if (group === "report") {
          sorted.sort((a, b) => compareIds(a.value, b.value));
        }
        return sorted.map((tag) => ({ value: tag.value, label: tag.label }));
      };

      const reportTagRows = toTagRows(library.tags?.report || [], "report")
        .filter((tag) => !isObservationReportTag(tag.value));
      const observationTagRows = toTagRows(library.tags?.observations || [], "observations");
      const trainingTagRows = toTagRows(library.tags?.training || [], "training")
        .sort((a, b) => a.label.localeCompare(b.label, localeBase, { numeric: true }));

      const nowIso = new Date().toISOString();
      appendSheet("Meta", [{
        schemaVersion: library.schemaVersion || "",
        locale: library.meta?.locale || "",
        sourceLibrary: stateHelpers.getLibraryFileName(state.project.meta || {}),
        exportedAt: nowIso,
      }]);
      appendSheet("Structure", structureRows);
      appendSheet("LibraryEntries", libraryRows);
      appendSheet("ChapterPositives", chapterPositiveRows);
      appendSheet("TagsReport", reportTagRows);
      appendSheet("TagsObservations", observationTagRows);
      appendSheet("TagsTraining", trainingTagRows);

      const stamp = nowIso.slice(0, 10);
      const jsonName = stateHelpers.getLibraryFileName(state.project.meta || {});
      const baseName = jsonName.replace(/\.json$/i, "");
      const xlsxName = `${baseName}_${stamp}.xlsx`;
      const arrayBuffer = window.XLSX.write(workbook, { type: "array", bookType: "xlsx" });
      const handle = await runtime.dirHandle.getFileHandle(xlsxName, { create: true });
      const writable = await handle.createWritable();
      await writable.write(arrayBuffer);
      await writable.close();
      setStatus(`Library Excel exported: ${xlsxName}`);
      debug.logLine("info", `Library Excel exported: ${xlsxName}`);
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
      const existingChapterPositives = (
        knowledgeBase.library
        && typeof knowledgeBase.library.chapterPositives === "object"
        && !Array.isArray(knowledgeBase.library.chapterPositives)
      ) ? structuredClone(knowledgeBase.library.chapterPositives) : {};
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

        if (typeof normalizeHelpers.ensureChapterMetaDefaults === "function") {
          normalizeHelpers.ensureChapterMetaDefaults(chapter);
        }
        const chapterMeta = chapter?.meta || {};
        const chapterAction = String(chapterMeta.positivesLibraryAction || "off").toLowerCase();
        if (!["append", "replace"].includes(chapterAction)) return;
        const chapterId = String(chapter?.id || "").trim();
        if (!chapterId) return;
        const text = stateHelpers.toText(chapterMeta.positivesText).trim();
        if (!text) return;
        const currentHash = stateHelpers.hashText(text);
        if (chapterMeta.positivesLibraryHash === currentHash) return;
        const currentEntry = existingChapterPositives[chapterId];
        const previousText = typeof currentEntry === "string"
          ? currentEntry
          : stateHelpers.toText(currentEntry?.text);
        let nextText = text;
        if (chapterAction === "append") {
          const existing = previousText.trim();
          nextText = existing ? `${existing}\n\n${text}` : text;
        }
        existingChapterPositives[chapterId] = {
          text: nextText,
          lastUsed: lastUsedAt,
        };
        chapterMeta.positivesLibraryHash = currentHash;
        applied += 1;
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
      output.library.chapterPositives = existingChapterPositives;
      output.tags = output.tags || { report: [], observations: [], training: [] };

      const applySummaryOrder = () => {
        const chapter = state.project.chapters.find((item) => item.id === "0");
        if (!chapter || !Array.isArray(output.structure?.items)) return;
        const order = chapter.meta?.order;
        if (!Array.isArray(order) || !order.length) return;
        const items = output.structure.items;
        const summaryItems = items.filter((item) => {
          const id = String(item?.id || "");
          const chapterId = String(item?.chapter || "");
          return id.startsWith("0.") || chapterId === "0";
        });
        const otherItems = items.filter((item) => !summaryItems.includes(item));
        const byId = new Map(summaryItems.map((item) => [String(item.id || ""), item]));
        const ordered = [];
        order.forEach((id) => {
          const item = byId.get(String(id));
          if (item && !ordered.includes(item)) ordered.push(item);
        });
        summaryItems.forEach((item) => {
          if (!ordered.includes(item)) ordered.push(item);
        });
        output.structure.items = [...ordered, ...otherItems];
      };

      const applyObservationTagOrder = () => {
        if (!Array.isArray(output.tags?.observations)) return;
        const localeBase = String(state.project?.meta?.locale || "de-CH").toLowerCase().split("-")[0] || "de";
        output.tags.observations = sortTagOptionsAlpha(output.tags.observations, localeBase);
      };

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
        output.tags.report = (normalizedTags.report || output.tags.report || [])
          .map(normalizeTagOption)
          .filter(Boolean)
          .filter((tag) => !isObservationReportTag(tag.value));
        output.tags.observations = normalizedTags.observations || output.tags.observations || [];
        output.tags.training = normalizedTags.training || output.tags.training || [];
      }

      applySummaryOrder();
      applyObservationTagOrder();

      await saveLibraryFile(output, timestamp);
      await saveSidecar();
      if (applied === 0) {
        setStatus("Library updated (0 changes). Set Library to Append/Replace on rows or chapter positives.");
      } else {
        setStatus(`Library updated (${applied} changes).`);
      }
      debug.logLine("info", `Library updated (${applied} changes).`);
    };

    const bootstrapProjectFromSeed = async (locale, options = {}) => {
      if (!runtime.dirHandle) {
        throw new Error("Open project folder first.");
      }
      const requestedLocale = String(locale || "").trim();
      if (!requestedLocale) {
        throw new Error("Select a report language before loading seed content.");
      }
      const seedBase = await loadSeedKnowledgeBase(requestedLocale);
      if (!seedBase) {
        throw new Error(`Knowledge base seed not found for locale ${requestedLocale}.`);
      }
      let project = seeds.buildProjectFromKnowledgeBase(seedBase);
      project = normalizeHelpers.normalizeProject(project, ctx.i18n.setLocale);
      state.project = project;
      normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
      state.spiderOverrides = runtime.sidecarDoc?.spider?.overrides || {};
      if (!state.selectedChapterId) {
        state.selectedChapterId = state.project.chapters[0]?.id || "";
      }
      runtime.awaitingLocaleBootstrap = false;
      runtime.pendingBootstrapWrite = options.deferSave === true;
      renderApi.buildPhotoIndex();
      renderApi.render();
      if (options.deferSave !== true) {
        await saveSidecar();
      }
      const deferredMsg = options.deferSave === true
        ? "Seed loaded from language. Sidecar will be saved when leaving Project page."
        : "Seed loaded and saved to project_sidecar.json.";
      setStatus(deferredMsg);
      debug.logLine("info", deferredMsg);
      return { ok: true, deferred: options.deferSave === true };
    };

    return {
      extractReportProject,
      mergeSidecar,
      loadProjectFromFolder,
      saveSidecar,
      backupSidecar,
      loadLibraryFile,
      saveLibraryFile,
      generateLibrary,
      exportLibraryExcel,
      bootstrapProjectFromSeed,
    };
  };

  window.AutoBerichtSidecar = { init };
})();
