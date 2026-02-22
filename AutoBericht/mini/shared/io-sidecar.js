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

    const backfillRowItemsFromSeed = (project, seedBase) => {
      const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => 0);
      const normalizeId = (raw) => String(raw || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/\.+$/g, "");
      const deriveSectionId = (value) => {
        const parts = String(value || "").split(".");
        if (parts.length < 2) return "";
        return `${parts[0]}.${parts[1]}`;
      };
      const seedItems = seedBase?.structure?.items || [];
      if (!project?.chapters || !seedItems.length) {
        return {
          added: 0,
          touched: 0,
          sectionLabelsBackfilled: 0,
          sectionsRebuilt: 0,
        };
      }

      const itemsByGroup = new Map();
      const sectionLabelsById = new Map();
      seedItems.forEach((item) => {
        const groupId = item.collapsedId || item.groupId || item.id;
        if (!groupId) return;
        const group = itemsByGroup.get(groupId) || [];
        group.push(item);
        itemsByGroup.set(groupId, group);

        const sectionId = deriveSectionId(groupId);
        const sectionLabel = String(item.sectionLabel || "").trim();
        if (sectionId && sectionLabel && !sectionLabelsById.has(sectionId)) {
          sectionLabelsById.set(sectionId, sectionLabel);
        }
      });
      itemsByGroup.forEach((group) => group.sort((a, b) => compareIdSegments(a.id, b.id)));

      let added = 0;
      let touched = 0;
      let sectionLabelsBackfilled = 0;
      let sectionsRebuilt = 0;
      project.chapters.forEach((chapter) => {
        const nonSectionRows = [];
        (chapter.rows || []).forEach((row) => {
          if (!row || row.kind === "section") return;
          nonSectionRows.push(row);
          const rowId = String(row.id || "");
          if (!rowId) return;
          if (row.type === "field_observation" || rowId === "4.8" || rowId.startsWith("4.8.")) return;

          const rowSectionId = String(row.sectionId || "").trim() || deriveSectionId(rowId);
          if (rowSectionId && !row.sectionId) row.sectionId = rowSectionId;
          if (rowSectionId && !String(row.sectionLabel || "").trim()) {
            const seededLabel = sectionLabelsById.get(rowSectionId);
            if (seededLabel) {
              row.sectionLabel = seededLabel;
              sectionLabelsBackfilled += 1;
            }
          }

          const groupSeed = itemsByGroup.get(rowId);
          if (!groupSeed || !groupSeed.length) return;

          row.customer = row.customer || {};
          if (!Array.isArray(row.customer.items)) row.customer.items = [];
          const existingIds = new Set(row.customer.items.map((item) => normalizeId(item?.id)));

          let rowAdded = 0;
          groupSeed.forEach((seedItem) => {
            const id = String(seedItem?.id || "");
            if (!id) return;
            const key = normalizeId(id);
            if (!key) return;
            if (existingIds.has(key)) return;
            const clone = structuredClone(seedItem);
            clone.answer = null;
            clone.comment = "";
            clone.evidence = "";
            row.customer.items.push(clone);
            existingIds.add(key);
            rowAdded += 1;
          });
          if (!rowAdded) return;
          row.customer.items.sort((a, b) => compareIdSegments(a.id, b.id));
          if (row.customer.items.length > 1 && row.customer.items[0]?.question) {
            // For tiroir questions, the first sub-question carries the full stem.
            row.titleOverride = row.customer.items[0].question;
          }
          added += rowAdded;
          touched += 1;
        });

        if (chapter.id === "0" || chapter.id === "4.8") return;

        let lastSectionId = "";
        const rebuilt = [];
        nonSectionRows.forEach((row) => {
          const sectionId = String(row?.sectionId || "").trim();
          const sectionLabel = String(row?.sectionLabel || "").trim();
          if (sectionId && sectionLabel && sectionId !== lastSectionId) {
            rebuilt.push({
              kind: "section",
              id: sectionId,
              title: sectionLabel,
            });
            lastSectionId = sectionId;
          }
          rebuilt.push(row);
        });

        const oldSections = (chapter.rows || []).filter((row) => row?.kind === "section");
        const newSections = rebuilt.filter((row) => row?.kind === "section");
        const sectionRowsChanged = (
          oldSections.length !== newSections.length
          || oldSections.some((row, index) => (
            String(row?.id || "") !== String(newSections[index]?.id || "")
            || String(row?.title || "") !== String(newSections[index]?.title || "")
          ))
        );
        if (sectionRowsChanged) {
          chapter.rows = rebuilt;
          sectionsRebuilt += 1;
        }
      });

      return {
        added,
        touched,
        sectionLabelsBackfilled,
        sectionsRebuilt,
      };
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
        let backfill = null;
        try {
          const seedBase = await loadSeedKnowledgeBase(state.project?.meta?.locale);
          if (seedBase) {
            const result = backfillRowItemsFromSeed(state.project, seedBase);
            backfill = result;
            if (result.added) {
              debug.logLine("info", `Backfilled missing questions from seed: ${result.added} items across ${result.touched} rows.`);
            }
            if (result.sectionLabelsBackfilled || result.sectionsRebuilt) {
              debug.logLine(
                "info",
                `Backfilled section metadata from seed: ${result.sectionLabelsBackfilled || 0} labels; rebuilt sections in ${result.sectionsRebuilt || 0} chapters.`,
              );
            }
          }
        } catch (err) {
          debug.logLine("warn", `Seed backfill skipped: ${err.message || err}`);
        }
        normalizeHelpers.syncObservationChapterRows(state.project, runtime.sidecarDoc);
        state.spiderOverrides = runtime.sidecarDoc?.spider?.overrides || {};
        renderApi.buildPhotoIndex();
        state.selectedChapterId = state.project.chapters[0]?.id || "";
        renderApi.render();
        if (backfill?.added || backfill?.sectionLabelsBackfilled || backfill?.sectionsRebuilt) {
          try {
            await saveSidecar();
          } catch (err) {
            debug.logLine("warn", `Seed backfill save failed: ${err.message || err}`);
          }
        }
        const backfillParts = [];
        if (backfill?.added) backfillParts.push(`questions +${backfill.added}`);
        if (backfill?.sectionLabelsBackfilled) backfillParts.push(`section labels +${backfill.sectionLabelsBackfilled}`);
        if (backfill?.sectionsRebuilt) backfillParts.push(`section headers rebuilt ${backfill.sectionsRebuilt}`);
        const backfillNote = backfillParts.length ? ` (${backfillParts.join("; ")})` : "";
        setStatus(`Loaded project_sidecar.json${backfillNote}`);
        debug.logLine("info", `Loaded project_sidecar.json${backfillNote}`);
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
        const chapter = state.project.chapters.find((item) => item.id === "4.8");
        if (!chapter || !Array.isArray(output.tags?.observations)) return;
        const order = chapter.meta?.order;
        if (!Array.isArray(order) || !order.length) return;
        const tags = output.tags.observations;
        const byKey = new Map();
        tags.forEach((tag) => {
          const value = String(tag?.value || tag?.label || "").trim();
          const label = String(tag?.label || tag?.value || "").trim();
          if (value) byKey.set(value, tag);
          if (label) byKey.set(label, tag);
        });
        const ordered = [];
        order.forEach((rowId) => {
          const row = (chapter.rows || []).find((r) => r.id === rowId);
          if (!row) return;
          const key = String(row.tag || row.titleOverride || "").trim();
          if (!key) return;
          const tag = byKey.get(key);
          if (tag && !ordered.includes(tag)) ordered.push(tag);
        });
        tags.forEach((tag) => {
          if (!ordered.includes(tag)) ordered.push(tag);
        });
        output.tags.observations = ordered;
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
        output.tags.report = normalizedTags.report || output.tags.report || [];
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

    return {
      extractReportProject,
      mergeSidecar,
      loadProjectFromFolder,
      saveSidecar,
      backupSidecar,
      loadLibraryFile,
      saveLibraryFile,
      generateLibrary,
    };
  };

  window.AutoBerichtSidecar = { init };
})();
