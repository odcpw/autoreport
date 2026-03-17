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
    const textDecoder = new TextDecoder();
    const textEncoder = new TextEncoder();
    const zipTools = window.AutoBerichtWordDocxZip || {};
    const reportRows = window.AutoBerichtReportRows || {};
    const unzipAllEntries = zipTools?.unzipAllEntries;
    const buildZipStore = zipTools?.buildZipStore;
    const XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const XML_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const ACTION_PLAN_TEMPLATE_BY_LOCALE = {
      de: "templates/Vorlage Aktionsplan d.V01.xlsx",
      fr: "templates/Vorlage Aktionsplan f.V01.xlsx",
      it: "templates/Vorlage Aktionsplan i.V01.xlsx",
    };
    const ACTION_PLAN_SHEET_NAMES = {
      de: { plan: "Aktionsplan", report: "Berichtsbasis" },
      fr: { plan: "Plan d'action", report: "Base du rapport" },
      it: { plan: "Piano d'azione", report: "Base del rapporto" },
    };
    const ACTION_PLAN_TEMPLATE_MAX_QUARTERS = 11;
    const ACTION_PLAN_PLAN_QUARTER_START_COL = 9; // J
    const ACTION_PLAN_PLAN_YEAR_ROW = 3;
    const ACTION_PLAN_PLAN_QUARTER_ROW = 4;

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

    const findFileCaseInsensitive = async (parentHandle, name) => {
      if (!parentHandle) return null;
      try {
        return await parentHandle.getFileHandle(name);
      } catch (err) {
        // probe entries case-insensitively
      }
      if (!parentHandle.entries) return null;
      for await (const [entryName, handle] of parentHandle.entries()) {
        if (handle.kind !== "file") continue;
        if (String(entryName).toLowerCase() === String(name).toLowerCase()) {
          return handle;
        }
      }
      return null;
    };

    const parseRelativePath = (value) => {
      const parts = String(value || "")
        .split("/")
        .map((part) => part.trim())
        .filter(Boolean);
      if (!parts.length) throw new Error("Invalid empty path in project template manifest.");
      if (parts.some((part) => part === "." || part === "..")) {
        throw new Error(`Invalid path in project template manifest: ${value}`);
      }
      return parts;
    };

    const loadBundledProjectTemplateManifest = async () => {
      const manifestUrl = new URL("../project-template/manifest.json", window.location.href).toString();
      const response = await fetch(manifestUrl, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`Bundled project template manifest not found (${response.status}).`);
      }
      const manifest = await response.json();
      if (!manifest || typeof manifest !== "object") {
        throw new Error("Bundled project template manifest is invalid.");
      }
      if (!Array.isArray(manifest.directories) || !Array.isArray(manifest.files)) {
        throw new Error("Bundled project template manifest must define directories and files.");
      }
      return manifest;
    };

    const fetchBundledProjectTemplateBytes = async (relativePath) => {
      const sourceParts = parseRelativePath(relativePath);
      const sourceUrl = new URL(`../project-template/${sourceParts.join("/")}`, window.location.href).toString();
      const response = await fetch(sourceUrl, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`Bundled template file missing (${response.status}): ${sourceParts.join("/")}`);
      }
      return new Uint8Array(await response.arrayBuffer());
    };

    const copyBundledProjectTemplateToFolder = async (projectHandle, ensureDir, options = {}) => {
      const manifest = await loadBundledProjectTemplateManifest();
      const overwrite = options.overwrite === true;

      for (let i = 0; i < manifest.directories.length; i += 1) {
        const dirParts = parseRelativePath(manifest.directories[i]);
        let parent = projectHandle;
        for (let p = 0; p < dirParts.length; p += 1) {
          // eslint-disable-next-line no-await-in-loop
          parent = await ensureDir(parent, dirParts[p]);
        }
      }

      for (let i = 0; i < manifest.files.length; i += 1) {
        const entry = manifest.files[i];
        const sourceRel = String(entry?.source || "").trim();
        const targetRel = String(entry?.target || sourceRel).trim();
        const sourceParts = parseRelativePath(sourceRel);
        const targetParts = parseRelativePath(targetRel);
        const fileName = targetParts[targetParts.length - 1];
        if (!fileName) throw new Error(`Invalid target file path in project template manifest: ${targetRel}`);

        let parent = projectHandle;
        for (let p = 0; p < targetParts.length - 1; p += 1) {
          // eslint-disable-next-line no-await-in-loop
          parent = await ensureDir(parent, targetParts[p]);
        }
        // eslint-disable-next-line no-await-in-loop
        const existingFile = await findFileCaseInsensitive(parent, fileName);
        if (existingFile && !overwrite) continue;
        // eslint-disable-next-line no-await-in-loop
        const bytes = await fetchBundledProjectTemplateBytes(sourceParts.join("/"));
        // eslint-disable-next-line no-await-in-loop
        const fileHandle = existingFile || await parent.getFileHandle(fileName, { create: true });
        // eslint-disable-next-line no-await-in-loop
        const writable = await fileHandle.createWritable();
        // eslint-disable-next-line no-await-in-loop
        await writable.write(bytes);
        // eslint-disable-next-line no-await-in-loop
        await writable.close();
      }
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
      await ensureDir(projectHandle, "templates");
      const photosDir = await ensureDir(projectHandle, "photos");
      const rawDir = await ensureDir(photosDir, "raw");
      await ensureDir(rawDir, "pm1");
      await ensureDir(rawDir, "pm2");
      await ensureDir(rawDir, "pm3");
      await ensureDir(photosDir, "resized");
      await ensureDir(photosDir, "export");
      await copyBundledProjectTemplateToFolder(projectHandle, ensureDir, { overwrite: false });
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
      runtime.awaitingLocaleBootstrap = source === "empty";
      runtime.pendingBootstrapWrite = false;
      renderApi.buildPhotoIndex();
      renderApi.render();

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

    const getLocaleBase = (locale) => {
      const base = String(locale || "").toLowerCase().split("-")[0];
      return ["de", "fr", "it"].includes(base) ? base : "de";
    };

    const toFileSafeSlug = (value, fallback = "Company") => {
      const raw = String(value || "").trim();
      const hyphenated = raw.replace(/\s+/g, "-");
      const cleaned = hyphenated
        .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
        .replace(/-+/g, "-")
        .replace(/^[.\-\s]+|[.\-\s]+$/g, "");
      return cleaned || fallback;
    };

    const getNestedDirectory = async (parentHandle, parts, options = {}) => {
      let current = parentHandle;
      const list = Array.isArray(parts) ? parts : String(parts || "").split("/").filter(Boolean);
      for (let i = 0; i < list.length; i += 1) {
        const name = String(list[i] || "").trim();
        if (!name) continue;
        // eslint-disable-next-line no-await-in-loop
        current = options.create
          ? await current.getDirectoryHandle(name, { create: true })
          : await current.getDirectoryHandle(name);
      }
      return current;
    };

    const getFileHandleFromPath = async (projectHandle, relativePath) => {
      const parts = String(relativePath || "").split("/").map((part) => part.trim()).filter(Boolean);
      if (!parts.length) return null;
      let dir = projectHandle;
      for (let i = 0; i < parts.length - 1; i += 1) {
        // eslint-disable-next-line no-await-in-loop
        dir = await findDirectoryCaseInsensitive(dir, parts[i]);
        if (!dir) return null;
      }
      const fileName = parts[parts.length - 1];
      try {
        return await dir.getFileHandle(fileName);
      } catch (err) {
        if (dir?.entries) {
          for await (const [name, handle] of dir.entries()) {
            if (handle.kind !== "file") continue;
            if (name.toLowerCase() === fileName.toLowerCase()) return handle;
          }
        }
        if (err?.name === "NotFoundError") return null;
        throw err;
      }
    };

    const readProjectBinaryFile = async (projectHandle, relativePath) => {
      const handle = await getFileHandleFromPath(projectHandle, relativePath);
      if (!handle) return null;
      return handle.getFile();
    };

    const announceTemplateStatus = (message) => {
      setStatus(message);
      debug.logLine("info", message);
    };

    const resolveProjectOrBundledTemplateFile = async (projectHandle, relativePath, label, mimeType) => {
      const projectFile = await readProjectBinaryFile(projectHandle, relativePath);
      if (projectFile) {
        announceTemplateStatus(`${label}: using project template ${relativePath}`);
        return projectFile;
      }

      announceTemplateStatus(`${label}: project template missing, trying bundled template ${relativePath}`);
      let bundledBytes;
      try {
        bundledBytes = await fetchBundledProjectTemplateBytes(relativePath);
      } catch (err) {
        throw new Error(`Missing ${label.toLowerCase()} template: ${relativePath}`);
      }

      try {
        const pathParts = parseRelativePath(relativePath);
        const fileName = pathParts[pathParts.length - 1];
        const parent = await getNestedDirectory(projectHandle, pathParts.slice(0, -1), { create: true });
        await writeBinaryFile(parent, fileName, bundledBytes);
        const copiedFile = await readProjectBinaryFile(projectHandle, relativePath);
        if (copiedFile) {
          announceTemplateStatus(`${label}: copied bundled template to project templates and using it`);
          return copiedFile;
        }
      } catch (err) {
        announceTemplateStatus(`${label}: bundled template copy failed (${err.message || err}); using bundled template for this run`);
      }

      announceTemplateStatus(`${label}: using bundled template for this run`);
      return new Blob([bundledBytes], { type: mimeType });
    };

    const writeBinaryFile = async (dirHandle, name, data) => {
      const handle = await dirHandle.getFileHandle(name, { create: true });
      const writable = await handle.createWritable();
      await writable.write(data);
      await writable.close();
      return handle;
    };

    const getOutputsDirectory = async (projectHandle) => {
      try {
        return await getNestedDirectory(projectHandle, ["outputs"], { create: false });
      } catch (err) {
        return getNestedDirectory(projectHandle, ["outputs"], { create: true });
      }
    };

    const normalizeZipPartName = (value) => String(value || "").replace(/^\/+/, "");

    const resolveZipPartTarget = (baseDir, target) => {
      const normalizedBase = `${normalizeZipPartName(baseDir).replace(/\/?$/, "/")}`;
      const resolved = new URL(String(target || ""), `https://zip.invalid/${normalizedBase}`).pathname;
      return normalizeZipPartName(resolved);
    };

    const getEntryText = (map, name) => {
      const entry = map.get(normalizeZipPartName(name));
      return entry ? textDecoder.decode(entry.data) : "";
    };

    const setEntryText = (map, name, xml) => {
      const key = normalizeZipPartName(name);
      map.set(key, {
        name: key,
        data: textEncoder.encode(xml),
        flags: map.get(key)?.flags || 0,
      });
    };

    const parseXml = (xml, label) => {
      const doc = new DOMParser().parseFromString(String(xml || ""), "application/xml");
      if (doc.getElementsByTagName("parsererror").length) {
        throw new Error(`Invalid XML in ${label}.`);
      }
      return doc;
    };

    const serializeXml = (doc) => new XMLSerializer().serializeToString(doc);

    const ensureWorkbookNamespaceAliases = (xml) => {
      const source = String(xml || "");
      if (!source.includes("<workbook")) return source;
      const ignorableMatch = source.match(/\bmc:Ignorable="([^"]+)"/);
      if (!ignorableMatch) return source;
      const ignorable = new Set(
        String(ignorableMatch[1] || "")
          .split(/\s+/)
          .map((item) => item.trim())
          .filter(Boolean),
      );
      const aliasMap = {
        x15: "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
        xr6: "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
        xr10: "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
      };
      const missing = Object.entries(aliasMap)
        .filter(([alias]) => ignorable.has(alias) && !source.includes(`xmlns:${alias}=`))
        .map(([alias, uri]) => ` xmlns:${alias}="${uri}"`)
        .join("");
      if (!missing) return source;
      return source.replace("<workbook", `<workbook${missing}`);
    };

    const getChildElements = (node, localName) => Array.from(node?.getElementsByTagNameNS?.("*", localName) || []);

    const getFirstChild = (node, localName) => getChildElements(node, localName)[0] || null;

    const getAttributeByLocalName = (node, localName) => {
      if (!node?.attributes) return "";
      const attr = Array.from(node.attributes).find((item) => item.localName === localName);
      return attr ? attr.value : "";
    };

    const colLetterToIndex = (value) => {
      let out = 0;
      const input = String(value || "").trim().toUpperCase();
      for (let i = 0; i < input.length; i += 1) {
        out = (out * 26) + (input.charCodeAt(i) - 64);
      }
      return out - 1;
    };

    const indexToColLetter = (index) => {
      let n = Number(index);
      if (!Number.isFinite(n) || n < 0) return "A";
      let out = "";
      while (n >= 0) {
        out = String.fromCharCode((n % 26) + 65) + out;
        n = Math.floor(n / 26) - 1;
      }
      return out;
    };

    const parseCellRef = (ref) => {
      const match = String(ref || "").match(/^([A-Z]+)(\d+)$/i);
      if (!match) return null;
      return {
        col: colLetterToIndex(match[1]),
        row: Number(match[2]),
      };
    };

    const compareCellRefs = (a, b) => {
      const cellA = parseCellRef(a);
      const cellB = parseCellRef(b);
      if (!cellA || !cellB) return String(a || "").localeCompare(String(b || ""));
      if (cellA.row !== cellB.row) return cellA.row - cellB.row;
      return cellA.col - cellB.col;
    };

    const ensureSheetRow = (sheetDoc, rowNumber) => {
      const sheetData = getFirstChild(sheetDoc.documentElement, "sheetData");
      if (!sheetData) throw new Error("Worksheet is missing sheetData.");
      const rows = Array.from(sheetData.getElementsByTagNameNS("*", "row"));
      const existing = rows.find((row) => Number(row.getAttribute("r")) === Number(rowNumber));
      if (existing) return existing;
      const row = sheetDoc.createElementNS(XML_NS, "row");
      row.setAttribute("r", String(rowNumber));
      const insertBefore = rows.find((item) => Number(item.getAttribute("r")) > Number(rowNumber));
      if (insertBefore) {
        sheetData.insertBefore(row, insertBefore);
      } else {
        sheetData.appendChild(row);
      }
      return row;
    };

    const ensureSheetCell = (sheetDoc, rowNumber, colLetter) => {
      const ref = `${String(colLetter || "").toUpperCase()}${Number(rowNumber)}`;
      const row = ensureSheetRow(sheetDoc, rowNumber);
      const cells = Array.from(row.getElementsByTagNameNS("*", "c"));
      const existing = cells.find((cell) => String(cell.getAttribute("r") || "").toUpperCase() === ref);
      if (existing) return existing;
      const cell = sheetDoc.createElementNS(XML_NS, "c");
      cell.setAttribute("r", ref);
      const insertBefore = cells.find((item) => compareCellRefs(item.getAttribute("r"), ref) > 0);
      if (insertBefore) {
        row.insertBefore(cell, insertBefore);
      } else {
        row.appendChild(cell);
      }
      return cell;
    };

    const setInlineStringCellValue = (sheetDoc, rowNumber, colLetter, value) => {
      const cell = ensureSheetCell(sheetDoc, rowNumber, colLetter);
      Array.from(cell.childNodes).forEach((child) => cell.removeChild(child));
      cell.setAttribute("t", "inlineStr");
      const is = sheetDoc.createElementNS(XML_NS, "is");
      const text = sheetDoc.createElementNS(XML_NS, "t");
      text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      text.textContent = String(value == null ? "" : value);
      is.appendChild(text);
      cell.appendChild(is);
      return cell;
    };

    const setFormulaCellValue = (sheetDoc, rowNumber, colLetter, formula, styleId = "") => {
      const cell = ensureSheetCell(sheetDoc, rowNumber, colLetter);
      Array.from(cell.childNodes).forEach((child) => cell.removeChild(child));
      cell.removeAttribute("t");
      if (styleId) {
        cell.setAttribute("s", styleId);
      }
      const formulaEl = sheetDoc.createElementNS(XML_NS, "f");
      formulaEl.textContent = String(formula || "");
      cell.appendChild(formulaEl);
      const valueEl = sheetDoc.createElementNS(XML_NS, "v");
      valueEl.textContent = "";
      cell.appendChild(valueEl);
      return cell;
    };

    const replaceCellRowNumber = (ref, rowNumber) => {
      const parsed = parseCellRef(ref);
      if (!parsed) return String(ref || "");
      return `${indexToColLetter(parsed.col)}${Number(rowNumber)}`;
    };

    const rewriteRowNumber = (rowEl, rowNumber) => {
      rowEl.setAttribute("r", String(rowNumber));
      Array.from(rowEl.getElementsByTagNameNS("*", "c")).forEach((cell) => {
        const currentRef = String(cell.getAttribute("r") || "");
        if (currentRef) {
          cell.setAttribute("r", replaceCellRowNumber(currentRef, rowNumber));
        }
      });
      return rowEl;
    };

    const cloneSheetRow = (rowEl, rowNumber) => {
      if (!rowEl) return null;
      const clone = rowEl.cloneNode(true);
      return rewriteRowNumber(clone, rowNumber);
    };

    const getSheetDataNode = (sheetDoc) => {
      const sheetData = getFirstChild(sheetDoc.documentElement, "sheetData");
      if (!sheetData) throw new Error("Worksheet is missing sheetData.");
      return sheetData;
    };

    const getSheetRows = (sheetDoc) => Array.from(getSheetDataNode(sheetDoc).getElementsByTagNameNS("*", "row"));

    const insertSheetRowSorted = (sheetDoc, rowEl) => {
      const sheetData = getSheetDataNode(sheetDoc);
      const rowNumber = Number(rowEl?.getAttribute("r"));
      const insertBefore = getSheetRows(sheetDoc)
        .find((item) => Number(item.getAttribute("r")) > rowNumber);
      if (insertBefore) {
        sheetData.insertBefore(rowEl, insertBefore);
      } else {
        sheetData.appendChild(rowEl);
      }
      return rowEl;
    };

    const shiftSheetRows = (sheetDoc, startRow, delta) => {
      const start = Number(startRow);
      const offset = Number(delta);
      if (!Number.isFinite(start) || !Number.isFinite(offset) || !offset) return;
      const rows = getSheetRows(sheetDoc)
        .filter((row) => Number(row.getAttribute("r")) >= start)
        .sort((a, b) => (
          offset > 0
            ? Number(b.getAttribute("r")) - Number(a.getAttribute("r"))
            : Number(a.getAttribute("r")) - Number(b.getAttribute("r"))
        ));
      rows.forEach((row) => {
        rewriteRowNumber(row, Number(row.getAttribute("r")) + offset);
      });
    };

    const ensureReportRows = (sheetDoc, firstDataRow, currentLastDataRow, targetLastDataRow) => {
      const sheetData = getSheetDataNode(sheetDoc);
      const rows = getSheetRows(sheetDoc);
      const templateRow = rows.find((row) => Number(row.getAttribute("r")) === Number(firstDataRow));
      if (!templateRow) throw new Error(`Template worksheet is missing data row ${firstDataRow}.`);

      const currentLast = Number(currentLastDataRow);
      const targetLast = Number(targetLastDataRow);
      if (!Number.isFinite(currentLast) || !Number.isFinite(targetLast)) {
        throw new Error("Invalid report table bounds.");
      }

      if (targetLast > currentLast) {
        shiftSheetRows(sheetDoc, currentLast + 1, targetLast - currentLast);
        for (let rowNo = currentLast + 1; rowNo <= targetLast; rowNo += 1) {
          const clone = cloneSheetRow(templateRow, rowNo);
          if (!clone) continue;
          insertSheetRowSorted(sheetDoc, clone);
        }
        return;
      }

      if (targetLast < currentLast) {
        getSheetRows(sheetDoc).forEach((row) => {
          const rowNo = Number(row.getAttribute("r"));
          if (rowNo > targetLast && rowNo <= currentLast) {
            sheetData.removeChild(row);
          }
        });
        shiftSheetRows(sheetDoc, currentLast + 1, targetLast - currentLast);
      }
    };

    const updateSheetDimension = (sheetDoc, ref) => {
      let dimension = getFirstChild(sheetDoc.documentElement, "dimension");
      if (!dimension) {
        dimension = sheetDoc.createElementNS(XML_NS, "dimension");
        sheetDoc.documentElement.insertBefore(dimension, sheetDoc.documentElement.firstChild);
      }
      dimension.setAttribute("ref", String(ref || "A1"));
    };

    const resolveSheetParts = (templateMap) => {
      const workbookDoc = parseXml(getEntryText(templateMap, "xl/workbook.xml"), "xl/workbook.xml");
      const relsDoc = parseXml(getEntryText(templateMap, "xl/_rels/workbook.xml.rels"), "xl/_rels/workbook.xml.rels");
      const relMap = new Map(
        getChildElements(relsDoc.documentElement, "Relationship").map((rel) => [
          rel.getAttribute("Id"),
          resolveZipPartTarget("xl/", getAttributeByLocalName(rel, "Target")),
        ]),
      );
      const out = new Map();
      const sheetsNode = getFirstChild(workbookDoc.documentElement, "sheets");
      const sheets = sheetsNode ? Array.from(sheetsNode.children) : [];
      sheets.forEach((sheet) => {
        if (sheet.localName !== "sheet") return;
        const name = sheet.getAttribute("name");
        const relId = Array.from(sheet.attributes).find((attr) => attr.localName === "id")?.value || "";
        const part = relMap.get(relId);
        if (name && part) out.set(name, part);
      });
      return out;
    };

    const relsPartNameForPart = (partName) => {
      const normalized = normalizeZipPartName(partName);
      const parts = normalized.split("/").filter(Boolean);
      const file = parts.pop() || "";
      const dir = parts.join("/");
      return dir ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
    };

    const resolveWorksheetTablePart = (templateMap, sheetPart) => {
      const sheetDoc = parseXml(getEntryText(templateMap, sheetPart), sheetPart);
      const tablePartsNode = getFirstChild(sheetDoc.documentElement, "tableParts");
      const tablePartNode = getChildElements(tablePartsNode, "tablePart")[0];
      const relId = getAttributeByLocalName(tablePartNode, "id");
      if (!relId) return "";
      const relsPart = relsPartNameForPart(sheetPart);
      if (!templateMap.has(relsPart)) return "";
      const relsDoc = parseXml(getEntryText(templateMap, relsPart), relsPart);
      const relNode = getChildElements(relsDoc.documentElement, "Relationship")
        .find((rel) => rel.getAttribute("Id") === relId);
      if (!relNode) return "";
      const baseDir = `${normalizeZipPartName(sheetPart).split("/").slice(0, -1).join("/")}/`;
      return resolveZipPartTarget(baseDir, getAttributeByLocalName(relNode, "Target"));
    };

    const parseRangeRef = (ref) => {
      const raw = String(ref || "").trim();
      if (!raw) return null;
      const [startRef, endRef = startRef] = raw.split(":");
      const start = parseCellRef(startRef);
      const end = parseCellRef(endRef);
      if (!start || !end) return null;
      return { start, end };
    };

    const getMaxSheetRowNumber = (sheetDoc) => (
      getSheetRows(sheetDoc).reduce((max, row) => {
        const rowNo = Number(row.getAttribute("r"));
        return Number.isFinite(rowNo) ? Math.max(max, rowNo) : max;
      }, 1)
    );

    const toAbsoluteRangeRef = (range) => {
      const parsed = typeof range === "string" ? parseRangeRef(range) : range;
      if (!parsed) return "";
      const start = `$${indexToColLetter(parsed.start.col)}$${parsed.start.row}`;
      const end = `$${indexToColLetter(parsed.end.col)}$${parsed.end.row}`;
      return start === end ? start : `${start}:${end}`;
    };

    const quoteSheetRangeRef = (sheetName, range) => (
      `'${String(sheetName || "").replace(/'/g, "''")}'!${toAbsoluteRangeRef(range)}`
    );

    const quarterWindowFromDate = (date = new Date()) => {
      const currentQuarter = Math.floor(date.getMonth() / 3) + 1;
      let year = date.getFullYear();
      let quarter = currentQuarter + 1;
      if (quarter === 5) {
        quarter = 1;
        year += 1;
      }
      const total = Math.max(0, 4 - currentQuarter) + 8;
      const out = [];
      for (let i = 0; i < total; i += 1) {
        out.push({ year, quarter });
        quarter += 1;
        if (quarter === 5) {
          quarter = 1;
          year += 1;
        }
      }
      return out;
    };

    const actionPlanOutputStem = (locale) => {
      const base = getLocaleBase(locale);
      if (base === "fr") return "Plan-Action-Securite-Integree";
      if (base === "it") return "Piano-Azione-Sicurezza-Integrata";
      return "Aktionsplan-Integrierte-Sicherheit";
    };

    const resolveActionPlanTemplatePath = (locale) => {
      const base = getLocaleBase(locale);
      return ACTION_PLAN_TEMPLATE_BY_LOCALE[base] || ACTION_PLAN_TEMPLATE_BY_LOCALE.de;
    };

    const resolveActionPlanSheetNames = (locale) => {
      const base = getLocaleBase(locale);
      return ACTION_PLAN_SHEET_NAMES[base] || ACTION_PLAN_SHEET_NAMES.de;
    };

    const buildActionPlanReportRows = () => {
      const locale = String(state.project?.meta?.locale || "de-CH");
      const compareIds = typeof stateHelpers.compareIdSegments === "function"
        ? stateHelpers.compareIdSegments
        : ((a, b) => String(a || "").localeCompare(String(b || ""), "de", { numeric: true }));
      const chapters = [...(state.project?.chapters || [])].sort((a, b) => compareIds(a?.id, b?.id));
      const formatChapterLabel = typeof stateHelpers.formatChapterLabel === "function"
        ? stateHelpers.formatChapterLabel
        : (chapter) => String(chapter?.id || "");
      const out = [];

      chapters.forEach((chapter) => {
        const chapterId = String(chapter?.id || "").trim();
        if (!chapterId || chapterId === "0") return;
        const entries = typeof reportRows.buildChapterRows === "function"
          ? reportRows.buildChapterRows(chapter, { toText: stateHelpers.toText })
          : [];
        let currentSection = "";
        entries.forEach((entry) => {
          if (entry?.kind === "section") {
            currentSection = String(entry.title || "").trim();
            return;
          }
          if (entry?.kind !== "finding") return;
          out.push({
            reportRef: String(entry.id || "").trim(),
            chapter: formatChapterLabel(chapter, locale),
            theme: chapterId === "4.8"
              ? String(entry.title || currentSection || "").trim()
              : String(currentSection || "").trim(),
            finding: String(entry.finding || ""),
            recommendation: String(entry.recommendation || ""),
            priority: String(entry.priority || ""),
          });
        });
      });

      return out;
    };

    const updateQuarterHeaderMerges = (sheetDoc, quarterWindow) => {
      const mergesNode = getFirstChild(sheetDoc.documentElement, "mergeCells");
      if (!mergesNode) return;
      const quarterStart = ACTION_PLAN_PLAN_QUARTER_START_COL;
      const quarterEnd = ACTION_PLAN_PLAN_QUARTER_START_COL + ACTION_PLAN_TEMPLATE_MAX_QUARTERS - 1;
      Array.from(mergesNode.getElementsByTagNameNS("*", "mergeCell")).forEach((mergeCell) => {
        const ref = String(mergeCell.getAttribute("ref") || "");
        const parts = ref.split(":");
        const start = parseCellRef(parts[0]);
        const end = parseCellRef(parts[1] || parts[0]);
        if (!start || !end) return;
        const isQuarterYearMerge = start.row === ACTION_PLAN_PLAN_YEAR_ROW
          && end.row === ACTION_PLAN_PLAN_YEAR_ROW
          && start.col >= quarterStart
          && end.col <= quarterEnd;
        if (isQuarterYearMerge) mergesNode.removeChild(mergeCell);
      });

      let previousYear = null;
      let previousStart = 0;
      const groups = [];
      quarterWindow.forEach((item, index) => {
        if (previousYear == null) {
          previousYear = item.year;
          previousStart = index;
          return;
        }
        if (item.year !== previousYear) {
          groups.push({ year: previousYear, start: previousStart, end: index - 1 });
          previousYear = item.year;
          previousStart = index;
        }
      });
      if (previousYear != null) {
        groups.push({ year: previousYear, start: previousStart, end: quarterWindow.length - 1 });
      }

      groups.forEach((group) => {
        if (group.end <= group.start) return;
        const mergeCell = sheetDoc.createElementNS(XML_NS, "mergeCell");
        mergeCell.setAttribute(
          "ref",
          `${indexToColLetter(quarterStart + group.start)}${ACTION_PLAN_PLAN_YEAR_ROW}:${indexToColLetter(quarterStart + group.end)}${ACTION_PLAN_PLAN_YEAR_ROW}`,
        );
        mergesNode.appendChild(mergeCell);
      });
      mergesNode.setAttribute("count", String(mergesNode.getElementsByTagNameNS("*", "mergeCell").length));
    };

    const updateActionPlanTableHeaders = (templateMap, tablePart, quarterWindow) => {
      if (!tablePart || !templateMap.has(tablePart)) return;
      const tableDoc = parseXml(getEntryText(templateMap, tablePart), tablePart);
      const tableColumns = getFirstChild(tableDoc.documentElement, "tableColumns");
      if (!tableColumns) return;
      const cols = getChildElements(tableColumns, "tableColumn");
      for (let i = 0; i < ACTION_PLAN_TEMPLATE_MAX_QUARTERS; i += 1) {
        const col = cols[9 + i];
        if (!col) continue;
        const item = quarterWindow[i];
        col.setAttribute("name", item ? `${item.year}-Q${item.quarter}` : "");
      }
      setEntryText(templateMap, tablePart, serializeXml(tableDoc));
    };

    const updateWorkbookFilterRanges = (templateMap, refsBySheetName) => {
      const workbookPart = "xl/workbook.xml";
      if (!templateMap.has(workbookPart)) return;
      const workbookDoc = parseXml(getEntryText(templateMap, workbookPart), workbookPart);
      let definedNames = getFirstChild(workbookDoc.documentElement, "definedNames");
      if (!definedNames) {
        definedNames = workbookDoc.createElementNS(XML_NS, "definedNames");
        workbookDoc.documentElement.appendChild(definedNames);
      }
      const filterNames = getChildElements(definedNames, "definedName")
        .filter((node) => node.getAttribute("name") === "_xlnm._FilterDatabase");
      const sheetList = Array.from(getChildElements(getFirstChild(workbookDoc.documentElement, "sheets"), "sheet"));
      Object.entries(refsBySheetName || {}).forEach(([sheetName, ref]) => {
        if (!sheetName || !ref) return;
        const sheetIndex = sheetList.findIndex((sheet) => sheet.getAttribute("name") === sheetName);
        if (sheetIndex < 0) return;
        let targetNode = filterNames.find((node) => Number(node.getAttribute("localSheetId")) === sheetIndex);
        if (!targetNode) {
          targetNode = workbookDoc.createElementNS(XML_NS, "definedName");
          targetNode.setAttribute("name", "_xlnm._FilterDatabase");
          targetNode.setAttribute("localSheetId", String(sheetIndex));
          targetNode.setAttribute("hidden", "1");
          definedNames.appendChild(targetNode);
          filterNames.push(targetNode);
        }
        targetNode.textContent = quoteSheetRangeRef(sheetName, ref);
      });
      setEntryText(templateMap, workbookPart, ensureWorkbookNamespaceAliases(serializeXml(workbookDoc)));
    };

    const buildActionPlanLookupFormula = (planSheetName, rowNumber) => (
      `IF($H${rowNumber}="","",IFERROR(INDEX('${String(planSheetName || "").replace(/'/g, "''")}'!$B:$B,MATCH($H${rowNumber},'${String(planSheetName || "").replace(/'/g, "''")}'!$A:$A,0)),""))`
    );

    const updateReportTableRange = (templateMap, tablePart, tableRange, lastDataRow, planSheetName) => {
      if (!tablePart || !templateMap.has(tablePart) || !tableRange) return;
      const tableDoc = parseXml(getEntryText(templateMap, tablePart), tablePart);
      const endColLetter = indexToColLetter(tableRange.end.col);
      const ref = `${indexToColLetter(tableRange.start.col)}${tableRange.start.row}:${endColLetter}${Number(lastDataRow)}`;
      tableDoc.documentElement.setAttribute("ref", ref);
      const autoFilter = getFirstChild(tableDoc.documentElement, "autoFilter");
      if (autoFilter) autoFilter.setAttribute("ref", ref);
      const tableColumns = getFirstChild(tableDoc.documentElement, "tableColumns");
      const formulaColumn = getChildElements(tableColumns, "tableColumn")[8];
      if (formulaColumn) {
        let formulaNode = getFirstChild(formulaColumn, "calculatedColumnFormula");
        if (!formulaNode) {
          formulaNode = tableDoc.createElementNS(XML_NS, "calculatedColumnFormula");
          formulaColumn.appendChild(formulaNode);
        }
        formulaNode.textContent = buildActionPlanLookupFormula(planSheetName, tableRange.start.row + 1);
      }
      setEntryText(templateMap, tablePart, serializeXml(tableDoc));
    };

    const patchActionPlanWorkbook = (templateMap, sourceRows, locale) => {
      const sheetNames = resolveActionPlanSheetNames(locale);
      const sheetParts = resolveSheetParts(templateMap);
      const reportPart = sheetParts.get(sheetNames.report);
      const planPart = sheetParts.get(sheetNames.plan);
      if (!reportPart) throw new Error(`Template sheet '${sheetNames.report}' was not found.`);
      if (!planPart) throw new Error(`Template sheet '${sheetNames.plan}' was not found.`);

      const reportDoc = parseXml(getEntryText(templateMap, reportPart), reportPart);
      const planDoc = parseXml(getEntryText(templateMap, planPart), planPart);
      const planTablePart = resolveWorksheetTablePart(templateMap, planPart);
      const reportTablePart = resolveWorksheetTablePart(templateMap, reportPart);
      if (!planTablePart) throw new Error(`Template sheet '${sheetNames.plan}' is missing its table definition.`);
      if (!reportTablePart) throw new Error(`Template sheet '${sheetNames.report}' is missing its table definition.`);
      const planTableDoc = parseXml(getEntryText(templateMap, planTablePart), planTablePart);
      const planTableRange = parseRangeRef(planTableDoc.documentElement.getAttribute("ref"));
      const reportTableDoc = parseXml(getEntryText(templateMap, reportTablePart), reportTablePart);
      const reportTableRange = parseRangeRef(reportTableDoc.documentElement.getAttribute("ref"));
      if (!planTableRange) throw new Error(`Template sheet '${sheetNames.plan}' has an invalid table range.`);
      if (!reportTableRange) throw new Error(`Template sheet '${sheetNames.report}' has an invalid table range.`);
      const reportHeaderRow = Number(reportTableRange.start.row);
      const reportFirstDataRow = reportHeaderRow + 1;
      const currentLastReportRow = Number(reportTableRange.end.row);
      const reportRowCount = Math.max(sourceRows.length, 1);
      const lastReportRow = reportHeaderRow + reportRowCount;

      ensureReportRows(reportDoc, reportFirstDataRow, currentLastReportRow, lastReportRow);
      updateSheetDimension(reportDoc, `A1:${indexToColLetter(reportTableRange.end.col)}${getMaxSheetRowNumber(reportDoc)}`);

      const formulaStyleId = String(ensureSheetCell(reportDoc, reportFirstDataRow, "I").getAttribute("s") || "");
      for (let rowNo = reportFirstDataRow; rowNo <= lastReportRow; rowNo += 1) {
        const index = rowNo - reportFirstDataRow;
        const source = sourceRows[index] || {
          reportRef: "",
          chapter: "",
          theme: "",
          finding: "",
          recommendation: "",
          priority: "",
        };
        setInlineStringCellValue(reportDoc, rowNo, "A", source.reportRef);
        setInlineStringCellValue(reportDoc, rowNo, "B", source.chapter);
        setInlineStringCellValue(reportDoc, rowNo, "C", source.theme);
        setInlineStringCellValue(reportDoc, rowNo, "D", source.finding);
        setInlineStringCellValue(reportDoc, rowNo, "E", source.recommendation);
        setInlineStringCellValue(reportDoc, rowNo, "F", source.priority);
        setInlineStringCellValue(reportDoc, rowNo, "G", "");
        setInlineStringCellValue(reportDoc, rowNo, "H", "");
        setFormulaCellValue(
          reportDoc,
          rowNo,
          "I",
          buildActionPlanLookupFormula(sheetNames.plan, rowNo),
          formulaStyleId,
        );
        setInlineStringCellValue(reportDoc, rowNo, "J", "");
      }
      updateReportTableRange(templateMap, reportTablePart, reportTableRange, lastReportRow, sheetNames.plan);

      const quarterWindow = quarterWindowFromDate(new Date());
      for (let i = 0; i < ACTION_PLAN_TEMPLATE_MAX_QUARTERS; i += 1) {
        const item = quarterWindow[i];
        const colLetter = indexToColLetter(ACTION_PLAN_PLAN_QUARTER_START_COL + i);
        setInlineStringCellValue(planDoc, ACTION_PLAN_PLAN_YEAR_ROW, colLetter, item ? String(item.year) : "");
        setInlineStringCellValue(planDoc, ACTION_PLAN_PLAN_QUARTER_ROW, colLetter, item ? `${item.year}-Q${item.quarter}` : "");
      }
      updateQuarterHeaderMerges(planDoc, quarterWindow);
      updateActionPlanTableHeaders(templateMap, planTablePart, quarterWindow);
      updateWorkbookFilterRanges(templateMap, {
        [sheetNames.plan]: planTableRange,
        [sheetNames.report]: {
          start: reportTableRange.start,
          end: { col: reportTableRange.end.col, row: lastReportRow },
        },
      });

      setEntryText(templateMap, reportPart, serializeXml(reportDoc));
      setEntryText(templateMap, planPart, serializeXml(planDoc));
    };

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

    const exportActionPlanExcel = async () => {
      if (!runtime.dirHandle) {
        setStatus("Open project folder first.");
        return null;
      }
      if (typeof unzipAllEntries !== "function" || typeof buildZipStore !== "function") {
        throw new Error("ZIP workbook helpers are unavailable.");
      }

      const locale = String(state.project?.meta?.locale || "de-CH");
      const templatePath = resolveActionPlanTemplatePath(locale);
      const templateFile = await resolveProjectOrBundledTemplateFile(
        runtime.dirHandle,
        templatePath,
        "Action plan",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );

      const sourceRows = buildActionPlanReportRows();
      if (!sourceRows.length) {
        throw new Error("No included report items available for action plan export.");
      }

      const entries = await unzipAllEntries(await templateFile.arrayBuffer());
      const templateMap = new Map(entries.map((entry) => [normalizeZipPartName(entry.name), entry]));
      patchActionPlanWorkbook(templateMap, sourceRows, locale);

      const outputBytes = buildZipStore(Array.from(templateMap.values()));
      const outputsDir = await getOutputsDirectory(runtime.dirHandle);
      const stamp = new Date().toISOString().slice(0, 10);
      const companySlug = toFileSafeSlug(
        state.project?.meta?.company || state.project?.meta?.projectName || "",
        "Company",
      );
      const outName = `${stamp}-${companySlug}-${actionPlanOutputStem(locale)}.xlsx`;
      await writeBinaryFile(outputsDir, outName, outputBytes);
      debug.logLine("info", `Action plan exported: outputs/${outName}`);
      return { savedAs: `outputs/${outName}` };
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
      const existingMeta = isPlainObject(state.project?.meta) ? structuredClone(state.project.meta) : {};
      let project = seeds.buildProjectFromKnowledgeBase(seedBase);
      project.meta = {
        ...(isPlainObject(project?.meta) ? project.meta : {}),
        ...existingMeta,
        locale: requestedLocale,
        createdAt: existingMeta.createdAt || project?.meta?.createdAt || new Date().toISOString(),
      };
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
      exportActionPlanExcel,
      bootstrapProjectFromSeed,
    };
  };

  window.AutoBerichtSidecar = { init };
})();
