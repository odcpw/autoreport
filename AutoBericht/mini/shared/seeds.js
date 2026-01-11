(() => {
  const stateHelpers = window.AutoBerichtState || {};
  const normalizeHelpers = window.AutoBerichtNormalize || {};
  const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => 0);

  const resolveLocaleKey = (locale) => {
    const base = String(locale || "de-CH").toLowerCase();
    if (base.startsWith("fr")) return "fr";
    if (base.startsWith("it")) return "it";
    return "de";
  };

  const getKnowledgeBaseFilename = (locale) => `knowledge_base_${resolveLocaleKey(locale)}.json`;

  const seedPathCandidates = (filename) => ([
    ["AutoBericht", "data", "seed", filename],
  ]);

  const seedHttpCandidates = (filename) => {
    if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
      return [];
    }
    return [new URL(`../data/seed/${filename}`, window.location.href).toString()];
  };

  const readJsonFromPath = async (rootHandle, pathParts) => {
    let currentHandle = rootHandle;
    for (let i = 0; i < pathParts.length - 1; i += 1) {
      currentHandle = await currentHandle.getDirectoryHandle(pathParts[i]);
    }
    const fileHandle = await currentHandle.getFileHandle(pathParts[pathParts.length - 1]);
    const file = await fileHandle.getFile();
    const text = await file.text();
    return JSON.parse(text);
  };

  const readJsonIfExists = async (rootHandle, pathParts) => {
    try {
      return await readJsonFromPath(rootHandle, pathParts);
    } catch (err) {
      return null;
    }
  };

  const readJsonFromHttp = async (url) => {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Failed to load ${url}`);
    }
    return response.json();
  };

  const readSeedFromProject = async (rootHandle, filename) => {
    if (!rootHandle) return null;
    for (const parts of seedPathCandidates(filename)) {
      const data = await readJsonIfExists(rootHandle, parts);
      if (data) return data;
    }
    return null;
  };

  const readSeedFromHttp = async (filename) => {
    const candidates = seedHttpCandidates(filename);
    for (const url of candidates) {
      try {
        return await readJsonFromHttp(url);
      } catch (err) {
        // try next candidate
      }
    }
    return null;
  };

  const normalizeLibraryEntries = (libraryData) => {
    if (!libraryData) return [];
    if (Array.isArray(libraryData.library?.entries)) return libraryData.library.entries;
    if (Array.isArray(libraryData.entries)) return libraryData.entries;
    if (libraryData.entries && typeof libraryData.entries === "object") {
      return Object.entries(libraryData.entries).map(([id, levels]) => ({ id, levels }));
    }
    return [];
  };

  const normalizeTagGroups = (tags) => {
    if (!tags || typeof tags !== "object") {
      return { report: [], observations: [], training: [] };
    }
    return {
      report: tags.report || tags.bericht || [],
      observations: tags.observations || tags.topic || [],
      training: tags.training || tags.seminar || [],
    };
  };

  const buildLibraryMap = (masterData, userData) => {
    const masterEntries = normalizeLibraryEntries(masterData);
    const userEntries = normalizeLibraryEntries(userData);
    const map = new Map(masterEntries.map((entry) => [entry.id, entry]));
    userEntries.forEach((entry) => {
      const existing = map.get(entry.id) || { id: entry.id, levels: {}, finding: "" };
      const mergedLevels = { ...(existing.levels || {}) };
      Object.entries(entry.levels || {}).forEach(([key, value]) => {
        const text = stateHelpers.toText(value).trim();
        if (text) mergedLevels[key] = text;
      });
      const merged = {
        ...existing,
        levels: mergedLevels,
      };
      if (entry.finding) merged.finding = entry.finding;
      if (entry.lastUsed) merged.lastUsed = entry.lastUsed;
      if (!merged.lastUsed && existing.lastUsed) merged.lastUsed = existing.lastUsed;
      map.set(entry.id, merged);
    });
    return map;
  };

  const normalizeKnowledgeBase = (knowledge, fallback) => {
    const base = fallback && typeof fallback === "object"
      ? structuredClone(fallback)
      : {
        schemaVersion: "1.0",
        meta: { locale: "de-CH", updatedAt: new Date().toISOString(), source: "generated" },
        structure: { items: [] },
        library: { entries: [] },
        tags: { report: [], observations: [], training: [] },
      };

    if (!knowledge || typeof knowledge !== "object") return base;

    const looksNew = knowledge.schemaVersion && knowledge.structure && knowledge.library;
    if (looksNew) {
      const normalized = structuredClone(knowledge);
      normalized.schemaVersion = normalized.schemaVersion || "1.0";
      normalized.meta = normalized.meta || {};
      normalized.structure = normalized.structure || { items: [] };
      normalized.structure.items = normalized.structure.items || [];
      normalized.library = normalized.library || { entries: [] };
      normalized.library.entries = normalized.library.entries || [];
      normalized.tags = normalizeTagGroups(normalized.tags);
      return normalized;
    }

    const legacyEntries = normalizeLibraryEntries(knowledge);
    const merged = structuredClone(base);
    if (legacyEntries.length) {
      const mergedMap = buildLibraryMap(base, { entries: legacyEntries });
      merged.library.entries = Array.from(mergedMap.values());
    }
    if (knowledge.meta) merged.meta = { ...merged.meta, ...knowledge.meta };
    return merged;
  };

  const buildProjectFromKnowledgeBase = (knowledgeBase) => {
    const libraryMap = buildLibraryMap(knowledgeBase, null);
    const groups = new Map();

    (knowledgeBase?.structure?.items || []).forEach((item) => {
      const collapsedId = item.collapsedId || item.groupId || item.id;
      if (!collapsedId) return;
      const rawId = String(collapsedId);
      let chapterId = rawId.split(".")[0];
      if (rawId === "4.8" || rawId.startsWith("4.8.")) {
        chapterId = "4.8";
      }
      const group = groups.get(collapsedId) || {
        id: collapsedId,
        chapterId,
        items: [],
        chapterLabel: item.chapterLabel || "",
      };
      group.items.push(item);
      if (!group.chapterLabel && item.chapterLabel) {
        group.chapterLabel = item.chapterLabel;
      }
      groups.set(collapsedId, group);
    });

    const chaptersById = new Map();
    groups.forEach((group) => {
      const chapterId = group.chapterId || "0";
      const chapter = chaptersById.get(chapterId) || {
        id: chapterId,
        title: {
          de: group.chapterLabel || (chapterId === "4.8" ? "Beobachtungen" : `Kapitel ${chapterId}`),
        },
        rows: [],
      };
      const master = libraryMap.get(group.id) || null;
      const sectionLabel = group.items.find((item) => item.sectionLabel)?.sectionLabel || "";
      const sectionId = String(group.id).split(".").slice(0, 2).join(".");
      chapter.rows.push({
        id: group.id,
        type: chapterId === "4.8" ? "field_observation" : "standard",
        sectionId,
        sectionLabel,
        titleOverride: group.items[0]?.question || "",
        master,
        customer: {
          answer: null,
          remark: "",
          items: group.items,
        },
        workstate: {
          selectedLevel: 1,
          includeFinding: true,
          includeRecommendation: true,
          done: false,
          useFindingOverride: false,
          findingOverride: "",
          useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
          levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
        },
      });
      chaptersById.set(chapterId, chapter);
    });

    const chapters = Array.from(chaptersById.values()).sort((a, b) => compareIdSegments(a.id, b.id));
    chapters.forEach((chapter) => {
      chapter.rows.sort((a, b) => compareIdSegments(a.id, b.id));
      const withSections = [];
      let lastSection = "";
      chapter.rows.forEach((row) => {
        if (row.sectionLabel && row.sectionId !== lastSection) {
          withSections.push({
            kind: "section",
            id: row.sectionId,
            title: row.sectionLabel,
          });
          lastSection = row.sectionId;
        }
        withSections.push(row);
      });
      chapter.rows = withSections;
    });

    if (!chaptersById.has("4.8")) {
      chapters.push({
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [],
      });
      chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    }

    return {
      meta: {
        projectId: "seed-import",
        company: "",
        locale: knowledgeBase?.meta?.locale || "de-CH",
        author: "",
        createdAt: new Date().toISOString(),
      },
      chapters,
    };
  };

  const loadSeedsForProject = async ({ dirHandle, getLibraryFileName, locale }) => {
    const resolvedLocale = locale || stateHelpers.defaultProject?.meta?.locale || "de-CH";
    const seedFilename = getKnowledgeBaseFilename(resolvedLocale);
    let seedKnowledge = null;
    let userKnowledge = null;
    if (dirHandle) {
      seedKnowledge = await readSeedFromProject(dirHandle, seedFilename);
      userKnowledge = await readJsonIfExists(dirHandle, [getLibraryFileName()]);
    }
    if (!seedKnowledge) {
      seedKnowledge = await readSeedFromHttp(seedFilename);
    }
    if (!seedKnowledge) throw new Error("Knowledge base seed not found.");
    const knowledgeBase = normalizeKnowledgeBase(userKnowledge, seedKnowledge);
    const project = buildProjectFromKnowledgeBase(knowledgeBase);
    return normalizeHelpers.ensureObservationChapter(project);
  };

  window.AutoBerichtSeeds = {
    resolveLocaleKey,
    getKnowledgeBaseFilename,
    seedPathCandidates,
    seedHttpCandidates,
    readJsonFromPath,
    readJsonIfExists,
    readJsonFromHttp,
    readSeedFromProject,
    readSeedFromHttp,
    normalizeLibraryEntries,
    normalizeTagGroups,
    buildLibraryMap,
    normalizeKnowledgeBase,
    buildProjectFromKnowledgeBase,
    loadSeedsForProject,
  };
})();
