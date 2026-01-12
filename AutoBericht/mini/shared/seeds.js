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

  const normalizeTagGroups = (tags) => {
    if (!tags || typeof tags !== "object") {
      return { report: [], observations: [], training: [] };
    }
    return {
      report: tags.report || [],
      observations: tags.observations || [],
      training: tags.training || [],
    };
  };

  const validateKnowledgeBase = (knowledgeBase) => {
    if (!knowledgeBase || typeof knowledgeBase !== "object") {
      throw new Error("Knowledge base missing or invalid.");
    }
    if (!knowledgeBase.schemaVersion) {
      throw new Error("Knowledge base missing schemaVersion.");
    }
    if (!knowledgeBase.structure || !Array.isArray(knowledgeBase.structure.items)) {
      throw new Error("Knowledge base structure is missing items.");
    }
    if (!knowledgeBase.library || !Array.isArray(knowledgeBase.library.entries)) {
      throw new Error("Knowledge base library entries missing.");
    }
    if (!knowledgeBase.tags || typeof knowledgeBase.tags !== "object") {
      throw new Error("Knowledge base tags missing.");
    }
    return knowledgeBase;
  };

  const buildLibraryMap = (knowledgeBase) => {
    const entries = knowledgeBase?.library?.entries || [];
    return new Map(entries.map((entry) => [entry.id, entry]));
  };

  const buildProjectFromKnowledgeBase = (knowledgeBase) => {
    validateKnowledgeBase(knowledgeBase);
    const libraryMap = buildLibraryMap(knowledgeBase);
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

    // Inject 4.8 Beobachtungen rows from tag list if none exist
    const obsTags = knowledgeBase?.tags?.observations || [];
    const obsChapterExisting = chapters.find((c) => c.id === "4.8");
    const needObsRows = !obsChapterExisting || (obsChapterExisting.rows || []).length === 0;
    if (needObsRows && obsTags.length) {
      const obsRows = obsTags.map((tag, idx) => ({
        id: `4.8.${idx + 1}`,
        type: "field_observation",
        sectionId: "4.8",
        sectionLabel: "4.8 Beobachtungen",
        titleOverride: tag.label || tag.value || `4.8.${idx + 1}`,
        master: null,
        customer: { answer: null, remark: "", items: [] },
        workstate: {
          selectedLevel: 1,
          includeFinding: true,
          includeRecommendation: true,
          done: false,
          useFindingOverride: true,
          findingOverride: tag.label || tag.value || "Beobachtung",
          useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
          levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
        },
      }));
      const obsChapter = obsChapterExisting || {
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [],
      };
      obsChapter.rows = obsRows;
      if (!obsChapterExisting) {
        chapters.push(obsChapter);
      }
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
    let knowledgeBase = null;
    if (dirHandle) {
      knowledgeBase = await readJsonIfExists(dirHandle, [getLibraryFileName()]);
      if (!knowledgeBase) {
        knowledgeBase = await readSeedFromProject(dirHandle, seedFilename);
      }
    }
    if (!knowledgeBase) {
      knowledgeBase = await readSeedFromHttp(seedFilename);
    }
    if (!knowledgeBase) throw new Error("Knowledge base seed not found.");
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
    normalizeTagGroups,
    validateKnowledgeBase,
    buildLibraryMap,
    buildProjectFromKnowledgeBase,
    loadSeedsForProject,
  };
})();
