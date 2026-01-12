(() => {
  const stateHelpers = window.AutoBerichtState || {};
  const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => 0);

  const ensureWorkstateDefaults = (row) => {
    if (!row.workstate) row.workstate = {};
    const ws = row.workstate;
    if (ws.selectedLevel == null) ws.selectedLevel = 1;
    if (ws.includeFinding == null) ws.includeFinding = true;
    if (ws.includeRecommendation == null) ws.includeRecommendation = true;
    if (ws.done == null) ws.done = false;
    if (!ws.useFindingOverride) ws.useFindingOverride = false;
    if (!ws.findingOverride) ws.findingOverride = "";
    ws.useLevelOverride = ws.useLevelOverride || { "1": false, "2": false, "3": false, "4": false };
    ws.levelOverrides = ws.levelOverrides || { "1": "", "2": "", "3": "", "4": "" };
    ws.libraryActions = ws.libraryActions || { "1": "off", "2": "off", "3": "off", "4": "off" };
    ws.libraryHashes = ws.libraryHashes || { "1": "", "2": "", "3": "", "4": "" };
  };

  const ensureProjectMeta = (project, setLocale) => {
    if (!project.meta) project.meta = {};
    if (!project.meta.locale) project.meta.locale = "de-CH";
    if (!project.meta.company) project.meta.company = "";
    if (!project.meta.companyId) project.meta.companyId = "";
    if (!project.meta.author) project.meta.author = "";
    if (!project.meta.initials) project.meta.initials = "";
    if (setLocale) setLocale(project.meta.locale);
  };

  const ensureObservationChapter = (project) => {
    if (!project?.chapters) return project;
    const hasObservation = project.chapters.some((chapter) => chapter.id === "4.8");
    if (!hasObservation) {
      project.chapters.push({
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [],
      });
    }
    project.chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    return project;
  };

  const ensureObservationRowsFromTags = (project, tagOptions) => {
    if (!project?.chapters) return project;
    const chapter = project.chapters.find((c) => c.id === "4.8");
    if (!chapter) return project;
    const existingIds = new Set((chapter.rows || []).map((r) => String(r.id || "")));
    const obsTags = tagOptions?.observations || [];
    const newRows = [];
    obsTags.forEach((tag, index) => {
      const rowId = `4.8.${index + 1}`;
      if (existingIds.has(rowId)) return;
      newRows.push({
        id: rowId,
        type: "field_observation",
        sectionId: "4.8",
        sectionLabel: "4.8 Beobachtungen",
        titleOverride: tag.label || tag.value || rowId,
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
      });
    });
    if (newRows.length) {
      chapter.rows = [...(chapter.rows || []), ...newRows];
      chapter.rows.sort((a, b) => compareIdSegments(a.id, b.id));
    }
    return project;
  };

  const ensureManagementSummaryChapter = (project) => {
    if (!project?.chapters) return project;
    const hasSummary = project.chapters.some((chapter) => chapter.id === "0");
    if (!hasSummary) {
      project.chapters.unshift({
        id: "0",
        title: { de: "Management Summary" },
        rows: Array.from({ length: 8 }).map((_, index) => ({
          id: `0.${index + 1}`,
          type: "summary",
          titleOverride: "",
          master: {
            finding: "",
            levels: {
              "1": "",
              "2": "",
              "3": "",
              "4": "",
            },
          },
          customer: {
            answer: null,
            remark: "",
            items: [],
          },
          workstate: {
            selectedLevel: 1,
            includeFinding: true,
            includeRecommendation: true,
            done: false,
            useFindingOverride: true,
            findingOverride: "",
            useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
            levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
          },
        })),
      });
    }
    project.chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    return project;
  };

  const normalizeProject = (project, setLocale) => {
    if (!project) return project;
    ensureObservationChapter(project);
    ensureManagementSummaryChapter(project);
    ensureProjectMeta(project, setLocale);
    (project.chapters || []).forEach((chapter) => {
      (chapter.rows || []).forEach((row) => {
        if (row.kind === "section") return;
        ensureWorkstateDefaults(row);
      });
    });
    return project;
  };

  const slugifyTag = (value) => String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/^-+|-+$/g, "") || "tag";

  const getObservationTagsFromSidecar = (doc) => {
    const options = doc?.photos?.photoTagOptions?.observations || [];
    return options
      .map((opt) => (typeof opt === "string" ? opt : opt.value || opt.label))
      .map((val) => String(val || "").trim())
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, "de", { numeric: true }));
  };

  const buildObservationRow = (tag) => ({
    id: `4.8:${slugifyTag(tag)}`,
    type: "field_observation",
    tag,
    titleOverride: tag,
    master: {
      finding: "Es wurden folgende unsichere Situationen beobachtet.",
      levels: {
        "1": "",
        "2": "",
        "3": "",
        "4": "",
      },
    },
    customer: {
      answer: null,
      remark: "",
      items: [],
    },
    workstate: {
      selectedLevel: 1,
      includeFinding: true,
      includeRecommendation: true,
      done: false,
      useFindingOverride: true,
      findingOverride: "Es wurden folgende unsichere Situationen beobachtet.",
      useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
      levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
    },
  });

  const syncObservationChapterRows = (project, doc) => {
    if (!project?.chapters) return;
    const chapter = project.chapters.find((item) => item.id === "4.8");
    if (!chapter) return;
    const tags = getObservationTagsFromSidecar(doc);
    if (!tags.length) return;

    const existingRows = (chapter.rows || []).filter((row) => row.kind !== "section");
    const byTag = new Map();
    existingRows.forEach((row) => {
      const tag = row.tag || row.titleOverride;
      if (tag) {
        row.tag = tag;
        row.titleOverride = tag;
        byTag.set(tag, row);
      }
    });

    const nextRows = tags.map((tag) => {
      const existing = byTag.get(tag);
      if (existing) {
        existing.titleOverride = tag;
        return existing;
      }
      return buildObservationRow(tag);
    });

    if (!chapter.meta) chapter.meta = {};
    if (!Array.isArray(chapter.meta.order)) {
      chapter.meta.order = nextRows.map((row) => row.id);
    } else {
      const order = chapter.meta.order.filter((id) => nextRows.some((row) => row.id === id));
      nextRows.forEach((row) => {
        if (!order.includes(row.id)) order.push(row.id);
      });
      chapter.meta.order = order;
    }

    chapter.rows = nextRows;
  };

  const orderObservationRows = (chapter) => {
    const rows = (chapter.rows || []).filter((row) => row.kind !== "section");
    const order = chapter.meta?.order;
    if (!order || !order.length) return rows;
    const byId = new Map(rows.map((row) => [row.id, row]));
    const ordered = [];
    order.forEach((id) => {
      const match = byId.get(id);
      if (match) ordered.push(match);
    });
    rows.forEach((row) => {
      if (!ordered.includes(row)) ordered.push(row);
    });
    return ordered;
  };

  const moveObservationRow = (chapter, rowId, delta) => {
    if (!chapter.meta || !Array.isArray(chapter.meta.order)) return false;
    const order = [...chapter.meta.order];
    const index = order.indexOf(rowId);
    if (index === -1) return false;
    const next = index + delta;
    if (next < 0 || next >= order.length) return false;
    [order[index], order[next]] = [order[next], order[index]];
    chapter.meta.order = order;
    return true;
  };

  window.AutoBerichtNormalize = {
    ensureWorkstateDefaults,
    ensureProjectMeta,
    ensureObservationChapter,
    ensureManagementSummaryChapter,
    normalizeProject,
    slugifyTag,
    getObservationTagsFromSidecar,
    buildObservationRow,
    syncObservationChapterRows,
    orderObservationRows,
    moveObservationRow,
  };
})();
