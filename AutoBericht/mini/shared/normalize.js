(() => {
  const stateHelpers = window.AutoBerichtState || {};
  const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => 0);
  const toText = stateHelpers.toText || ((value) => {
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  });
  const OBS_FINDING_TEXT = "Folgende unsichere Situationen wurden beobachtet:";

  const ensureWorkstateDefaults = (row) => {
    if (!row.workstate) row.workstate = {};
    const ws = row.workstate;
    if (ws.selectedLevel == null) ws.selectedLevel = 1;
    if (ws.includeFinding == null) ws.includeFinding = false;
    if (ws.includeRecommendation == null) ws.includeRecommendation = true;
    if (ws.done == null) ws.done = false;
    if (ws.findingText == null) {
      if (row.type === "field_observation") {
        ws.findingText = OBS_FINDING_TEXT;
      } else {
        ws.findingText = toText(row.master?.finding);
      }
    }
    if (ws.recommendationText == null) {
      ws.recommendationText = toText(row.master?.recommendation);
    }
    if (!ws.libraryAction) ws.libraryAction = "off";
    if (ws.libraryHash == null) ws.libraryHash = "";
  };

  const ensureProjectMeta = (project, setLocale) => {
    if (!project.meta) project.meta = {};
    if (!project.meta.locale) project.meta.locale = "de-CH";
    if (!project.meta.company) project.meta.company = "";
    if (!project.meta.companyId) project.meta.companyId = "";
    if (!project.meta.moderator && project.meta.author) project.meta.moderator = project.meta.author;
    if (!project.meta.moderator) project.meta.moderator = "";
    if (!project.meta.moderatorInitials && project.meta.initials) project.meta.moderatorInitials = project.meta.initials;
    if (!project.meta.moderatorInitials) project.meta.moderatorInitials = "";
    if (!project.meta.coModerator) project.meta.coModerator = "";
    if (!project.meta.coModeratorInitials) project.meta.coModeratorInitials = "";
    if (project.meta.autobackupMinutes == null) project.meta.autobackupMinutes = 30;
    // keep legacy fields aligned for backward compatibility
    project.meta.author = project.meta.moderator;
    project.meta.initials = project.meta.moderatorInitials;
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
          includeFinding: false,
          includeRecommendation: true,
          done: false,
          findingText: OBS_FINDING_TEXT,
          recommendationText: "",
          libraryAction: "off",
          libraryHash: "",
        },
      });
    });
    if (newRows.length) {
      chapter.rows = [...(chapter.rows || []), ...newRows];
      chapter.rows.sort((a, b) => compareIdSegments(a.id, b.id));
    }
    if (!chapter.meta) chapter.meta = {};
    if (!Array.isArray(chapter.meta.order)) {
      chapter.meta.order = (chapter.rows || [])
        .filter((row) => row.kind !== "section")
        .map((row) => row.id);
    } else {
      const order = chapter.meta.order.filter((id) => (chapter.rows || []).some((row) => row.id === id));
      (chapter.rows || []).forEach((row) => {
        if (row.kind === "section") return;
        if (!order.includes(row.id)) order.push(row.id);
      });
      chapter.meta.order = order;
    }
    return project;
  };

  const ensureManagementSummaryChapter = (project) => {
    if (!project?.chapters) return project;
    const hasSummary = project.chapters.some((chapter) => chapter.id === "0");
    if (!hasSummary) {
      const rows = Array.from({ length: 8 }).map((_, index) => ({
        id: `0.${index + 1}`,
        type: "summary",
        titleOverride: "",
        master: {
          finding: "",
          recommendation: "",
        },
        customer: {
          answer: null,
          remark: "",
          items: [],
        },
        workstate: {
          selectedLevel: 1,
          includeFinding: false,
          includeRecommendation: true,
          done: false,
          findingText: "",
          recommendationText: "",
          libraryAction: "off",
          libraryHash: "",
        },
      }));
      project.chapters.unshift({
        id: "0",
        title: { de: "Management Summary" },
        rows,
        meta: {
          order: rows.map((row) => row.id),
        },
      });
    }
    const summaryChapter = project.chapters.find((chapter) => chapter.id === "0");
    if (summaryChapter) {
      summaryChapter.meta = summaryChapter.meta || {};
      if (!Array.isArray(summaryChapter.meta.order)) {
        summaryChapter.meta.order = (summaryChapter.rows || [])
          .filter((row) => row.kind !== "section")
          .map((row) => row.id);
      } else {
        const order = summaryChapter.meta.order
          .filter((id) => (summaryChapter.rows || []).some((row) => row.id === id));
        (summaryChapter.rows || []).forEach((row) => {
          if (row.kind === "section") return;
          if (!order.includes(row.id)) order.push(row.id);
        });
        summaryChapter.meta.order = order;
      }
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

  const buildObservationRow = (tag, idOverride = "") => ({
    id: idOverride || `4.8:${slugifyTag(tag)}`,
    type: "field_observation",
    tag,
    titleOverride: tag,
    master: {
      finding: OBS_FINDING_TEXT,
      recommendation: "",
    },
    customer: {
      answer: null,
      remark: "",
      items: [],
    },
    workstate: {
      selectedLevel: 1,
      includeFinding: false,
      includeRecommendation: true,
      done: false,
      findingText: OBS_FINDING_TEXT,
      recommendationText: "",
      libraryAction: "off",
      libraryHash: "",
    },
  });

  const getNextObservationIndex = (rows = []) => {
    let max = 0;
    rows.forEach((row) => {
      const match = /^4\.8\.(\d+)$/.exec(String(row?.id || ""));
      if (!match) return;
      const value = Number(match[1]);
      if (Number.isFinite(value)) max = Math.max(max, value);
    });
    return max + 1;
  };

  const syncObservationChapterRows = (project, doc) => {
    if (!project?.chapters) return;
    const chapter = project.chapters.find((item) => item.id === "4.8");
    if (!chapter) return;
    const tags = getObservationTagsFromSidecar(doc);
    if (!tags.length) return;

    const existingRows = (chapter.rows || []).filter((row) => row.kind !== "section");
    let nextIndex = getNextObservationIndex(existingRows);
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
      const rowId = `4.8.${nextIndex}`;
      nextIndex += 1;
      return buildObservationRow(tag, rowId);
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
