(() => {
  const stateHelpers = window.AutoBerichtState || {};
  const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => 0);
  const toText = stateHelpers.toText || ((value) => {
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  });
  const getLocaleBase = (locale) => {
    const base = String(locale || "de-CH").toLowerCase().split("-")[0];
    return ["de", "fr", "it"].includes(base) ? base : "de";
  };
  const OBS_CHAPTER_TITLE = {
    de: "Beobachtungen",
    fr: "Observations",
    it: "Osservazioni",
  };
  const OBS_FINDING_TEXT = {
    de: "Folgende unsichere Situationen wurden beobachtet:",
    fr: "Les situations dangereuses suivantes ont ete observees:",
    it: "Sono state osservate le seguenti situazioni pericolose:",
  };
  const getObservationChapterTitle = (locale) => OBS_CHAPTER_TITLE[getLocaleBase(locale)] || OBS_CHAPTER_TITLE.de;
  const getObservationSectionLabel = (locale) => `4.8 ${getObservationChapterTitle(locale)}`;
  const getObservationFindingText = (locale) => OBS_FINDING_TEXT[getLocaleBase(locale)] || OBS_FINDING_TEXT.de;
  const OBS_DEFAULT_SECTION_LABELS = new Set(Object.values(OBS_CHAPTER_TITLE).map((title) => `4.8 ${title}`));
  const OBS_DEFAULT_FINDING_TEXTS = new Set(Object.values(OBS_FINDING_TEXT));
  const clampSelectedLevel = (value) => {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) return 1;
    return Math.max(1, Math.min(4, Math.round(parsed)));
  };

  const getSelfAnswerState = (row) => {
    const direct = row?.customer?.answer;
    if (direct === 1 || direct === "1" || direct === true) return 1;
    if (direct === 0 || direct === "0" || direct === false) return 0;
    const items = Array.isArray(row?.customer?.items) ? row.customer.items : [];
    const answers = new Set();
    items.forEach((item) => {
      const value = item?.answer;
      if (value === 1 || value === "1" || value === true) answers.add(1);
      if (value === 0 || value === "0" || value === false) answers.add(0);
    });
    if (answers.size === 1) return Array.from(answers)[0];
    if (answers.size > 1) return "mixed";
    return null;
  };

  const ensureWorkstateDefaults = (row) => {
    if (!row.workstate) row.workstate = {};
    const ws = row.workstate;
    ws.selectedLevel = clampSelectedLevel(ws.selectedLevel == null ? 1 : ws.selectedLevel);
    if (!Object.prototype.hasOwnProperty.call(ws, "priority")) {
      ws.priority = 0;
    } else {
      const parsedPriority = Number(ws.priority);
      ws.priority = Number.isFinite(parsedPriority) && parsedPriority >= 0 && parsedPriority <= 4
        ? parsedPriority
        : 0;
    }
    if (ws.includeFinding == null) ws.includeFinding = false;
    if (ws.includeRecommendation == null) ws.includeRecommendation = true;
    if (ws.done == null) ws.done = false;
    if (ws.scoreTouched == null) {
      // Legacy rows: values other than historical defaults are treated as user-edited.
      ws.scoreTouched = ws.selectedLevel !== 1 || ws.done === true || ws.includeFinding === true;
    } else {
      ws.scoreTouched = ws.scoreTouched === true;
    }
    if (ws.autoScoreApplied == null) ws.autoScoreApplied = false;
    ws.autoScoreApplied = ws.autoScoreApplied === true;
    ws.autoScoreLevel = clampSelectedLevel(ws.autoScoreLevel == null ? ws.selectedLevel : ws.autoScoreLevel);
    if (ws.findingText == null) {
      if (row.type === "field_observation") {
        ws.findingText = getObservationFindingText("de-CH");
      } else {
        ws.findingText = toText(row.master?.finding);
      }
    }
    if (ws.recommendationText == null) {
      ws.recommendationText = toText(row.master?.recommendation);
    }
    const findingAction = String(ws.findingLibraryAction || "off").toLowerCase();
    ws.findingLibraryAction = ["off", "append", "replace"].includes(findingAction) ? findingAction : "off";
    if (ws.findingLibraryHash == null) ws.findingLibraryHash = "";
    const recommendationAction = String(ws.libraryAction || "off").toLowerCase();
    ws.libraryAction = ["off", "append", "replace"].includes(recommendationAction) ? recommendationAction : "off";
    if (ws.libraryHash == null) ws.libraryHash = "";

    if (!ws.scoreTouched && row?.type !== "field_observation" && row?.type !== "summary") {
      const answerState = getSelfAnswerState(row);
      const desiredLevel = answerState === 1 ? 4 : 1;
      if (ws.autoScoreApplied && ws.selectedLevel !== ws.autoScoreLevel) {
        // Row was moved away from the last auto-default value -> treat as manual.
        ws.scoreTouched = true;
      }
      if (!ws.scoreTouched) {
        ws.selectedLevel = desiredLevel;
        ws.autoScoreApplied = true;
        ws.autoScoreLevel = desiredLevel;
      }
    }

    if (ws.scoreTouched) {
      // Keep marker aligned if manual rows are later changed by script/import.
      ws.autoScoreLevel = ws.selectedLevel;
    } else if (row?.type === "field_observation" || row?.type === "summary") {
      ws.autoScoreLevel = ws.selectedLevel;
      ws.autoScoreApplied = false;
    }
  };

  const ensureProjectMeta = (project, setLocale) => {
    if (!project.meta) project.meta = {};
    if (!project.meta.locale) project.meta.locale = "de-CH";
    if (!project.meta.company) project.meta.company = "";
    if (!project.meta.companyId) project.meta.companyId = "";
    if (!project.meta.address) project.meta.address = "";
    if (!project.meta.plz) project.meta.plz = "";
    if (!project.meta.city) project.meta.city = "";
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

  const ensureChapterMetaDefaults = (chapter) => {
    if (!chapter || typeof chapter !== "object") return;
    if (!chapter.meta || typeof chapter.meta !== "object") chapter.meta = {};
    const meta = chapter.meta;
    if (meta.positivesText == null) meta.positivesText = "";
    if (meta.positivesInclude == null) {
      meta.positivesInclude = false;
    } else {
      meta.positivesInclude = meta.positivesInclude === true;
    }
    if (meta.positivesDone == null) {
      meta.positivesDone = false;
    } else {
      meta.positivesDone = meta.positivesDone === true;
    }
    const action = String(meta.positivesLibraryAction || "off").toLowerCase();
    meta.positivesLibraryAction = ["off", "append", "replace"].includes(action) ? action : "off";
    if (meta.positivesLibraryHash == null) meta.positivesLibraryHash = "";
  };

  const ensureObservationChapter = (project) => {
    if (!project?.chapters) return project;
    const hasObservation = project.chapters.some((chapter) => chapter.id === "4.8");
    if (!hasObservation) {
      project.chapters.push({
        id: "4.8",
        title: { de: OBS_CHAPTER_TITLE.de },
        rows: [],
      });
    }
    project.chapters.sort((a, b) => compareIdSegments(a.id, b.id));
    return project;
  };

  const localizeObservationChapter = (project) => {
    if (!project?.chapters) return project;
    const chapter = project.chapters.find((item) => item.id === "4.8");
    if (!chapter) return project;
    const locale = project?.meta?.locale || "de-CH";
    chapter.title = {
      de: OBS_CHAPTER_TITLE.de,
      fr: OBS_CHAPTER_TITLE.fr,
      it: OBS_CHAPTER_TITLE.it,
    };
    (chapter.rows || []).forEach((row) => {
      if (!row || row.kind === "section") {
        if (row?.id === "4.8" || row?.sectionId === "4.8") {
          row.title = getObservationSectionLabel(locale);
        }
        return;
      }
      row.sectionId = "4.8";
      row.sectionLabel = getObservationSectionLabel(locale);
      const ws = row.workstate || {};
      const master = row.master || {};
      if (!String(row.titleOverride || "").trim()) {
        row.titleOverride = row.tag || row.id || "";
      }
      if (!String(master.finding || "").trim() || OBS_DEFAULT_FINDING_TEXTS.has(String(master.finding || "").trim())) {
        master.finding = getObservationFindingText(locale);
      }
      row.master = master;
      if (!String(ws.findingText || "").trim() || OBS_DEFAULT_FINDING_TEXTS.has(String(ws.findingText || "").trim())) {
        ws.findingText = getObservationFindingText(locale);
      }
      row.workstate = ws;
    });
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
        sectionLabel: getObservationSectionLabel(project?.meta?.locale || "de-CH"),
        titleOverride: tag.label || tag.value || rowId,
        master: null,
        customer: { answer: null, remark: "", items: [] },
        workstate: {
          selectedLevel: 1,
          includeFinding: false,
          includeRecommendation: true,
          done: false,
          findingText: getObservationFindingText(project?.meta?.locale || "de-CH"),
          recommendationText: "",
          findingLibraryAction: "off",
          findingLibraryHash: "",
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
          findingLibraryAction: "off",
          findingLibraryHash: "",
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
      (summaryChapter.rows || []).forEach((row) => {
        if (!row || row.kind === "section") return;
        // Chapter 0 is edited/exported as single-text summary items.
        row.type = "summary";
        ensureWorkstateDefaults(row);
      });
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
    localizeObservationChapter(project);
    (project.chapters || []).forEach((chapter) => {
      ensureChapterMetaDefaults(chapter);
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

  const normalizeObservationTagOption = (option) => {
    if (!option) return null;
    if (typeof option === "string") {
      const value = String(option || "").trim();
      if (!value) return null;
      return { value, label: value };
    }
    if (typeof option === "object") {
      const value = String(option.value || option.label || "").trim();
      if (!value) return null;
      const label = String(option.label || value).trim() || value;
      return { value, label };
    }
    return null;
  };

  const getObservationTagOptionsFromSidecar = (doc) => {
    const options = doc?.photos?.photoTagOptions?.observations || [];
    return options
      .map(normalizeObservationTagOption)
      .filter(Boolean)
      .sort((a, b) => String(a.label || a.value || "").localeCompare(String(b.label || b.value || ""), "de", { numeric: true }));
  };

  const getObservationTagsFromSidecar = (doc) => {
    return getObservationTagOptionsFromSidecar(doc)
      .map((option) => option.value)
      .filter(Boolean);
  };

  const buildObservationRow = (tag, idOverride = "") => ({
    id: idOverride || `4.8:${slugifyTag(tag)}`,
    type: "field_observation",
    tag,
    titleOverride: tag,
    master: {
      finding: getObservationFindingText("de-CH"),
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
      findingText: getObservationFindingText("de-CH"),
      recommendationText: "",
      findingLibraryAction: "off",
      findingLibraryHash: "",
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

  const getObservationPhotoTagKeys = (doc) => {
    const keys = new Set();
    const photos = doc?.photos?.photos || {};
    Object.values(photos).forEach((photo) => {
      const tags = photo?.tags?.observations || [];
      tags.forEach((tag) => {
        const key = String(tag || "").trim();
        if (key) keys.add(key);
      });
    });
    return keys;
  };

  // Legacy sidecars are inconsistent here: some still use the original seed
  // value as the observation key, others already carry a localized key because
  // they were patched manually. Keep whichever key is actually referenced by
  // the photo assignments today; once sidecars are migrated to a canonical
  // observation-key schema, this branch can be removed entirely.
  const resolveObservationStorageKey = (existingRow, option, usedKeys) => {
    const existingKey = String(existingRow?.tag || "").trim();
    if (existingKey && usedKeys.has(existingKey)) return existingKey;
    if (usedKeys.has(option.value)) return option.value;
    if (usedKeys.has(option.label)) return option.label;
    if (existingKey) return existingKey;
    return option.value;
  };

  const syncObservationChapterRows = (project, doc) => {
    if (!project?.chapters) return;
    const chapter = project.chapters.find((item) => item.id === "4.8");
    if (!chapter) return;
    const locale = project?.meta?.locale || "de-CH";
    const tagOptions = getObservationTagOptionsFromSidecar(doc);
    if (!tagOptions.length) return;
    const usedKeys = getObservationPhotoTagKeys(doc);

    const existingRows = (chapter.rows || []).filter((row) => row.kind !== "section");
    let nextIndex = getNextObservationIndex(existingRows);
    const byTag = new Map();
    const registerRow = (key, row) => {
      const normalized = String(key || "").trim();
      if (!normalized || byTag.has(normalized)) return;
      byTag.set(normalized, row);
    };
    existingRows.forEach((row) => {
      registerRow(row?.tag, row);
      registerRow(row?.titleOverride, row);
    });

    const nextRows = tagOptions.map((option) => {
      const value = String(option?.value || "").trim();
      if (!value) return null;
      const label = String(option?.label || value).trim() || value;
      const existing = byTag.get(value) || byTag.get(label);
      if (existing) {
        existing.tag = resolveObservationStorageKey(existing, { value, label }, usedKeys);
        existing.titleOverride = label;
        return existing;
      }
      const rowId = `4.8.${nextIndex}`;
      nextIndex += 1;
      const row = buildObservationRow(value, rowId);
      row.tag = value;
      row.titleOverride = label;
      row.sectionId = "4.8";
      row.sectionLabel = getObservationSectionLabel(locale);
      row.master.finding = getObservationFindingText(locale);
      row.workstate.findingText = getObservationFindingText(locale);
      return row;
    }).filter(Boolean);

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
    ensureChapterMetaDefaults,
    ensureObservationChapter,
    localizeObservationChapter,
    ensureManagementSummaryChapter,
    normalizeProject,
    slugifyTag,
    normalizeObservationTagOption,
    getObservationTagOptionsFromSidecar,
    getObservationTagsFromSidecar,
    buildObservationRow,
    syncObservationChapterRows,
    orderObservationRows,
    moveObservationRow,
  };
})();
