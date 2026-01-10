(() => {
  const stateHelpers = window.AutoBerichtPhotoSorterState || {};

  const isUnsortedLabel = (value) => String(value || "").trim().toLowerCase() === "unsorted";

  const normalizeIncomingOptions = (options) => {
    if (!options) return {};
    const mapped = { ...options };
    if (!mapped.report && mapped.bericht) mapped.report = mapped.bericht;
    if (!mapped.training && mapped.seminar) mapped.training = mapped.seminar;
    if (!mapped.observations && mapped.topic) mapped.observations = mapped.topic;
    return mapped;
  };

  const normalizeTagOption = (option) => {
    if (!option) return null;
    if (typeof option === "string") {
      if (isUnsortedLabel(option)) return null;
      return { value: option, label: option };
    }
    if (typeof option === "object") {
      const value = option.value || option.label;
      if (!value) return null;
      if (isUnsortedLabel(value)) return null;
      return { value, label: option.label || value };
    }
    return null;
  };

  const isNumericTag = (value) => /^\d+(?:\.\d+)*$/.test(value);

  const compareNumericTags = (a, b) => {
    const left = String(a.value || a.label || "");
    const right = String(b.value || b.label || "");
    const leftIsNum = isNumericTag(left);
    const rightIsNum = isNumericTag(right);
    if (leftIsNum && rightIsNum) {
      const leftParts = left.split(".").map((part) => Number(part));
      const rightParts = right.split(".").map((part) => Number(part));
      const max = Math.max(leftParts.length, rightParts.length);
      for (let i = 0; i < max; i += 1) {
        const l = leftParts[i] ?? -1;
        const r = rightParts[i] ?? -1;
        if (l !== r) return l - r;
      }
      return 0;
    }
    if (leftIsNum) return -1;
    if (rightIsNum) return 1;
    return left.localeCompare(right, "de", { numeric: true });
  };

  const sortOptionsForGroup = (group, options) => {
    const list = [...options];
    if (group === "report") {
      return list.sort(compareNumericTags);
    }
    return list.sort((a, b) => String(a.label || a.value || "")
      .localeCompare(String(b.label || b.value || ""), "de", { numeric: true }));
  };

  const dedupeOptions = (options) => {
    const seen = new Set();
    return options.filter((opt) => {
      const key = opt.value;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const normalizeTagOptions = (options) => {
    const incoming = normalizeIncomingOptions(options);
    return {
      report: (incoming.report || []).map(normalizeTagOption).filter(Boolean),
      observations: (incoming.observations || []).map(normalizeTagOption).filter(Boolean),
      training: (incoming.training || []).map(normalizeTagOption).filter(Boolean),
    };
  };

  const normalizePhotoTags = (tags) => {
    const incoming = normalizeIncomingOptions(tags || {});
    const sanitize = (list) => Array.from(list || []).filter((value) => !isUnsortedLabel(value));
    return {
      report: sanitize(incoming.report),
      observations: sanitize(incoming.observations),
      training: sanitize(incoming.training),
    };
  };

  const ensureTagOptions = (options) => {
    const incoming = normalizeIncomingOptions(options);
    const merged = {
      report: [...(stateHelpers.DEFAULT_TAGS?.report || []), ...(incoming.report || [])],
      observations: [...(stateHelpers.DEFAULT_TAGS?.observations || []), ...(incoming.observations || [])],
      training: [...(stateHelpers.DEFAULT_TAGS?.training || []), ...(incoming.training || [])],
    };
    const normalized = normalizeTagOptions(merged);
    normalized.report = sortOptionsForGroup("report", dedupeOptions(normalized.report));
    normalized.observations = sortOptionsForGroup("observations", dedupeOptions(normalized.observations));
    normalized.training = sortOptionsForGroup("training", dedupeOptions(normalized.training));
    return normalized;
  };

  const buildReportTagOptionsFromProject = (project) => {
    if (!project?.chapters) return null;
    const tags = new Map();
    const addTag = (value, label) => {
      const val = String(value || "").trim();
      if (!val || isUnsortedLabel(val)) return;
      let lbl = String(label || val).trim();
      if (lbl && !(lbl === val || lbl.startsWith(`${val} `) || lbl.startsWith(`${val}.`))) {
        lbl = `${val} ${lbl}`;
      }
      if (!tags.has(val)) tags.set(val, lbl || val);
    };
    const shouldSkipSection = (sectionId) => {
      const topLevel = String(sectionId || "").split(".")[0];
      return ["11", "12", "13", "14"].includes(topLevel);
    };
    project.chapters.forEach((chapter) => {
      addTag(chapter.id, chapter.title?.de || chapter.title || chapter.id);
      (chapter.rows || []).forEach((row) => {
        if (row.kind === "section") {
          if (!shouldSkipSection(row.id)) {
            addTag(row.id, row.title || row.id);
          }
          return;
        }
        if (row.sectionId) {
          if (!shouldSkipSection(row.sectionId)) {
            addTag(row.sectionId, row.sectionLabel || row.sectionId);
          }
        }
      });
    });
    const options = Array.from(tags.entries()).map(([value, label]) => ({ value, label }));
    return sortOptionsForGroup("report", options);
  };

  const splitChapterOptions = (options) => {
    const chapters = [];
    const rest = [];
    options.forEach((option) => {
      const value = String(option.value || "");
      if (isNumericTag(value) && !value.includes(".")) {
        chapters.push(option);
      } else {
        rest.push(option);
      }
    });
    return { chapters, rest };
  };

  const SEED_TAG_OPTIONS = window.PS_CATEGORY_LABELS_SEED
    ? ensureTagOptions(window.PS_CATEGORY_LABELS_SEED)
    : ensureTagOptions(stateHelpers.DEFAULT_TAGS);

  window.AutoBerichtPhotoSorterTags = {
    isUnsortedLabel,
    normalizeIncomingOptions,
    normalizeTagOptions,
    normalizePhotoTags,
    ensureTagOptions,
    buildReportTagOptionsFromProject,
    splitChapterOptions,
    sortOptionsForGroup,
    dedupeOptions,
    SEED_TAG_OPTIONS,
  };
})();
