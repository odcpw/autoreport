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
  const getSortLocale = () => document.documentElement?.getAttribute("lang") || "de";

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
    return left.localeCompare(right, getSortLocale(), { numeric: true });
  };

  const sortOptionsForGroup = (group, options) => {
    const list = [...options];
    if (group === "report") {
      return list.sort(compareNumericTags);
    }
    const locale = getSortLocale();
    return list.sort((a, b) => String(a.label || a.value || "")
      .localeCompare(String(b.label || b.value || ""), locale, { numeric: true }));
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

  const isTopLevelChapterTag = (value) => {
    const val = String(value || "").trim();
    return !val.includes("."); // anything without a dot is treated as top-level and excluded
  };
  const isZeroChapterTag = (value) => String(value || "").trim().startsWith("0.");

  const normalizeTagOptions = (options) => {
    const incoming = normalizeIncomingOptions(options);
    const report = (incoming.report || [])
      .map(normalizeTagOption)
      .filter(Boolean)
      // Keep only 1.x style entries; drop anything without a dot (top-level)
      .filter((opt) => !isTopLevelChapterTag(opt.value))
      // Drop 0.x chapters (Management Summary) from tag options
      .filter((opt) => !isZeroChapterTag(opt.value));
    return {
      report,
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
    const normalized = normalizeTagOptions(options);
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
      // Skip top-level (0,1,2,3,4...) and specific exclusion chapters
      return /^[0-9]+$/.test(topLevel) || ["11", "12", "13", "14"].includes(topLevel);
    };
    project.chapters.forEach((chapter) => {
      // Do not include whole-chapter tags; we only want 1.1-style sections
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

  const createEmptyTagOptions = () => ensureTagOptions({});
  const EMPTY_TAG_OPTIONS = createEmptyTagOptions();

  const buildReportTagOptionsFromStructure = (items) => {
    if (!Array.isArray(items)) return [];
    const tags = new Map();
    const addTag = (value, label) => {
      const val = String(value || "").trim();
      if (!val || isUnsortedLabel(val)) return;
      if (!val.includes(".")) return; // only 1.x style tags
      if (isZeroChapterTag(val)) return; // exclude 0.x
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
    items.forEach((item) => {
      if (!item) return;
      if (item.chapter) {
        addTag(String(item.chapter), item.chapterLabel || "");
      }
      const sectionId = String(item.id || "").split(".").slice(0, 2).join(".");
      if (sectionId && item.sectionLabel && !shouldSkipSection(sectionId)) {
        addTag(sectionId, item.sectionLabel);
      }
    });
    if (!tags.has("4.8")) {
      tags.set("4.8", "4.8 Beobachtungen");
    }
    const options = Array.from(tags.entries()).map(([value, label]) => ({ value, label }));
    return sortOptionsForGroup("report", options);
  };

  window.AutoBerichtPhotoSorterTags = {
    isUnsortedLabel,
    normalizeIncomingOptions,
    normalizeTagOptions,
    normalizePhotoTags,
    ensureTagOptions,
    buildReportTagOptionsFromProject,
    buildReportTagOptionsFromStructure,
    splitChapterOptions,
    sortOptionsForGroup,
    dedupeOptions,
    createEmptyTagOptions,
    EMPTY_TAG_OPTIONS,
  };
})();
