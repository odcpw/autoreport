import { cloneEntry, cloneOverride } from '../utils/cloning.js';

/**
 * Project building and tree assembly functions.
 * Handles the construction of the report tree and project structure from master and self-evaluation data.
 */

/**
 * Build the complete report tree from master chapters.
 *
 * @param {Array} chapters - Master chapter definitions
 * @param {Map} selfEvalMap - Map of self-evaluation responses by ID
 * @param {Map} overrides - Map of project overrides by ID
 * @returns {Array} The constructed report tree
 */
export function buildReportTree(chapters, selfEvalMap, overrides = new Map()) {
  return chapters.map((chapter) => {
    const children = buildReportTree(chapter.children || [], selfEvalMap, overrides);
    const isFinding = Boolean(
      chapter.findingTemplate || (chapter.recommendations && Object.keys(chapter.recommendations).length)
    );
    let reportEntry = isFinding
      ? createReportEntry(chapter, selfEvalMap.get(chapter.id))
      : null;

    if (reportEntry) {
      const override = overrides?.get(chapter.id);
      if (override) {
        reportEntry = applyOverrideToEntry(reportEntry, override);
      }
    }

    return {
      id: chapter.id,
      title: chapter.title,
      isFinding,
      reportEntry,
      children,
    };
  });
}

/**
 * Create a report entry from a master chapter and optional client input.
 *
 * @param {Object} chapter - The master chapter definition
 * @param {Object} clientInput - Optional client input from self-evaluation
 * @returns {Object} The constructed report entry
 */
export function createReportEntry(chapter, clientInput) {
  const recommendations = {};
  for (let idx = 1; idx <= 4; idx += 1) {
    const key = String(idx);
    recommendations[key] = {
      master: chapter.recommendations?.[key] || '',
      adjusted: '',
      useAdjusted: false,
    };
  }

  return {
    id: chapter.id,
    title: chapter.title,
    clientInput: {
      yesNo: clientInput?.yesNo || 'n/a',
      remarks: clientInput?.remarks || '',
    },
    finding: {
      masterText: chapter.findingTemplate || '',
      adjustedText: '',
      useAdjusted: false,
    },
    recommendations,
    level: 2,
    includeInReport: true,
    proposeMasterUpdate: { action: 'none', scope: '' },
  };
}

/**
 * Collect all findings from a report tree into a flat array.
 *
 * @param {Array} tree - The report tree structure
 * @returns {Array} Flat array of all report entries
 */
export function collectFindings(tree) {
  const bucket = [];
  tree.forEach((node) => {
    if (node.isFinding && node.reportEntry) {
      bucket.push(node.reportEntry);
    }
    bucket.push(...collectFindings(node.children || []));
  });
  return bucket;
}

/**
 * Collect chapter options for selection dropdowns.
 *
 * @param {Array} tree - The report tree structure
 * @returns {Array<{value: string, label: string}>} Array of chapter options
 */
export function collectChapterOptions(tree) {
  const options = [];
  const visit = (nodes) => {
    nodes.forEach((node) => {
      if (node.isFinding) {
        options.push({ value: node.id, label: `${node.id} — ${node.title}` });
      }
      if (node.children?.length) {
        visit(node.children);
      }
    });
  };
  visit(tree);
  return options;
}

/**
 * Apply an override object to a report entry.
 *
 * @param {Object} entry - The original report entry
 * @param {Object} override - The override values to apply
 * @returns {Object} A new entry with overrides applied
 */
export function applyOverrideToEntry(entry, override) {
  const result = cloneEntry(entry);

  if (override.finding) {
    if (Object.prototype.hasOwnProperty.call(override.finding, 'adjustedText')) {
      result.finding.adjustedText = override.finding.adjustedText;
    }
    if (Object.prototype.hasOwnProperty.call(override.finding, 'useAdjusted')) {
      result.finding.useAdjusted = override.finding.useAdjusted;
    }
  }

  if (override.recommendations) {
    Object.keys(result.recommendations).forEach((key) => {
      const overrideRecommendation = override.recommendations[key];
      if (!overrideRecommendation) return;
      if (Object.prototype.hasOwnProperty.call(overrideRecommendation, 'adjusted')) {
        result.recommendations[key].adjusted = overrideRecommendation.adjusted;
      }
      if (Object.prototype.hasOwnProperty.call(overrideRecommendation, 'useAdjusted')) {
        result.recommendations[key].useAdjusted = overrideRecommendation.useAdjusted;
      }
    });
  }

  if (Object.prototype.hasOwnProperty.call(override, 'level')) {
    result.level = override.level;
  }

  if (Object.prototype.hasOwnProperty.call(override, 'includeInReport')) {
    result.includeInReport = override.includeInReport;
  }

  if (override.proposeMasterUpdate) {
    result.proposeMasterUpdate = {
      action: override.proposeMasterUpdate.action || 'none',
      scope: override.proposeMasterUpdate.scope || '',
    };
  }

  return result;
}

/**
 * Build an override object from a report entry.
 *
 * @param {Object} entry - The report entry to extract overrides from
 * @returns {Object} The override object
 */
export function buildOverrideFromEntry(entry) {
  const override = {
    finding: {
      adjustedText: entry.finding?.adjustedText || '',
      useAdjusted: Boolean(entry.finding?.useAdjusted),
    },
    recommendations: {},
    level: entry.level ?? 2,
    includeInReport: entry.includeInReport !== false,
    proposeMasterUpdate: {
      action: entry.proposeMasterUpdate?.action || 'none',
      scope: entry.proposeMasterUpdate?.scope || '',
    },
  };

  Object.keys(entry.recommendations || {}).forEach((key) => {
    const recommendation = entry.recommendations[key];
    override.recommendations[key] = {
      adjusted: recommendation.adjusted || '',
      useAdjusted: Boolean(recommendation.useAdjusted),
    };
  });

  return override;
}

/**
 * Ensure a recommendation exists in an override object.
 *
 * @param {Object} override - The override object
 * @param {string} key - The recommendation key
 */
export function ensureRecommendation(override, key) {
  if (!override.recommendations[key]) {
    override.recommendations[key] = { adjusted: '', useAdjusted: false };
  }
}

/**
 * Serialize photos Map to plain object.
 *
 * @param {Map} photoMap - Map of photo path to photo metadata
 * @returns {Object} Serialized photos object
 */
export function serializePhotos(photoMap) {
  const result = {};
  photoMap.forEach((value, key) => {
    result[key] = {
      notes: value.notes,
      tags: {
        bericht: Array.from(value.tags?.bericht || []),
        seminar: Array.from(value.tags?.seminar || []),
        topic: Array.from(value.tags?.topic || []),
      },
    };
  });
  return result;
}

/**
 * Normalize a tag list to ensure unique, trimmed values.
 *
 * @param {Array} values - Array of tag values (may be strings or objects)
 * @returns {Array<string>} Normalized array of unique tag strings
 */
export function normalizeTagList(values) {
  if (!Array.isArray(values)) return [];
  const seen = new Set();
  const normalized = [];
  values.forEach((entry) => {
    let value = entry;
    if (entry && typeof entry === 'object') {
      value = entry.value ?? entry.label ?? entry.id ?? '';
    }
    if (typeof value !== 'string') {
      value = value != null ? String(value) : '';
    }
    const trimmed = value.trim();
    if (!trimmed) return;
    const key = trimmed.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    normalized.push(trimmed);
  });
  return normalized;
}

/**
 * Get a finding by ID from the project.
 *
 * @param {Object} project - The project object
 * @param {string} id - The finding ID to search for
 * @returns {Object|null} The finding entry or null if not found
 */
export function getFindingById(project, id) {
  if (!project?.report?.chapters) return null;
  return project.report.chapters.find((chapter) => chapter.id === id) || null;
}

/**
 * Validate project snapshot structure.
 *
 * @param {Object} data - The project snapshot to validate
 * @returns {Array<string>} Array of error messages (empty if valid)
 */
export function validateProjectSnapshotStructure(data) {
  const errors = [];
  if (typeof data !== 'object' || data === null) {
    errors.push('project.json is malformed.');
    return errors;
  }

  if (!Number.isInteger(data.version)) {
    errors.push('project.json missing numeric "version" field.');
  }

  if (!data.meta || typeof data.meta !== 'object') {
    errors.push('project.json missing "meta" section.');
  }

  if (!data.report || !Array.isArray(data.report.chapters)) {
    errors.push('project.json missing report chapters array.');
  } else {
    const invalidChapter = data.report.chapters.find((chapter) => !chapter?.id);
    if (invalidChapter) {
      errors.push('project.json contains a report chapter without an "id" field.');
    }
  }

  if (
    !data.lists ||
    !Array.isArray(data.lists.berichtList) ||
    !Array.isArray(data.lists.seminarList) ||
    !Array.isArray(data.lists.topicList)
  ) {
    errors.push('project.json missing tag lists (berichtList/seminarList/topicList).');
  }

  return errors;
}
