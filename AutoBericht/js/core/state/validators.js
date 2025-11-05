import { schemaValidator } from '../../schema/validator.js';
import { VALIDATION_MESSAGES } from '../constants/validation.js';

/**
 * Validation strategies for different data scopes.
 * Each validator returns { ok: boolean|null, messages: string[] }.
 */

/**
 * Validate master data against schema.
 *
 * @param {Object} data - The master data to validate
 * @param {Object} _context - Context object (unused for master validation)
 * @returns {{ok: boolean|null, messages: string[]}} Validation result
 */
function validateMaster(data, _context) {
  if (!data) {
    return { ok: null, messages: [VALIDATION_MESSAGES.MASTER_PENDING] };
  }

  const ok = Boolean(schemaValidator.validate('master', data));
  const messages = ok ? [] : schemaValidator.getErrors();
  return { ok, messages };
}

/**
 * Validate self-evaluation data against schema.
 *
 * @param {Object} data - The self-evaluation data to validate
 * @param {Object} _context - Context object (unused for self-eval validation)
 * @returns {{ok: boolean|null, messages: string[]}} Validation result
 */
function validateSelfEval(data, _context) {
  if (!data) {
    return { ok: null, messages: [VALIDATION_MESSAGES.SELF_EVAL_PENDING] };
  }

  const ok = Boolean(schemaValidator.validate('selfEval', data));
  const messages = ok ? [] : schemaValidator.getErrors();
  return { ok, messages };
}

/**
 * Validate photo library count.
 *
 * @param {Object} data - Object with count property
 * @param {Object} _context - Context object (unused for photo validation)
 * @returns {{ok: boolean|null, messages: string[]}} Validation result
 */
function validatePhotos(data, _context) {
  if (!data || typeof data.count !== 'number') {
    return { ok: null, messages: [VALIDATION_MESSAGES.PHOTOS_PENDING] };
  }

  if (data.count > 0) {
    return { ok: true, messages: [] };
  }

  return { ok: null, messages: [VALIDATION_MESSAGES.PHOTOS_NONE] };
}

/**
 * Validate project data with cross-reference checks.
 *
 * @param {Object} data - The project data to validate
 * @param {Object} context - Context object containing master and selfEval
 * @returns {{ok: boolean|null, messages: string[]}} Validation result
 */
function validateProject(data, context) {
  if (!data) {
    return { ok: null, messages: [VALIDATION_MESSAGES.PROJECT_PENDING] };
  }

  if (!context.master || !context.selfEval) {
    return { ok: null, messages: [VALIDATION_MESSAGES.PROJECT_PENDING] };
  }

  let ok = Boolean(schemaValidator.validate('project', data));
  let messages = ok ? [] : schemaValidator.getErrors();

  const crossCheck = validateCrossRefs(data, context.master);
  if (!crossCheck.ok) {
    ok = false;
    messages = messages.concat(crossCheck.messages);
  }

  return { ok, messages };
}

/**
 * Validate cross-references between project data and master data.
 *
 * @param {Object} project - The project data
 * @param {Object} master - The master data
 * @returns {{ok: boolean, messages: string[]}} Validation result
 */
function validateCrossRefs(project, master) {
  const messages = [];
  const masterIds = new Set();

  traverseMaster(master.chapters || [], (chapter) => masterIds.add(chapter.id));

  const normalizeValue = (value) => (value != null ? String(value).trim() : '');

  const berichtSet = new Set(
    (project.lists?.berichtList || []).map((entry) => {
      if (entry && typeof entry === 'object') {
        return normalizeValue(entry.value ?? entry.id ?? entry.label ?? '');
      }
      return normalizeValue(entry);
    }).filter(Boolean),
  );

  const seminarSet = new Set((project.lists?.seminarList || []).map(normalizeValue).filter(Boolean));
  const topicSet = new Set((project.lists?.topicList || []).map(normalizeValue).filter(Boolean));

  Object.entries(project.photos || {}).forEach(([path, info]) => {
    (info.tags.bericht || []).forEach((chapterId) => {
      const normalized = normalizeValue(chapterId);
      if (!normalized) return;
      if (!masterIds.has(normalized)) {
        messages.push(`Photo "${path}" references unknown bericht ${normalized}.`);
      }
      if (berichtSet.size > 0 && !berichtSet.has(normalized)) {
        messages.push(`Photo "${path}" uses bericht tag "${normalized}" not present in lists.`);
      }
    });
    (info.tags.seminar || []).forEach((seminar) => {
      const normalized = normalizeValue(seminar);
      if (!normalized) return;
      if (!seminarSet.has(normalized)) {
        messages.push(`Photo "${path}" uses unknown seminar tag "${normalized}".`);
      }
    });
    (info.tags.topic || []).forEach((topic) => {
      const normalized = normalizeValue(topic);
      if (!normalized) return;
      if (!topicSet.has(normalized)) {
        messages.push(`Photo "${path}" uses unknown topic tag "${normalized}".`);
      }
    });
  });

  (project.report?.chapters || []).forEach((chapter) => {
    if (chapter?.id && !masterIds.has(chapter.id)) {
      messages.push(`Report chapter ${chapter.id} is missing in master.json.`);
    }
  });

  return { ok: messages.length === 0, messages };
}

/**
 * Traverse master chapters recursively.
 *
 * @param {Array} chapters - Array of chapter objects
 * @param {Function} callback - Function to call for each chapter
 */
function traverseMaster(chapters, callback) {
  chapters.forEach((chapter) => {
    callback(chapter);
    if (chapter.children?.length) {
      traverseMaster(chapter.children, callback);
    }
  });
}

/**
 * Map of validation functions by scope.
 */
export const VALIDATORS = {
  master: validateMaster,
  selfEval: validateSelfEval,
  photos: validatePhotos,
  project: validateProject,
};

/**
 * Run validation for a given scope.
 *
 * @param {string} scope - The validation scope (master, selfEval, photos, project)
 * @param {Object} data - The data to validate
 * @param {Object} context - Context object with master and selfEval references
 * @returns {{ok: boolean|null, messages: string[]}} Validation result
 */
export function runValidation(scope, data, context = {}) {
  const validator = VALIDATORS[scope];
  if (!validator) {
    return { ok: Boolean(data), messages: [] };
  }
  return validator(data, context);
}
