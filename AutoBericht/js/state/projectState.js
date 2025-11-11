import { parseJsonFile } from '../utils/fileUtils.js';
import { EVENTS } from '../core/constants/events.js';
import { runValidation } from '../core/state/validators.js';
import {
  buildReportTree,
  collectFindings,
  collectChapterOptions,
  serializePhotos,
  normalizeTagList,
  buildOverrideFromEntry,
  ensureRecommendation,
  getFindingById,
  validateProjectSnapshotStructure,
} from '../core/state/projectBuilder.js';
import { cloneOverride } from '../core/utils/cloning.js';
import { loadWorkbookPayload, readWorkbookFromArray, writeWorkbookSnapshot } from '../utils/workbookLoader.js';

/**
 * ProjectState manages in-memory application data derived from uploads.
 * Exposes events to enable reactive UI updates without direct DOM manipulation.
 * Single source of truth for master data, self-evaluations, photos, and project state.
 */
class ProjectState extends EventTarget {
  constructor() {
    super();
    this.reset();
  }

  reset() {
    this.master = null;
    this.selfEval = null;
    this.masterSource = null;
    this.selfEvalSource = null;
    this.sources = { master: null, selfEval: null, photos: null, project: null, workbook: null };
    this.photos = new Map();
    this.reportTree = [];
    this.project = null;
    this.projectOverrides = new Map();
    this.branding = { left: null, right: null };
    this.tagLists = {
      seminar: ['Sicherheitsgrundlagen', 'PSA'],
      topic: ['Gefährdung', 'Brandschutz'],
    };
    this.tagOptions = {
      bericht: [],
      seminar: [],
      topic: [],
    };
    this.applyTagList('seminar', this.tagLists.seminar);
    this.applyTagList('topic', this.tagLists.topic);
    this.config = {
      locale: 'en',
      pdfMargin: 25,
      headerFooter: 'show',
      footerText: '{company} — {date}',
    };
    this.workbookBinary = null;
    this.workbookName = null;
    this.validation = {
      master: { ok: null, messages: [] },
      selfEval: { ok: null, messages: [] },
      photos: { ok: null, messages: [] },
      project: { ok: null, messages: [] },
    };
    this.dispatch(EVENTS.STATE_RESET);
  }

  async ingestMaster(file) {
    const data = await parseJsonFile(file);
    this.setMasterData(data, file.name);
  }

  async ingestSelfEval(file) {
    const data = await parseJsonFile(file);
    this.setSelfEvalData(data, file.name);
  }

  async ingestProjectSnapshot(file) {
    const data = await parseJsonFile(file);
    this.setProjectSnapshot(data, file.name);
  }

  async ingestWorkbook(file) {
    const result = await loadWorkbookPayload(file);
    const payload = result.payload;
    this.workbookBinary = result.arrayBuffer;
    this.workbookName = file?.name || 'project.xlsm';
    const sourceLabel = file?.name ? `${file.name} (Excel)` : 'Workbook';
    this.sources.workbook = sourceLabel;
    this.master = payload.master;
    this.masterSource = sourceLabel;
    this.sources.master = sourceLabel;
    this.selfEval = payload.selfEval;
    this.selfEvalSource = sourceLabel;
    this.sources.selfEval = sourceLabel;
    this.projectOverrides.clear();
    this.runValidation('master', this.master);
    this.runValidation('selfEval', this.selfEval);
    this.setProjectSnapshot(payload.project, sourceLabel);
  }

  setMasterData(data, source = null) {
    this.master = data;
    this.masterSource = source || 'Loaded programmatically';
    this.sources.master = this.masterSource;
    this.projectOverrides.clear();
    this.runValidation('master', data);
    this.rebuildProject();
  }

  setSelfEvalData(data, source = null) {
    this.selfEval = data;
    this.selfEvalSource = source || 'Loaded programmatically';
    this.sources.selfEval = this.selfEvalSource;
    this.runValidation('selfEval', data);
    this.rebuildProject();
  }

  setProjectSnapshot(data, source = null) {
    if (!data || typeof data !== 'object') {
      throw new Error('Invalid project snapshot.');
    }

    const structureErrors = validateProjectSnapshotStructure(data);
    if (structureErrors.length) {
      window.dispatchEvent?.(
        new CustomEvent(EVENTS.SESSION_WARNING, {
          detail: { messages: structureErrors },
        }),
      );
      throw new Error(structureErrors[0]);
    }

    this.sources.project = source || 'Loaded programmatically';

    if (data.meta?.locale) {
      this.config = { ...this.config, locale: data.meta.locale };
    }

    if (data.branding) {
      this.branding = {
        left: data.branding.left ?? null,
        right: data.branding.right ?? null,
      };
    }

    if (data.lists) {
      if (Array.isArray(data.lists.seminarList)) {
        const seminar = normalizeTagList(data.lists.seminarList);
        this.tagLists.seminar = seminar;
        this.tagOptions.seminar = seminar.slice();
      }
      if (Array.isArray(data.lists.topicList)) {
        const topic = normalizeTagList(data.lists.topicList);
        this.tagLists.topic = topic;
        this.tagOptions.topic = topic.slice();
      }
      if (Array.isArray(data.lists.berichtList)) {
        this.tagOptions.bericht = data.lists.berichtList.map((entry) => {
          if (entry && typeof entry === 'object') {
            const value = entry.value ?? entry.id ?? '';
            const label = entry.label ?? entry.title ?? value;
            return { value, label };
          }
          const value = entry != null ? String(entry).trim() : '';
          return value ? { value, label: value } : null;
        }).filter(Boolean);
      }
    }

    this.photos.clear();
    if (data.photos && typeof data.photos === 'object') {
      Object.entries(data.photos).forEach(([path, info]) => {
        const berichtTags = Array.isArray(info?.tags?.bericht)
          ? info.tags.bericht
              .map((tag) => (tag != null ? String(tag).trim() : ''))
              .filter(Boolean)
          : [];
        const seminarTags = normalizeTagList(info?.tags?.seminar || []);
        const topicTags = normalizeTagList(info?.tags?.topic || []);

        this.photos.set(path, {
          file: null,
          notes: info?.notes || '',
          tags: {
            bericht: berichtTags,
            seminar: seminarTags,
            topic: topicTags,
          },
        });
      });
    }
    this.runValidation('photos', { count: this.photos.size });

    this.projectOverrides.clear();
    if (data.report?.chapters) {
      data.report.chapters.forEach((chapter) => {
        if (chapter?.id) {
          const override = buildOverrideFromEntry(chapter);
          this.projectOverrides.set(chapter.id, override);
        }
      });
    }

    this.rebuildProject();
  }

  ingestPhotos(fileList) {
    this.photos.clear();
    const allFiles = Array.from(fileList);
    const imageFiles = allFiles.filter(file => isImageFile(file.name));

    imageFiles.forEach((file) => {
      const relativePath = file.webkitRelativePath || file.name;
      this.photos.set(relativePath, {
        file,
        notes: '',
        tags: { bericht: [], seminar: [], topic: [] },
      });
    });

    const rejected = allFiles.length - imageFiles.length;
    this.sources.photos = rejected > 0
      ? `${imageFiles.length} images (${rejected} non-image files skipped)`
      : `${imageFiles.length} files`;

    this.runValidation('photos', { count: this.photos.size });
    this.rebuildProject();
  }

  setPhotoFixtures(entries) {
    this.photos.clear();
    entries.forEach((entry) => {
      const tags = entry.tags || {};
      this.photos.set(entry.path, {
        file: null,
        notes: entry.notes || '',
        tags: {
          bericht: Array.from(tags.bericht || []),
          seminar: Array.from(tags.seminar || []),
          topic: Array.from(tags.topic || []),
        },
      });
    });
    this.sources.photos = `${entries.length} fixtures`;
    this.runValidation('photos', { count: this.photos.size });
    this.rebuildProject();
  }

  updateConfig(partial) {
    this.config = { ...this.config, ...partial };
    this.rebuildProject();
  }

  updateBranding(side, dataUrl) {
    this.branding = { ...this.branding, [side]: dataUrl };
    this.rebuildProject();
  }

  setTagOptions(partial) {
    let changed = false;
    if (partial && Object.prototype.hasOwnProperty.call(partial, 'seminar')) {
      this.applyTagList('seminar', partial.seminar);
      changed = true;
    }
    if (partial && Object.prototype.hasOwnProperty.call(partial, 'topic')) {
      this.applyTagList('topic', partial.topic);
      changed = true;
    }
    if (partial && Object.prototype.hasOwnProperty.call(partial, 'bericht')) {
      this.tagOptions.bericht = Array.isArray(partial.bericht)
        ? partial.bericht.map((option) => ({ ...option }))
        : [];
      changed = true;
    }
    if (changed) {
      this.rebuildProject();
    }
  }

  updateTagList(type, values) {
    if (!['seminar', 'topic'].includes(type)) return;
    this.applyTagList(type, values);
    this.rebuildProject();
  }

  applyTagList(type, values) {
    if (!['seminar', 'topic'].includes(type)) return;
    const normalized = normalizeTagList(values);
    this.tagLists[type] = normalized;
    this.tagOptions[type] = normalized.slice();
  }

  updatePhotoNotes(path, notes) {
    const photo = this.photos.get(path);
    if (!photo) return;
    photo.notes = notes;
    this.photos.set(path, photo);
    this.rebuildProject();
  }

  togglePhotoTag(path, group, value) {
    const allowedGroups = ['bericht', 'seminar', 'topic'];
    if (!allowedGroups.includes(group)) return;

    const photo = this.photos.get(path);
    if (!photo || !group) return;
    if (!photo.tags) {
      photo.tags = { bericht: [], seminar: [], topic: [] };
    }
    if (!Array.isArray(photo.tags[group])) {
      photo.tags[group] = [];
    }
    const current = new Set(photo.tags[group] || []);
    if (current.has(value)) {
      current.delete(value);
    } else {
      current.add(value);
    }
    photo.tags[group] = Array.from(current);
    this.photos.set(path, photo);
    this.rebuildProject();
  }

  updateFindingAdjusted(id, text) {
    this.applyOverride(id, (override) => {
      override.finding.adjustedText = text;
    });
  }

  setFindingUseAdjusted(id, value) {
    this.applyOverride(id, (override) => {
      override.finding.useAdjusted = Boolean(value);
    });
  }

  updateFindingLevel(id, level) {
    const numeric = Number(level);
    if (Number.isNaN(numeric)) return;
    this.applyOverride(id, (override, base) => {
      override.level = Math.min(Math.max(numeric, 1), 4);
      if (!base.includeInReport) {
        override.includeInReport = base.includeInReport;
      }
    });
  }

  setFindingInclude(id, value) {
    this.applyOverride(id, (override) => {
      override.includeInReport = Boolean(value);
    });
  }

  updateRecommendationAdjusted(id, index, text) {
    const key = String(index);
    this.applyOverride(id, (override) => {
      ensureRecommendation(override, key);
      override.recommendations[key].adjusted = text;
    });
  }

  setRecommendationUse(id, index, value) {
    const key = String(index);
    this.applyOverride(id, (override) => {
      ensureRecommendation(override, key);
      override.recommendations[key].useAdjusted = Boolean(value);
    });
  }

  setProposeMasterUpdate(id, action, scope) {
    this.applyOverride(id, (override) => {
      override.proposeMasterUpdate = {
        action: action || 'none',
        scope: scope || '',
      };
    });
  }

  applyOverride(id, mutate) {
    if (!this.project) return;
    const baseEntry = getFindingById(this.project, id);
    if (!baseEntry) return;
    const currentOverride = this.projectOverrides.get(id) || buildOverrideFromEntry(baseEntry);
    const cloned = cloneOverride(currentOverride);
    mutate(cloned, baseEntry);
    this.projectOverrides.set(id, cloned);
    this.rebuildProject();
  }

  getSnapshot() {
    return {
      master: this.master,
      selfEval: this.selfEval,
      reportTree: this.reportTree,
      project: this.project,
      photos: this.getPhotoList(),
      validation: this.validation,
      config: this.config,
      branding: this.branding,
      tagOptions: this.tagOptions,
      tagLists: {
        seminar: Array.from(this.tagLists.seminar),
        topic: Array.from(this.tagLists.topic),
      },
      sources: this.sources,
      ready: this.isReadyForExport(),
      workbookReady: this.hasWorkbookTemplate(),
      workbookName: this.workbookName,
    };
  }

  getStoragePayload() {
    const clone = (value) => {
      if (!value) return null;
      return JSON.parse(JSON.stringify(value));
    };
    return {
      master: clone(this.master),
      selfEval: clone(this.selfEval),
      project: clone(this.project),
      photos: serializePhotos(this.photos),
      tagLists: {
        seminar: Array.from(this.tagLists.seminar),
        topic: Array.from(this.tagLists.topic),
      },
      config: { ...this.config },
      branding: { ...this.branding },
    };
  }

  getPhotoList() {
    return Array.from(this.photos.entries()).map(([path, meta]) => ({
      path,
      file: meta.file,
      notes: meta.notes,
      tags: meta.tags,
    }));
  }

  isReadyForExport() {
    const { master, selfEval, project } = this.validation;
    return Boolean(this.project) &&
      master.ok === true &&
      selfEval.ok === true &&
      project.ok === true;
  }

  rebuildProject() {
    this.syncTagOptions();

    if (!this.hasRequiredData()) {
      this.clearProject();
      return;
    }

    const tree = this.buildTreeWithOverrides();
    const project = this.assembleProject(tree);
    this.updateProjectState(project, tree);
  }

  syncTagOptions() {
    this.tagOptions.seminar = Array.from(this.tagLists.seminar);
    this.tagOptions.topic = Array.from(this.tagLists.topic);
  }

  hasRequiredData() {
    return Boolean(this.master && this.selfEval);
  }

  clearProject() {
    this.project = null;
    this.reportTree = [];
    this.tagOptions.bericht = [];
    this.runValidation('project', null);
    this.dispatch(EVENTS.PROJECT_CLEARED);
  }

  buildTreeWithOverrides() {
    const selfEvalMap = new Map(
      (this.selfEval.responses || []).map((entry) => [entry.id, entry])
    );
    return buildReportTree(this.master.chapters || [], selfEvalMap, this.projectOverrides);
  }

  assembleProject(tree) {
    const findings = collectFindings(tree);
    this.tagOptions.bericht = collectChapterOptions(tree);

    return {
      version: this.master.version || 1,
      meta: {
        created: new Date().toISOString(),
        company: this.selfEval.company || '',
        locale: this.config.locale,
      },
      branding: {
        left: this.branding.left,
        right: this.branding.right,
      },
      lists: {
        berichtList: this.tagOptions.bericht.map((option) => ({
          value: option.value,
          label: option.label,
        })),
        seminarList: Array.from(this.tagLists.seminar),
        topicList: Array.from(this.tagLists.topic),
      },
      photos: serializePhotos(this.photos),
      report: {
        chapters: findings,
      },
      presentation: {
        reportPresentation: { layout: 'auto-grid', maxPerSlide: 6 },
        seminar: { mode: 'per-seminar-deck', groupBy: 'bericht', maxPerSlide: 6 },
        topic: { mode: 'per-topic-deck', groupBy: 'bericht', maxPerSlide: 6 },
      },
    };
  }

  updateProjectState(project, tree) {
    this.project = project;
    this.reportTree = tree;
    this.runValidation('project', project);
    this.dispatch(EVENTS.PROJECT_UPDATED, {
      project,
      tree,
      ready: this.isReadyForExport(),
    });
  }

  runValidation(scope, data) {
    const context = {
      master: this.master,
      selfEval: this.selfEval,
    };

    const result = runValidation(scope, data, context);
    this.validation[scope] = result;
    this.dispatch(EVENTS.STATE_VALIDATION, { scope, status: result });
  }

  dispatch(type, detail = {}) {
    this.dispatchEvent(new CustomEvent(type, { detail }));
    this.dispatchEvent(new CustomEvent(EVENTS.STATE_CHANGE, { detail: this.getSnapshot() }));
  }

  hasWorkbookTemplate() {
    return Boolean(this.workbookBinary);
  }

  generateWorkbookFile() {
    if (!this.hasWorkbookTemplate()) {
      throw new Error('Load an Excel workbook before exporting.');
    }
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS library is not available.');
    }
    const workbook = readWorkbookFromArray(this.workbookBinary);
    writeWorkbookSnapshot(workbook, this.getSnapshot());
    const arrayBuffer = XLSX.write(workbook, {
      bookType: 'xlsm',
      type: 'array',
      bookSST: true,
      bookVBA: true,
    });
    this.workbookBinary = arrayBuffer;
    return {
      arrayBuffer,
      fileName: this.workbookName || 'project.xlsm',
    };
  }
}

/**
 * Check if a filename represents an image file.
 * @param {string} filename - The filename to check
 * @returns {boolean} True if the file is a supported image type
 */
function isImageFile(filename) {
  if (!filename || typeof filename !== 'string') return false;
  const ext = filename.split('.').pop().toLowerCase();
  return ['jpg', 'jpeg', 'png'].includes(ext);
}

export const projectState = new ProjectState();
