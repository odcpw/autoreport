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
import {
  ensureUnifiedDefaults,
  isUnifiedProjectSnapshot,
} from '../core/state/unifiedProject.js';
import { getLocalizedTitle } from '../core/state/unifiedProject.js';

const PROJECT_REBUILD_DEBOUNCE_MS = 250;

/**
 * ProjectState manages in-memory application data derived from uploads.
 * Exposes events to enable reactive UI updates without direct DOM manipulation.
 * Single source of truth for master data, self-evaluations, photos, and project state.
 */
class ProjectState extends EventTarget {
  constructor() {
    super();
    this.projectDirty = false;
    this.projectRebuildTimer = null;
    this.reset();
  }

  reset() {
    this.clearPendingRebuild();
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
    this.requestProjectRebuild({ immediate: true });
  }

  setSelfEvalData(data, source = null) {
    this.selfEval = data;
    this.selfEvalSource = source || 'Loaded programmatically';
    this.sources.selfEval = this.selfEvalSource;
    this.runValidation('selfEval', data);
    this.requestProjectRebuild({ immediate: true });
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
    }

    this.sources.project = source || 'Loaded programmatically';

    if (!isUnifiedProjectSnapshot(data)) {
      throw new Error(structureErrors[0] || 'Unsupported project.json shape.');
    }
    const unified = ensureUnifiedDefaults(data);

    this.master = null;
    this.selfEval = null;
    this.masterSource = null;
    this.selfEvalSource = null;
    this.sources.master = null;
    this.sources.selfEval = null;
    this.runValidation('master', null);
    this.runValidation('selfEval', null);

    if (unified.meta?.locale) {
      this.config = { ...this.config, locale: unified.meta.locale };
    }

    if (unified.branding) {
      this.branding = {
        left: unified.branding.left ?? null,
        right: unified.branding.right ?? null,
      };
    }

    this.hydrateTagListsFromProjectLists(unified);

    this.photos.clear();
    if (unified.photos && typeof unified.photos === 'object') {
      Object.entries(unified.photos).forEach(([path, info]) => {
        const tags = info?.tags && typeof info.tags === 'object' ? info.tags : {};
        const berichtTags = Array.isArray(tags.bericht)
          ? tags.bericht.map((tag) => normalizeTagValue(tag)).filter(Boolean)
          : [];
        const seminarTags = normalizeTagList(tags.seminar || []);
        const topicTags = normalizeTagList(tags.topic || []);

        this.photos.set(String(path), {
          file: null,
          notes: info?.notes || '',
          tags: { bericht: berichtTags, seminar: seminarTags, topic: topicTags },
        });
      });
    }
    this.runValidation('photos', { count: this.photos.size });

    this.projectOverrides.clear();
    this.project = unified;

    this.requestProjectRebuild({ immediate: true });
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
    this.requestProjectRebuild({ immediate: true });
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
    this.requestProjectRebuild({ immediate: true });
  }

  updateConfig(partial) {
    this.config = { ...this.config, ...partial };
    this.requestProjectRebuild({ immediate: true });
  }

  updateBranding(side, dataUrl) {
    this.branding = { ...this.branding, [side]: dataUrl };
    this.requestProjectRebuild({ immediate: true });
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
      this.requestProjectRebuild({ immediate: true });
    }
  }

  updateTagList(type, values) {
    if (!['seminar', 'topic'].includes(type)) return;
    this.applyTagList(type, values);
    this.requestProjectRebuild({ immediate: true });
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
    this.emitStateChange();
    this.requestProjectRebuild();
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
    this.emitStateChange();
    this.requestProjectRebuild();
  }

  updateFindingAdjusted(id, text) {
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      const normalized = String(text ?? '');
      row.workstate.findingOverride = normalized;
      row.overrides.finding.text = normalized;
      this.requestProjectRebuild();
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      override.finding.adjustedText = text;
    });
    if (changed) this.requestProjectRebuild();
  }

  setFindingUseAdjusted(id, value) {
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      const enabled = Boolean(value);
      row.workstate.useFindingOverride = enabled;
      row.overrides.finding.enabled = enabled;
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      override.finding.useAdjusted = Boolean(value);
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  updateFindingLevel(id, level) {
    const numeric = Number(level);
    if (Number.isNaN(numeric)) return;
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      row.workstate.selectedLevel = Math.min(Math.max(Math.trunc(numeric), 1), 4);
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override, base) => {
      override.level = Math.min(Math.max(numeric, 1), 4);
      if (!base.includeInReport) {
        override.includeInReport = base.includeInReport;
      }
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  setFindingInclude(id, value) {
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      row.workstate.includeFinding = Boolean(value);
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      override.includeInReport = Boolean(value);
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  updateRecommendationAdjusted(id, index, text) {
    const key = String(index);
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      const normalized = String(text ?? '');
      row.workstate.levelOverrides[key] = normalized;
      if (row.overrides.levels[key]) {
        row.overrides.levels[key].text = normalized;
      }
      this.requestProjectRebuild();
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      ensureRecommendation(override, key);
      override.recommendations[key].adjusted = text;
    });
    if (changed) this.requestProjectRebuild();
  }

  setRecommendationUse(id, index, value) {
    const key = String(index);
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      const enabled = Boolean(value);
      row.workstate.useLevelOverride[key] = enabled;
      if (row.overrides.levels[key]) {
        row.overrides.levels[key].enabled = enabled;
      }
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      ensureRecommendation(override, key);
      override.recommendations[key].useAdjusted = Boolean(value);
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  setFindingDone(id, value) {
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      row.workstate.done = Boolean(value);
      if (value) row.workstate.needsReview = false;
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      override.done = Boolean(value);
      if (value) {
        override.needsReview = false;
      }
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  setFindingNeedsReview(id, value) {
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      row.workstate.needsReview = Boolean(value);
      if (value) row.workstate.done = false;
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      override.needsReview = Boolean(value);
      if (value) {
        override.done = false;
      }
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  setProposeMasterUpdate(id, action, scope) {
    if (this.project?.chapters) {
      const row = findUnifiedRow(this.project, id);
      if (!row) return;
      const normalized = action || 'none';
      if (normalized === 'append' || normalized === 'replace') {
        row.workstate.overwriteMode = normalized;
      }
      this.requestProjectRebuild({ immediate: true });
      return;
    }
    const changed = this.applyOverride(id, (override) => {
      override.proposeMasterUpdate = {
        action: action || 'none',
        scope: scope || '',
      };
    });
    if (changed) this.requestProjectRebuild({ immediate: true });
  }

  applyOverride(id, mutate) {
    if (!this.project) return false;
    const baseEntry = getFindingById(this.project, id);
    if (!baseEntry) return false;
    const currentOverride = this.projectOverrides.get(id) || buildOverrideFromEntry(baseEntry);
    const cloned = cloneOverride(currentOverride);
    mutate(cloned, baseEntry);
    this.projectOverrides.set(id, cloned);
    return true;
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
    const project = this.validation.project;
    return Boolean(this.project) && project.ok === true;
  }

  requestProjectRebuild({ immediate = false } = {}) {
    if (immediate) {
      this.clearPendingRebuild();
      this.rebuildProject();
      return;
    }

    this.projectDirty = true;
    if (this.projectRebuildTimer) return;

    this.projectRebuildTimer = setTimeout(() => {
      this.projectRebuildTimer = null;
      if (!this.projectDirty) return;
      this.projectDirty = false;
      this.rebuildProject();
    }, PROJECT_REBUILD_DEBOUNCE_MS);
  }

  clearPendingRebuild() {
    if (this.projectRebuildTimer) {
      clearTimeout(this.projectRebuildTimer);
      this.projectRebuildTimer = null;
    }
    this.projectDirty = false;
  }

  rebuildProject() {
    this.syncTagOptions();

    if (!this.hasRequiredData()) {
      this.clearProject();
      return;
    }

    if (this.project?.chapters) {
      this.syncUnifiedListsAndPhotos();
      const tree = buildUnifiedReportTree(this.project);
      this.updateProjectState(this.project, tree);
      return;
    }

    const tree = this.buildTreeWithOverrides();
    const project = this.assembleProject(tree);
    if (project?.chapters) {
      this.project = project;
      this.syncUnifiedListsAndPhotos();
      const unifiedTree = buildUnifiedReportTree(project);
      this.updateProjectState(project, unifiedTree);
      return;
    }
    this.updateProjectState(project, tree);
  }

  syncTagOptions() {
    this.tagOptions.seminar = Array.from(this.tagLists.seminar);
    this.tagOptions.topic = Array.from(this.tagLists.topic);
  }

  hasRequiredData() {
    return Boolean(this.project?.chapters || (this.master && this.selfEval));
  }

  clearProject() {
    this.project = null;
    this.reportTree = [];
    this.tagOptions.bericht = [];
    this.runValidation('project', null);
    this.dispatch(EVENTS.PROJECT_CLEARED);
  }

  syncUnifiedListsAndPhotos() {
    if (!this.project?.chapters) return;
    const berichtOptions = collectUnifiedRowOptions(this.project);
    this.tagOptions.bericht = berichtOptions;
    this.project.lists = this.project.lists || {};
    syncProjectList(this.project.lists, 'photo.bericht', berichtOptions.map((o) => o.value));
    syncProjectList(this.project.lists, 'photo.seminar', Array.from(this.tagLists.seminar));
    syncProjectList(this.project.lists, 'photo.topic', Array.from(this.tagLists.topic));
    this.project.photos = serializePhotos(this.photos);
    this.project.branding = this.project.branding || {};
    this.project.branding.left = this.branding.left;
    this.project.branding.right = this.branding.right;
    this.project.meta = this.project.meta || {};
    this.project.meta.locale = this.config.locale;
  }

  hydrateTagListsFromProjectLists(project) {
    const lists = project?.lists && typeof project.lists === 'object' ? project.lists : {};
    const seminarValues = extractListValues(lists, 'photo.seminar');
    const topicValues = extractListValues(lists, 'photo.topic');
    this.tagLists.seminar = normalizeTagList(seminarValues);
    this.tagLists.topic = normalizeTagList(topicValues);
    this.tagOptions.seminar = this.tagLists.seminar.slice();
    this.tagOptions.topic = this.tagLists.topic.slice();
  }

  buildTreeWithOverrides() {
    const selfEvalMap = new Map(
      (this.selfEval.responses || []).map((entry) => [entry.id, entry])
    );
    return buildReportTree(this.master.chapters || [], selfEvalMap, this.projectOverrides);
  }

  assembleProject(tree) {
    const titleById = new Map();
    const walkMaster = (nodes) => {
      (nodes || []).forEach((node) => {
        if (node?.id) {
          titleById.set(String(node.id), String(node.title || '').trim());
        }
        if (node?.children?.length) {
          walkMaster(node.children);
        }
      });
    };
    walkMaster(this.master?.chapters || []);

    const deriveChapterId = (findingId) => {
      const parts = String(findingId || '').split('.').filter(Boolean);
      if (parts.length <= 1) return String(findingId || '');
      return parts.slice(0, -1).join('.');
    };

    const ensureLocalizedTitle = (title) => {
      const value = String(title || '').trim();
      return { de: value, fr: value, it: value, en: value };
    };

    const chaptersById = new Map();
    const ensureChapter = (chapterId) => {
      const key = String(chapterId || '').trim();
      if (!key) return null;
      const existing = chaptersById.get(key);
      if (existing) return existing;
      const title = titleById.get(key) || key;
      const created = { id: key, title: ensureLocalizedTitle(title), rows: [] };
      chaptersById.set(key, created);
      return created;
    };

    const buildOverrideBlock = (text, enabled) => ({
      text: String(text || ''),
      enabled: Boolean(enabled),
    });

    const findings = collectFindings(tree);
    findings.forEach((entry) => {
      if (!entry?.id) return;
      const id = String(entry.id);
      const chapterId = deriveChapterId(id);
      const chapter = ensureChapter(chapterId);
      if (!chapter) return;

      const findingOverrideText = entry.finding?.adjustedText || '';
      const useFindingOverride = Boolean(entry.finding?.useAdjusted);

      const masterLevels = {};
      const overrideLevels = {};
      const workstateLevelOverrides = {};
      const workstateUseLevelOverride = {};
      for (let idx = 1; idx <= 4; idx += 1) {
        const key = String(idx);
        const level = entry.recommendations?.[key] || {};
        masterLevels[key] = String(level.master || '');
        workstateLevelOverrides[key] = String(level.adjusted || '');
        workstateUseLevelOverride[key] = Boolean(level.useAdjusted);
        overrideLevels[key] = buildOverrideBlock(level.adjusted || '', level.useAdjusted);
      }

      const include = entry.includeInReport !== false;
      const customer = entry.clientInput || {};

      chapter.rows.push({
        id,
        chapterId,
        titleOverride: String(entry.title || ''),
        master: {
          finding: String(entry.finding?.masterText || ''),
          levels: masterLevels,
        },
        overrides: {
          finding: buildOverrideBlock(findingOverrideText, useFindingOverride),
          levels: overrideLevels,
        },
        customer: {
          answer: customer.yesNo ?? 'n/a',
          remark: customer.remarks ?? '',
          priority: null,
        },
        workstate: {
          selectedLevel: entry.level ?? 2,
          includeFinding: include,
          includeRecommendation: include,
          overwriteMode: 'append',
          done: Boolean(entry.done),
          needsReview: Boolean(entry.needsReview),
          notes: '',
          lastEditedBy: '',
          lastEditedAt: '',
          findingOverride: String(findingOverrideText || ''),
          useFindingOverride,
          levelOverrides: workstateLevelOverrides,
          useLevelOverride: workstateUseLevelOverride,
        },
      });
    });

    const chapters = Array.from(chaptersById.values()).sort((a, b) =>
      a.id.localeCompare(b.id, undefined, { numeric: true }),
    );
    chapters.forEach((chapter) => {
      chapter.rows.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
    });

    const berichtItems = [];
    let berichtIndex = 1;
    chapters.forEach((chapter) => {
      chapter.rows.forEach((row) => {
        const value = String(row.id || '').trim();
        if (!value) return;
        berichtItems.push({
          value,
          label: value,
          labels: { de: value, en: value, fr: value, it: value },
          group: 'bericht',
          sortOrder: berichtIndex,
          chapterId: String(chapter.id || ''),
        });
        berichtIndex += 1;
      });
    });

    const makeListItems = (values, group) =>
      (values || [])
        .map((value, index) => {
          const trimmed = String(value || '').trim();
          return {
            value: trimmed,
            label: trimmed,
            labels: { de: trimmed, en: trimmed, fr: trimmed, it: trimmed },
            group,
            sortOrder: index + 1,
            chapterId: '',
          };
        })
        .filter((item) => item.value);

    const now = new Date().toISOString();
    return {
      version: this.master?.version || 1,
      meta: {
        projectId: '',
        company: this.selfEval?.company || '',
        createdAt: now,
        locale: this.config.locale,
        author: '',
      },
      branding: {
        left: this.branding.left,
        right: this.branding.right,
      },
      lists: {
        'photo.bericht': berichtItems,
        'photo.seminar': makeListItems(Array.from(this.tagLists.seminar), 'seminar'),
        'photo.topic': makeListItems(Array.from(this.tagLists.topic), 'topic'),
      },
      photos: serializePhotos(this.photos),
      chapters,
      history: [],
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
    this.emitStateChange();
  }

  emitStateChange() {
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

function normalizeTagValue(value) {
  if (value == null) return '';
  return String(value).trim();
}

function ensureUnifiedRowScaffolding(row) {
  row.master = row.master && typeof row.master === 'object' ? row.master : {};
  row.master.levels = row.master.levels && typeof row.master.levels === 'object'
    ? row.master.levels
    : { '1': '', '2': '', '3': '', '4': '' };

  row.overrides = row.overrides && typeof row.overrides === 'object' ? row.overrides : {};
  row.overrides.finding = row.overrides.finding && typeof row.overrides.finding === 'object'
    ? row.overrides.finding
    : { text: '', enabled: false };
  row.overrides.levels = row.overrides.levels && typeof row.overrides.levels === 'object'
    ? row.overrides.levels
    : {};
  ['1', '2', '3', '4'].forEach((key) => {
    const level = row.overrides.levels[key];
    if (!level || typeof level !== 'object') {
      row.overrides.levels[key] = { text: '', enabled: false };
      return;
    }
    if (level.text == null) level.text = '';
    if (level.enabled == null) level.enabled = false;
  });

  row.customer = row.customer && typeof row.customer === 'object' ? row.customer : {};
  row.workstate = row.workstate && typeof row.workstate === 'object' ? row.workstate : {};
  row.workstate.levelOverrides = row.workstate.levelOverrides && typeof row.workstate.levelOverrides === 'object'
    ? row.workstate.levelOverrides
    : { '1': '', '2': '', '3': '', '4': '' };
  row.workstate.useLevelOverride = row.workstate.useLevelOverride && typeof row.workstate.useLevelOverride === 'object'
    ? row.workstate.useLevelOverride
    : { '1': false, '2': false, '3': false, '4': false };

  if (row.workstate.includeFinding == null) row.workstate.includeFinding = true;
  if (row.workstate.includeRecommendation == null) row.workstate.includeRecommendation = true;
  if (row.workstate.useFindingOverride == null) row.workstate.useFindingOverride = Boolean(row.overrides.finding.enabled);
  if (row.workstate.findingOverride == null) row.workstate.findingOverride = String(row.overrides.finding.text || '');
  if (row.workstate.done == null) row.workstate.done = false;
  return row;
}

function findUnifiedRow(project, id) {
  if (!project?.chapters || !id) return null;
  for (const chapter of project.chapters) {
    const rows = Array.isArray(chapter?.rows) ? chapter.rows : [];
    const match = rows.find((row) => row?.id === id);
    if (match) return ensureUnifiedRowScaffolding(match);
  }
  return null;
}

function mapCustomerAnswerToYesNo(answer) {
  if (answer == null || answer === '') return 'n/a';
  const normalized = String(answer).trim().toLowerCase();
  if (['yes', 'ja', 'y', 'true', '1'].includes(normalized)) return 'yes';
  if (['no', 'nein', 'n', 'false', '0'].includes(normalized)) return 'no';
  if (['partial', 'teilweise'].includes(normalized)) return 'partial';
  if (['n/a', 'na'].includes(normalized)) return 'n/a';
  return normalized;
}

function mapUnifiedRowToReportEntry(row) {
  const ws = row.workstate || {};
  const levels = row.master?.levels || {};
  const overrideLevels = ws.levelOverrides || {};
  const useOverrideLevels = ws.useLevelOverride || {};
  const includeInReport = ws.includeFinding ?? true;
  const selectedLevel = Number.isFinite(Number(ws.selectedLevel)) ? Number(ws.selectedLevel) : 2;

  return {
    id: row.id,
    title: row.titleOverride || row.id,
    clientInput: {
      yesNo: mapCustomerAnswerToYesNo(row.customer?.answer),
      remarks: row.customer?.remark || '',
    },
    finding: {
      masterText: row.master?.finding || '',
      adjustedText: ws.findingOverride || '',
      useAdjusted: Boolean(ws.useFindingOverride),
    },
    recommendations: {
      '1': {
        master: levels['1'] || '',
        adjusted: overrideLevels['1'] || '',
        useAdjusted: Boolean(useOverrideLevels['1']),
      },
      '2': {
        master: levels['2'] || '',
        adjusted: overrideLevels['2'] || '',
        useAdjusted: Boolean(useOverrideLevels['2']),
      },
      '3': {
        master: levels['3'] || '',
        adjusted: overrideLevels['3'] || '',
        useAdjusted: Boolean(useOverrideLevels['3']),
      },
      '4': {
        master: levels['4'] || '',
        adjusted: overrideLevels['4'] || '',
        useAdjusted: Boolean(useOverrideLevels['4']),
      },
    },
    level: selectedLevel,
    includeInReport,
    proposeMasterUpdate: {
      action: ws.overwriteMode === 'replace' ? 'replace' : ws.overwriteMode === 'append' ? 'append' : 'none',
      scope: 'all',
    },
    done: Boolean(ws.done),
    needsReview: Boolean(ws.needsReview),
  };
}

function buildUnifiedReportTree(project) {
  const chaptersByTop = new Map();
  const locale = project?.meta?.locale || '';
  project.chapters.forEach((chapter) => {
    const top = String(chapter.id || '').split('.')[0] || '';
    if (!top) return;
    const bucket = chaptersByTop.get(top) || [];
    bucket.push(chapter);
    chaptersByTop.set(top, bucket);
  });

  const topKeys = Array.from(chaptersByTop.keys()).sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true }),
  );

  return topKeys.map((topKey) => {
    const chapters = chaptersByTop.get(topKey) || [];
    chapters.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
    const children = chapters.map((chapter) => ({
      id: chapter.id,
      title: getLocalizedTitle(chapter.title, locale) || chapter.id,
      isFinding: false,
      reportEntry: null,
      children: (chapter.rows || []).map((row) => ({
        id: row.id,
        title: row.titleOverride || row.id,
        isFinding: true,
        reportEntry: mapUnifiedRowToReportEntry(ensureUnifiedRowScaffolding(row)),
        children: [],
      })),
    }));

    const topTitleRaw = chapters.find((c) => c.id === topKey)?.title;
    const topTitle = getLocalizedTitle(topTitleRaw, locale) || topKey;
    return {
      id: topKey,
      title: topTitle,
      isFinding: false,
      reportEntry: null,
      children,
    };
  });
}

function collectUnifiedRowOptions(project) {
  const options = [];
  (project.chapters || []).forEach((chapter) => {
    (chapter.rows || []).forEach((row) => {
      if (!row?.id) return;
      const title = row.titleOverride ? ` — ${row.titleOverride}` : '';
      options.push({ value: row.id, label: `${row.id}${title}` });
    });
  });
  return options.sort((a, b) => a.value.localeCompare(b.value, undefined, { numeric: true }));
}

function extractListValues(lists, listName) {
  const entries = Array.isArray(lists?.[listName]) ? lists[listName] : [];
  return entries.map((entry) => {
    if (entry && typeof entry === 'object') {
      return entry.value ?? entry.label ?? '';
    }
    return entry;
  }).map((value) => String(value || '').trim()).filter(Boolean);
}

function syncProjectList(lists, listName, values) {
  const desired = new Set((values || []).map((v) => String(v || '').trim()).filter(Boolean));
  const existing = Array.isArray(lists[listName]) ? lists[listName] : [];
  const kept = [];
  const seen = new Set();

  existing.forEach((entry) => {
    const value = entry && typeof entry === 'object' ? String(entry.value || '').trim() : String(entry || '').trim();
    if (!value) return;
    if (!desired.has(value)) return;
    if (seen.has(value)) return;
    seen.add(value);
    if (entry && typeof entry === 'object') {
      kept.push(entry);
    } else {
      kept.push({ value, label: value, labels: { de: value, en: value, fr: value, it: value } });
    }
  });

  desired.forEach((value) => {
    if (seen.has(value)) return;
    kept.push({ value, label: value, labels: { de: value, en: value, fr: value, it: value } });
  });

  lists[listName] = kept;
}
