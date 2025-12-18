import { normalizeTagList } from '../core/state/projectBuilder.js';

const SHEETS = {
  META: 'Meta',
  CHAPTERS: 'Chapters',
  ROWS: 'Rows',
  PHOTOS: 'Photos',
  PHOTO_TAGS: 'PhotoTags',
  LISTS: 'Lists',
};

export async function loadWorkbookPayload(file) {
  ensureSheetJs();
  const arrayBuffer = await file.arrayBuffer();
  const workbook = readWorkbookFromArray(arrayBuffer);
  const payload = extractWorkbookPayload(workbook);
  return { payload, arrayBuffer };
}

export function readWorkbookFromArray(arrayBuffer) {
  ensureSheetJs();
  return XLSX.read(arrayBuffer, {
    type: 'array',
    cellStyles: false,
    cellFormula: false,
    bookVBA: true,
  });
}

export function writeWorkbookSnapshot(workbook, snapshot) {
  ensureSheetJs();
  if (!snapshot) throw new Error('Missing snapshot for workbook export.');
  const metaRows = buildMetaRows(snapshot);
  const chaptersRows = buildChapterRows(snapshot);
  const rowRows = buildRowsSheet(snapshot);
  const photoRows = buildPhotoRows(snapshot);
  const photoTagRows = buildPhotoTagRows(snapshot);
  const listRows = buildListRows(snapshot);

  setSheet(workbook, SHEETS.META, metaRows);
  setSheet(workbook, SHEETS.CHAPTERS, chaptersRows);
  setSheet(workbook, SHEETS.ROWS, rowRows);
  setSheet(workbook, SHEETS.PHOTOS, photoRows);
  setSheet(workbook, SHEETS.PHOTO_TAGS, photoTagRows);
  setSheet(workbook, SHEETS.LISTS, listRows);
}

function ensureSheetJs() {
  if (typeof XLSX === 'undefined') {
    throw new Error('SheetJS (XLSX) library is not loaded.');
  }
}

function sheetToJson(workbook, name, options = {}) {
  const sheet = workbook.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, {
    defval: '',
    raw: false,
    ...options,
  });
}

function extractWorkbookPayload(workbook) {
  const meta = readMeta(workbook);
  const chapterRows = readChapterRows(workbook);
  const rows = readRowEntries(workbook);
  const lists = readLists(workbook);
  const photos = readPhotos(workbook);

  const master = buildMasterPayload(meta, chapterRows, rows);
  const selfEval = buildSelfEvalPayload(meta, rows);
  const project = buildProjectSnapshot(meta, chapterRows, rows, lists, photos);

  return { master, selfEval, project };
}

function readMeta(workbook) {
  const rows = sheetToJson(workbook, SHEETS.META, { header: 1 });
  const meta = {};
  rows.slice(1).forEach((row) => {
    const key = String(row?.[0] ?? '').trim();
    if (!key) return;
    meta[key] = row?.[1] ?? '';
  });
  return meta;
}

function readChapterRows(workbook) {
  return sheetToJson(workbook, SHEETS.CHAPTERS).map((row) => ({
    id: String(row.chapterId || '').trim(),
    parentId: String(row.parentId || '').trim(),
    orderIndex: toNumber(row.orderIndex, 0),
    title: String(row.defaultTitle_de || row.defaultTitle_en || row.chapterId || '').trim(),
    titles: {
      de: String(row.defaultTitle_de || '').trim(),
      fr: String(row.defaultTitle_fr || '').trim(),
      it: String(row.defaultTitle_it || '').trim(),
      en: String(row.defaultTitle_en || '').trim(),
    },
    pageSize: toNumber(row.pageSize, null),
    isActive: toBool(row.isActive, true),
  })).filter((row) => row.id);
}

function readRowEntries(workbook) {
  const rows = sheetToJson(workbook, SHEETS.ROWS);
  return rows
    .map((row, index) => ({
      rowId: String(row.rowId || '').trim(),
      chapterId: String(row.chapterId || '').trim(),
      titleOverride: String(row.titleOverride || '').trim(),
      masterFinding: String(row.masterFinding || '').trim(),
      masterLevels: [
        String(row.masterLevel1 || '').trim(),
        String(row.masterLevel2 || '').trim(),
        String(row.masterLevel3 || '').trim(),
        String(row.masterLevel4 || '').trim(),
      ],
      overrideFinding: String(row.overrideFinding || '').trim(),
      useOverrideFinding: toBool(row.useOverrideFinding, false),
      overrideLevels: [
        String(row.overrideLevel1 || '').trim(),
        String(row.overrideLevel2 || '').trim(),
        String(row.overrideLevel3 || '').trim(),
        String(row.overrideLevel4 || '').trim(),
      ],
      useOverrideLevels: [
        toBool(row.useOverrideLevel1, false),
        toBool(row.useOverrideLevel2, false),
        toBool(row.useOverrideLevel3, false),
        toBool(row.useOverrideLevel4, false),
      ],
      customerAnswer: row.customerAnswer,
      customerRemark: String(row.customerRemark || '').trim(),
      customerPriority: row.customerPriority,
      includeFinding: toBool(row.includeFinding, true),
      includeRecommendation: toBool(row.includeRecommendation, true),
      overwriteMode: String(row.overwriteMode || '').trim(),
      selectedLevel: toNumber(row.selectedLevel, 2),
      done: toBool(row.done, false),
      notes: String(row.notes || '').trim(),
      lastEditedBy: String(row.lastEditedBy || '').trim(),
      lastEditedAt: String(row.lastEditedAt || '').trim(),
      position: index,
    }))
    .filter((row) => row.rowId);
}

function readLists(workbook) {
  const rows = sheetToJson(workbook, SHEETS.LISTS);
  const result = {};
  rows.forEach((row) => {
    const listName = String(row.listName || '').trim();
    if (!listName) return;
    const item = {
      value: pickFirst(row.value, ''),
      label: pickFirst(row.label_de, row.label_en, row.label_fr, row.label_it, row.value),
      labels: {
        de: pickFirst(row.label_de, row.value),
        fr: pickFirst(row.label_fr, row.value),
        it: pickFirst(row.label_it, row.value),
        en: pickFirst(row.label_en, row.value),
      },
      group: pickFirst(row.group, ''),
      sortOrder: row.sortOrder ?? '',
      chapterId: pickFirst(row.chapterId, ''),
    };
    if (!item.value) return;
    if (!result[listName]) result[listName] = [];
    result[listName].push(item);
  });
  return result;
}

function readPhotos(workbook) {
  const photoRows = sheetToJson(workbook, SHEETS.PHOTOS);
  const photos = {};
  photoRows.forEach((row) => {
    const fileName = String(row.fileName || '').trim();
    if (!fileName) return;
    photos[fileName] = {
      notes: String(row.notes || '').trim(),
      preferredLocale: String(row.preferredLocale || '').trim(),
      tags: {
        bericht: [],
        seminar: [],
        topic: [],
      },
    };
  });

  const tagRows = sheetToJson(workbook, SHEETS.PHOTO_TAGS);
  const mapListName = (listName) => {
    if (!listName) return null;
    if (String(listName).toLowerCase().includes('bericht')) return 'bericht';
    if (String(listName).toLowerCase().includes('seminar')) return 'seminar';
    if (String(listName).toLowerCase().includes('topic')) return 'topic';
    return null;
  };

  tagRows.forEach((row) => {
    const fileName = String(row.fileName || '').trim();
    const listName = mapListName(row.listName);
    const tagValue = String(row.tagValue || '').trim();
    if (!fileName || !listName || !tagValue) return;
    if (!photos[fileName]) {
      photos[fileName] = {
        notes: '',
        preferredLocale: '',
        tags: { bericht: [], seminar: [], topic: [] },
      };
    }
    photos[fileName].tags[listName].push(tagValue);
  });

  // Normalize tag lists
  Object.values(photos).forEach((photo) => {
    photo.tags.bericht = normalizeTagList(photo.tags.bericht);
    photo.tags.seminar = normalizeTagList(photo.tags.seminar);
    photo.tags.topic = normalizeTagList(photo.tags.topic);
  });

  return photos;
}

function buildMasterPayload(meta, chapters, rows) {
  const nodeMap = new Map();
  chapters.forEach((chapter) => {
    nodeMap.set(chapter.id, {
      ...chapter,
      children: [],
      findingTemplate: '',
      recommendations: null,
    });
  });

  rows.forEach((row) => {
    const existing = nodeMap.get(row.rowId);
    const baseNode = existing || {
      id: row.rowId,
      parentId: row.chapterId,
      orderIndex: row.position + 1000,
      title: row.titleOverride || row.rowId,
      titles: {},
      children: [],
    };
    baseNode.findingTemplate = row.masterFinding || '';
    baseNode.recommendations = {
      1: row.masterLevels[0] || '',
      2: row.masterLevels[1] || '',
      3: row.masterLevels[2] || '',
      4: row.masterLevels[3] || '',
    };
    baseNode.title = row.titleOverride || baseNode.title || row.rowId;
    nodeMap.set(baseNode.id, baseNode);
    ensureParentPlaceholder(nodeMap, row.chapterId);
  });

  const roots = [];
  nodeMap.forEach((node) => {
    if (!node.parentId || !nodeMap.has(node.parentId)) {
      roots.push(node);
      return;
    }
    const parent = nodeMap.get(node.parentId);
    parent.children.push(node);
  });

  const sortChildren = (list) => {
    list.sort((a, b) => (a.orderIndex ?? 0) - (b.orderIndex ?? 0));
    list.forEach((child) => {
      if (child.children?.length) sortChildren(child.children);
    });
  };
  sortChildren(roots);

  const serializeNode = (node) => {
    const payload = {
      id: node.id,
      title: node.title || node.id,
      orderIndex: node.orderIndex,
      pageSize: node.pageSize,
      isActive: node.isActive,
    };
    if (node.findingTemplate) {
      payload.findingTemplate = node.findingTemplate;
    }
    if (node.recommendations) {
      payload.recommendations = { ...node.recommendations };
    }
    if (node.children?.length) {
      payload.children = node.children.map(serializeNode);
    }
    return payload;
  };

  return {
    version: Number(meta.version) || 1,
    chapters: roots.map(serializeNode),
  };
}

function buildSelfEvalPayload(meta, rows) {
  const responses = rows.map((row) => ({
    id: row.rowId,
    yesNo: mapAnswer(row.customerAnswer),
    remarks: row.customerRemark || '',
  }));
  return {
    company: String(meta.company || '').trim(),
    responses,
  };
}

function buildProjectSnapshot(meta, chapterRows, rows, lists, photos) {
  const chapterMap = new Map();
  (chapterRows || []).forEach((chapter) => {
    chapterMap.set(chapter.id, {
      id: chapter.id,
      parentId: chapter.parentId || '',
      orderIndex: chapter.orderIndex || 0,
      title: chapter.titles || { de: '', fr: '', it: '', en: '' },
      pageSize: chapter.pageSize ?? '',
      isActive: typeof chapter.isActive === 'boolean' ? chapter.isActive : true,
      rows: [],
    });
  });

  (rows || []).forEach((row) => {
    const chapterId = String(row.chapterId || '').trim();
    if (!chapterId) return;
    if (!chapterMap.has(chapterId)) {
      chapterMap.set(chapterId, {
        id: chapterId,
        parentId: deriveParentId(chapterId),
        orderIndex: chapterMap.size + 1,
        title: { de: chapterId, fr: chapterId, it: chapterId, en: chapterId },
        pageSize: '',
        isActive: true,
        rows: [],
      });
    }
    chapterMap.get(chapterId).rows.push(mapRowToVbaRowEntry(row));
  });

  const chapters = Array.from(chapterMap.values()).sort(
    (a, b) => Number(a.orderIndex || 0) - Number(b.orderIndex || 0),
  );

  return {
    version: Number(meta.version) || 1,
    meta: {
      projectId: String(meta.projectId || '').trim(),
      company: String(meta.company || '').trim(),
      locale: String(meta.locale || 'de-CH').trim(),
      createdAt: meta.createdAt || new Date().toISOString(),
      author: String(meta.author || '').trim(),
    },
    branding: {},
    lists,
    photos,
    chapters,
  };
}

function mapRowToVbaRowEntry(row) {
  const levels = {
    '1': row.masterLevels?.[0] || '',
    '2': row.masterLevels?.[1] || '',
    '3': row.masterLevels?.[2] || '',
    '4': row.masterLevels?.[3] || '',
  };
  const overrideLevelText = {
    '1': row.overrideLevels?.[0] || '',
    '2': row.overrideLevels?.[1] || '',
    '3': row.overrideLevels?.[2] || '',
    '4': row.overrideLevels?.[3] || '',
  };
  const overrideLevelUse = {
    '1': Boolean(row.useOverrideLevels?.[0]),
    '2': Boolean(row.useOverrideLevels?.[1]),
    '3': Boolean(row.useOverrideLevels?.[2]),
    '4': Boolean(row.useOverrideLevels?.[3]),
  };

  return {
    id: row.rowId,
    chapterId: row.chapterId,
    titleOverride: row.titleOverride || '',
    master: {
      finding: row.masterFinding || '',
      levels,
    },
    overrides: {
      finding: { text: row.overrideFinding || '', enabled: Boolean(row.useOverrideFinding) },
      levels: {
        '1': { text: overrideLevelText['1'], enabled: overrideLevelUse['1'] },
        '2': { text: overrideLevelText['2'], enabled: overrideLevelUse['2'] },
        '3': { text: overrideLevelText['3'], enabled: overrideLevelUse['3'] },
        '4': { text: overrideLevelText['4'], enabled: overrideLevelUse['4'] },
      },
    },
    customer: {
      answer: row.customerAnswer ?? null,
      remark: row.customerRemark || '',
      priority: row.customerPriority ?? null,
    },
    workstate: {
      selectedLevel: row.selectedLevel || null,
      useFindingOverride: row.useOverrideFinding,
      findingOverride: row.overrideFinding || '',
      levelOverrides: overrideLevelText,
      useLevelOverride: overrideLevelUse,
      includeFinding: row.includeFinding !== false,
      includeRecommendation: row.includeRecommendation !== false,
      overwriteMode: row.overwriteMode || 'append',
      done: Boolean(row.done),
      notes: row.notes || '',
      lastEditedBy: row.lastEditedBy || '',
      lastEditedAt: row.lastEditedAt || '',
    },
  };
}

function ensureParentPlaceholder(nodeMap, chapterId) {
  if (!chapterId) return;
  if (nodeMap.has(chapterId)) return;
  nodeMap.set(chapterId, {
    id: chapterId,
    parentId: deriveParentId(chapterId),
    orderIndex: 0,
    title: chapterId,
    children: [],
    findingTemplate: '',
    recommendations: null,
  });
}

function deriveParentId(childId) {
  if (!childId) return '';
  const parts = childId.split('.');
  if (parts.length <= 1) return '';
  return parts.slice(0, -1).join('.');
}

function toBool(value, fallback) {
  if (value === true || value === false) return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true' || normalized === 'yes' || normalized === '1') return true;
    if (normalized === 'false' || normalized === 'no' || normalized === '0') return false;
  }
  return fallback;
}

function toNumber(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function mapAnswer(value) {
  if (value === '' || value === null || typeof value === 'undefined') return 'n/a';
  const num = Number(value);
  if (Number.isFinite(num)) return String(num);
  return String(value);
}

function pickFirst(...values) {
  for (const value of values) {
    const text = String(value || '').trim();
    if (text) return text;
  }
  return '';
}

function buildMetaRows(snapshot) {
  const headers = ['key', 'value'];
  const meta = snapshot.project?.meta || {};
  const entries = Object.keys(meta).length ? Object.entries(meta) : [];
  const rows = entries.length
    ? entries.map(([key, value]) => [key, value ?? ''])
    : [
        ['projectId', ''],
        ['company', ''],
      ];
  return [headers, ...rows];
}

function buildChapterRows(snapshot) {
  const headers = [
    'chapterId',
    'parentId',
    'orderIndex',
    'defaultTitle_de',
    'defaultTitle_fr',
    'defaultTitle_it',
    'defaultTitle_en',
    'pageSize',
    'isActive',
  ];
  const chapters = [];
  const projectChapters = Array.isArray(snapshot.project?.chapters) ? snapshot.project.chapters : [];
  if (projectChapters.length) {
    projectChapters.forEach((chapter, idx) => {
      const id = String(chapter?.id || '').trim();
      if (!id) return;
      const titles = normalizeTitles(chapter.title);
      chapters.push({
        id,
        parentId: String(chapter.parentId || '').trim(),
        orderIndex: Number.isFinite(Number(chapter.orderIndex)) ? Number(chapter.orderIndex) : idx + 1,
        titles,
        pageSize: chapter.pageSize ?? '',
        isActive: typeof chapter.isActive === 'boolean' ? chapter.isActive : true,
      });
    });
  } else {
    const master = snapshot.master?.chapters || [];
    flattenChapters(master, '', chapters, new Map());
  }

  if (!chapters.length) return [headers];

  const aoa = chapters.map((chapter) => [
    chapter.id,
    chapter.parentId,
    chapter.orderIndex,
    chapter.titles.de,
    chapter.titles.fr,
    chapter.titles.it,
    chapter.titles.en,
    chapter.pageSize ?? '',
    typeof chapter.isActive === 'boolean' ? chapter.isActive : true,
  ]);
  return [headers, ...aoa];
}

function buildRowsSheet(snapshot) {
  const headers = [
    'rowId',
    'chapterId',
    'titleOverride',
    'masterFinding',
    'masterLevel1',
    'masterLevel2',
    'masterLevel3',
    'masterLevel4',
    'overrideFinding',
    'useOverrideFinding',
    'overrideLevel1',
    'overrideLevel2',
    'overrideLevel3',
    'overrideLevel4',
    'useOverrideLevel1',
    'useOverrideLevel2',
    'useOverrideLevel3',
    'useOverrideLevel4',
    'customerAnswer',
    'customerRemark',
    'customerPriority',
    'selectedLevel',
    'includeFinding',
    'includeRecommendation',
    'overwriteMode',
    'done',
    'notes',
    'lastEditedBy',
    'lastEditedAt',
  ];

  const chapterBuckets = Array.isArray(snapshot.project?.chapters) ? snapshot.project.chapters : [];
  const projectRows = [];
  chapterBuckets.forEach((chapter) => {
    (chapter?.rows || []).forEach((entry) => projectRows.push(entry));
  });
  if (!projectRows.length) return [headers];

  const aoa = projectRows.map((entry) => {
    const rowId = entry.id || '';
    const chapterId = entry.chapterId || deriveParentId(rowId);
    const master = entry.master || {};
    const masterLevels = master.levels || {};
    const overrides = entry.overrides || {};
    const overrideFinding = overrides.finding?.text ?? entry.workstate?.findingOverride ?? '';
    const useOverrideFinding = Boolean(overrides.finding?.enabled ?? entry.workstate?.useFindingOverride);
    const overrideLevels = overrides.levels || {};
    const levelText = (k) => overrideLevels?.[k]?.text ?? entry.workstate?.levelOverrides?.[k] ?? '';
    const levelUse = (k) => Boolean(overrideLevels?.[k]?.enabled ?? entry.workstate?.useLevelOverride?.[k]);
    const customer = entry.customer || {};
    const workstate = entry.workstate || {};

    return [
      rowId,
      chapterId,
      entry.titleOverride || '',
      master.finding || '',
      masterLevels['1'] || '',
      masterLevels['2'] || '',
      masterLevels['3'] || '',
      masterLevels['4'] || '',
      overrideFinding,
      useOverrideFinding,
      levelText('1'),
      levelText('2'),
      levelText('3'),
      levelText('4'),
      levelUse('1'),
      levelUse('2'),
      levelUse('3'),
      levelUse('4'),
      customer.answer ?? '',
      customer.remark ?? '',
      customer.priority ?? '',
      workstate.selectedLevel ?? '',
      workstate.includeFinding ?? true,
      workstate.includeRecommendation ?? true,
      workstate.overwriteMode ?? '',
      Boolean(workstate.done),
      workstate.notes ?? '',
      workstate.lastEditedBy ?? '',
      workstate.lastEditedAt ?? '',
    ];
  });

  return [headers, ...aoa];
}

function buildPhotoRows(snapshot) {
  const headers = [
    'fileName',
    'notes',
    'preferredLocale',
  ];
  const photos = Array.isArray(snapshot.photos) ? snapshot.photos : [];
  if (!photos.length) return [headers];
  const aoa = photos.map((photo) => {
    const path = photo.path || photo.file?.name || '';
    return [
      path,
      photo.notes || '',
      photo.preferredLocale || '',
    ];
  });
  return [headers, ...aoa];
}

function buildPhotoTagRows(snapshot) {
  const headers = ['fileName', 'listName', 'tagValue'];
  const photos = Array.isArray(snapshot.photos) ? snapshot.photos : [];
  if (!photos.length) return [headers];
  const rows = [];
  photos.forEach((photo) => {
    const path = photo.path || photo.file?.name || '';
    const tags = photo.tags || {};
    const pushTags = (listName, values) => {
      if (!Array.isArray(values)) return;
      values.forEach((val) => {
        if (!val) return;
        rows.push([path, listName, val]);
      });
    };
    pushTags('photo.bericht', tags.bericht);
    pushTags('photo.seminar', tags.seminar);
    pushTags('photo.topic', tags.topic);
  });
  return [headers, ...rows];
}

function buildListRows(snapshot) {
  const headers = [
    'listName',
    'value',
    'label_de',
    'label_fr',
    'label_it',
    'label_en',
    'group',
    'sortOrder',
    'chapterId',
  ];
  const rows = [];

  const lists =
    snapshot?.project?.lists && typeof snapshot.project.lists === 'object'
      ? snapshot.project.lists
      : {};

  const inferGroup = (listName) => {
    switch (listName) {
      case 'photo.bericht':
        return 'bericht';
      case 'photo.seminar':
        return 'seminar';
      case 'photo.topic':
        return 'topic';
      default:
        return '';
    }
  };

  Object.keys(lists)
    .sort()
    .forEach((listName) => {
      const entries = Array.isArray(lists[listName]) ? lists[listName] : [];
      entries.forEach((entry, index) => {
        const isObject = entry && typeof entry === 'object' && !Array.isArray(entry);
        const rawValue = isObject ? entry.value : entry;
        const value = rawValue != null ? String(rawValue) : '';
        const label = isObject ? String(entry.label ?? entry.value ?? '') : value;
        const labels = isObject && entry.labels && typeof entry.labels === 'object' ? entry.labels : null;
        const group = isObject ? String(entry.group ?? inferGroup(listName)) : inferGroup(listName);
        const sortOrder = isObject && entry.sortOrder != null ? entry.sortOrder : index + 1;
        const chapterId = isObject ? String(entry.chapterId ?? '') : '';

        rows.push([
          listName,
          value,
          String(labels?.de ?? label),
          String(labels?.fr ?? label),
          String(labels?.it ?? label),
          String(labels?.en ?? label),
          group,
          sortOrder,
          chapterId,
        ]);
      });
    });

  return [headers, ...rows];
}

function setSheet(workbook, name, aoa) {
  const sheet = XLSX.utils.aoa_to_sheet(aoa);
  workbook.Sheets[name] = sheet;
  if (!workbook.SheetNames.includes(name)) {
    workbook.SheetNames.push(name);
  }
}

function flattenChapters(chapters, parentId = '', bucket = [], masterIndex = new Map()) {
  if (!Array.isArray(chapters)) return bucket;
  chapters.forEach((chapter, idx) => {
    const id = String(chapter?.id || '').trim();
    if (!id) return;
    const titles = normalizeTitles(chapter.title);
    const entry = {
      id,
      parentId,
      orderIndex: Number.isFinite(chapter.orderIndex)
        ? chapter.orderIndex
        : idx + 1,
      titles,
      pageSize: chapter.pageSize ?? '',
      isActive: typeof chapter.isActive === 'boolean' ? chapter.isActive : true,
      findingTemplate: chapter.findingTemplate || '',
      recommendations: chapter.recommendations || {},
    };
    bucket.push(entry);
    masterIndex.set(id, {
      parentId,
      findingTemplate: entry.findingTemplate,
      recommendations: entry.recommendations,
    });
    if (Array.isArray(chapter.children) && chapter.children.length) {
      flattenChapters(chapter.children, id, bucket, masterIndex);
    }
  });
  return bucket;
}

function normalizeTitles(input) {
  if (!input) {
    return { de: '', fr: '', it: '', en: '' };
  }
  if (typeof input === 'string') {
    return { de: input, fr: input, it: input, en: input };
  }
  return {
    de: input.de || input.de_CH || '',
    fr: input.fr || input.fr_CH || '',
    it: input.it || input.it_CH || '',
    en: input.en || input.en_CH || '',
  };
}

function normalizeCustomerAnswer(answer) {
  if (answer === null || typeof answer === 'undefined' || answer === '') return '';
  const num = Number(answer);
  if (Number.isFinite(num)) return num;
  return String(answer);
}
