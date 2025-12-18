function normalizeString(value) {
  if (value == null) return '';
  return String(value).trim();
}

function toBoolean(value, fallback = false) {
  if (value === true || value === false) return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (['true', 'yes', '1'].includes(normalized)) return true;
    if (['false', 'no', '0'].includes(normalized)) return false;
  }
  return fallback;
}

function toInteger(value, fallback = null) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.trunc(numeric);
}

export function isUnifiedProjectSnapshot(data) {
  return Boolean(data && typeof data === 'object' && Array.isArray(data.chapters));
}

export function getLocalizedTitle(title, locale) {
  if (!title) return '';
  if (typeof title === 'string') return title.trim();
  if (typeof title !== 'object') return normalizeString(title);

  const normalizedLocale = normalizeString(locale).toLowerCase();
  const candidates = [];
  if (normalizedLocale) {
    candidates.push(normalizedLocale);
    candidates.push(normalizedLocale.split('-')[0]);
  }
  candidates.push('de', 'en', 'fr', 'it');

  for (const key of candidates) {
    if (!key) continue;
    const value = title[key];
    const text = normalizeString(value);
    if (text) return text;
  }

  const firstValue = Object.values(title).find((value) => normalizeString(value));
  return normalizeString(firstValue);
}

export function ensureUnifiedDefaults(project) {
  const meta = project.meta && typeof project.meta === 'object' ? project.meta : {};
  const lists = project.lists && typeof project.lists === 'object' ? project.lists : {};
  const photos = project.photos && typeof project.photos === 'object' ? project.photos : {};
  const chapters = Array.isArray(project.chapters) ? project.chapters : [];

  return {
    ...project,
    version: Number.isInteger(project.version) ? project.version : 1,
    meta: {
      projectId: meta.projectId ?? meta.projectID ?? undefined,
      company: meta.company ?? '',
      locale: meta.locale ?? 'de-CH',
      createdAt: meta.createdAt ?? meta.created ?? new Date().toISOString(),
      author: meta.author ?? '',
      ...meta,
    },
    branding: project.branding && typeof project.branding === 'object' ? project.branding : {},
    lists,
    photos,
    chapters: chapters.map((chapter) => ({
      ...chapter,
      id: normalizeString(chapter.id),
      title: chapter.title ?? '',
      rows: Array.isArray(chapter.rows) ? chapter.rows : [],
    })).filter((chapter) => chapter.id),
    history: Array.isArray(project.history) ? project.history : [],
  };
}

export const unifiedProjectUtils = {
  normalizeString,
  toBoolean,
  toInteger,
};
