const STORAGE_KEY = 'autobericht.session.v1';

export function isStorageAvailable() {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return false;
    }
    const testKey = '__autobericht_test__';
    window.localStorage.setItem(testKey, '1');
    window.localStorage.removeItem(testKey);
    return true;
  } catch (error) {
    console.warn('Local storage unavailable', error);
    return false;
  }
}

export function saveSession(data, options = {}) {
  if (!isStorageAvailable() || !data) return;
  try {
    const timestamp = new Date().toISOString();
    const payload = JSON.stringify({ timestamp, data });
    window.localStorage.setItem(STORAGE_KEY, payload);
    dispatchWindowEvent('autobericht:session-saved', {
      timestamp,
      source: options.source || 'autosave',
    });
    return timestamp;
  } catch (error) {
    console.error('Failed to save session', error);
    throw error;
  }
}

export function loadSession() {
  if (!isStorageAvailable()) return null;
  try {
    const stored = window.localStorage.getItem(STORAGE_KEY);
    if (!stored) return null;
    const parsed = JSON.parse(stored);
    if (!parsed || typeof parsed !== 'object') return null;
    return {
      timestamp: parsed.timestamp || null,
      data: parsed.data || null,
    };
  } catch (error) {
    console.error('Failed to load session', error);
    return null;
  }
}

export function clearSession() {
  if (!isStorageAvailable()) return;
  try {
    window.localStorage.removeItem(STORAGE_KEY);
    dispatchWindowEvent('autobericht:session-cleared');
  } catch (error) {
    console.error('Failed to clear session', error);
  }
}

export function getSessionMeta() {
  const session = loadSession();
  if (!session) return null;
  return { timestamp: session.timestamp || null };
}

function dispatchWindowEvent(type, detail) {
  if (typeof window === 'undefined') return;
  window.dispatchEvent(new CustomEvent(type, { detail }));
}
