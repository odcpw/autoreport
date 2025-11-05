import { TabController } from './ui/tabController.js';
import { projectState } from './state/projectState.js';
import { ImportPanel } from './modules/import/ImportPanel.js';
import { ExportPanel } from './modules/export/ExportPanel.js';
import { SettingsPanel } from './modules/settings/SettingsPanel.js';
import { AutoBerichtPanel } from './ui/autoberichtPanel.js';
import { PhotoSorterPanel } from './modules/photosorter/PhotoSorterPanel.js';
import { EVENTS } from './core/constants/events.js';
import { loadFixtures } from './state/fixtures.js';
import { loadSession, saveSession, clearSession } from './storage/session.js';

const SAVE_DELAY_MS = 600;
let autosaveTimer = null;
let suspendAutosave = false;
let tabsController = null;
let lastFocusedElement = null;

window.addEventListener('DOMContentLoaded', async () => {
  const tabs = new TabController({
    tabs: document.querySelectorAll('.app-tab'),
    panels: document.querySelectorAll('.tab-panel'),
  });
  tabsController = tabs;
  tabs.focusFirstTab();

  new ImportPanel({ state: projectState });
  new PhotoSorterPanel({
    state: projectState,
    container: document.getElementById('photosorter-content'),
  });
  new AutoBerichtPanel({
    state: projectState,
    treeContainer: document.getElementById('autobericht-tree'),
    detailContainer: document.getElementById('autobericht-detail'),
  });
  new ExportPanel({ state: projectState });
  new SettingsPanel({ state: projectState });

  setFsStatus();
  await bootstrapData();
  projectState.addEventListener(EVENTS.STATE_CHANGE, handleStateChange);
  window.addEventListener(EVENTS.RESTORE_SESSION, handleRestoreRequest);
  window.addEventListener(EVENTS.RESTORE_SESSION_PAYLOAD, handleRestorePayload);
  window.addEventListener(EVENTS.RESET_PROJECT, handleResetProject);
  setupKeyboardShortcuts();
});

function setFsStatus() {
  const statusElement = document.getElementById('status-storage');
  if (!statusElement) return;
  const available = 'showOpenFilePicker' in window;
  statusElement.textContent = available
    ? 'FS Access: available'
    : 'FS Access: unavailable';
  statusElement.className = available
    ? 'status-pill status-pill--ok'
    : 'status-pill';
}

async function bootstrapData() {
  const stored = loadSession();
  let restoredFromSession = false;
  if (stored?.data) {
    restoredFromSession = restoreFromPayload(stored.data, stored.timestamp);
    if (restoredFromSession) {
      window.dispatchEvent(
        new CustomEvent(EVENTS.SESSION_RESTORED, {
          detail: { success: true, timestamp: stored.timestamp || null },
        }),
      );
    }
  }

  if (!restoredFromSession) {
    try {
      await loadFixtures(projectState);
    } catch (error) {
      console.warn('Fixture loading skipped:', error);
    }
  }
}

function handleStateChange() {
  scheduleAutosave();
}

function scheduleAutosave() {
  if (suspendAutosave) return;
  if (typeof window === 'undefined') return;
  if (autosaveTimer) {
    window.clearTimeout(autosaveTimer);
  }
  autosaveTimer = window.setTimeout(() => {
    autosaveTimer = null;
    const payload = projectState.getStoragePayload();
    if (!payload?.master || !payload?.selfEval) return;
    const startTimestamp = new Date().toISOString();
    window.dispatchEvent(
      new CustomEvent(EVENTS.AUTOSAVE_START, {
        detail: { timestamp: startTimestamp },
      }),
    );
    try {
      const savedAt = saveSession(payload, { source: 'autosave' }) || startTimestamp;
      window.dispatchEvent(
        new CustomEvent(EVENTS.AUTOSAVE_SUCCESS, {
          detail: { timestamp: savedAt },
        }),
      );
    } catch (error) {
      window.dispatchEvent(
        new CustomEvent(EVENTS.AUTOSAVE_FAILED, {
          detail: { error: error.message },
        }),
      );
    }
  }, SAVE_DELAY_MS);
}

function handleRestoreRequest() {
  const stored = loadSession();
  if (!stored?.data) {
    window.dispatchEvent(
      new CustomEvent(EVENTS.SESSION_RESTORED, {
        detail: { success: false, reason: 'missing' },
      }),
    );
    return;
  }

  const ok = restoreFromPayload(stored.data, stored.timestamp);
  window.dispatchEvent(
    new CustomEvent(EVENTS.SESSION_RESTORED, {
      detail: {
        success: ok,
        timestamp: stored.timestamp || null,
      },
    }),
  );
  if (ok) {
    scheduleAutosave();
  }
}

function handleRestorePayload(event) {
  const payload = event.detail?.payload;
  if (!payload) {
    window.dispatchEvent(
      new CustomEvent(EVENTS.SESSION_RESTORED, {
        detail: { success: false, reason: 'invalid' },
      }),
    );
    return;
  }

  const data = payload.data || payload;
  const ok = restoreFromPayload(data, payload.timestamp || null);
  window.dispatchEvent(
    new CustomEvent(EVENTS.SESSION_RESTORED, {
      detail: {
        success: ok,
        timestamp: payload.timestamp || null,
      },
    }),
  );
  if (ok) {
    scheduleAutosave();
  }
}

function handleResetProject() {
  suspendAutosave = true;
  try {
    if (autosaveTimer) {
      window.clearTimeout(autosaveTimer);
      autosaveTimer = null;
    }
    projectState.reset();
    clearSession();
    window.dispatchEvent(
      new CustomEvent(EVENTS.PROJECT_RESET, {
        detail: { timestamp: new Date().toISOString() },
      }),
    );
  } finally {
    suspendAutosave = false;
  }
}

function setupKeyboardShortcuts() {
  const overlay = document.getElementById('shortcut-overlay');
  const closeBtn = document.getElementById('btn-shortcut-close');
  if (!overlay || !closeBtn) return;

  const setOverlayVisibility = (open) => {
    overlay.setAttribute('aria-hidden', open ? 'false' : 'true');
    if (open) {
      lastFocusedElement = document.activeElement;
      closeBtn.focus();
    } else if (lastFocusedElement instanceof HTMLElement) {
      lastFocusedElement.focus();
      lastFocusedElement = null;
    }
  };

  closeBtn.addEventListener('click', () => setOverlayVisibility(false));
  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) {
      setOverlayVisibility(false);
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && overlay.getAttribute('aria-hidden') === 'false') {
      event.preventDefault();
      setOverlayVisibility(false);
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.ctrlKey && !event.shiftKey && !event.altKey && event.key === '/') {
      event.preventDefault();
      const open = overlay.getAttribute('aria-hidden') === 'false';
      setOverlayVisibility(!open);
    }
  });

  document.addEventListener('keydown', (event) => {
    if (!event.altKey || event.ctrlKey || event.shiftKey) return;
    if (!tabsController) return;
    if (overlay.getAttribute('aria-hidden') === 'false') return;
    const targetTag = (event.target || {}).tagName;
    if (['INPUT', 'TEXTAREA', 'SELECT'].includes(targetTag)) return;
    switch (event.key) {
      case '1':
        event.preventDefault();
        tabsController.activate(0);
        break;
      case '2':
        event.preventDefault();
        tabsController.activate(1);
        break;
      case '3':
        event.preventDefault();
        tabsController.activate(2);
        break;
      case '4':
        event.preventDefault();
        tabsController.activate(3);
        break;
      default:
        break;
    }
  });
}

function restoreFromPayload(payload, timestamp = null) {
  if (!payload) return false;
  suspendAutosave = true;
  try {
    if (payload.tagLists?.topic) {
      projectState.updateTagList('topic', payload.tagLists.topic);
    }
    if (payload.tagLists?.seminar) {
      projectState.updateTagList('seminar', payload.tagLists.seminar);
    }
    if (payload.master) {
      projectState.setMasterData(payload.master, sourceLabel(timestamp));
    }
    if (payload.selfEval) {
      projectState.setSelfEvalData(payload.selfEval, sourceLabel(timestamp));
    }
    if (payload.project) {
      projectState.setProjectSnapshot(payload.project, sourceLabel(timestamp));
    }
    if (payload.config) {
      projectState.updateConfig(payload.config);
    }
    const brandingHandledByProject = Boolean(payload.project?.branding);
    if (payload.branding && !brandingHandledByProject) {
      if (payload.branding.left !== undefined) {
        projectState.updateBranding('left', payload.branding.left);
      }
      if (payload.branding.right !== undefined) {
        projectState.updateBranding('right', payload.branding.right);
      }
    }
  } catch (error) {
    console.error('Failed to restore session', error);
    return false;
  } finally {
    suspendAutosave = false;
  }
  const validation = projectState.validation?.project;
  if (validation && (validation.ok === false || (validation.messages?.length))) {
    window.dispatchEvent(
      new CustomEvent(EVENTS.SESSION_WARNING, {
        detail: { messages: validation.messages || [] },
      }),
    );
  }
  return true;
}

function sourceLabel(timestamp) {
  if (!timestamp) return 'Stored session';
  return `Stored session (${timestamp})`;
}
