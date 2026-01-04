import { readFileAsDataUrl } from '../../utils/fileUtils.js';
import { EVENTS } from '../../core/constants/events.js';
import { createActivityLog } from '../../shared/activityLog.js';
import { renderValidationMessages, renderValidationSummary, updateStatusBanner } from '../../shared/validationUI.js';
import { loadSession, getSessionMeta, isStorageAvailable } from '../../storage/session.js';

/**
 * ImportPanel manages data imports, validation display, and session restoration.
 * Handles uploads for master.json, self_eval.json, photos, project.json, and logos.
 */
export class ImportPanel {
  constructor({ state }) {
    this.state = state;
    this.currentSnapshot = null;
    this.storageAvailable = isStorageAvailable();
    this.prevProjectValidationKey = null;

    this.dataUploadFields = {
      master: document.getElementById('input-master'),
      selfEval: document.getElementById('input-self-eval'),
      photos: document.getElementById('input-photos'),
      workbook: document.getElementById('input-workbook'),
      project: document.getElementById('input-project'),
    };

    this.logoInputs = {
      left: document.getElementById('input-logo-left'),
      right: document.getElementById('input-logo-right'),
    };

    this.logoPreviews = {
      left: document.getElementById('preview-logo-left'),
      right: document.getElementById('preview-logo-right'),
    };

    this.validationContainer = document.getElementById('validation-messages');
    this.validationSummary = document.getElementById('validation-summary');
    this.sessionInfo = document.getElementById('session-info');
    this.sessionButtons = {
      restore: document.getElementById('btn-session-restore'),
      upload: document.getElementById('btn-session-upload'),
      clear: document.getElementById('btn-session-clear'),
    };
    this.sessionImportInput = document.getElementById('input-session-import');

    const activityLogElement = document.getElementById('activity-log-import');
    this.activityLog = activityLogElement ? createActivityLog(activityLogElement, 12) : null;

    this.bindDataUploads();
    this.bindSessionControls();
    this.bindEventListeners();

    this.state.addEventListener(EVENTS.STATE_CHANGE, (event) => {
      this.renderState(event.detail);
    });

    this.refreshSessionMeta();
    if (this.activityLog) {
      this.activityLog.render();
    }
  }

  bindDataUploads() {
    this.dataUploadFields.master?.addEventListener('change', (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      this.state.ingestMaster(file);
      this.reflectFileSelection(event.target, file.name);
      if (this.activityLog) {
        this.activityLog.add(`Uploaded master.json (${file.name})`);
      }
    });

    this.dataUploadFields.selfEval?.addEventListener('change', (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      this.state.ingestSelfEval(file);
      this.reflectFileSelection(event.target, file.name);
      if (this.activityLog) {
        this.activityLog.add(`Uploaded self_eval.json (${file.name})`);
      }
    });

    this.dataUploadFields.workbook?.addEventListener('change', async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        await this.state.ingestWorkbook(file);
        this.reflectFileSelection(event.target, file.name);
        if (this.activityLog) {
          this.activityLog.add(`Imported workbook (${file.name})`);
        }
      } catch (error) {
        console.error('Workbook import failed', error);
        if (this.activityLog) {
          this.activityLog.add('Workbook import failed');
        }
      }
    });

    this.dataUploadFields.photos?.addEventListener('change', (event) => {
      const files = event.target.files;
      if (!files || files.length === 0) return;
      this.state.ingestPhotos(files);
      this.reflectFileSelection(
        event.target,
        `${files.length} file${files.length === 1 ? '' : 's'}`
      );
      if (this.activityLog) {
        this.activityLog.add(`Uploaded ${files.length} photo(s)`);
      }
    });

    this.dataUploadFields.project?.addEventListener('change', async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        await this.state.ingestProjectSnapshot(file);
        this.reflectFileSelection(event.target, file.name);
        if (this.activityLog) {
          this.activityLog.add(`Imported project.json (${file.name})`);
        }
        const validation = this.state.validation?.project || this.currentSnapshot?.validation?.project;
        if (validation && Array.isArray(validation.messages) && validation.messages.length && validation.ok === false) {
          if (this.activityLog) {
            this.activityLog.add(`Import warning: ${validation.messages[0]}`);
          }
        }
      } catch (error) {
        console.error('Project import failed', error);
        if (this.activityLog) {
          this.activityLog.add('project.json import failed');
        }
      }
    });

    Object.entries(this.logoInputs).forEach(([side, input]) => {
      input?.addEventListener('change', async (event) => {
        const file = event.target.files?.[0];
        if (!file) return;
        const dataUrl = await readFileAsDataUrl(file);
        this.state.updateBranding(side, dataUrl);
        const preview = this.logoPreviews[side];
        if (preview) {
          preview.src = dataUrl;
          preview.alt = `${side} logo`;
        }
        if (this.activityLog) {
          this.activityLog.add(`Updated ${side} logo`);
        }
      });
    });
  }

  bindSessionControls() {
    this.sessionButtons.restore?.addEventListener('click', () => {
      if (!this.storageAvailable) {
        if (this.activityLog) {
          this.activityLog.add('Browser storage unavailable');
        }
        return;
      }
      window.dispatchEvent(new CustomEvent(EVENTS.RESTORE_SESSION));
    });

    this.sessionButtons.upload?.addEventListener('click', () => {
      this.sessionImportInput?.click();
    });

    this.sessionImportInput?.addEventListener('change', async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const payload = JSON.parse(text);
        window.dispatchEvent(
          new CustomEvent(EVENTS.RESTORE_SESSION_PAYLOAD, {
            detail: { payload },
          }),
        );
        if (this.activityLog) {
          this.activityLog.add(`Session uploaded (${file.name})`);
        }
      } catch (error) {
        console.error('Session upload failed', error);
        if (this.activityLog) {
          this.activityLog.add('Session upload failed');
        }
      } finally {
        event.target.value = '';
      }
    });

    this.sessionButtons.clear?.addEventListener('click', () => {
      if (!this.storageAvailable) {
        if (this.activityLog) {
          this.activityLog.add('Browser storage unavailable');
        }
        return;
      }
      window.dispatchEvent(new CustomEvent(EVENTS.RESET_PROJECT));
    });

    if (!this.storageAvailable) {
      this.sessionButtons.restore?.setAttribute('disabled', 'true');
      this.sessionButtons.clear?.setAttribute('disabled', 'true');
    }
  }

  bindEventListeners() {
    window.addEventListener(EVENTS.SESSION_RESTORED, (event) => {
      if (event.detail?.success) {
        this.refreshSessionMeta(event.detail?.timestamp || null);
        if (this.activityLog) {
          this.activityLog.add('Session restored', event.detail?.timestamp || null);
        }
      } else {
        this.refreshSessionMeta();
        if (this.activityLog) {
          this.activityLog.add('Session restore failed');
        }
      }
    });

    window.addEventListener(EVENTS.SESSION_WARNING, (event) => {
      const messages = event.detail?.messages || [];
      if (messages.length && this.activityLog) {
        this.activityLog.add(`Import warning: ${messages[0]}`);
      }
    });

    window.addEventListener(EVENTS.PROJECT_RESET, () => {
      if (this.activityLog) {
        this.activityLog.add('Project reset to fresh state');
      }
    });
  }

  renderState(snapshot) {
    this.currentSnapshot = snapshot;
    this.renderUploads(snapshot.sources);
    this.renderValidation(snapshot.validation);
    this.updateStatusBanner(snapshot.ready, snapshot.validation);
    this.refreshSessionMeta();
    this.trackProjectValidation(snapshot.validation?.project || null);
  }

  renderUploads(sources) {
    if (!sources) return;
    if (sources.master) {
      this.setFileBadge('master', sources.master);
    }
    if (sources.selfEval) {
      this.setFileBadge('selfEval', sources.selfEval);
    }
    if (sources.photos) {
      this.setFileBadge('photos', sources.photos);
    }
    if (sources.project) {
      this.setFileBadge('project', sources.project);
    }
  }

  renderValidation(validation) {
    const errorCount = renderValidationMessages(
      this.validationContainer,
      validation,
      (scope) => this.getSourceLabel(scope)
    );
    renderValidationSummary(this.validationSummary, errorCount);
  }

  updateStatusBanner(isReady, validation) {
    const statusElement = document.getElementById('status-validation');
    updateStatusBanner(statusElement, isReady, validation);
  }

  reflectFileSelection(input, label) {
    const field = input.closest('.upload-field');
    if (!field) return;
    this.applyBadge(field, label);
  }

  setFileBadge(fieldName, label) {
    const input = this.dataUploadFields[fieldName];
    if (!input) return;
    const field = input.closest('.upload-field');
    if (!field) return;
    this.applyBadge(field, label);
  }

  applyBadge(field, label) {
    let badge = field.querySelector('.file-badge');
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'file-badge';
      field.append(badge);
    }
    badge.textContent = label;
  }

  refreshSessionMeta(prefetchedTimestamp = null) {
    if (!this.sessionInfo) return;
    if (!this.storageAvailable) {
      this.sessionInfo.textContent = 'Browser storage unavailable.';
      return;
    }
    const meta = prefetchedTimestamp
      ? { timestamp: prefetchedTimestamp }
      : getSessionMeta();
    if (!meta || !meta.timestamp) {
      this.sessionInfo.textContent = 'No saved session.';
      return;
    }
    this.sessionInfo.textContent = `Saved ${this.formatTimestamp(meta.timestamp)}.`;
  }

  formatTimestamp(timestamp) {
    const date = new Date(timestamp);
    if (Number.isNaN(date.getTime())) return 'recently';
    if (typeof Intl !== 'undefined' && Intl.DateTimeFormat) {
      const formatter = new Intl.DateTimeFormat(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      });
      return formatter.format(date);
    }
    return date.toLocaleString();
  }

  trackProjectValidation(validation) {
    const key = validation
      ? `${validation.ok}|${(validation.messages || []).join('||')}`
      : 'null';
    if (this.prevProjectValidationKey === key) return;

    if (validation) {
      if (validation.ok === false && validation.messages?.length) {
        if (this.activityLog) {
          this.activityLog.add(`Validation error: ${validation.messages[0]}`);
        }
      } else if (
        validation.ok === true &&
        this.prevProjectValidationKey &&
        this.prevProjectValidationKey.startsWith('false')
      ) {
        if (this.activityLog) {
          this.activityLog.add('Validation passed');
        }
      }
    }

    this.prevProjectValidationKey = key;
  }

  getSourceLabel(scope) {
    const sources = this.currentSnapshot?.sources || {};
    switch (scope) {
      case 'master':
        return sources.master || null;
      case 'selfEval':
        return sources.selfEval || null;
      case 'photos':
        return sources.photos || null;
      case 'project':
        return sources.project || sources.master || null;
      default:
        return null;
    }
  }
}
