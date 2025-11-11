import { EVENTS } from '../../core/constants/events.js';
import { createActivityLog } from '../../shared/activityLog.js';
import { saveSession, isStorageAvailable } from '../../storage/session.js';
import { exportReportPdf } from '../../exporter/pdfExporter.js';
import { exportReportPpt, exportSeminarPpt } from '../../exporter/pptxExporter.js';

/**
 * ExportPanel manages all export operations and session downloads.
 * Handles PDF, PPTX, and project.json exports.
 */
export class ExportPanel {
  constructor({ state }) {
    this.state = state;
    this.currentSnapshot = null;
    this.storageAvailable = isStorageAvailable();

    this.exportButtons = document.querySelectorAll('.export-button');
    this.statusElement = document.getElementById('export-status');
    this.defaultStatusMessage = '';
    this.isStatusPinned = false;
    this.statusResetTimer = null;

    this.sessionButtons = {
      save: document.getElementById('btn-session-save'),
      download: document.getElementById('btn-session-download'),
    };

    const activityLogElement = document.getElementById('activity-log-export');
    this.activityLog = activityLogElement ? createActivityLog(activityLogElement, 12) : null;

    this.bindExportButtons();
    this.bindSessionControls();
    this.bindEventListeners();

    this.state.addEventListener(EVENTS.STATE_CHANGE, (event) => {
      this.renderState(event.detail);
    });

    if (this.activityLog) {
      this.activityLog.render();
    }
  }

  bindExportButtons() {
    this.exportButtons.forEach((button) => {
      button.addEventListener('click', async () => {
        if (button.disabled) return;
        const type = button.dataset.export;
        await this.handleExport(type);
        button.blur();
      });
    });
  }

  bindSessionControls() {
    this.sessionButtons.save?.addEventListener('click', () => {
      const payload = this.state.getStoragePayload();
      if (!payload?.master || !payload?.selfEval) {
        this.showStatus('Load master and self-eval before saving session.');
        return;
      }
      if (!this.storageAvailable) {
        this.showStatus('Browser storage unavailable. Use download instead.');
        return;
      }
      try {
        const timestamp = saveSession(payload, { source: 'manual' });
        if (timestamp && this.activityLog) {
          this.activityLog.add('Session saved (manual)', timestamp);
        }
      } catch (error) {
        this.showStatus('Failed to save session.');
        if (this.activityLog) {
          this.activityLog.add('Session save failed');
        }
        return;
      }
      this.showStatus('Session saved.');
    });

    this.sessionButtons.download?.addEventListener('click', () => {
      const payload = this.state.getStoragePayload();
      if (!payload?.master || !payload?.selfEval) {
        this.showStatus('Nothing to download yet.');
        return;
      }
      this.downloadSessionPayload(payload);
      this.showStatus('Session downloaded.');
      if (this.activityLog) {
        this.activityLog.add('Session downloaded');
      }
    });

    if (!this.storageAvailable) {
      this.sessionButtons.save?.setAttribute('disabled', 'true');
    }
  }

  bindEventListeners() {
    window.addEventListener(EVENTS.AUTOSAVE_START, (event) => {
      const timestamp = event.detail?.timestamp || new Date().toISOString();
      this.showStatus('Autosaving…');
      if (this.activityLog) {
        this.activityLog.add('Autosave started', timestamp);
      }
    });

    window.addEventListener(EVENTS.AUTOSAVE_SUCCESS, (event) => {
      const timestamp = event.detail?.timestamp || new Date().toISOString();
      this.showStatus('Autosave complete.');
      if (this.activityLog) {
        this.activityLog.add('Autosave complete', timestamp);
      }
    });

    window.addEventListener(EVENTS.AUTOSAVE_FAILED, (event) => {
      const reason = event.detail?.error || 'unknown error';
      this.showStatus('Autosave failed.');
      if (this.activityLog) {
        this.activityLog.add(`Autosave failed: ${reason}`);
      }
    });

    window.addEventListener(EVENTS.SESSION_SAVED, (event) => {
      const source = event.detail?.source || (this.storageAvailable ? 'autosave' : 'manual');
      if (source === 'autosave' && this.activityLog) {
        this.activityLog.add(`Session saved (${source})`, event.detail?.timestamp || null);
      }
    });

    window.addEventListener(EVENTS.SESSION_CLEARED, () => {
      if (this.activityLog) {
        this.activityLog.add('Session cleared');
      }
    });
  }

  renderState(snapshot) {
    this.currentSnapshot = snapshot;
    this.updateExports(snapshot);
  }

  updateExports(snapshot) {
    const isReady = snapshot?.ready === true;
    const workbookReady = Boolean(snapshot?.workbookReady);
    this.exportButtons.forEach((button) => {
      const type = button.dataset.export;
      if (type === 'workbook') {
        button.disabled = !isReady || !workbookReady;
      } else {
        button.disabled = !isReady;
      }
    });
    this.defaultStatusMessage = isReady
      ? 'Validation passed. Exports ready.'
      : 'Validation pending. Upload inputs to enable exports.';
    if (!this.isStatusPinned && this.statusElement) {
      this.statusElement.textContent = this.defaultStatusMessage;
    }
  }

  async handleExport(type) {
    try {
      switch (type) {
        case 'project-json':
          this.downloadProjectJson();
          if (this.activityLog) {
            this.activityLog.add('project.json downloaded');
          }
          break;
        case 'workbook':
          this.ensureSnapshotReady();
          this.ensureWorkbookTemplate();
          this.downloadWorkbook();
          if (this.activityLog) {
            this.activityLog.add('Workbook downloaded');
          }
          break;
        case 'pdf':
          this.ensureSnapshotReady();
          exportReportPdf(this.currentSnapshot);
          this.showStatus('PDF export opened.');
          if (this.activityLog) {
            this.activityLog.add('PDF export generated');
          }
          break;
        case 'pptx-report':
          this.ensureSnapshotReady();
          await exportReportPpt(this.currentSnapshot);
          this.showStatus('Report PPTX export started.');
          if (this.activityLog) {
            this.activityLog.add('Report PPTX export generated');
          }
          break;
        case 'pptx-seminar':
          this.ensureSnapshotReady();
          await exportSeminarPpt(this.currentSnapshot);
          this.showStatus('Seminar PPTX export started.');
          if (this.activityLog) {
            this.activityLog.add('Seminar PPTX export generated');
          }
          break;
        default:
          this.showStatus('Unknown export type.');
          break;
      }
    } catch (error) {
      console.error('Export failed', error);
      this.showStatus(error?.message || 'Export failed.');
      if (this.activityLog) {
        this.activityLog.add(`Export failed: ${error?.message || error}`);
      }
    }
  }

  ensureSnapshotReady() {
    if (!this.currentSnapshot?.project) {
      throw new Error('No project data loaded.');
    }
    if (this.currentSnapshot.ready === false) {
      const projectValidation = this.currentSnapshot.validation?.project;
      if (projectValidation?.messages?.length) {
        throw new Error(projectValidation.messages[0]);
      }
      throw new Error('Resolve validation errors before exporting.');
    }
  }

  downloadProjectJson() {
    const project = this.currentSnapshot?.project;
    if (!project) {
      this.showStatus('No project data available to export.');
      return;
    }

    const json = JSON.stringify(project, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'project.json';
    document.body.append(link);
    link.click();
    link.remove();
    requestAnimationFrame(() => URL.revokeObjectURL(url));
    this.showStatus('project.json downloaded.');
  }

  downloadSessionPayload(payload) {
    const timestamp = new Date().toISOString().split('T')[0];
    const blob = new Blob(
      [JSON.stringify({ timestamp: new Date().toISOString(), data: payload }, null, 2)],
      { type: 'application/json' }
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${timestamp}-autobericht-session.json`;
    document.body.append(link);
    link.click();
    link.remove();
    requestAnimationFrame(() => URL.revokeObjectURL(url));
  }

  showStatus(message) {
    if (!this.statusElement) return;
    this.isStatusPinned = true;
    this.statusElement.textContent = message;
    if (this.statusResetTimer) {
      clearTimeout(this.statusResetTimer);
    }
    this.statusResetTimer = window.setTimeout(() => {
      this.isStatusPinned = false;
      this.statusElement.textContent = this.defaultStatusMessage;
      this.statusResetTimer = null;
    }, 2500);
  }

  ensureWorkbookTemplate() {
    if (!this.state?.hasWorkbookTemplate || !this.state.hasWorkbookTemplate()) {
      throw new Error('Load a workbook before exporting.');
    }
  }

  downloadWorkbook() {
    const result = this.state.generateWorkbookFile();
    const blob = new Blob([result.arrayBuffer], {
      type: 'application/vnd.ms-excel.sheet.macroEnabled.12',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = result.fileName || 'project.xlsm';
    document.body.append(link);
    link.click();
    link.remove();
    requestAnimationFrame(() => URL.revokeObjectURL(url));
    this.showStatus('Workbook downloaded.');
  }
}
