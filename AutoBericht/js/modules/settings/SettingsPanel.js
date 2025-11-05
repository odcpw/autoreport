import { createActivityLog } from '../../shared/activityLog.js';

/**
 * SettingsPanel manages application configuration.
 * Handles PDF layout settings and tag list management (topic and seminar tags).
 */
export class SettingsPanel {
  constructor({ state }) {
    this.state = state;
    this.currentSnapshot = null;

    this.configInputs = {
      margin: document.getElementById('input-margin'),
      headerFooter: document.getElementById('input-header-footer'),
      locale: document.getElementById('input-locale'),
      footerText: document.getElementById('input-footer-text'),
      footerPreset: document.getElementById('input-footer-preset'),
    };

    this.listEditors = {};
    this.pdfDefaults = {
      pdfMargin: 25,
      headerFooter: 'show',
      locale: 'en',
      footerText: '{company} — {date}',
    };

    const activityLogElement = document.getElementById('activity-log-settings');
    this.activityLog = activityLogElement ? createActivityLog(activityLogElement, 12) : null;

    this.bindConfigInputs();
    this.bindListEditors();

    this.state.addEventListener('state:change', (event) => {
      this.renderState(event.detail);
    });

    if (this.activityLog) {
      this.activityLog.render();
    }
  }

  bindConfigInputs() {
    this.configInputs.margin?.addEventListener('change', (event) => {
      const value = Number.parseInt(event.target.value, 10);
      if (Number.isNaN(value)) return;
      this.state.updateConfig({ pdfMargin: value });
      if (this.activityLog) {
        this.activityLog.add(`PDF margin set to ${value}mm`);
      }
    });

    this.configInputs.headerFooter?.addEventListener('change', (event) => {
      this.state.updateConfig({ headerFooter: event.target.value });
      if (this.activityLog) {
        this.activityLog.add(`Header/footer ${event.target.value === 'hide' ? 'hidden' : 'shown'}`);
      }
    });

    this.configInputs.locale?.addEventListener('change', (event) => {
      this.state.updateConfig({ locale: event.target.value });
      if (this.activityLog) {
        this.activityLog.add(`Locale changed to ${event.target.value}`);
      }
    });

    this.configInputs.footerText?.addEventListener('change', (event) => {
      this.state.updateConfig({ footerText: event.target.value });
      if (this.activityLog) {
        this.activityLog.add('Footer text updated');
      }
    });

    this.configInputs.footerPreset?.addEventListener('change', (event) => {
      const value = event.target.value;
      if (!value) return;
      const formatted = value.replace(/\\n/g, '\n');
      this.state.updateConfig({ footerText: formatted });
      if (this.configInputs.footerText && this.configInputs.footerText.value !== formatted) {
        this.configInputs.footerText.value = formatted;
      }
      if (this.activityLog) {
        this.activityLog.add('Footer preset applied');
      }
    });

    const resetButton = document.getElementById('btn-reset-pdf');
    resetButton?.addEventListener('click', () => {
      this.state.updateConfig({
        pdfMargin: this.pdfDefaults.pdfMargin,
        headerFooter: this.pdfDefaults.headerFooter,
        locale: this.pdfDefaults.locale,
        footerText: this.pdfDefaults.footerText,
      });
      this.syncConfigInputs(this.state.config);
      if (this.activityLog) {
        this.activityLog.add('PDF layout settings reset');
      }
    });
  }

  bindListEditors() {
    document.querySelectorAll('[data-tag-list]').forEach((container) => {
      const type = container.dataset.tagList;
      if (!type) return;
      const items = container.querySelector('.list-editor__items');
      const form = container.querySelector('.list-editor__form');
      const input = form?.querySelector('input');

      this.listEditors[type] = { container, items, form, input };

      container.addEventListener('click', (event) => {
        const target = event.target;
        if (target instanceof HTMLElement && target.classList.contains('list-editor__remove')) {
          const value = target.dataset.value;
          if (!value) return;
          this.removeListEntry(type, value);
        }
      });

      form?.addEventListener('submit', (event) => {
        event.preventDefault();
        if (!input) return;
        const value = input.value.trim();
        if (!value) return;
        this.addListEntry(type, value);
        input.value = '';
      });
    });
  }

  renderState(snapshot) {
    this.currentSnapshot = snapshot;
    this.syncConfigInputs(snapshot.config);
    this.renderTagLists(snapshot.tagLists);
  }

  syncConfigInputs(config) {
    if (!config) return;
    const { margin, headerFooter, locale, footerText, footerPreset } = this.configInputs;
    if (margin && margin.value !== String(config.pdfMargin)) {
      margin.value = config.pdfMargin;
    }
    if (headerFooter && headerFooter.value !== config.headerFooter) {
      headerFooter.value = config.headerFooter;
    }
    if (locale && locale.value !== config.locale) {
      locale.value = config.locale;
    }
    if (footerText && footerText.value !== config.footerText) {
      footerText.value = config.footerText ?? '';
    }
    if (footerPreset) {
      const presetValue = footerPreset.value.replace(/\\n/g, '\n');
      if (config.footerText && presetValue === config.footerText) {
        footerPreset.value = footerPreset.querySelector(`option[value="${footerPreset.value}"]`)
          ? footerPreset.value
          : '';
      } else {
        footerPreset.value = '';
      }
    }
  }

  renderTagLists(tagLists) {
    if (!tagLists) return;
    Object.entries(this.listEditors).forEach(([type, editor]) => {
      const listElement = editor.items;
      if (!listElement) return;
      listElement.innerHTML = '';
      const items = tagLists[type] || [];
      if (items.length === 0) {
        const empty = document.createElement('li');
        empty.className = 'list-editor__empty';
        empty.textContent = 'No entries yet.';
        listElement.append(empty);
        return;
      }

      items.forEach((value) => {
        const item = document.createElement('li');
        item.className = 'list-editor__item';
        const label = document.createElement('span');
        label.textContent = value;
        const remove = document.createElement('button');
        remove.type = 'button';
        remove.className = 'list-editor__remove';
        remove.dataset.value = value;
        remove.dataset.list = type;
        remove.setAttribute('aria-label', `Remove ${value}`);
        remove.textContent = '×';
        item.append(label, remove);
        listElement.append(item);
      });
    });
  }

  addListEntry(type, value) {
    const existing = this.currentSnapshot?.tagLists?.[type] || [];
    const normalizedValue = value.trim();
    if (!normalizedValue) return;
    if (existing.some((item) => item.toLowerCase() === normalizedValue.toLowerCase())) {
      if (this.activityLog) {
        this.activityLog.add(`${capitalize(type)} already contains "${normalizedValue}".`);
      }
      return;
    }
    const updated = [...existing, normalizedValue];
    this.state.updateTagList(type, updated);
    if (this.activityLog) {
      this.activityLog.add(`${capitalize(type)} list updated.`);
    }
  }

  removeListEntry(type, value) {
    const existing = this.currentSnapshot?.tagLists?.[type] || [];
    const updated = existing.filter((item) => item !== value);
    if (updated.length === existing.length) return;
    this.state.updateTagList(type, updated);
    if (this.activityLog) {
      this.activityLog.add(`${capitalize(type)} entry removed.`);
    }
  }
}

function capitalize(value) {
  if (!value) return '';
  return value.charAt(0).toUpperCase() + value.slice(1);
}
