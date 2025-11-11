import { renderMarkdown, renderMarkdownInline } from '../utils/markdownRenderer.js';
import { MarkdownEditor } from './markdownEditor.js';

export class AutoBerichtPanel {
  constructor({ state, treeContainer, detailContainer }) {
    this.state = state;
    this.treeContainer = treeContainer;
    this.detailContainer = detailContainer;
    this.selectedId = null;
    this.currentSnapshot = null;

    this.chapterStrip = document.getElementById('autobericht-strip');
    this.subStrip = document.getElementById('autobericht-substrip');
    this.drawer = document.getElementById('autobericht-drawer');
    this.drawerToggle = document.getElementById('autobericht-drawer-toggle');
    this.drawerClose = document.getElementById('autobericht-drawer-close');

    this.state.addEventListener('state:change', (event) => {
      this.currentSnapshot = event.detail;
      this.ensureSelection();
      this.renderMainChapters();
      this.renderSubChapters();
      this.renderTree();
      this.renderDetail();
    });

    this.treeContainer?.addEventListener('keydown', (event) => {
      this.handleTreeKeydown(event);
    });

    this.drawerToggle?.addEventListener('click', () => this.setDrawerVisibility(true));
    this.drawerClose?.addEventListener('click', () => this.setDrawerVisibility(false));
    this.drawer?.addEventListener('click', (event) => {
      if (event.target === this.drawer) {
        this.setDrawerVisibility(false);
      }
    });
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && this.drawer?.getAttribute('aria-hidden') === 'false') {
        event.preventDefault();
        this.setDrawerVisibility(false);
      }
    });
  }

  ensureSelection() {
    if (!this.currentSnapshot?.reportTree?.length) {
      this.selectedId = null;
      return;
    }
    if (this.selectedId) {
      const exists = findNodeById(this.currentSnapshot.reportTree, this.selectedId);
      if (exists) return;
    }
    const firstTop = this.collectTopLevelChapters()[0];
    this.selectedId = this.pickFirstFinding(firstTop?.id) ?? null;
  }

  renderTree() {
    if (!this.treeContainer) return;

    if (!this.currentSnapshot?.reportTree?.length) {
      this.treeContainer.innerHTML = '';
      const empty = document.createElement('div');
      empty.className = 'empty-state';
      empty.innerHTML = '<p>Load master and self-evaluation JSON to view findings.</p>';
      this.treeContainer.append(empty);
      return;
    }

    const list = buildTreeList(this.currentSnapshot.reportTree, this.selectedId, (id) => {
      this.selectedId = id;
      this.renderTree();
      this.renderDetail();
    });

    this.treeContainer.innerHTML = '';
    this.treeContainer.append(list);
  }

  renderMainChapters() {
    if (!this.chapterStrip) return;
    const topChapters = this.collectTopLevelChapters();
    this.chapterStrip.innerHTML = '';
    if (!topChapters.length) {
      const empty = document.createElement('p');
      empty.className = 'chapter-strip__empty';
      empty.textContent = 'Load data to show chapters.';
      this.chapterStrip.append(empty);
      return;
    }

    topChapters.slice(0, 11).forEach((chapter, index) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'chapter-button';
      button.textContent = chapter.id || `#${index + 1}`;
      button.dataset.chapterId = chapter.id;
      const isActive = this.selectedId?.startsWith?.(chapter.id);
      button.classList.toggle('chapter-button--active', isActive);
      button.setAttribute('aria-pressed', String(isActive));
      button.addEventListener('click', () => {
        this.selectedId = this.pickFirstFinding(chapter.id);
        this.renderMainChapters();
        this.renderSubChapters();
        this.renderTree();
        this.renderDetail();
      });
      button.addEventListener('keydown', (event) => this.handleStripKeydown(event, index));
      this.chapterStrip.append(button);
    });
  }

  renderSubChapters() {
    if (!this.subStrip) return;
    this.subStrip.innerHTML = '';
    const parentId = this.getSelectedTopLevelId();
    if (!parentId) return;
    const secondLevel = this.collectSecondLevelChapters(parentId);
    if (!secondLevel.length) return;

    secondLevel.forEach((chapter, index) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'chapter-button';
      button.textContent = chapter.id;
      const isActive = this.selectedId?.startsWith?.(chapter.id);
      button.classList.toggle('chapter-button--active', isActive);
      button.addEventListener('click', () => {
        this.selectedId = this.pickFirstFinding(chapter.id);
        this.renderMainChapters();
        this.renderSubChapters();
        this.renderTree();
        this.renderDetail();
      });
      button.addEventListener('keydown', (event) => this.handleSubStripKeydown(event, index));
      this.subStrip.append(button);
    });
  }

  renderDetail() {
    if (!this.detailContainer) return;

    const rows = this.collectRowEntries();
    if (!rows.length) {
      this.detailContainer.innerHTML = '';
      const placeholder = document.createElement('div');
      placeholder.className = 'detail-placeholder';
      placeholder.innerHTML = '<p>Select a chapter to view findings.</p>';
      this.detailContainer.append(placeholder);
      return;
    }

    if (!this.selectedId) {
      this.selectedId = rows[0]?.id ?? null;
    }

    const focusMeta = this.captureFocusMeta();
    this.detailContainer.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'finding-row-container';

    rows.forEach((entry) => {
      const row = renderFindingRow(entry, {
        onFindingChange: (text) => this.state.updateFindingAdjusted(entry.id, text),
        onFindingToggle: (value) => this.state.setFindingUseAdjusted(entry.id, value),
        onLevelChange: (level) => this.state.updateFindingLevel(entry.id, level),
        onIncludeChange: (value) => this.state.setFindingInclude(entry.id, value),
        onRecommendationChange: (index, text) =>
          this.state.updateRecommendationAdjusted(entry.id, index, text),
        onRecommendationToggle: (index, value) =>
          this.state.setRecommendationUse(entry.id, index, value),
        onMasterUpdateChange: (update) =>
          this.state.setProposeMasterUpdate(entry.id, update.action, update.scope),
      });
      container.append(row);
    });

    this.detailContainer.append(container);
    this.restoreFocusMeta(focusMeta);
  }

  handleTreeKeydown(event) {
    if (!this.currentSnapshot?.reportTree?.length) return;
    const focusable = Array.from(this.treeContainer.querySelectorAll('[data-node-id]'));
    const currentIndex = focusable.findIndex((el) => el.dataset.nodeId === this.selectedId);
    if (currentIndex === -1) return;

    const moveSelection = (delta) => {
      const nextIndex = Math.min(Math.max(currentIndex + delta, 0), focusable.length - 1);
      const next = focusable[nextIndex];
      if (next) {
        this.selectedId = next.dataset.nodeId;
        this.renderTree();
        this.renderDetail();
        next.focus();
      }
    };

    switch (event.key) {
      case 'ArrowUp':
        event.preventDefault();
        moveSelection(-1);
        break;
      case 'ArrowDown':
        event.preventDefault();
        moveSelection(1);
        break;
      case 'Enter':
      case ' ':
        event.preventDefault();
        if (focusable[currentIndex]) {
          focusable[currentIndex].click();
        }
        break;
      default:
        break;
    }
  }

  captureFocusMeta() {
    if (!this.detailContainer) return null;
    const active = document.activeElement;
    if (!active || !this.detailContainer.contains(active)) return null;
    const target = active.closest('[data-field]');
    if (!target) return null;
    const meta = {
      field: target.dataset.field || null,
      recommendation: target.dataset.recommendation || null,
      finding: target.dataset.finding || null,
      selectionStart: null,
      selectionEnd: null,
    };

    if (target.dataset.editorWrapper === 'true') {
      return meta;
    }

    if ('selectionStart' in active && 'selectionEnd' in active) {
      meta.selectionStart = active.selectionStart;
      meta.selectionEnd = active.selectionEnd;
    }

    if (!meta.field) return null;
    if (meta.finding && meta.finding !== this.selectedId) return null;
    return meta;
  }

  restoreFocusMeta(meta) {
    if (!meta || !meta.field || !this.detailContainer) return;
    const attributeSelector = [`[data-field="${meta.field}"]`];
    if (meta.recommendation) {
      attributeSelector.push(`[data-recommendation="${meta.recommendation}"]`);
    }
    attributeSelector.push(`[data-finding="${this.selectedId}"]`);
    const selector = attributeSelector.join('');
    const target = this.detailContainer.querySelector(selector);
    if (!target) return;

    const editorInstance =
      MarkdownEditor.getInstance(target) ||
      MarkdownEditor.getInstance(target.querySelector?.('textarea'));
    if (editorInstance) {
      editorInstance.focus();
      return;
    }

    target.focus({ preventScroll: true });
    if (
      typeof meta.selectionStart === 'number' &&
      typeof meta.selectionEnd === 'number' &&
      'setSelectionRange' in target
    ) {
      try {
        target.setSelectionRange(meta.selectionStart, meta.selectionEnd);
      } catch (error) {
        // Ignore selection errors
      }
    }
  }

  handleStripKeydown(event, index) {
    const buttons = Array.from(this.chapterStrip?.querySelectorAll('.chapter-button') || []);
    if (!buttons.length) return;
    const moveFocus = (delta) => {
      const nextIndex = (index + delta + buttons.length) % buttons.length;
      buttons[nextIndex].focus();
    };
    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowDown':
        event.preventDefault();
        moveFocus(1);
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
        event.preventDefault();
        moveFocus(-1);
        break;
      case 'Home':
        event.preventDefault();
        buttons[0].focus();
        break;
      case 'End':
        event.preventDefault();
        buttons[buttons.length - 1].focus();
        break;
      case 'Enter':
      case ' ':
        event.preventDefault();
        buttons[index].click();
        break;
      default:
        break;
    }
  }

  handleSubStripKeydown(event, index) {
    const buttons = Array.from(this.subStrip?.querySelectorAll('.chapter-button') || []);
    if (!buttons.length) return;
    const moveFocus = (delta) => {
      const nextIndex = (index + delta + buttons.length) % buttons.length;
      buttons[nextIndex].focus();
    };
    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowDown':
        event.preventDefault();
        moveFocus(1);
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
        event.preventDefault();
        moveFocus(-1);
        break;
      case 'Home':
        event.preventDefault();
        buttons[0].focus();
        break;
      case 'End':
        event.preventDefault();
        buttons[buttons.length - 1].focus();
        break;
      case 'Enter':
      case ' ':
        event.preventDefault();
        buttons[index].click();
        break;
      default:
        break;
    }
  }

  setDrawerVisibility(open) {
    if (!this.drawer) return;
    this.drawer.setAttribute('aria-hidden', String(!open));
    if (open) {
      this.drawerClose?.focus();
    } else {
      this.drawerToggle?.focus();
    }
  }

  collectTopLevelChapters() {
    if (!this.currentSnapshot?.reportTree) return [];
    return this.currentSnapshot.reportTree.filter((node) => node.isFinding || node.children?.length);
  }

  collectSecondLevelChapters(parentId) {
    const parent = findNodeByIdAny(this.currentSnapshot.reportTree, parentId);
    if (!parent?.children) return [];
    return parent.children.filter((node) => node.isFinding || node.children?.length);
  }

  pickFirstFinding(prefix) {
    if (!prefix) {
      const first = findFirstFinding(this.currentSnapshot.reportTree);
      return first?.id ?? null;
    }
    const node = findNodeByIdAny(this.currentSnapshot.reportTree, prefix);
    if (!node) return null;
    if (node.isFinding && node.reportEntry) return node.id;
    const childFinding = findFirstFinding(node.children || []);
    return childFinding?.id ?? node.id;
  }

  getSelectedTopLevelId() {
    if (!this.selectedId) return null;
    return this.selectedId.split('.')[0] || null;
  }

  collectRowEntries() {
    const parentId = this.getActiveParentId();
    if (!parentId) return [];
    const parentNode = findNodeByIdAny(this.currentSnapshot.reportTree, parentId);
    if (!parentNode) return [];
    const findings = flattenFindings(parentNode);
    return findings
      .map((node) => getFindingById(this.currentSnapshot.project, node.id))
      .filter(Boolean);
  }

  getActiveParentId() {
    if (!this.selectedId) {
      const top = this.collectTopLevelChapters()[0];
      return top?.id || null;
    }
    const parts = this.selectedId.split('.');
    if (parts.length >= 3) return `${parts[0]}.${parts[1]}`;
    if (parts.length === 2) return this.selectedId;
    return parts[0];
  }
}

function buildTreeList(nodes, selectedId, onSelect) {
  const list = document.createElement('ul');
  list.className = 'chapter-tree';

  nodes.forEach((node) => {
    const item = document.createElement('li');
    item.className = 'chapter-node';
    if (node.isFinding && node.id === selectedId) {
      item.classList.add('chapter-node--selected');
    }

    const header = document.createElement(node.isFinding ? 'button' : 'div');
    if (node.isFinding) {
      header.type = 'button';
      header.dataset.nodeId = node.id;
      header.addEventListener('click', () => onSelect(node.id));
    }
    header.className = 'chapter-node__header';

    const title = document.createElement('div');
    title.className = 'chapter-node__title';
    title.textContent = `${node.id} — ${node.title}`;
    header.append(title);

    if (node.isFinding && node.reportEntry) {
      const chip = document.createElement('span');
      chip.className = 'client-chip';
      chip.dataset.state = node.reportEntry.clientInput.yesNo || 'n/a';
      chip.textContent = (node.reportEntry.clientInput.yesNo || 'n/a').toUpperCase();
      header.append(chip);
    }

    item.append(header);

    if (node.children?.length) {
      const childList = buildTreeList(node.children, selectedId, onSelect);
      childList.classList.add('chapter-node__children');
      item.append(childList);
    }

    list.append(item);
  });

  return list;
}

function renderDetailHeader(entry) {
  const wrapper = document.createElement('div');
  wrapper.className = 'detail-header';

  const title = document.createElement('h3');
  title.textContent = `${entry.id} — ${entry.title}`;
  wrapper.append(title);

  const chip = document.createElement('span');
  chip.className = 'client-chip';
  chip.dataset.state = entry.clientInput.yesNo || 'n/a';
  chip.textContent = (entry.clientInput.yesNo || 'n/a').toUpperCase();
  wrapper.append(chip);

  return wrapper;
}

function renderFindingRow(entry, handlers) {
  const row = document.createElement('article');
  row.className = 'finding-row';
  row.dataset.findingId = entry.id;

  const header = document.createElement('header');
  header.className = 'finding-row__header';

  const title = document.createElement('div');
  title.className = 'finding-row__title';
  title.innerHTML = `<strong>${entry.id}</strong> ${entry.title || ''}`;
  header.append(title);

  const levelGroup = document.createElement('div');
  levelGroup.className = 'finding-row__levels';
  ['1', '2', '3', '4'].forEach((value) => {
    const label = document.createElement('label');
    label.className = 'level-radio';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = `level-${entry.id}`;
    radio.value = value;
    radio.checked = Number(entry.level) === Number(value);
    radio.addEventListener('change', () => handlers.onLevelChange(Number(value)));
    label.append(radio, document.createTextNode(value));
    levelGroup.append(label);
  });
  header.append(levelGroup);

  const controls = document.createElement('div');
  controls.className = 'finding-row__controls';

  const includeLabel = document.createElement('label');
  includeLabel.className = 'checkbox-row';
  const includeCheckbox = document.createElement('input');
  includeCheckbox.type = 'checkbox';
  includeCheckbox.checked = Boolean(entry.includeInReport);
  includeCheckbox.addEventListener('change', () => handlers.onIncludeChange(includeCheckbox.checked));
  includeLabel.append(includeCheckbox, document.createTextNode('Bericht'));
  controls.append(includeLabel);

  const updateSelect = document.createElement('select');
  updateSelect.innerHTML = `
    <option value="none">No master update</option>
    <option value="finding">Finding</option>
    <option value="recommendation">Recommendations</option>
    <option value="all">Entire record</option>
  `;
  updateSelect.value = entry.proposeMasterUpdate?.scope || '';
  updateSelect.addEventListener('change', () =>
    handlers.onMasterUpdateChange({
      action: updateSelect.value ? 'append' : 'none',
      scope: updateSelect.value,
    }),
  );
  controls.append(updateSelect);

  header.append(controls);
  row.append(header);

  const columns = document.createElement('div');
  columns.className = 'finding-row__columns';

  // Finding column
  const findingColumn = document.createElement('div');
  findingColumn.className = 'finding-row__column finding-row__column--finding';
  findingColumn.innerHTML = `
    <label>Master Text</label>
    <div class="preview-pane">${renderMarkdown(entry.finding.masterText || '')}</div>
  `;
  const adjustedField = document.createElement('div');
  adjustedField.className = 'detail-field';
  adjustedField.innerHTML = '<label>Adjusted Text</label>';
  const adjustArea = document.createElement('textarea');
  adjustArea.className = 'markdown-textarea';
  adjustArea.dataset.field = 'finding-adjusted';
  adjustArea.dataset.finding = entry.id;
  adjustedField.append(adjustArea);
  new MarkdownEditor(adjustArea, {
    value: entry.finding.adjustedText || '',
    onChange: handlers.onFindingChange,
  });
  const checkboxRow = document.createElement('label');
  checkboxRow.className = 'checkbox-row';
  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.checked = Boolean(entry.finding.useAdjusted);
  checkbox.addEventListener('change', () => handlers.onFindingToggle(checkbox.checked));
  checkboxRow.append(checkbox, document.createTextNode('Use adjusted text'));
  adjustedField.append(checkboxRow);
  findingColumn.append(adjustedField);

  columns.append(findingColumn);

  const recommendationsColumn = document.createElement('div');
  recommendationsColumn.className = 'finding-row__column finding-row__column--recommendations';

  Object.keys(entry.recommendations || {}).forEach((key) => {
    const recommendation = entry.recommendations[key];
    const card = document.createElement('div');
    card.className = 'recommendation-card';
    card.innerHTML = `
      <header>
        <strong>${key}</strong>
        <label class="checkbox-row">
          <input type="checkbox" ${recommendation.useAdjusted ? 'checked' : ''} />
          <span>Use adjusted</span>
        </label>
      </header>
      <div class="recommendation-master">${renderMarkdown(recommendation.master || '')}</div>
    `;
    const checkbox = card.querySelector('input[type="checkbox"]');
    checkbox.addEventListener('change', () => handlers.onRecommendationToggle(key, checkbox.checked));

    const textarea = document.createElement('textarea');
    textarea.className = 'markdown-textarea';
    textarea.dataset.field = 'recommendation-adjusted';
    textarea.dataset.finding = entry.id;
    textarea.dataset.recommendation = key;
    card.append(textarea);
    new MarkdownEditor(textarea, {
      value: recommendation.adjusted || '',
      onChange: (value) => handlers.onRecommendationChange(key, value),
    });

    recommendationsColumn.append(card);
  });

  columns.append(recommendationsColumn);
  row.append(columns);

  return row;
}

function findNodeById(nodes, id) {
  for (const node of nodes) {
    if (node.id === id && node.isFinding) {
      return node;
    }
    if (node.children?.length) {
      const child = findNodeById(node.children, id);
      if (child) return child;
    }
  }
  return null;
}

function findNodeByIdAny(nodes, id) {
  for (const node of nodes) {
    if (node.id === id) {
      return node;
    }
    if (node.children?.length) {
      const child = findNodeByIdAny(node.children, id);
      if (child) return child;
    }
  }
  return null;
}

function findFirstFinding(nodes) {
  for (const node of nodes) {
    if (node.isFinding) return node;
    if (node.children?.length) {
      const nested = findFirstFinding(node.children);
      if (nested) return nested;
    }
  }
  return null;
}

function getFindingById(project, id) {
  if (!project?.report?.chapters) return null;
  return project.report.chapters.find((chapter) => chapter.id === id) || null;
}

function flattenFindings(node) {
  const results = [];
  if (!node) return results;
  const stack = Array.isArray(node) ? [...node] : [node];
  while (stack.length) {
    const current = stack.shift();
    if (!current) continue;
    if (current.isFinding && current.reportEntry) {
      results.push(current);
    }
    if (current.children?.length) {
      stack.unshift(...current.children);
    }
  }
  return results;
}
