import { renderMarkdown } from '../utils/markdownRenderer.js';
import { saveSession, isStorageAvailable } from '../storage/session.js';

export class AutoBerichtPanel {
  constructor({ state, treeContainer, detailContainer }) {
    this.state = state;
    this.treeContainer = treeContainer;
    this.detailContainer = detailContainer;
    this.selectedId = null;
    this.currentSnapshot = null;
    this.isEditing = false;
    this.pendingRender = false;
    this.editIdleTimer = null;
    this.filters = { status: 'all', hideExcluded: false, globalPreview: false };
    this.previewByFinding = new Map();
    this.editModeByFinding = new Map();
    this.photoModal = null;
    this.photoModalState = { photos: [], index: 0 };
    this.storageAvailable = isStorageAvailable();

    this.chapterStrip = document.getElementById('autobericht-strip');
    this.subStrip = document.getElementById('autobericht-substrip');
    this.drawer = document.getElementById('autobericht-drawer');
    this.drawerToggle = document.getElementById('autobericht-drawer-toggle');
    this.drawerClose = document.getElementById('autobericht-drawer-close');
    this.statusRadios = document.querySelectorAll('input[name="filter-status"]');
    this.hideExcludedToggle = document.getElementById('filter-hide-excluded');
    this.previewGlobalToggle = document.getElementById('filter-preview-global');
    this.saveStatusElement = document.getElementById('autobericht-save-status');
    this.saveNowButton = document.getElementById('autobericht-save-now');

    this.state.addEventListener('state:change', (event) => {
      this.currentSnapshot = event.detail;
      if (this.isEditing) {
        this.pendingRender = true;
        return;
      }
      this.ensureSelection();
      this.renderMainChapters();
      this.renderSubChapters();
      this.renderTree();
      this.renderDetail();
    });

    this.bindFilters();
    this.bindSaveNow();

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

    const filteredRows = this.filterRows(rows);
    const focusMeta = this.captureFocusMeta();
    this.detailContainer.innerHTML = '';

    if (!filteredRows.length) {
      const placeholder = document.createElement('div');
      placeholder.className = 'detail-placeholder';
      placeholder.innerHTML = '<p>No findings match the current filters.</p>';
      this.detailContainer.append(placeholder);
      return;
    }

    const table = document.createElement('div');
    table.className = 'finding-table';
    table.append(this.buildTableHead());

    const container = document.createElement('div');
    container.className = 'finding-row-container';

    filteredRows.forEach((entry) => {
      const photos = this.getPhotosForFinding(entry.id);
      const row = renderFindingRow(entry, {
        editEnabled: this.isEditEnabled(entry.id),
        onTogglePreview: (value) => this.setRowPreview(entry.id, value),
        onToggleEdit: () => this.toggleEdit(entry.id),
        onFindingChange: (text) => this.handleFindingChange(entry.id, text),
        onLevelChange: (level) => this.handleLevelChange(entry, level),
        onIncludeChange: (value) => this.state.setFindingInclude(entry.id, value),
        onRecommendationChange: (text) => this.handleRecommendationChange(entry, text),
        onDoneChange: (value) => this.state.setFindingDone(entry.id, value),
        onReviewChange: (value) => this.state.setFindingNeedsReview(entry.id, value),
        onMasterUpdateChange: (update) =>
          this.state.setProposeMasterUpdate(entry.id, update.action, update.scope),
        onPhotos: () => this.openPhotoModal(entry.id),
        hasPhotos: photos.length > 0,
        photoCount: photos.length,
      });
      container.append(row);
    });

    table.append(container);
    this.detailContainer.append(table);
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
      .map((node) => node.reportEntry)
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

  buildTableHead() {
    const head = document.createElement('div');
    head.className = 'finding-table__head';
    ['Finding', 'Report Text', 'Recommendation', 'Controls'].forEach((label) => {
      const cell = document.createElement('div');
      cell.className = 'finding-table__cell';
      cell.textContent = label;
      head.append(cell);
    });
    return head;
  }

  filterRows(rows) {
    return rows.filter((entry) => {
      if (this.filters.hideExcluded && entry.includeInReport === false) {
        return false;
      }
      if (this.filters.status === 'done' && !entry.done) {
        return false;
      }
      if (this.filters.status === 'todo' && entry.done) {
        return false;
      }
      return true;
    });
  }

  isPreviewEnabled(id) {
    return this.filters.globalPreview || this.previewByFinding.get(id) === true;
  }

  setRowPreview(id, enabled) {
    if (enabled) {
      this.previewByFinding.set(id, true);
    } else {
      this.previewByFinding.delete(id);
    }
    this.renderDetail();
  }

  isEditEnabled(id) {
    return this.editModeByFinding.get(id) === true;
  }

  toggleEdit(id) {
    if (this.isEditEnabled(id)) {
      this.editModeByFinding.delete(id);
    } else {
      this.editModeByFinding.set(id, true);
    }
    this.renderDetail();
  }

  handleFindingChange(id, text) {
    this.markEditing();
    this.state.updateFindingAdjusted(id, text);
    this.state.setFindingUseAdjusted?.(id, true);
  }

  handleRecommendationChange(entry, text) {
    this.markEditing();
    const levelKey = entry.level || '1';
    this.state.updateRecommendationAdjusted(entry.id, levelKey, text);
    this.state.setRecommendationUse(entry.id, levelKey, true);
  }

  markEditing() {
    this.isEditing = true;
    if (this.editIdleTimer) {
      clearTimeout(this.editIdleTimer);
    }
    this.editIdleTimer = setTimeout(() => {
      this.isEditing = false;
      this.editIdleTimer = null;
      if (this.pendingRender && this.currentSnapshot) {
        this.pendingRender = false;
        this.ensureSelection();
        this.renderMainChapters();
        this.renderSubChapters();
        this.renderTree();
        this.renderDetail();
      }
    }, 350);
  }

  handleLevelChange(entry, level) {
    this.state.updateFindingLevel(entry.id, level);
  }

  bindFilters() {
    this.statusRadios?.forEach((radio) => {
      radio.addEventListener('change', () => {
        if (!radio.checked) return;
        this.filters.status = radio.value;
        this.renderDetail();
      });
    });
    this.hideExcludedToggle?.addEventListener('change', () => {
      this.filters.hideExcluded = this.hideExcludedToggle.checked;
      this.renderDetail();
    });
    this.previewGlobalToggle?.addEventListener('change', () => {
      this.filters.globalPreview = this.previewGlobalToggle.checked;
      this.renderDetail();
    });
  }

  bindSaveNow() {
    this.saveNowButton?.addEventListener('click', () => this.handleSaveNow());
  }

  handleSaveNow() {
    const payload = this.state.getStoragePayload();
    if (!payload?.project) {
      this.setSaveStatus('Load project.json to save.');
      return;
    }
    if (this.storageAvailable) {
      try {
        saveSession(payload, { source: 'manual' });
        this.setSaveStatus('Saved locally.');
      } catch (error) {
        this.setSaveStatus('Local save failed.');
      }
    } else {
      this.setSaveStatus('Download only.');
    }
    this.downloadSnapshot(payload);
  }

  downloadSnapshot(payload) {
    if (!payload) return;
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${timestamp}-autobericht-session.json`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    this.setSaveStatus('Backup downloaded.');
  }

  setSaveStatus(message) {
    if (this.saveStatusElement) {
      this.saveStatusElement.textContent = message || '';
    }
  }

  getPhotosForFinding(findingId) {
    const photos = Array.isArray(this.currentSnapshot?.photos) ? this.currentSnapshot.photos : [];
    if (!photos.length || !findingId) return [];
    const target = findingId.toLowerCase();
    return photos.filter((photo) => (photo.path || '').toLowerCase().includes(target));
  }

  openPhotoModal(findingId) {
    const photos = this.getPhotosForFinding(findingId);
    if (!photos.length) {
      this.setSaveStatus('No photos for this finding.');
      return;
    }
    this.photoModalState = { photos, index: 0 };
    const modal = this.ensurePhotoModal();
    this.renderPhoto();
    modal.overlay.setAttribute('aria-hidden', 'false');
    modal.close?.focus();
  }

  closePhotoModal() {
    const modal = this.photoModal;
    if (!modal) return;
    modal.overlay.setAttribute('aria-hidden', 'true');
    if (modal.img && modal.img.src.startsWith('blob:')) {
      URL.revokeObjectURL(modal.img.src);
    }
  }

  renderPhoto() {
    const modal = this.ensurePhotoModal();
    const { photos, index } = this.photoModalState;
    if (!photos.length) {
      modal.caption.textContent = 'No photos available.';
      modal.img.removeAttribute('src');
      return;
    }
    const boundedIndex = Math.max(0, Math.min(index, photos.length - 1));
    this.photoModalState.index = boundedIndex;
    const photo = photos[boundedIndex];
    modal.caption.textContent = photo.path || `Photo ${boundedIndex + 1}`;
    modal.counter.textContent = `${boundedIndex + 1} / ${photos.length}`;
    if (photo.file instanceof File) {
      const objectUrl = URL.createObjectURL(photo.file);
      modal.img.src = objectUrl;
      modal.img.alt = photo.path || `Photo ${boundedIndex + 1}`;
    } else if (photo.path) {
      modal.img.src = photo.path;
      modal.img.alt = photo.path;
    } else {
      modal.img.removeAttribute('src');
    }
    modal.prev.disabled = boundedIndex === 0;
    modal.next.disabled = boundedIndex >= photos.length - 1;
  }

  ensurePhotoModal() {
    if (this.photoModal) return this.photoModal;
    const overlay = document.createElement('div');
    overlay.className = 'photo-modal';
    overlay.setAttribute('aria-hidden', 'true');
    const content = document.createElement('div');
    content.className = 'photo-modal__content';

    const header = document.createElement('div');
    header.className = 'photo-modal__header';
    const title = document.createElement('div');
    title.textContent = 'Photos';
    const counter = document.createElement('div');
    counter.className = 'photo-modal__counter';
    header.append(title, counter);

    const body = document.createElement('div');
    body.className = 'photo-modal__body';
    const img = document.createElement('img');
    img.alt = 'Photo preview';
    body.append(img);

    const caption = document.createElement('div');
    caption.className = 'photo-modal__caption';

    const controls = document.createElement('div');
    controls.className = 'photo-modal__controls';
    const prev = document.createElement('button');
    prev.type = 'button';
    prev.textContent = 'Prev';
    const next = document.createElement('button');
    next.type = 'button';
    next.textContent = 'Next';
    const close = document.createElement('button');
    close.type = 'button';
    close.textContent = 'Close';
    controls.append(prev, next, close);

    content.append(header, body, caption, controls);
    overlay.append(content);
    document.body.append(overlay);

    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) {
        this.closePhotoModal();
      }
    });
    close.addEventListener('click', () => this.closePhotoModal());
    prev.addEventListener('click', () => {
      this.photoModalState.index = Math.max(0, this.photoModalState.index - 1);
      this.renderPhoto();
    });
    next.addEventListener('click', () => {
      this.photoModalState.index = Math.min(
        this.photoModalState.photos.length - 1,
        this.photoModalState.index + 1,
      );
      this.renderPhoto();
    });

    document.addEventListener('keydown', (event) => {
      if (overlay.getAttribute('aria-hidden') === 'true') return;
      if (event.key === 'Escape') {
        this.closePhotoModal();
      } else if (event.key === 'ArrowLeft') {
        prev.click();
      } else if (event.key === 'ArrowRight') {
        next.click();
      }
    });

    this.photoModal = { overlay, content, img, caption, prev, next, close, counter };
    return this.photoModal;
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
    title.textContent = node.title ? `${node.id} — ${node.title}` : node.id;
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

function renderFindingRow(entry, handlers) {
  const row = document.createElement('article');
  row.className = 'finding-row';
  row.dataset.findingId = entry.id;
  if (!entry.includeInReport) {
    row.classList.add('finding-row--excluded');
  }
  if (entry.done) {
    row.classList.add('finding-row--done');
  } else if (entry.needsReview) {
    row.classList.add('finding-row--review');
  }

  const levelKey = String(entry.level || 1);
  const activeRecommendation = entry.recommendations?.[levelKey] || {
    master: '',
    adjusted: '',
    useAdjusted: false,
  };
  const recommendationDisplay =
    (activeRecommendation.useAdjusted && activeRecommendation.adjusted) ||
    activeRecommendation.adjusted ||
    activeRecommendation.master ||
    '';
  const recommendationText = activeRecommendation.adjusted || activeRecommendation.master || '';
  const findingDisplay =
    entry.finding.adjustedText ||
    entry.finding.masterText ||
    '';
  const findingText = entry.finding.adjustedText || entry.finding.masterText || '';
  const editOn = Boolean(handlers.editEnabled);

  const metaCell = document.createElement('div');
  metaCell.className = 'finding-row__meta';

  const title = document.createElement('p');
  title.className = 'finding-row__title';
  title.textContent = `${entry.id} — ${entry.title || ''}`;
  metaCell.append(title);

  const chip = document.createElement('span');
  chip.className = 'client-chip';
  chip.dataset.state = entry.clientInput.yesNo || 'n/a';
  chip.textContent = (entry.clientInput.yesNo || 'n/a').toUpperCase();
  metaCell.append(chip);

  if (entry.clientInput.remarks) {
    const remarkWrapper = document.createElement('div');
    remarkWrapper.className = 'client-remark';
    const remarkLabel = document.createElement('label');
    remarkLabel.textContent = 'Client comment';
    const remarkBody = document.createElement('div');
    remarkBody.className = 'preview-pane preview-pane--compact';
    remarkBody.innerHTML = renderMarkdown(entry.clientInput.remarks || '');
    remarkWrapper.append(remarkLabel, remarkBody);
    metaCell.append(remarkWrapper);
  }

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
    const visual = document.createElement('span');
    visual.textContent = value;
    label.append(radio, visual);
    levelGroup.append(label);
  });
  metaCell.append(levelGroup);

  const findingCell = document.createElement('div');
  findingCell.className = 'finding-row__cell finding-row__cell--finding';
  const adjustedLabel = document.createElement('label');
  adjustedLabel.textContent = 'Finding (Markdown)';
  findingCell.append(adjustedLabel);
  if (editOn) {
    const adjustArea = document.createElement('textarea');
    adjustArea.className = 'markdown-textarea';
    adjustArea.dataset.field = 'finding-adjusted';
    adjustArea.dataset.finding = entry.id;
    adjustArea.value = findingText;
    adjustArea.addEventListener('input', () => handlers.onFindingChange(adjustArea.value));
    findingCell.append(adjustArea);
  } else {
    const findingPreview = document.createElement('div');
    findingPreview.className = 'preview-pane preview-pane--compact';
    findingPreview.innerHTML = renderMarkdown(findingDisplay);
    findingCell.append(findingPreview);
  }

  const recommendationCell = document.createElement('div');
  recommendationCell.className = 'finding-row__cell finding-row__cell--recommendation';
  const recLabel = document.createElement('label');
  recLabel.textContent = `Recommendation (Level ${levelKey})`;
  recommendationCell.append(recLabel);

  if (editOn) {
    const recTextarea = document.createElement('textarea');
    recTextarea.className = 'markdown-textarea';
    recTextarea.dataset.field = 'recommendation-adjusted';
    recTextarea.dataset.finding = entry.id;
    recTextarea.dataset.recommendation = levelKey;
    recTextarea.value = recommendationText;
    recTextarea.addEventListener('input', () => handlers.onRecommendationChange(recTextarea.value));
    recommendationCell.append(recTextarea);
  } else {
    const recPreview = document.createElement('div');
    recPreview.className = 'preview-pane preview-pane--compact';
    recPreview.innerHTML = renderMarkdown(recommendationDisplay || '');
    recommendationCell.append(recPreview);
  }

  const controlsCell = document.createElement('div');
  controlsCell.className = 'finding-row__cell finding-row__controls';

  const includeLabel = document.createElement('label');
  includeLabel.className = 'checkbox-row';
  const includeCheckbox = document.createElement('input');
  includeCheckbox.type = 'checkbox';
  includeCheckbox.checked = Boolean(entry.includeInReport);
  includeCheckbox.addEventListener('change', () => handlers.onIncludeChange(includeCheckbox.checked));
  includeLabel.append(includeCheckbox, document.createTextNode('Include'));
  controlsCell.append(includeLabel);

  const doneLabel = document.createElement('label');
  doneLabel.className = 'checkbox-row';
  const doneCheckbox = document.createElement('input');
  doneCheckbox.type = 'checkbox';
  doneCheckbox.checked = Boolean(entry.done);
  doneCheckbox.addEventListener('change', () => handlers.onDoneChange(doneCheckbox.checked));
  doneLabel.append(doneCheckbox, document.createTextNode('Done'));
  controlsCell.append(doneLabel);

  const reviewLabel = document.createElement('label');
  reviewLabel.className = 'checkbox-row';
  const reviewCheckbox = document.createElement('input');
  reviewCheckbox.type = 'checkbox';
  reviewCheckbox.checked = Boolean(entry.needsReview);
  reviewCheckbox.addEventListener('change', () => handlers.onReviewChange(reviewCheckbox.checked));
  reviewLabel.append(reviewCheckbox, document.createTextNode('Needs review'));
  controlsCell.append(reviewLabel);

  const intentWrapper = document.createElement('div');
  intentWrapper.className = 'control-group';
  const intentLabel = document.createElement('div');
  intentLabel.className = 'control-group__label';
  intentLabel.textContent = 'Master intent';
  intentWrapper.append(intentLabel);
  const intents = [
    { value: 'none', label: 'None', scope: '' },
    { value: 'append', label: 'Append', scope: 'all' },
    { value: 'replace', label: 'Replace', scope: 'all' },
  ];
  intents.forEach((intent) => {
    const intentRow = document.createElement('label');
    intentRow.className = 'checkbox-row';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = `intent-${entry.id}`;
    radio.value = intent.value;
    radio.checked = (entry.proposeMasterUpdate?.action || 'none') === intent.value;
    radio.addEventListener('change', () =>
      handlers.onMasterUpdateChange({ action: intent.value, scope: intent.scope }),
    );
    intentRow.append(radio, document.createTextNode(intent.label));
    intentWrapper.append(intentRow);
  });
  controlsCell.append(intentWrapper);

  const photosButton = document.createElement('button');
  photosButton.type = 'button';
  photosButton.className = handlers.hasPhotos ? 'chip-button chip-button--active' : 'chip-button';
  photosButton.textContent = handlers.hasPhotos
    ? `Photos (${handlers.photoCount || 0})`
    : 'Photos';
  photosButton.addEventListener('click', () => handlers.onPhotos?.());
  controlsCell.append(photosButton);

  const editButton = document.createElement('button');
  editButton.type = 'button';
  editButton.className = 'chip-button';
  editButton.textContent = editOn ? 'Hide editors' : 'Edit';
  editButton.addEventListener('click', () => handlers.onToggleEdit());
  controlsCell.append(editButton);

  row.append(metaCell, findingCell, recommendationCell, controlsCell);

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
  if (!project) return null;
  if (Array.isArray(project.chapters)) {
    for (const chapter of project.chapters) {
      const rows = Array.isArray(chapter?.rows) ? chapter.rows : [];
      const row = rows.find((entry) => entry?.id === id);
      if (row) return row;
    }
  }
  return null;
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
