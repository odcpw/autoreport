import { ImageViewer } from './ImageViewer.js';
import { ThumbnailCarousel } from './ThumbnailCarousel.js';
import { ButtonPanel } from './ButtonPanel.js';
import { PhotoSorterState } from './PhotoSorterState.js';

/**
 * PhotoSorterPanel orchestrates the photo tagging interface.
 * Displays one photo at a time with thumbnail navigation and tag button panels.
 * Supports keyboard shortcuts: A/D for navigation, arrow keys for button navigation.
 */
export class PhotoSorterPanel {
  constructor({ state, container }) {
    this.state = state;
    this.container = container;
    this.localState = new PhotoSorterState();

    this.imageViewer = null;
    this.thumbnailCarousel = null;
    this.berichtPanel = null;
    this.seminarPanel = null;
    this.topicPanel = null;

    this.currentSnapshot = null;

    this.handleGlobalKeyDown = this.handleGlobalKeyDown.bind(this);
    this.renderInitialUI();

    this.state.addEventListener('state:change', (event) => {
      this.handleStateChange(event.detail);
    });

    window.addEventListener('keydown', this.handleGlobalKeyDown);
  }

  renderInitialUI() {
    if (!this.container) return;

    this.container.innerHTML = '';
    this.container.className = 'photosorter';

    const controls = document.createElement('div');
    controls.className = 'photosorter__controls';

    const filterToggle = document.createElement('button');
    filterToggle.type = 'button';
    filterToggle.className = 'photosorter__filter-toggle';
    filterToggle.textContent = 'Show All';
    filterToggle.addEventListener('click', () => this.toggleFilter());
    this.filterToggleButton = filterToggle;

    const counter = document.createElement('div');
    counter.className = 'photosorter__counter';
    counter.textContent = 'No photos';
    this.counterElement = counter;

    const prevButton = document.createElement('button');
    prevButton.type = 'button';
    prevButton.className = 'photosorter__nav-button';
    prevButton.textContent = 'Previous (A)';
    prevButton.addEventListener('click', () => this.navigatePrevious());
    this.prevButton = prevButton;

    const nextButton = document.createElement('button');
    nextButton.type = 'button';
    nextButton.className = 'photosorter__nav-button';
    nextButton.textContent = 'Next (D)';
    nextButton.addEventListener('click', () => this.navigateNext());
    this.nextButton = nextButton;

    controls.append(filterToggle, counter, prevButton, nextButton);

    const mainLayout = document.createElement('div');
    mainLayout.className = 'photosorter__layout';

    const viewerSection = document.createElement('div');
    viewerSection.className = 'photosorter__viewer-section';

    const viewerContainer = document.createElement('div');
    viewerContainer.className = 'photosorter__viewer-container';
    this.viewerContainer = viewerContainer;

    const carouselContainer = document.createElement('div');
    carouselContainer.className = 'photosorter__carousel-container';
    this.carouselContainer = carouselContainer;

    viewerSection.append(viewerContainer, carouselContainer);

    const sidebar = document.createElement('div');
    sidebar.className = 'photosorter__sidebar';

    const berichtContainer = document.createElement('div');
    berichtContainer.className = 'photosorter__panel-container';
    this.berichtContainer = berichtContainer;

    const seminarContainer = document.createElement('div');
    seminarContainer.className = 'photosorter__panel-container';
    this.seminarContainer = seminarContainer;

    const topicContainer = document.createElement('div');
    topicContainer.className = 'photosorter__panel-container';
    this.topicContainer = topicContainer;

    const notesContainer = document.createElement('div');
    notesContainer.className = 'photosorter__notes-container';

    const notesLabel = document.createElement('label');
    notesLabel.className = 'photosorter__notes-label';
    notesLabel.textContent = 'Notes';

    const notesTextarea = document.createElement('textarea');
    notesTextarea.className = 'photosorter__notes';
    notesTextarea.rows = 3;
    notesTextarea.placeholder = 'Add notes for this photo…';
    notesTextarea.addEventListener('change', () => this.handleNotesChange());
    this.notesTextarea = notesTextarea;

    notesContainer.append(notesLabel, notesTextarea);

    sidebar.append(berichtContainer, seminarContainer, topicContainer, notesContainer);

    mainLayout.append(viewerSection, sidebar);

    this.container.append(controls, mainLayout);

    this.imageViewer = new ImageViewer({ container: viewerContainer });
    this.thumbnailCarousel = new ThumbnailCarousel({
      container: carouselContainer,
      onThumbnailClick: (index) => this.jumpToPhoto(index),
    });

    this.renderEmpty();
  }

  renderEmpty() {
    if (this.imageViewer) {
      this.imageViewer.clear();
    }

    if (this.thumbnailCarousel) {
      this.thumbnailCarousel.setPhotos([], 0);
    }

    this.counterElement.textContent = 'No photos imported yet. Upload a directory in the Import tab.';
    this.prevButton.disabled = true;
    this.nextButton.disabled = true;
  }

  handleStateChange(snapshot) {
    this.currentSnapshot = snapshot;

    const photos = snapshot?.photos || [];
    const tagOptions = snapshot?.tagOptions || {};

    if (photos.length === 0) {
      this.renderEmpty();
      return;
    }

    this.localState.setPhotos(photos);

    if (!this.berichtPanel) {
      this.createButtonPanels(tagOptions);
    } else {
      this.updateButtonPanels(tagOptions);
    }

    this.renderCurrentPhoto();
    this.updateControls();
  }

  createButtonPanels(tagOptions) {
    this.berichtPanel = new ButtonPanel({
      container: this.berichtContainer,
      title: 'Bericht',
      description: 'Tag photo with relevant report sections',
      options: tagOptions.bericht || [],
      selectedValues: this.getCurrentPhotoTags('bericht'),
      onToggle: (value) => this.handleBerichtToggle(value),
    });

    this.seminarPanel = new ButtonPanel({
      container: this.seminarContainer,
      title: 'Seminar',
      description: 'Classify photo for seminars',
      options: tagOptions.seminar || [],
      selectedValues: this.getCurrentPhotoTags('seminar'),
      onToggle: (value) => this.handleSeminarToggle(value),
    });

    this.topicPanel = new ButtonPanel({
      container: this.topicContainer,
      title: 'Topic',
      description: 'Tag photo for topic folders',
      options: tagOptions.topic || [],
      selectedValues: this.getCurrentPhotoTags('topic'),
      onToggle: (value) => this.handleTopicToggle(value),
    });
  }

  updateButtonPanels(tagOptions) {
    if (this.berichtPanel) {
      this.berichtPanel.setOptions(tagOptions.bericht || []);
    }

    if (this.seminarPanel) {
      this.seminarPanel.setOptions(tagOptions.seminar || []);
    }

    if (this.topicPanel) {
      this.topicPanel.setOptions(tagOptions.topic || []);
    }
  }

  getCurrentPhotoTags(group) {
    const currentPhoto = this.localState.getCurrentPhoto();
    if (!currentPhoto || !currentPhoto.tags) return [];
    return currentPhoto.tags[group] || [];
  }

  renderCurrentPhoto() {
    const currentPhoto = this.localState.getCurrentPhoto();
    const currentIndex = this.localState.getCurrentIndex();
    const filteredPhotos = this.localState.getFilteredPhotos();

    if (!currentPhoto) {
      this.renderEmpty();
      return;
    }

    if (this.imageViewer) {
      this.imageViewer.setPhoto(currentPhoto.file, currentPhoto.path);
    }

    if (this.thumbnailCarousel) {
      const photosForCarousel = filteredPhotos.map(p => ({ file: p.file, path: p.path }));
      this.thumbnailCarousel.setPhotos(photosForCarousel, currentIndex);
    }

    if (this.notesTextarea) {
      this.notesTextarea.value = currentPhoto.notes || '';
    }

    this.updateButtonSelections();
    this.updateCounter();
  }

  updateButtonSelections() {
    if (this.berichtPanel) {
      this.berichtPanel.setSelectedValues(this.getCurrentPhotoTags('bericht'));
    }

    if (this.seminarPanel) {
      this.seminarPanel.setSelectedValues(this.getCurrentPhotoTags('seminar'));
    }

    if (this.topicPanel) {
      this.topicPanel.setSelectedValues(this.getCurrentPhotoTags('topic'));
    }
  }

  updateControls() {
    const stats = this.localState.getStatistics();
    const hasPhotos = stats.filtered > 0;

    this.prevButton.disabled = !hasPhotos;
    this.nextButton.disabled = !hasPhotos;
  }

  updateCounter() {
    const stats = this.localState.getStatistics();

    if (stats.filtered === 0) {
      this.counterElement.textContent = 'No photos match filter';
      return;
    }

    this.counterElement.textContent =
      `Image ${stats.current} of ${stats.filtered} - Total: ${stats.total} - Unsorted: ${stats.unsorted}`;
  }

  toggleFilter() {
    const currentMode = this.localState.filterMode;
    const newMode = currentMode === 'all' ? 'unsorted' : 'all';

    this.localState.setFilterMode(newMode);
    this.filterToggleButton.textContent = newMode === 'all' ? 'Show All' : 'Show Unsorted';

    this.renderCurrentPhoto();
    this.updateControls();
  }

  navigateNext() {
    this.localState.nextPhoto();
    this.renderCurrentPhoto();
  }

  navigatePrevious() {
    this.localState.previousPhoto();
    this.renderCurrentPhoto();
  }

  jumpToPhoto(index) {
    this.localState.goToIndex(index);
    this.renderCurrentPhoto();
  }

  handleBerichtToggle(value) {
    const currentPhoto = this.localState.getCurrentPhoto();
    if (!currentPhoto) return;

    this.state.togglePhotoTag(currentPhoto.path, 'bericht', value);
  }

  handleSeminarToggle(value) {
    const currentPhoto = this.localState.getCurrentPhoto();
    if (!currentPhoto) return;

    this.state.togglePhotoTag(currentPhoto.path, 'seminar', value);
  }

  handleTopicToggle(value) {
    const currentPhoto = this.localState.getCurrentPhoto();
    if (!currentPhoto) return;

    this.state.togglePhotoTag(currentPhoto.path, 'topic', value);
  }

  handleNotesChange() {
    const currentPhoto = this.localState.getCurrentPhoto();
    if (!currentPhoto) return;

    const notes = this.notesTextarea.value;
    this.state.updatePhotoNotes(currentPhoto.path, notes);
  }

  handleGlobalKeyDown(event) {
    const target = event.target;
    const isInput = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement;

    if (isInput) return;
    if (!this.isActivePanel()) return;

    const stats = this.localState.getStatistics();
    if (stats.filtered === 0) return;

    switch (event.key.toLowerCase()) {
      case 'a':
        event.preventDefault();
        this.navigatePrevious();
        break;

      case 'd':
        event.preventDefault();
        this.navigateNext();
        break;

      default:
        break;
    }
  }

  isActivePanel() {
    const panel = this.container?.closest('.tab-panel');
    return Boolean(panel?.classList.contains('tab-panel--active'));
  }

  destroy() {
    window.removeEventListener('keydown', this.handleGlobalKeyDown);

    if (this.imageViewer) {
      this.imageViewer.destroy();
    }

    if (this.thumbnailCarousel) {
      this.thumbnailCarousel.destroy();
    }

    if (this.berichtPanel) {
      this.berichtPanel.destroy();
    }

    if (this.seminarPanel) {
      this.seminarPanel.destroy();
    }

    if (this.topicPanel) {
      this.topicPanel.destroy();
    }

    if (this.container) {
      this.container.innerHTML = '';
    }
  }
}
