/**
 * PhotoSorterState manages local UI state for the PhotoSorter panel.
 * Handles current photo index, filtering, and photo navigation.
 */
export class PhotoSorterState {
  constructor() {
    this.currentIndex = 0;
    this.filterMode = 'all'; // 'all' or 'unsorted'
    this.allPhotos = [];
    this.filteredPhotos = [];
  }

  /**
   * Update the photo list and recompute filtered list.
   * @param {Array} photos - Array of photo objects from projectState
   */
  setPhotos(photos) {
    this.allPhotos = photos || [];
    this.recomputeFiltered();

    if (this.currentIndex >= this.filteredPhotos.length) {
      this.currentIndex = Math.max(0, this.filteredPhotos.length - 1);
    }
  }

  /**
   * Recompute the filtered photo list based on current filter mode.
   */
  recomputeFiltered() {
    if (this.filterMode === 'unsorted') {
      this.filteredPhotos = this.allPhotos.filter(photo => this.isPhotoUnsorted(photo));
    } else {
      this.filteredPhotos = this.allPhotos;
    }
  }

  /**
   * Check if a photo has no tags.
   * @param {Object} photo - Photo object with tags property
   * @returns {boolean} True if photo has no tags in any category
   */
  isPhotoUnsorted(photo) {
    if (!photo || !photo.tags) return true;

    const { bericht, seminar, topic } = photo.tags;
    const hasBericht = Array.isArray(bericht) && bericht.length > 0;
    const hasSeminar = Array.isArray(seminar) && seminar.length > 0;
    const hasTopic = Array.isArray(topic) && topic.length > 0;

    return !hasBericht && !hasSeminar && !hasTopic;
  }

  /**
   * Get the current photo object.
   * @returns {Object|null} The current photo or null if none available
   */
  getCurrentPhoto() {
    if (this.filteredPhotos.length === 0) return null;
    return this.filteredPhotos[this.currentIndex] || null;
  }

  /**
   * Navigate to the next photo.
   * @returns {Object|null} The new current photo
   */
  nextPhoto() {
    if (this.filteredPhotos.length === 0) return null;
    this.currentIndex = (this.currentIndex + 1) % this.filteredPhotos.length;
    return this.getCurrentPhoto();
  }

  /**
   * Navigate to the previous photo.
   * @returns {Object|null} The new current photo
   */
  previousPhoto() {
    if (this.filteredPhotos.length === 0) return null;
    this.currentIndex = (this.currentIndex - 1 + this.filteredPhotos.length) % this.filteredPhotos.length;
    return this.getCurrentPhoto();
  }

  /**
   * Jump to a specific photo index.
   * @param {number} index - The target index
   * @returns {Object|null} The photo at the target index
   */
  goToIndex(index) {
    if (this.filteredPhotos.length === 0) return null;
    if (index < 0 || index >= this.filteredPhotos.length) return null;
    this.currentIndex = index;
    return this.getCurrentPhoto();
  }

  /**
   * Set the filter mode and recompute filtered photos.
   * @param {string} mode - 'all' or 'unsorted'
   */
  setFilterMode(mode) {
    if (mode !== 'all' && mode !== 'unsorted') return;
    this.filterMode = mode;
    this.recomputeFiltered();

    if (this.currentIndex >= this.filteredPhotos.length) {
      this.currentIndex = Math.max(0, this.filteredPhotos.length - 1);
    }
  }

  /**
   * Get statistics about the photo collection.
   * @returns {Object} Statistics object with total, filtered, and unsorted counts
   */
  getStatistics() {
    const total = this.allPhotos.length;
    const filtered = this.filteredPhotos.length;
    const unsorted = this.allPhotos.filter(photo => this.isPhotoUnsorted(photo)).length;

    return {
      total,
      filtered,
      unsorted,
      current: this.currentIndex + 1,
    };
  }

  /**
   * Get the filtered photo list.
   * @returns {Array} The filtered photo array
   */
  getFilteredPhotos() {
    return this.filteredPhotos;
  }

  /**
   * Get the current index.
   * @returns {number} The current photo index
   */
  getCurrentIndex() {
    return this.currentIndex;
  }

  /**
   * Reset state to initial values.
   */
  reset() {
    this.currentIndex = 0;
    this.filterMode = 'all';
    this.allPhotos = [];
    this.filteredPhotos = [];
  }
}
