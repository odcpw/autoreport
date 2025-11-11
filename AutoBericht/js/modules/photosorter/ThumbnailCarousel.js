/**
 * ThumbnailCarousel displays a horizontal strip of photo thumbnails.
 * Users can click thumbnails to navigate between photos.
 * Manages object URLs for thumbnail display and cleanup.
 */
export class ThumbnailCarousel {
  constructor({ container, onThumbnailClick }) {
    this.container = container;
    this.onThumbnailClick = onThumbnailClick;
    this.photos = [];
    this.currentIndex = -1;
    this.thumbnailUrls = new Map();

    this.render();
  }

  render() {
    if (!this.container) return;

    this.container.innerHTML = '';
    this.container.className = 'thumbnail-carousel';

    const scrollContainer = document.createElement('div');
    scrollContainer.className = 'thumbnail-carousel__scroll';
    this.scrollContainer = scrollContainer;

    this.container.append(scrollContainer);
  }

  /**
   * Update the list of photos and regenerate thumbnails.
   * @param {Array<{file: File|null, path: string}>} photos - Array of photo objects
   * @param {number} activeIndex - Currently active photo index
   */
  setPhotos(photos, activeIndex = 0) {
    this.clearThumbnails();

    this.photos = photos || [];
    this.currentIndex = activeIndex;

    if (!this.scrollContainer) {
      this.render();
    }

    this.scrollContainer.innerHTML = '';

    if (this.photos.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'thumbnail-carousel__empty';
      empty.textContent = 'No photos';
      this.scrollContainer.append(empty);
      return;
    }

    this.photos.forEach((photo, index) => {
      const thumb = document.createElement('button');
      thumb.type = 'button';
      thumb.className = 'thumbnail-carousel__item';
      thumb.dataset.index = index;

      if (index === activeIndex) {
        thumb.classList.add('thumbnail-carousel__item--active');
      }

      if (photo.file && photo.file instanceof File) {
        const objectUrl = URL.createObjectURL(photo.file);
        this.thumbnailUrls.set(index, objectUrl);

        const img = document.createElement('img');
        img.src = objectUrl;
        img.alt = photo.path || `Photo ${index + 1}`;
        img.className = 'thumbnail-carousel__img';
        thumb.append(img);
      } else {
        thumb.textContent = photo.path ? photo.path.split('/').pop() : `Photo ${index + 1}`;
        thumb.classList.add('thumbnail-carousel__item--no-preview');
      }

      thumb.addEventListener('click', () => {
        if (this.onThumbnailClick) {
          this.onThumbnailClick(index);
        }
      });

      this.scrollContainer.append(thumb);
    });

    this.scrollToActive();
  }

  /**
   * Update the active thumbnail index.
   * @param {number} index - The new active index
   */
  setActiveIndex(index) {
    if (index === this.currentIndex) return;

    const items = this.scrollContainer?.querySelectorAll('.thumbnail-carousel__item');
    if (!items) return;

    items.forEach((item, i) => {
      if (i === index) {
        item.classList.add('thumbnail-carousel__item--active');
      } else {
        item.classList.remove('thumbnail-carousel__item--active');
      }
    });

    this.currentIndex = index;
    this.scrollToActive();
  }

  /**
   * Scroll the active thumbnail into view.
   */
  scrollToActive() {
    if (!this.scrollContainer) return;

    const activeItem = this.scrollContainer.querySelector('.thumbnail-carousel__item--active');
    if (activeItem) {
      activeItem.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
        inline: 'center',
      });
    }
  }

  /**
   * Clear all thumbnail object URLs.
   */
  clearThumbnails() {
    this.thumbnailUrls.forEach((url) => {
      URL.revokeObjectURL(url);
    });
    this.thumbnailUrls.clear();
  }

  /**
   * Clean up resources when component is destroyed.
   */
  destroy() {
    this.clearThumbnails();
    if (this.container) {
      this.container.innerHTML = '';
    }
  }
}
