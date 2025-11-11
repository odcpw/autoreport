/**
 * ImageViewer displays a large preview of a photo.
 * Manages object URLs for File objects and handles cleanup to prevent memory leaks.
 */
export class ImageViewer {
  constructor({ container }) {
    this.container = container;
    this.currentFile = null;
    this.objectUrl = null;
    this.imgElement = null;

    this.render();
  }

  render() {
    if (!this.container) return;

    this.container.innerHTML = '';
    this.container.className = 'image-viewer';

    this.imgElement = document.createElement('img');
    this.imgElement.className = 'image-viewer__img';
    this.imgElement.alt = 'Photo preview';

    this.container.append(this.imgElement);
  }

  /**
   * Display a photo from a File object or show placeholder.
   * @param {File|null} file - The image file to display
   * @param {string} path - The path/name of the file for fallback display
   */
  setPhoto(file, path = '') {
    if (this.objectUrl) {
      URL.revokeObjectURL(this.objectUrl);
      this.objectUrl = null;
    }

    this.currentFile = file;

    if (!this.imgElement) {
      this.render();
    }

    if (file && file instanceof File) {
      this.objectUrl = URL.createObjectURL(file);
      this.imgElement.src = this.objectUrl;
      this.imgElement.alt = path || 'Photo preview';
      this.imgElement.style.display = 'block';
    } else {
      this.imgElement.src = '';
      this.imgElement.alt = path || 'No preview available';
      this.imgElement.style.display = 'none';
    }
  }

  /**
   * Clear the current image and clean up resources.
   */
  clear() {
    if (this.objectUrl) {
      URL.revokeObjectURL(this.objectUrl);
      this.objectUrl = null;
    }

    if (this.imgElement) {
      this.imgElement.src = '';
      this.imgElement.style.display = 'none';
    }

    this.currentFile = null;
  }

  /**
   * Clean up resources when component is destroyed.
   */
  destroy() {
    this.clear();
    if (this.container) {
      this.container.innerHTML = '';
    }
  }
}
