/**
 * ButtonPanel displays a grid of toggle buttons with keyboard navigation.
 * Supports multi-select and roving tabindex pattern for accessibility.
 * Arrow keys navigate, Space toggles buttons.
 */
export class ButtonPanel {
  constructor({ container, title, description, options, selectedValues, onToggle }) {
    this.container = container;
    this.title = title;
    this.description = description;
    this.options = options || [];
    this.selectedValues = new Set(selectedValues || []);
    this.onToggle = onToggle;
    this.currentFocusIndex = 0;
    this.buttons = [];

    this.handleKeyDown = this.handleKeyDown.bind(this);

    this.render();
  }

  render() {
    if (!this.container) return;

    this.container.innerHTML = '';
    this.container.className = 'button-panel';

    if (this.title) {
      const header = document.createElement('div');
      header.className = 'button-panel__header';

      const titleElement = document.createElement('h4');
      titleElement.className = 'button-panel__title';
      titleElement.textContent = this.title;
      header.append(titleElement);

      if (this.description) {
        const desc = document.createElement('p');
        desc.className = 'button-panel__description';
        desc.textContent = this.description;
        header.append(desc);
      }

      this.container.append(header);
    }

    const grid = document.createElement('div');
    grid.className = 'button-panel__grid';
    grid.setAttribute('role', 'group');
    grid.setAttribute('aria-label', this.title || 'Button group');
    this.gridElement = grid;

    this.renderButtons();

    this.container.append(grid);
  }

  renderButtons() {
    if (!this.gridElement) return;

    this.gridElement.innerHTML = '';
    this.buttons = [];

    if (this.options.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'button-panel__empty';
      empty.textContent = 'No options available';
      this.gridElement.append(empty);
      return;
    }

    this.options.forEach((option, index) => {
      const value = typeof option === 'string' ? option : option.value;
      const label = typeof option === 'string' ? option : option.label || option.value;

      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'button-panel__button';
      button.textContent = label;
      button.dataset.value = value;
      button.dataset.index = index;

      const isActive = this.selectedValues.has(value);
      if (isActive) {
        button.classList.add('button-panel__button--active');
        button.setAttribute('aria-pressed', 'true');
      } else {
        button.setAttribute('aria-pressed', 'false');
      }

      button.tabIndex = index === this.currentFocusIndex ? 0 : -1;

      button.addEventListener('click', () => {
        this.toggleButton(value);
      });

      button.addEventListener('keydown', this.handleKeyDown);

      button.addEventListener('focus', () => {
        this.currentFocusIndex = index;
        this.updateTabIndices();
      });

      this.buttons.push(button);
      this.gridElement.append(button);
    });
  }

  /**
   * Handle keyboard navigation within the button grid.
   * @param {KeyboardEvent} event - The keyboard event
   */
  handleKeyDown(event) {
    const { key } = event;

    switch (key) {
      case 'ArrowRight':
      case 'ArrowDown':
        event.preventDefault();
        this.focusNextButton();
        break;

      case 'ArrowLeft':
      case 'ArrowUp':
        event.preventDefault();
        this.focusPreviousButton();
        break;

      case ' ':
      case 'Enter':
        event.preventDefault();
        const value = event.target.dataset.value;
        if (value) {
          this.toggleButton(value);
        }
        break;

      case 'Home':
        event.preventDefault();
        this.focusFirstButton();
        break;

      case 'End':
        event.preventDefault();
        this.focusLastButton();
        break;

      default:
        break;
    }
  }

  focusNextButton() {
    const nextIndex = (this.currentFocusIndex + 1) % this.buttons.length;
    this.buttons[nextIndex]?.focus();
  }

  focusPreviousButton() {
    const prevIndex = (this.currentFocusIndex - 1 + this.buttons.length) % this.buttons.length;
    this.buttons[prevIndex]?.focus();
  }

  focusFirstButton() {
    this.buttons[0]?.focus();
  }

  focusLastButton() {
    this.buttons[this.buttons.length - 1]?.focus();
  }

  updateTabIndices() {
    this.buttons.forEach((button, index) => {
      button.tabIndex = index === this.currentFocusIndex ? 0 : -1;
    });
  }

  /**
   * Toggle a button's active state.
   * @param {string} value - The value to toggle
   */
  toggleButton(value) {
    const isActive = this.selectedValues.has(value);

    if (isActive) {
      this.selectedValues.delete(value);
    } else {
      this.selectedValues.add(value);
    }

    this.updateButtonState(value, !isActive);

    if (this.onToggle) {
      this.onToggle(value, !isActive);
    }
  }

  updateButtonState(value, isActive) {
    const button = this.buttons.find(b => b.dataset.value === value);
    if (!button) return;

    if (isActive) {
      button.classList.add('button-panel__button--active');
      button.setAttribute('aria-pressed', 'true');
    } else {
      button.classList.remove('button-panel__button--active');
      button.setAttribute('aria-pressed', 'false');
    }
  }

  /**
   * Update the available options and re-render.
   * @param {Array<string|object>} options - New option list
   */
  setOptions(options) {
    this.options = options || [];
    this.renderButtons();
  }

  /**
   * Update the selected values.
   * @param {Array<string>} values - Array of selected values
   */
  setSelectedValues(values) {
    this.selectedValues = new Set(values || []);
    this.buttons.forEach((button) => {
      const value = button.dataset.value;
      const isActive = this.selectedValues.has(value);
      this.updateButtonState(value, isActive);
    });
  }

  /**
   * Get the current selected values.
   * @returns {Array<string>} Array of selected values
   */
  getSelectedValues() {
    return Array.from(this.selectedValues);
  }

  /**
   * Clear all selections.
   */
  clearSelection() {
    this.selectedValues.clear();
    this.buttons.forEach((button) => {
      const value = button.dataset.value;
      this.updateButtonState(value, false);
    });
  }

  /**
   * Clean up event listeners.
   */
  destroy() {
    this.buttons.forEach((button) => {
      button.removeEventListener('keydown', this.handleKeyDown);
    });

    if (this.container) {
      this.container.innerHTML = '';
    }
  }
}
