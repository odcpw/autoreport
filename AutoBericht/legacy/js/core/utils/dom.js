/**
 * DOM utility functions for creating and manipulating DOM elements.
 * Reduces boilerplate in UI rendering code.
 */

/**
 * Create a DOM element with optional className and content.
 *
 * @param {string} tag - The HTML tag name
 * @param {string} [className] - CSS class name(s)
 * @param {string} [textContent] - Text content for the element
 * @returns {HTMLElement} The created element
 */
export function createElement(tag, className = '', textContent = '') {
  const element = document.createElement(tag);
  if (className) {
    element.className = className;
  }
  if (textContent) {
    element.textContent = textContent;
  }
  return element;
}

/**
 * Create a button element with text and click handler.
 *
 * @param {string} text - Button text
 * @param {Function} onClick - Click handler function
 * @param {string} [className] - CSS class name(s)
 * @returns {HTMLButtonElement} The created button
 */
export function createButton(text, onClick, className = '') {
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = text;
  if (className) {
    button.className = className;
  }
  if (onClick) {
    button.addEventListener('click', onClick);
  }
  return button;
}

/**
 * Create a label with an input element.
 *
 * @param {string} labelText - The label text
 * @param {HTMLElement} inputElement - The input element to wrap
 * @param {string} [className] - CSS class name(s) for the label
 * @returns {HTMLLabelElement} The created label
 */
export function createLabel(labelText, inputElement, className = '') {
  const label = document.createElement('label');
  if (className) {
    label.className = className;
  }
  const span = document.createElement('span');
  span.textContent = labelText;
  label.append(span, inputElement);
  return label;
}

/**
 * Create a checkbox with label.
 *
 * @param {string} labelText - The checkbox label text
 * @param {boolean} checked - Initial checked state
 * @param {Function} onChange - Change handler function
 * @param {string} [className] - CSS class name(s)
 * @returns {HTMLLabelElement} The created label containing the checkbox
 */
export function createCheckbox(labelText, checked, onChange, className = 'checkbox-row') {
  const label = document.createElement('label');
  label.className = className;

  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.checked = checked;
  if (onChange) {
    checkbox.addEventListener('change', () => onChange(checkbox.checked));
  }

  const span = document.createElement('span');
  span.textContent = labelText;

  label.append(checkbox, span);
  return label;
}

/**
 * Create a select dropdown with options.
 *
 * @param {Array<{value: string, label: string}>} options - Option definitions
 * @param {string} selectedValue - The initially selected value
 * @param {Function} onChange - Change handler function
 * @param {string} [className] - CSS class name(s)
 * @returns {HTMLSelectElement} The created select element
 */
export function createSelect(options, selectedValue, onChange, className = '') {
  const select = document.createElement('select');
  if (className) {
    select.className = className;
  }

  options.forEach(({ value, label }) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    if (value === selectedValue) {
      option.selected = true;
    }
    select.append(option);
  });

  if (onChange) {
    select.addEventListener('change', () => onChange(select.value));
  }

  return select;
}

/**
 * Remove all child nodes from an element.
 *
 * @param {HTMLElement} element - The element to clear
 */
export function clearElement(element) {
  if (!element) return;
  element.innerHTML = '';
}

/**
 * Create an empty state placeholder message.
 *
 * @param {string} message - The placeholder message
 * @param {string} [className] - CSS class name(s)
 * @returns {HTMLDivElement} The placeholder element
 */
export function createEmptyState(message, className = 'empty-state') {
  const div = createElement('div', className);
  const p = createElement('p', '', message);
  div.append(p);
  return div;
}
