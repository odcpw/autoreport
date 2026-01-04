import { VALIDATION_LABELS } from '../core/constants/validation.js';

/**
 * Validation UI rendering utilities.
 * Creates and updates validation message displays and summaries.
 */

/**
 * Render validation messages for all scopes.
 *
 * @param {HTMLElement} container - The container element for validation messages
 * @param {Object} validation - Validation status object with scopes
 * @param {Function} getSourceLabel - Function to get source label for a scope
 */
export function renderValidationMessages(container, validation, getSourceLabel) {
  if (!container) return;

  const scopes = Object.keys(VALIDATION_LABELS);
  const fragment = document.createDocumentFragment();
  let errorCount = 0;

  scopes.forEach((scope) => {
    const status = validation?.[scope] ?? { ok: null, messages: [] };
    const wrapper = document.createElement('div');
    wrapper.className = 'message';

    const statusClass = status.ok === true
      ? 'message--ok'
      : status.ok === false
      ? 'message--error'
      : 'message--pending';
    wrapper.classList.add(statusClass);

    const title = document.createElement('strong');
    const source = getSourceLabel ? getSourceLabel(scope) : null;
    title.textContent = source
      ? `${VALIDATION_LABELS[scope] || scope} — ${source}`
      : VALIDATION_LABELS[scope] || scope;

    const header = document.createElement('div');
    header.className = 'message__header';
    header.append(title);

    if (status.messages && status.messages.length) {
      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.textContent = 'Details';
      toggle.className = 'message__toggle';
      toggle.addEventListener('click', () => {
        const open = wrapper.getAttribute('data-open') === 'true';
        wrapper.setAttribute('data-open', String(!open));
      });
      header.append(toggle);

      const list = document.createElement('ul');
      status.messages.forEach((msg) => {
        const item = document.createElement('li');
        item.textContent = msg;
        list.append(item);
      });

      const content = document.createElement('div');
      content.className = 'message__content';
      content.append(list);
      wrapper.append(header, content);
      wrapper.setAttribute('data-open', 'false');

      if (status.ok === false) {
        errorCount += status.messages.length;
      }
    } else {
      wrapper.append(header);
      const p = document.createElement('p');
      p.textContent = status.ok === true ? 'Validation OK.' : 'Awaiting input.';
      wrapper.append(p);
    }

    fragment.append(wrapper);
  });

  container.innerHTML = '';
  container.append(fragment);

  return errorCount;
}

/**
 * Render a validation summary showing total error count.
 *
 * @param {HTMLElement} summaryElement - The summary container element
 * @param {number} errorCount - Total number of validation errors
 */
export function renderValidationSummary(summaryElement, errorCount) {
  if (!summaryElement) return;
  if (!Number.isFinite(errorCount)) {
    summaryElement.textContent = '';
    return;
  }

  if (errorCount > 0) {
    summaryElement.textContent = `${errorCount} validation issue(s) detected.`;
    summaryElement.className = 'validation-summary validation-summary--error';
  } else {
    summaryElement.textContent = 'Validation passed.';
    summaryElement.className = 'validation-summary validation-summary--ok';
  }
}

/**
 * Update the status banner pill with validation state.
 *
 * @param {HTMLElement} statusElement - The status pill element
 * @param {boolean} isReady - Whether exports are ready
 * @param {Object} validation - Validation status object
 */
export function updateStatusBanner(statusElement, isReady, validation) {
  if (!statusElement) return;

  const hasError = validation
    ? Object.values(validation).some((item) => item?.ok === false)
    : false;

  if (hasError) {
    statusElement.textContent = 'Validation errors';
    statusElement.className = 'status-pill status-pill--error';
  } else if (isReady) {
    statusElement.textContent = 'Ready to export';
    statusElement.className = 'status-pill status-pill--ok';
  } else {
    statusElement.textContent = 'Validation pending';
    statusElement.className = 'status-pill';
  }
}
