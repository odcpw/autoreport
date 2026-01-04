/**
 * Activity log management for tracking user actions and system events.
 * Maintains a chronological log of activities with timestamp formatting.
 */

/**
 * Activity log entry
 * @typedef {{message: string, timestamp: string}} ActivityEntry
 */

/**
 * Create an activity log manager.
 *
 * @param {HTMLElement} logElement - The container element for the activity log
 * @param {number} maxEntries - Maximum number of entries to keep (default: 12)
 * @returns {{add: Function, render: Function, clear: Function, getEntries: Function}}
 */
export function createActivityLog(logElement, maxEntries = 12) {
  const entries = [];

  /**
   * Add an activity entry to the log.
   *
   * @param {string} message - The activity message
   * @param {string} [timestamp] - ISO timestamp (defaults to now)
   */
  function add(message, timestamp = null) {
    const entry = {
      message,
      timestamp: timestamp || new Date().toISOString(),
    };
    entries.unshift(entry);
    if (entries.length > maxEntries) {
      entries.length = maxEntries;
    }
    render();
  }

  /**
   * Render the activity log to the DOM.
   */
  function render() {
    if (!logElement) return;

    logElement.innerHTML = '';

    if (entries.length === 0) {
      const empty = document.createElement('li');
      empty.className = 'activity-log__empty';
      empty.textContent = 'No activity yet.';
      logElement.append(empty);
      return;
    }

    entries.forEach((entry) => {
      const li = document.createElement('li');
      li.className = 'activity-log__item';

      const label = document.createElement('span');
      label.textContent = entry.message;

      const time = document.createElement('time');
      time.dateTime = entry.timestamp;
      time.textContent = formatTimestamp(entry.timestamp);

      li.append(label, time);
      logElement.append(li);
    });
  }

  /**
   * Clear all activity log entries.
   */
  function clear() {
    entries.length = 0;
    render();
  }

  /**
   * Get all activity log entries.
   *
   * @returns {Array<ActivityEntry>} Array of activity entries
   */
  function getEntries() {
    return entries.slice();
  }

  return {
    add,
    render,
    clear,
    getEntries,
  };
}

/**
 * Format a timestamp for display.
 *
 * @param {string} timestamp - ISO timestamp string
 * @returns {string} Formatted timestamp
 */
export function formatTimestamp(timestamp) {
  const date = new Date(timestamp);
  if (Number.isNaN(date.getTime())) return 'recently';

  if (typeof Intl !== 'undefined' && Intl.DateTimeFormat) {
    const formatter = new Intl.DateTimeFormat(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
    return formatter.format(date);
  }

  return date.toLocaleString();
}
