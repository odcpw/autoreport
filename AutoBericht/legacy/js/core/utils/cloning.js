/**
 * Cloning utilities for deep copying objects.
 * Prefers structuredClone when available for better performance and accuracy.
 * Falls back to JSON serialization for older browsers.
 */

/**
 * Deep clone an object, preserving most types including Date objects.
 * Uses structuredClone when available, otherwise falls back to JSON serialization.
 *
 * @param {any} obj - The object to clone
 * @returns {any} A deep copy of the object
 */
export function deepClone(obj) {
  if (obj === null || obj === undefined) {
    return obj;
  }

  if (typeof structuredClone === 'function') {
    try {
      return structuredClone(obj);
    } catch (error) {
      console.warn('structuredClone failed, falling back to JSON serialization', error);
    }
  }

  try {
    return JSON.parse(JSON.stringify(obj));
  } catch (error) {
    console.error('Failed to clone object', error);
    return obj;
  }
}

/**
 * Clone a simple entry object (optimized for report entries).
 *
 * @param {Object} entry - The entry object to clone
 * @returns {Object} A deep copy of the entry
 */
export function cloneEntry(entry) {
  return deepClone(entry);
}

/**
 * Clone an override object (optimized for project overrides).
 *
 * @param {Object} override - The override object to clone
 * @returns {Object} A deep copy of the override
 */
export function cloneOverride(override) {
  return deepClone(override);
}
