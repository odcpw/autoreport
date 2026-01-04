/**
 * Event name constants for consistent event handling across the application.
 * Centralizing event names prevents typos and makes event usage discoverable.
 */

export const EVENTS = {
  // State events
  STATE_CHANGE: 'state:change',
  STATE_RESET: 'state:reset',
  STATE_VALIDATION: 'state:validation',

  // Project events
  PROJECT_UPDATED: 'project:updated',
  PROJECT_CLEARED: 'project:cleared',
  PROJECT_RESET: 'autobericht:project-reset',

  // Session events
  SESSION_SAVED: 'autobericht:session-saved',
  SESSION_CLEARED: 'autobericht:session-cleared',
  SESSION_RESTORED: 'autobericht:session-restored',
  SESSION_WARNING: 'autobericht:session-warning',

  // Session action events
  RESTORE_SESSION: 'autobericht:restore-session',
  RESTORE_SESSION_PAYLOAD: 'autobericht:restore-session-payload',
  RESET_PROJECT: 'autobericht:reset-project',

  // Autosave events
  AUTOSAVE_START: 'autobericht:autosave-start',
  AUTOSAVE_SUCCESS: 'autobericht:autosave-success',
  AUTOSAVE_FAILED: 'autobericht:autosave-failed',
};
