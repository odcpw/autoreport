/**
 * Validation labels and messages for data validation across the application.
 * Centralizes validation-related text for consistency and easy updates.
 */

export const VALIDATION_LABELS = {
  master: 'Master data',
  selfEval: 'Self-evaluation',
  photos: 'Photo library',
  project: 'Project state',
};

export const VALIDATION_MESSAGES = {
  MASTER_PENDING: 'Awaiting master.json upload.',
  SELF_EVAL_PENDING: 'Awaiting self_eval.json upload.',
  PHOTOS_PENDING: 'Select a photo directory when you are ready.',
  PHOTOS_NONE: 'No photos have been imported yet.',
  PROJECT_PENDING: 'Waiting for master and self-evaluation uploads.',
};
