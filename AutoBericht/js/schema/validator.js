import { masterSchema, selfEvalSchema, projectSchema } from './definitions.js';

const SCHEMA_MAP = {
  master: masterSchema,
  selfEval: selfEvalSchema,
  project: projectSchema,
};

class SchemaValidator {
  constructor() {
    this.validatorInstances = {};
    this.errors = [];
    this.ajv = null;
    this.warned = false;
  }

  ensureAjv() {
    if (this.ajv) return this.ajv;
    if (typeof window !== 'undefined' && typeof window.Ajv === 'function') {
      try {
        this.ajv = new window.Ajv({
          allErrors: true,
          allowUnionTypes: true,
          verbose: false,
        });
        this.ajv.addFormat('date-time', { type: 'string', validate: () => true });
      } catch (error) {
        console.error('Failed to initialise AJV', error);
        this.ajv = null;
      }
    }
    if (!this.ajv && !this.warned) {
      console.warn('AJV not found. Schema validation is disabled.');
      this.warned = true;
    }
    return this.ajv;
  }

  getValidator(scope) {
    if (!SCHEMA_MAP[scope]) return null;
    if (this.validatorInstances[scope]) return this.validatorInstances[scope];
    const ajv = this.ensureAjv();
    if (!ajv) return null;
    try {
      this.validatorInstances[scope] = ajv.compile(SCHEMA_MAP[scope]);
      return this.validatorInstances[scope];
    } catch (error) {
      console.error(`Failed to compile schema for ${scope}`, error);
      return null;
    }
  }

  validate(scope, data) {
    const validator = this.getValidator(scope);
    if (!validator) {
      this.errors = [];
      return true;
    }
    const valid = validator(data);
    if (!valid) {
      this.errors = (validator.errors || []).map((err) => this.formatError(err));
    } else {
      this.errors = [];
    }
    return valid;
  }

  getErrors() {
    return this.errors.slice();
  }

  formatError(error) {
    const path = error.instancePath || error.dataPath || '';
    const location = path ? `${path} ` : '';
    return `${location}${error.message || 'schema error'}`.trim();
  }
}

export const schemaValidator = new SchemaValidator();
