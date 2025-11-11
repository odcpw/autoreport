export const masterSchema = {
  $id: 'master.json',
  type: 'object',
  required: ['version', 'chapters'],
  properties: {
    version: { type: 'integer', minimum: 1 },
    chapters: {
      type: 'array',
      items: { $ref: '#/$defs/chapter' },
    },
    pptTaxonomies: {
      type: 'object',
      properties: {
        reportChapters: {
          type: 'array',
          items: { type: 'string', pattern: '^\\d+(\\.\\d+)*$' },
        },
        seminarChapters: {
          type: 'array',
          items: { type: 'string', pattern: '^\\d+(\\.\\d+)*$' },
        },
      },
    },
  },
  $defs: {
    chapter: {
      type: 'object',
      required: ['id', 'title'],
      properties: {
        id: { type: 'string', pattern: '^\\d+(\\.\\d+)*$' },
        title: { type: 'string' },
        findingTemplate: { type: 'string' },
        recommendations: {
          type: 'object',
          patternProperties: {
            '^[1-4]$': { type: 'string' },
          },
          additionalProperties: false,
        },
        children: {
          type: 'array',
          items: { $ref: '#/$defs/chapter' },
        },
      },
    },
  },
};

export const selfEvalSchema = {
  $id: 'self_eval.json',
  type: 'object',
  required: ['company', 'responses'],
  properties: {
    company: { type: 'string', minLength: 1 },
    responses: {
      type: 'array',
      items: {
        type: 'object',
        required: ['id', 'yesNo'],
        properties: {
          id: { type: 'string', pattern: '^\\d+(\\.\\d+)*$' },
          yesNo: { enum: ['yes', 'no', 'partial', 'n/a'] },
          remarks: { type: 'string' },
        },
      },
    },
  },
};

export const projectSchema = {
  $id: 'project.json',
  type: 'object',
  required: ['version', 'meta', 'photos', 'report', 'presentation'],
  properties: {
    version: { type: 'integer', minimum: 1 },
    meta: {
      type: 'object',
      required: ['company', 'locale'],
      properties: {
        created: { type: 'string', format: 'date-time' },
        company: { type: 'string' },
        locale: { enum: ['en', 'de', 'fr', 'it'] },
      },
    },
    branding: {
      type: 'object',
      properties: {
        left: { type: ['string', 'null'] },
        right: { type: ['string', 'null'] },
      },
    },
    lists: {
      type: 'object',
      required: ['berichtList', 'seminarList', 'topicList'],
      properties: {
        berichtList: {
          type: 'array',
          items: {
            anyOf: [
              { type: 'string' },
              {
                type: 'object',
                properties: {
                  value: { type: 'string' },
                  label: { type: 'string' },
                },
                required: ['value'],
              },
            ],
          },
        },
        seminarList: {
          type: 'array',
          items: { type: 'string' },
        },
        topicList: {
          type: 'array',
          items: { type: 'string' },
        },
      },
    },
    photos: {
      type: 'object',
      additionalProperties: {
        type: 'object',
        properties: {
          notes: { type: 'string' },
          tags: {
            type: 'object',
            properties: {
              bericht: {
                type: 'array',
                items: { type: 'string', pattern: '^\\d+(\\.\\d+)*$' },
              },
              seminar: { type: 'array', items: { type: 'string' } },
              topic: { type: 'array', items: { type: 'string' } },
            },
          },
        },
      },
    },
    report: { type: 'object' },
    presentation: { type: 'object' },
  },
};
