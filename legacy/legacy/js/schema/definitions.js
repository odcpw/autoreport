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
  required: ['version', 'meta', 'chapters', 'photos', 'lists'],
  properties: {
    version: { type: 'integer', minimum: 1 },
    meta: {
      type: 'object',
      required: ['company', 'locale'],
      properties: {
        projectId: { type: 'string' },
        company: { type: 'string' },
        createdAt: { type: 'string' },
        locale: { type: 'string' },
        author: { type: 'string' },
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
      additionalProperties: {
        type: 'array',
        items: {
          type: 'object',
          required: ['value'],
          properties: {
            value: { type: 'string' },
            label: { type: 'string' },
            labels: { type: 'object' },
            group: { type: 'string' },
            sortOrder: { type: ['number', 'integer', 'string'] },
            chapterId: { type: 'string' },
          },
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
                items: { type: 'string' },
              },
              seminar: { type: 'array', items: { type: 'string' } },
              topic: { type: 'array', items: { type: 'string' } },
            },
          },
        },
      },
    },
    chapters: {
      type: 'array',
      items: {
        type: 'object',
        required: ['id', 'rows'],
        properties: {
          id: { type: 'string' },
          parentId: { type: 'string' },
          orderIndex: { type: ['number', 'integer'] },
          title: {
            anyOf: [
              { type: 'string' },
              {
                type: 'object',
                properties: {
                  de: { type: 'string' },
                  fr: { type: 'string' },
                  it: { type: 'string' },
                  en: { type: 'string' },
                },
                additionalProperties: true,
              },
            ],
          },
          pageSize: { type: ['integer', 'null'] },
          isActive: { type: ['boolean', 'null'] },
          rows: {
            type: 'array',
            items: {
              type: 'object',
              required: ['id', 'chapterId', 'master', 'overrides', 'customer', 'workstate'],
              properties: {
                id: { type: 'string' },
                chapterId: { type: 'string' },
                titleOverride: { type: 'string' },
                master: {
                  type: 'object',
                  required: ['finding', 'levels'],
                  properties: {
                    finding: { type: 'string' },
                    levels: { type: 'object' },
                  },
                },
                overrides: { type: 'object' },
                customer: { type: 'object' },
                workstate: { type: 'object' },
              },
            },
          },
        },
      },
    },
    history: { type: 'array' },
  },
};
