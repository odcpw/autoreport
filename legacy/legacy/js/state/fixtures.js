const EMBEDDED_FIXTURES = {
  project: {
    version: 1,
    meta: {
      projectId: '2025-ACME-001',
      company: 'ACME AG',
      createdAt: '2025-02-11T08:45:00Z',
      locale: 'en',
      author: 'Fixture',
    },
    branding: {},
    lists: {
      'photo.bericht': [
        { value: '1.1.1', label: '1.1.1', labels: { en: '1.1.1', de: '1.1.1' }, group: 'bericht', sortOrder: 1, chapterId: '1.1' },
        { value: '4.6.1', label: '4.6.1', labels: { en: '4.6.1', de: '4.6.1' }, group: 'bericht', sortOrder: 2, chapterId: '4.6' },
      ],
      'photo.seminar': [
        { value: 'PSA Basics', label: 'PSA Basics', labels: { en: 'PSA Basics', de: 'PSA Basics' }, group: 'seminar', sortOrder: 1, chapterId: '' },
        { value: 'Leadership Coaching', label: 'Leadership Coaching', labels: { en: 'Leadership Coaching', de: 'Leadership Coaching' }, group: 'seminar', sortOrder: 2, chapterId: '' },
      ],
      'photo.topic': [
        { value: 'Policy', label: 'Policy', labels: { en: 'Policy', de: 'Policy' }, group: 'topic', sortOrder: 1, chapterId: '' },
        { value: 'Leadership', label: 'Leadership', labels: { en: 'Leadership', de: 'Leadership' }, group: 'topic', sortOrder: 2, chapterId: '' },
        { value: 'Training', label: 'Training', labels: { en: 'Training', de: 'Training' }, group: 'topic', sortOrder: 3, chapterId: '' },
      ],
    },
    photos: {
      'photos/site/psa-check.jpg': {
        notes: 'Missing signage near north entrance',
        tags: {
          bericht: ['4.6.1'],
          seminar: ['PSA Basics'],
          topic: ['Training'],
        },
      },
    },
    chapters: [
      {
        id: '1.1',
        title: { en: 'Policies', de: 'Policies' },
        rows: [
          {
            id: '1.1.1',
            chapterId: '1.1',
            titleOverride: 'Safety Policy',
            master: {
              finding: 'The organisation maintains an up-to-date safety policy.',
              levels: {
                '1': 'Adopt a site-wide communication plan for safety policy updates.',
                '2': 'Introduce quarterly training refreshers focused on leadership accountability.',
                '3': 'Document all policy exceptions and review them during the annual audit.',
                '4': 'Escalate non-compliance cases to senior management within 5 working days.',
              },
            },
            overrides: {
              finding: { text: '', enabled: false },
              levels: {
                '1': { text: '', enabled: false },
                '2': { text: '', enabled: false },
                '3': { text: '', enabled: false },
                '4': { text: '', enabled: false },
              },
            },
            customer: { answer: 'yes', remark: 'Policies reviewed in March; pending communication update.', priority: null },
            workstate: {
              selectedLevel: 2,
              includeFinding: true,
              includeRecommendation: true,
              overwriteMode: 'append',
              done: false,
              notes: '',
              lastEditedBy: '',
              lastEditedAt: '',
              findingOverride: '',
              useFindingOverride: false,
              levelOverrides: { '1': '', '2': '', '3': '', '4': '' },
              useLevelOverride: { '1': false, '2': false, '3': false, '4': false },
            },
          },
        ],
      },
      {
        id: '4.6',
        title: { en: 'On-Site Training', de: 'On-Site Training' },
        rows: [
          {
            id: '4.6.1',
            chapterId: '4.6',
            titleOverride: 'PSA Usage',
            master: {
              finding: 'Assess compliance with on-site personal safety equipment usage.',
              levels: {
                '1': 'Roll out refresher modules for PSA usage.',
                '2': 'Update signage across high-risk areas.',
                '3': 'Track PSA training completion monthly.',
                '4': 'Ensure supervisors perform spot checks twice per shift.',
              },
            },
            overrides: {
              finding: { text: '', enabled: false },
              levels: {
                '1': { text: '', enabled: false },
                '2': { text: '', enabled: false },
                '3': { text: '', enabled: false },
                '4': { text: '', enabled: false },
              },
            },
            customer: { answer: 'partial', remark: 'PSA availability is high, but compliance audits show gaps.', priority: null },
            workstate: {
              selectedLevel: 3,
              includeFinding: true,
              includeRecommendation: true,
              overwriteMode: 'append',
              done: false,
              notes: '',
              lastEditedBy: '',
              lastEditedAt: '',
              findingOverride: '',
              useFindingOverride: false,
              levelOverrides: { '1': '', '2': '', '3': '', '4': '' },
              useLevelOverride: { '1': false, '2': false, '3': false, '4': false },
              needsReview: true
            },
          },
        ],
      },
    ],
    history: [],
  },
};

const FIXTURE_PATHS = {
  master: 'fixtures/master.sample.json',
  selfEval: 'fixtures/self_eval.sample.json',
  project: 'fixtures/project.sample.json',
};

export async function loadFixtures(state) {
  if (!state || state.project) return;

  const shouldSkipFetch = window.location.protocol === 'file:';
  let project = EMBEDDED_FIXTURES.project;

  if (!shouldSkipFetch) {
    try {
      const projectResponse = await fetch(FIXTURE_PATHS.project);
      if (projectResponse.ok) {
        project = await projectResponse.json();
      }
    } catch (error) {
      console.warn('Fixture fetch failed, using embedded defaults:', error);
    }
  }

  try {
    state.setProjectSnapshot(project, FIXTURE_PATHS.project);
  } catch (error) {
    console.warn('Fixture project snapshot invalid:', error);
  }
}
