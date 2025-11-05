const EMBEDDED_FIXTURES = {
  master: {
    version: 1,
    chapters: [
      {
        id: '1',
        title: 'Leadership & Culture',
        children: [
          {
            id: '1.1',
            title: 'Policies',
            children: [
              {
                id: '1.1.1',
                title: 'Safety Policy',
                findingTemplate:
                  'The organisation maintains an up-to-date safety policy. Describe deviations here.',
                recommendations: {
                  1: 'Adopt a site-wide communication plan for safety policy updates.',
                  2: 'Introduce quarterly training refreshers focused on leadership accountability.',
                  3: 'Document all policy exceptions and review them during the annual audit.',
                  4: 'Escalate non-compliance cases to senior management within 5 working days.',
                },
              },
            ],
          },
        ],
      },
      {
        id: '4',
        title: 'Training & Competence',
        children: [
          {
            id: '4.6',
            title: 'On-Site Training',
            children: [
              {
                id: '4.6.1',
                title: 'PSA Usage',
                findingTemplate:
                  'Assess compliance with on-site personal safety equipment usage.',
                recommendations: {
                  1: 'Roll out refresher modules for PSA usage.',
                  2: 'Update signage across high-risk areas.',
                  3: 'Track PSA training completion monthly.',
                  4: 'Ensure supervisors perform spot checks twice per shift.',
                },
              },
            ],
          },
        ],
      },
    ],
    pptTaxonomies: {
      reportChapters: ['1.1.1'],
      seminarChapters: ['4.6.1'],
    },
  },
  selfEval: {
    company: 'ACME AG',
    responses: [
      {
        id: '1.1.1',
        yesNo: 'yes',
        remarks: 'Policies reviewed in March; pending communication update.',
      },
      {
        id: '4.6.1',
        yesNo: 'partial',
        remarks: 'PSA availability is high, but compliance audits show gaps.',
      },
    ],
  },
  project: {
    lists: {
      berichtList: [{ value: '4.6.1', label: '4.6.1' }],
      seminarList: ['PSA Basics', 'Leadership Coaching'],
      topicList: ['Policy', 'Leadership', 'Training'],
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
  },
};

const FIXTURE_PATHS = {
  master: 'fixtures/master.sample.json',
  selfEval: 'fixtures/self_eval.sample.json',
  project: 'fixtures/project.sample.json',
};

export async function loadFixtures(state) {
  if (!state || state.master || state.selfEval) return;

  const shouldSkipFetch = window.location.protocol === 'file:';
  let master = EMBEDDED_FIXTURES.master;
  let selfEval = EMBEDDED_FIXTURES.selfEval;
  let project = EMBEDDED_FIXTURES.project;

  if (!shouldSkipFetch) {
    try {
      const [masterResponse, selfEvalResponse, projectResponse] = await Promise.all([
        fetch(FIXTURE_PATHS.master),
        fetch(FIXTURE_PATHS.selfEval),
        fetch(FIXTURE_PATHS.project),
      ]);

      if (masterResponse.ok) {
        master = await masterResponse.json();
      }
      if (selfEvalResponse.ok) {
        selfEval = await selfEvalResponse.json();
      }
      if (projectResponse.ok) {
        project = await projectResponse.json();
      }
    } catch (error) {
      console.warn('Fixture fetch failed, using embedded defaults:', error);
    }
  }

  state.setMasterData(master, FIXTURE_PATHS.master);
  state.setSelfEvalData(selfEval, FIXTURE_PATHS.selfEval);
  try {
    state.setProjectSnapshot(project, FIXTURE_PATHS.project);
  } catch (error) {
    console.warn('Fixture project snapshot invalid:', error);
  }
}
