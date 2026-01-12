(() => {
  const defaultProject = {
    meta: {
      projectId: "2026-TEST-001",
      company: "ACME AG",
      companyId: "CH-000000",
      locale: "de-CH",
      moderator: "consultant@example.com",
      moderatorInitials: "MM",
      coModerator: "",
      coModeratorInitials: "",
      createdAt: new Date().toISOString(),
    },
    chapters: [
      {
        id: "1",
        title: { de: "Leitbild" },
        rows: [
          {
            id: "1.1.1",
            type: "standard",
            titleOverride: "Unternehmensleitbild",
            master: {
              finding: "Das Unternehmen verfuegt nicht ueber ein Leitbild.",
              levels: {
                "1": "Eine Sicherheitscharta als ersten Schritt etablieren.",
                "2": "Die vorhandenen Werte in die Fuhrung integrieren.",
                "3": "Die dokumentierte Charta verbreiten und leben.",
                "4": "Die Charta als aktives Fuehrungsinstrument nutzen.",
              },
            },
            customer: {
              answer: 1,
              remark: "Leitbild vorhanden",
              items: [
                {
                  id: "1.1.1",
                  question: "Gibt es ein Sicherheitsleitbild?",
                  collapsedId: "1.1.1",
                },
              ],
            },
            workstate: {
              selectedLevel: 2,
              includeFinding: true,
              includeRecommendation: true,
              done: false,
              useFindingOverride: false,
              findingOverride: "",
              useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
            },
          },
          {
            id: "1.1.2",
            type: "standard",
            titleOverride: "Strategie",
            master: {
              finding: "Es fehlt eine dokumentierte Sicherheitsstrategie.",
              levels: {
                "1": "Strategie-Grundsaetze definieren.",
                "2": "Strategie in Ziele uebersetzen.",
                "3": "Strategie regelmaessig pruefen.",
                "4": "Strategie in allen Bereichen verankern.",
              },
            },
            customer: {
              answer: 0,
              remark: "",
              items: [
                {
                  id: "1.1.2",
                  question: "Gibt es eine dokumentierte Sicherheitsstrategie?",
                  collapsedId: "1.1.2",
                },
              ],
            },
            workstate: {
              selectedLevel: 3,
              includeFinding: true,
              includeRecommendation: true,
              done: true,
              useFindingOverride: true,
              findingOverride: "Eine Strategie besteht, wird aber nicht aktiv kommuniziert.",
              useLevelOverride: { "1": false, "2": true, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "Strategie sichtbar machen.", "3": "", "4": "" },
            },
          },
        ],
      },
      {
        id: "4.8",
        title: { de: "Beobachtungen" },
        rows: [
          {
            id: "4.8.1",
            type: "field_observation",
            titleOverride: "Regale",
            master: {
              finding: "Regale sind nicht gegen Kippen gesichert.",
              levels: {
                "1": "Regale sichern und Sichtkontrolle definieren.",
                "2": "Regalinspektionen regelmaessig durchfuehren.",
                "3": "Sicherheitschecks dokumentieren.",
                "4": "Regalmanagement im Sicherheitsprogramm verankern.",
              },
            },
            customer: {
              answer: 0,
              remark: "",
              items: [
                {
                  id: "4.8.1",
                  question: "Sind Regale gegen Kippen gesichert?",
                  collapsedId: "4.8.1",
                },
              ],
            },
            workstate: {
              selectedLevel: 1,
              includeFinding: true,
              includeRecommendation: true,
              done: false,
              useFindingOverride: false,
              findingOverride: "",
              useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
            },
          },
        ],
      },
    ],
  };

  const createState = (projectTemplate = defaultProject) => ({
    project: structuredClone(projectTemplate),
    selectedChapterId: projectTemplate.chapters[0]?.id || "",
    filters: {
      mode: "all",
    },
    spiderOverrides: {},
    photoIndex: {
      report: new Map(),
      observations: new Map(),
    },
    photoRoot: "",
    photoOverlay: {
      tag: "",
      items: [],
      index: 0,
      url: "",
    },
    checklistOverlay: {
      items: [],
      locale: "de",
      requested: "",
      fallback: false,
      industryFilter: "",
      categoryFilter: "",
    },
  });

  const compareIdSegments = (a, b) => {
    const aParts = String(a || "").split(".");
    const bParts = String(b || "").split(".");
    const maxLen = Math.max(aParts.length, bParts.length);
    for (let i = 0; i < maxLen; i += 1) {
      const aPart = aParts[i];
      const bPart = bParts[i];
      if (aPart == null) return -1;
      if (bPart == null) return 1;
      const aNum = Number(aPart);
      const bNum = Number(bPart);
      const aIsNum = !Number.isNaN(aNum);
      const bIsNum = !Number.isNaN(bNum);
      if (aIsNum && bIsNum) {
        if (aNum !== bNum) return aNum - bNum;
      } else {
        const cmp = String(aPart).localeCompare(String(bPart), "de", { numeric: true });
        if (cmp !== 0) return cmp;
      }
    }
    return 0;
  };

  const toText = (value) => {
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  };

  const calculateScore = (row) => {
    if (row?.type === "field_observation" || row?.type === "summary") return null;
    const ws = row?.workstate || {};
    if (ws.includeFinding === false) return 100;
    const level = Number(ws.selectedLevel || 1);
    return Math.max(0, Math.min(100, (level - 1) * 25));
  };

  const sanitizeFilename = (value) => String(value || "")
    .trim()
    .replace(/[^a-zA-Z0-9._-]/g, "_");

  const getLibraryFileName = (meta = {}) => {
    const locale = sanitizeFilename(meta.locale || "de-CH");
    const initials = sanitizeFilename(
      meta.moderatorInitials
      || meta.initials
      || meta.moderator
      || ""
    );
    if (initials) return `library_user_${initials}_${locale}.json`;
    return `library_user_${locale}.json`;
  };

  const hashText = (value) => {
    const text = String(value || "");
    let hash = 2166136261;
    for (let i = 0; i < text.length; i += 1) {
      hash ^= text.charCodeAt(i);
      hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
    }
    return String(hash >>> 0);
  };

  const getChapterTitle = (chapter) => {
    if (!chapter) return "";
    if (typeof chapter.title === "string") return chapter.title;
    if (chapter.title && chapter.title.de) return chapter.title.de;
    if (chapter.id === "4.8") return "Beobachtungen";
    return chapter.id || "";
  };

  const formatChapterLabel = (chapter) => {
    if (!chapter) return "";
    const title = getChapterTitle(chapter);
    const id = chapter.id || "";
    if (!title) return id;
    if (!id) return title;
    const normalized = title.trim();
    if (normalized === id || normalized.startsWith(`${id} `) || normalized.startsWith(`${id}.`)) {
      return normalized;
    }
    return `${id} ${title}`.trim();
  };

  const getFindingText = (row) => {
    const ws = row.workstate;
    if (ws.useFindingOverride && ws.findingOverride) return ws.findingOverride;
    return toText(row.master?.finding);
  };

  const getRecommendationText = (row, level) => {
    const ws = row.workstate;
    const levelKey = String(level);
    if (ws.useLevelOverride?.[levelKey] && ws.levelOverrides?.[levelKey]) {
      return ws.levelOverrides[levelKey];
    }
    return toText(row.master?.levels?.[levelKey]);
  };

  const getAnswerState = (row) => {
    const direct = row.customer?.answer;
    if (direct === 0 || direct === 1) return direct;
    const items = row.customer?.items || [];
    const answers = new Set();
    items.forEach((item) => {
      if (item.answer === 0 || item.answer === 1) answers.add(item.answer);
    });
    if (answers.size === 1) return Array.from(answers)[0];
    if (answers.size > 1) return "mixed";
    return null;
  };

  const getAnswerComments = (row) => {
    const comments = [];
    const items = row.customer?.items || [];
    items.forEach((item) => {
      if (item.comment) {
        comments.push(`${item.id}: ${item.comment}`);
      }
    });
    if (row.customer?.comment && !comments.length) {
      comments.push(row.customer.comment);
    }
    return comments;
  };

  const getAnswerEvidence = (row) => {
    const evidence = [];
    const items = row.customer?.items || [];
    items.forEach((item) => {
      if (item.evidence) {
        evidence.push(`${item.id}: ${item.evidence}`);
      }
    });
    if (row.customer?.evidence && !evidence.length) {
      evidence.push(row.customer.evidence);
    }
    return evidence;
  };

  window.AutoBerichtState = {
    defaultProject,
    createState,
    compareIdSegments,
    toText,
    calculateScore,
    sanitizeFilename,
    getLibraryFileName,
    hashText,
    getChapterTitle,
    formatChapterLabel,
    getFindingText,
    getRecommendationText,
    getAnswerState,
    getAnswerComments,
    getAnswerEvidence,
  };
})();
