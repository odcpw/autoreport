(() => {
  const stateHelpers = window.AutoBerichtState || {};
  const LEVEL_TO_PCT = stateHelpers.levelToPct
    || ((level) => {
      const raw = Number(level || 1);
      const clamped = Math.max(1, Math.min(4, Number.isFinite(raw) ? raw : 1));
      const idx = Math.round(clamped) - 1;
      return [0, 33, 66, 100][Math.max(0, Math.min(3, idx))];
    });

  const readJsonFromHandle = async (dirHandle, filename) => {
    if (!dirHandle) return null;
    try {
      const fileHandle = await dirHandle.getFileHandle(filename);
      const file = await fileHandle.getFile();
      return JSON.parse(await file.text());
    } catch (err) {
      return null;
    }
  };

  const fetchBundledWeights = async () => {
    try {
      const res = await fetch("../data/weights.json");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.json();
    } catch (err) {
      return null;
    }
  };

  const loadWeights = async (dirHandle) => {
    // Project override
    const projectWeights = await readJsonFromHandle(dirHandle, "weights.json");
    if (projectWeights) return { weights: projectWeights, source: "project/weights.json" };
    // Bundled
    const bundled = await fetchBundledWeights();
    if (bundled) return { weights: bundled, source: "bundle:data/weights.json" };
    throw new Error("weights.json not found (project or bundle).");
  };

  const deriveChapters = (weights) => {
    if (Array.isArray(weights?.chapters) && weights.chapters.length) return weights.chapters;
    // fallback: derive from items
    const sums = {};
    (weights?.items || []).forEach((item) => {
      const ch = String(item.id).split(".")[0];
      sums[ch] = (sums[ch] || 0) + Number(item.weight || 0);
    });
    return Object.entries(sums)
      .map(([id, w]) => ({
        id,
        includeIn11: Number(id) <= 11,
        includeIn14: Number(id) <= 14,
        weightSum: w,
      }))
      .sort((a, b) => Number(a.id) - Number(b.id));
  };

  const normalizeOverrides = (overrides = {}) => {
    const out = {};
    Object.entries(overrides).forEach(([id, val]) => {
      const useCompany = !!val?.useCompany;
      const useConsultant = !!val?.useConsultant;
      const company = Number.isFinite(Number(val?.company)) ? Number(val.company) : null;
      const consultant = Number.isFinite(Number(val?.consultant)) ? Number(val.consultant) : null;
      out[id] = { useCompany, useConsultant, company, consultant };
    });
    return out;
  };

  const normalizeWeightId = (raw) => String(raw || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/\.+$/g, "");

  // Legacy self-assessment imports keep workbook leaf ids on `originalId`.
  // Newer sidecars may only have the collapsed project row ids. Keep lookup permissive.
  const buildWeightLookup = (weights) => {
    const map = new Map();
    (weights?.items || []).forEach((item) => {
      const id = String(item?.id || "");
      const weight = Number(item?.weight || 0);
      const normalized = normalizeWeightId(id);
      if (normalized) map.set(normalized, weight);
      const collapsed = normalized.replace(/\./g, "");
      if (collapsed) map.set(collapsed, weight);
    });
    return map;
  };

  // These row-level overrides mirror the workbook Analyse formulas where some
  // project rows collapse multiple workbook leaves into one report row.
  const ROW_RULES = {
    "2.1.1": { mode: "max" },
    "2.1.4": { mode: "max" },
    "5.1.1": { mode: "max" },
    "2.2.4": {
      mode: "sum",
      maxGroups: [
        ["2.2.1.4.c", "2.2.1.4.d"],
      ],
    },
  };

  const REVERSE_SCORE_IDS = (() => {
    const ids = [
      "5.4.6",
      "9.1.4",
      "9.1.5",
      "9.1.6",
      "9.2.1",
      "9.2.2",
      "9.2.3",
      "9.2.4",
      "9.2.5.a",
      "9.2.5.b",
      "9.2.5.c",
      "9.2.5.d",
      "9.5.4",
      "9.5.5",
      "9.5.6",
      "9.5.7",
      "9.7.1.c",
      "9.7.1.e",
      "9.8.1",
      "9.8.2",
      "9.9.1",
      "9.9.3",
      "9.9.5",
    ];
    for (let code = "a".charCodeAt(0); code <= "z".charCodeAt(0); code += 1) {
      ids.push(`9.3.1.${String.fromCharCode(code)}`);
    }
    for (let code = "a".charCodeAt(0); code <= "o".charCodeAt(0); code += 1) {
      ids.push(`9.4.1.${String.fromCharCode(code)}`);
    }
    return new Set(ids.map((id) => normalizeWeightId(id)));
  })();

  // The Selbstbeurteilung workbook's Analyse sheet contains a small set of
  // reverse-scored questions, encoded as `(1-answer) * weight`. Keep that
  // mapping explicit here until the project data model carries row polarity.
  const applyScoreDirection = (pct, rawId) => {
    if (pct == null) return pct;
    return REVERSE_SCORE_IDS.has(normalizeWeightId(rawId)) ? 100 - pct : pct;
  };

  const pctFromCustomer = (row) => {
    // Prefer row.customer.answer, else average of customer.items[*].answer
    const ans = row?.customer?.answer;
    if (ans === 1 || ans === "1" || ans === true) return 100;
    if (ans === 0 || ans === "0" || ans === false) return 0;
    const items = row?.customer?.items;
    if (Array.isArray(items) && items.length) {
      let total = 0; let count = 0;
      items.forEach((it) => {
        if (it && (it.answer === 0 || it.answer === 1 || it.answer === "0" || it.answer === "1")) {
          total += Number(it.answer);
          count += 1;
        }
      });
      if (count > 0) return (total / count) * 100;
    }
    return 0;
  };

  const pctFromAnswerValue = (answer) => {
    if (answer === 1 || answer === "1" || answer === true) return 100;
    if (answer === 0 || answer === "0" || answer === false) return 0;
    return null;
  };

  const getRowRule = (rowId) => ROW_RULES[String(rowId || "")] || { mode: "sum" };

  const getWeightForId = (weightLookup, rawId) => {
    const normalized = normalizeWeightId(rawId);
    if (!normalized) return null;
    if (weightLookup.has(normalized)) return weightLookup.get(normalized);
    const collapsed = normalized.replace(/\./g, "");
    if (weightLookup.has(collapsed)) return weightLookup.get(collapsed);
    return null;
  };

  const collapseMaxGroups = (contributions, rule) => {
    const groups = Array.isArray(rule?.maxGroups) ? rule.maxGroups : [];
    if (!groups.length) return contributions;
    const used = new Set();
    const out = [];
    groups.forEach((ids) => {
      const normalizedIds = new Set((ids || []).map((id) => normalizeWeightId(id)));
      const matches = contributions.filter((entry) => normalizedIds.has(entry.key));
      matches.forEach((entry) => used.add(entry));
      if (!matches.length) return;
      out.push({
        key: [...normalizedIds].join("|"),
        weight: Math.max(...matches.map((entry) => entry.weight)),
        companyScore: Math.max(...matches.map((entry) => entry.companyScore)),
        consultantScore: Math.max(...matches.map((entry) => entry.consultantScore)),
      });
    });
    contributions.forEach((entry) => {
      if (!used.has(entry)) out.push(entry);
    });
    return out;
  };

  const getRowContribution = (row, weightLookup) => {
    const consultantPct = (() => {
      const ws = row?.workstate || {};
      return LEVEL_TO_PCT(ws.selectedLevel || 1);
    })();
    const items = Array.isArray(row?.customer?.items) ? row.customer.items : [];
    const contributions = items.flatMap((item) => {
      const rawId = item?.originalId || item?.id;
      const weight = getWeightForId(weightLookup, rawId);
      const companyPct = applyScoreDirection(pctFromAnswerValue(item?.answer), rawId);
      if (!Number.isFinite(weight) || companyPct == null) return [];
      const adjustedConsultantPct = applyScoreDirection(consultantPct, rawId);
      return [{
        key: normalizeWeightId(rawId),
        weight,
        companyScore: weight * companyPct,
        consultantScore: weight * adjustedConsultantPct,
      }];
    });
    const rule = getRowRule(row?.id);
    const collapsed = collapseMaxGroups(contributions, rule);
    if (collapsed.length) {
      if (rule.mode === "max") {
        const weight = Math.max(...collapsed.map((entry) => entry.weight));
        return {
          weight,
          companyScore: Math.max(...collapsed.map((entry) => entry.companyScore)),
          consultantScore: Math.max(...collapsed.map((entry) => entry.consultantScore)),
        };
      }
      return collapsed.reduce((acc, entry) => ({
        weight: acc.weight + entry.weight,
        companyScore: acc.companyScore + entry.companyScore,
        consultantScore: acc.consultantScore + entry.consultantScore,
      }), { weight: 0, companyScore: 0, consultantScore: 0 });
    }

    const directWeight = getWeightForId(weightLookup, row?.id);
    if (!Number.isFinite(directWeight) || directWeight <= 0) return null;
    const adjustedCompanyPct = applyScoreDirection(pctFromCustomer(row), row?.id);
    const adjustedConsultantPct = applyScoreDirection(consultantPct, row?.id);
    return {
      weight: directWeight,
      companyScore: directWeight * adjustedCompanyPct,
      consultantScore: directWeight * adjustedConsultantPct,
    };
  };

  const chapterDisplayLabel = (chapter) => {
    const titleObj = chapter?.title;
    if (titleObj && typeof titleObj === "object") {
      const values = Object.values(titleObj).filter((value) => typeof value === "string" && value.trim());
      if (values.length) return values[0].trim();
    }
    return String(chapter?.id || "");
  };

  const buildSections = (chapter) => {
    const sections = [];
    let current = null;
    (chapter?.rows || []).forEach((row) => {
      if (row?.kind === "section") {
        current = { id: String(row.id || ""), rows: [] };
        sections.push(current);
        return;
      }
      if (!current) {
        current = { id: null, rows: [] };
        sections.push(current);
      }
      current.rows.push(row);
    });
    return sections;
  };

  const sumRowContributions = (rows, weightLookup) => rows.reduce((acc, row) => {
    if (row?.type === "field_observation" || String(row?.id || "").startsWith("4.8")) return acc;
    const contribution = getRowContribution(row, weightLookup);
    if (!contribution) return acc;
    acc.weight += contribution.weight;
    acc.companyScore += contribution.companyScore;
    acc.consultantScore += contribution.consultantScore;
    return acc;
  }, { weight: 0, companyScore: 0, consultantScore: 0 });

  const computeChapterScores = (project, weights) => {
    const totals = new Map(); // chapter -> {wSum, compSum, consSum, title}
    const weightLookup = buildWeightLookup(weights);
    const sectionWeights = new Map();
    (weights.items || []).forEach((item) => {
      const id = String(item.id || "");
      if (id.split(".").length === 2) {
        sectionWeights.set(id, Number(item.weight || 0));
      }
    });

    (project?.chapters || []).forEach((chapter) => {
      const chapterId = String(chapter?.id || "");
      if (!/^\d+$/.test(chapterId)) return;
      const sections = buildSections(chapter);
      const hasWeightedSections = sections.some((section) => section.id && sectionWeights.has(section.id));
      const acc = { w: 0, comp: 0, cons: 0, title: chapterDisplayLabel(chapter) };

      if (hasWeightedSections) {
        sections.forEach((section) => {
          if (!section.id || !sectionWeights.has(section.id)) return;
          const sectionWeight = sectionWeights.get(section.id);
          const rowTotals = sumRowContributions(section.rows, weightLookup);
          const companyPct = rowTotals.weight ? rowTotals.companyScore / rowTotals.weight : 0;
          const consultantPct = rowTotals.weight ? rowTotals.consultantScore / rowTotals.weight : 0;
          acc.w += sectionWeight;
          acc.comp += sectionWeight * companyPct;
          acc.cons += sectionWeight * consultantPct;
        });
        // The workbook Analyse sheet scores chapter 4 with a fixed 4.8 sample
        // contribution (`4.8.1 = 100%`, weighted by section 4.8). Mirror that
        // here until the self-assessment model exposes an explicit 4.8 score.
        if (chapterId === "4" && sectionWeights.has("4.8")) {
          const sectionWeight = sectionWeights.get("4.8");
          acc.w += sectionWeight;
          acc.comp += sectionWeight * 100;
          acc.cons += 0;
        }
      } else {
        const rowTotals = sumRowContributions(chapter.rows || [], weightLookup);
        acc.w = rowTotals.weight;
        acc.comp = rowTotals.companyScore;
        acc.cons = rowTotals.consultantScore;
      }
      totals.set(chapterId, acc);
    });

    const chapters = deriveChapters(weights);
    const round5 = (value) => {
      const v = Number(value) || 0;
      return Math.round(v / 5) * 5;
    };

    const build = (predicate) => chapters
      .filter(predicate)
      .map((ch) => {
        const acc = totals.get(String(ch.id)) || { w: 0, comp: 0, cons: 0, title: "" };
        const denom = acc.w || ch.weightSum || 0;
        const comp = denom ? acc.comp / denom : 0;
        const cons = denom ? acc.cons / denom : 0;
        return {
          id: String(ch.id),
          label: acc.title || `${ch.id}`,
          weightSum: denom,
          company: round5(comp),
          consultant: round5(cons),
        };
      });

    return {
      chapters11: build((ch) => ch.includeIn11 !== false),
      chapters14: build((ch) => ch.includeIn14 !== false),
    };
  };

  const applyOverrides = (chapters, overrides) => chapters.map((row) => {
    const ov = overrides[row.id] || {};
    return {
      ...row,
      company: ov.useCompany && Number.isFinite(ov.company) ? ov.company : row.company,
      consultant: ov.useConsultant && Number.isFinite(ov.consultant) ? ov.consultant : row.consultant,
    };
  });

  const computeSpider = async ({ project, overrides = {}, dirHandle = null }) => {
    const { weights, source } = await loadWeights(dirHandle);
    const normOverrides = normalizeOverrides(overrides);
    const baseline = computeChapterScores(project, weights);
    const effective11 = applyOverrides(baseline.chapters11, normOverrides);
    const effective14 = applyOverrides(baseline.chapters14, normOverrides);
    return {
      schemaVersion: "1",
      weightsSource: source,
      generatedAt: new Date().toISOString(),
      overrides: normOverrides,
      baseline: {
        chapters_1_11: baseline.chapters11,
        chapters_1_14: baseline.chapters14,
      },
      effective: {
        chapters_1_11: effective11,
        chapters_1_14: effective14,
      },
    };
  };

  window.AutoBerichtSpider = {
    computeSpider,
    loadWeights,
    normalizeOverrides,
  };
})();
