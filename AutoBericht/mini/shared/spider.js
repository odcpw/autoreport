(() => {
  const LEVEL_TO_PCT = (level) => Math.max(0, Math.min(100, (Number(level || 1) - 1) * 25));

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

  const computeChapterScores = (project, weights) => {
    const totals = new Map(); // chapter -> {wSum, compSum, consSum}
    const itemWeights = new Map((weights.items || []).map((it) => [String(it.id), Number(it.weight || 0)]));
    (project?.chapters || []).forEach((chapter) => {
      (chapter.rows || []).forEach((row) => {
        if (row.kind === "section") return;
        if (row.type === "field_observation" || String(row.id || "").startsWith("4.8")) return;
        const id = String(row.id || "");
        const weight = itemWeights.get(id);
        if (!weight || Number.isNaN(weight)) return;
        const chapterId = id.split(".")[0];
        const chapterTitle = (() => {
          if (row.sectionLabel) return row.sectionLabel;
          if (chapter.title?.de) return `${chapterId}. ${chapter.title.de}`;
          return chapterId;
        })();
        const consPct = (() => {
          const ws = row.workstate || {};
          if (ws.includeFinding === false) return 100;
          return LEVEL_TO_PCT(ws.selectedLevel || 1);
        })();
        const compPct = pctFromCustomer(row);
        const acc = totals.get(chapterId) || { w: 0, comp: 0, cons: 0, title: chapterTitle };
        acc.w += weight;
        acc.comp += weight * compPct;
        acc.cons += weight * consPct;
        if (!acc.title && chapterTitle) acc.title = chapterTitle;
        totals.set(chapterId, acc);
      });
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
