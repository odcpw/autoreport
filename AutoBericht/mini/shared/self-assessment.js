/* Strict, dependency-free validation for imported self-assessment workbooks. */
(() => {
  const normalizeId = (raw) => String(raw || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/\.+$/g, "");

  const normalizeToken = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

  const findAssessmentSheetName = (sheetNames) => {
    const match = (sheetNames || []).find((name) => {
      const token = normalizeToken(name);
      return token.includes("selbstbeurteilung")
        || token.includes("autoevaluation")
        || token.includes("autovalutazione");
    });
    if (!match) {
      throw new Error("Workbook does not contain a recognized self-assessment sheet.");
    }
    return match;
  };

  const HEADER_ALIASES = {
    id: ["nr", "n", "no", "numero", "number", "id", "nrsb"],
    question: ["frage", "question", "domanda"],
    yes: ["ja", "oui", "si", "yes"],
    no: ["nein", "non", "no"],
    comment: ["bemerk", "remarque", "osserv", "comment"],
    evidence: ["nachweis", "beleg", "evidence", "document", "preuve"],
  };

  const findColumn = (header, aliases, exact = false, excluded = new Set()) => {
    // SheetJS preserves blank and merged Excel cells as sparse array slots.
    // Array#map preserves those holes while Array#findIndex visits them as
    // undefined, so normalize every slot through Array.from first.
    const tokens = Array.from(Array.isArray(header) ? header : [], normalizeToken);
    return tokens.findIndex((token, index) => !excluded.has(index) && aliases.some((alias) => (
      exact ? token === alias : token === alias || token.includes(alias)
    )));
  };

  const isMarked = (value, words) => {
    if (value === true || value === 1) return true;
    const token = normalizeToken(value);
    return token === "x" || token === "1" || token === "true" || words.includes(token);
  };

  const parseRows = (rows) => {
    if (!Array.isArray(rows)) throw new Error("Self-assessment sheet could not be read.");
    let headerRowIndex = -1;
    let columns = null;
    for (let index = 0; index < rows.length; index += 1) {
      const header = Array.isArray(rows[index]) ? rows[index] : [];
      const idColumn = findColumn(header, HEADER_ALIASES.id, true);
      const yesColumn = findColumn(header, HEADER_ALIASES.yes, true, new Set([idColumn]));
      const candidate = {
        id: idColumn,
        yes: yesColumn,
        no: findColumn(header, HEADER_ALIASES.no, true, new Set([idColumn, yesColumn])),
        question: findColumn(header, HEADER_ALIASES.question),
        comment: findColumn(header, HEADER_ALIASES.comment),
        evidence: findColumn(header, HEADER_ALIASES.evidence),
      };
      if (candidate.id >= 0 && candidate.yes >= 0 && candidate.no >= 0) {
        headerRowIndex = index;
        columns = candidate;
        break;
      }
    }
    if (!columns) {
      throw new Error("Self-assessment header must contain ID, Yes and No columns.");
    }

    const structuralIds = [];
    const entries = [];
    const seen = new Set();
    rows.slice(headerRowIndex + 1).forEach((row, offset) => {
      if (!Array.isArray(row)) return;
      const originalId = String(row[columns.id] || "").trim();
      const id = normalizeId(originalId);
      if (!/^[0-9]/.test(id)) return;
      if (seen.has(id)) {
        throw new Error(`Duplicate self-assessment ID ${originalId || id} (row ${headerRowIndex + offset + 2}).`);
      }
      seen.add(id);
      structuralIds.push(id);
      const yes = isMarked(row[columns.yes], ["ja", "oui", "si", "yes"]);
      const no = isMarked(row[columns.no], ["nein", "non", "no"]);
      if (yes && no) {
        throw new Error(`Conflicting Yes and No marks for self-assessment ID ${originalId || id}.`);
      }
      const question = columns.question >= 0 ? String(row[columns.question] || "").trim() : "";
      const comment = columns.comment >= 0 ? String(row[columns.comment] || "").trim() : "";
      const evidence = columns.evidence >= 0 ? String(row[columns.evidence] || "").trim() : "";
      const answer = yes ? 1 : (no ? 0 : null);
      if (answer == null && !comment && !evidence) return;
      entries.push({ id, originalId, question, answer, comment, evidence });
    });
    if (!structuralIds.length) throw new Error("Self-assessment sheet contains no numbered items.");
    if (!entries.length) throw new Error("Self-assessment sheet contains no answers, comments, or evidence to import.");
    return { headerRowIndex, columns, structuralIds, entries };
  };

  const validateProjectCoverage = (parsed, knownIds) => {
    const known = new Set(Array.from(knownIds || [], normalizeId).filter(Boolean));
    if (!known.size) throw new Error("Current project contains no self-assessment IDs.");
    const matched = parsed.structuralIds.filter((id) => known.has(id)).length;
    const requiredCount = Math.min(10, parsed.structuralIds.length);
    const ratio = matched / parsed.structuralIds.length;
    if (matched < requiredCount || ratio < 0.7) {
      throw new Error(
        `Workbook structure does not match this project (${matched}/${parsed.structuralIds.length} IDs recognized).`,
      );
    }
    return { matched, total: parsed.structuralIds.length, ratio };
  };

  window.AutoBerichtSelfAssessment = {
    normalizeId,
    findAssessmentSheetName,
    parseRows,
    validateProjectCoverage,
  };
})();
