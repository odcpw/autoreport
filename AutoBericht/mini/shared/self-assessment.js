/* Shared parsing for imported self-assessment workbooks. */
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

  // Prefer the sheet named like the Suva workbook; otherwise the first sheet.
  const findAssessmentSheetName = (sheetNames) => {
    const names = Array.isArray(sheetNames) ? sheetNames : [];
    const match = names.find((name) => {
      const token = normalizeToken(name);
      return token.includes("selbstbeurteilung")
        || token.includes("autoevaluation")
        || token.includes("autovalutazione");
    });
    return match || names[0] || "";
  };

  const HEADER_ALIASES = {
    id: ["nr", "n", "no", "numero", "number", "id", "nrsb"],
    question: ["frage", "question", "domanda"],
    yes: ["ja", "oui", "si", "yes"],
    no: ["nein", "non", "no"],
    comment: ["bemerk", "remarque", "osserv", "comment"],
    evidence: ["nachweis", "beleg", "evidence", "document", "preuve"],
  };

  // Column layout of the Suva workbook, used when no header row is recognized.
  const DEFAULT_COLUMNS = { id: 0, question: 2, yes: 3, no: 4, comment: 5, evidence: -1 };

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
    if (!columns) columns = { ...DEFAULT_COLUMNS };

    const structuralIds = new Set();
    const entries = new Map();
    rows.slice(headerRowIndex + 1).forEach((row) => {
      if (!Array.isArray(row)) return;
      const originalId = String(row[columns.id] || "").trim();
      const id = normalizeId(originalId);
      if (!/^[0-9]/.test(id)) return;
      structuralIds.add(id);
      const yes = isMarked(row[columns.yes], ["ja", "oui", "si", "yes"]);
      const no = isMarked(row[columns.no], ["nein", "non", "no"]);
      // A row marked both Ja and Nein counts as Nein.
      const answer = no ? 0 : (yes ? 1 : null);
      const question = columns.question >= 0 ? String(row[columns.question] || "").trim() : "";
      const comment = columns.comment >= 0 ? String(row[columns.comment] || "").trim() : "";
      const evidence = columns.evidence >= 0 ? String(row[columns.evidence] || "").trim() : "";
      if (answer == null && !comment && !evidence) return;
      // A repeated ID keeps the later row, as the April importer did.
      entries.set(id, { id, originalId, question, answer, comment, evidence });
    });
    if (!structuralIds.size) throw new Error("Self-assessment sheet contains no numbered items.");
    if (!entries.size) throw new Error("Self-assessment sheet contains no answers, comments, or evidence to import.");
    return {
      headerRowIndex,
      columns,
      structuralIds: Array.from(structuralIds),
      entries: Array.from(entries.values()),
    };
  };

  window.AutoBerichtSelfAssessment = {
    normalizeId,
    findAssessmentSheetName,
    parseRows,
  };
})();
