(() => {
  const rowToText = (value, toText) => {
    if (typeof toText === "function") return toText(value);
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  };

  const stripLeadingNumber = (value) => String(value || "")
    .replace(/^\s*\d+(?:\.\d+)*(?:\s|[.:-]\s*)?/, "")
    .trim();

  const isSectionRow = (row) => String(row?.kind || "").toLowerCase() === "section";

  const isReportReadyRow = (row) => {
    const ws = row?.workstate;
    if (!ws || ws.includeFinding == null) return false;
    return ws.includeFinding === true && ws.done === true;
  };

  const resolveSectionId = (row, chapterId) => {
    const sectionId = String(row?.sectionId || "").trim();
    if (sectionId) return sectionId;
    const parts = String(row?.id || "").split(".");
    if (parts.length >= 2) return `${parts[0]}.${parts[1]}`;
    return `${chapterId}.1`;
  };

  const resolveSectionTitle = (row) => {
    const rawTitle = String(row?.title || row?.id || "");
    const cleaned = stripLeadingNumber(rawTitle);
    return cleaned || rawTitle;
  };

  const resolveFindingText = (row, toText) => {
    const ws = row?.workstate;
    if (ws && Object.prototype.hasOwnProperty.call(ws, "findingText")) {
      return rowToText(ws.findingText, toText);
    }
    return rowToText(row?.master?.finding, toText);
  };

  const resolveRecommendationText = (row, toText) => {
    const ws = row?.workstate || {};
    if (ws.includeRecommendation === false) return "";
    if (Object.prototype.hasOwnProperty.call(ws, "recommendationText")) {
      return rowToText(ws.recommendationText, toText);
    }
    return rowToText(row?.master?.recommendation, toText);
  };

  const resolvePriorityText = (row) => {
    const ws = row?.workstate || {};
    const raw = Number(ws.priority);
    if (!Number.isFinite(raw)) return "";
    const value = Math.round(raw);
    if (value < 1 || value > 4) return "";
    return String(value);
  };

  const isFieldObservationChapter = (chapterId) => String(chapterId || "").includes(".");

  const resolveSectionDisplayId = (sectionId, chapterId, sectionMap) => {
    const key = String(sectionId || "").trim();
    if (!key) return "";
    if (sectionMap && sectionMap.has(key)) return `${chapterId}.${sectionMap.get(key)}`;
    return key;
  };

  const buildIncludedSections = (rows, includeRow = isReportReadyRow) => {
    const included = new Set();
    rows.forEach((row) => {
      if (isSectionRow(row) || !includeRow(row)) return;
      const sectionId = String(row?.sectionId || "").trim();
      if (sectionId) included.add(sectionId);
    });
    return included;
  };

  const buildRenumberMap = (rows, chapterId, includeRow = isReportReadyRow) => {
    const rowMap = new Map();
    const sectionMap = new Map();
    const sectionCounts = new Map();
    let itemCount = 0;

    rows.forEach((row) => {
      if (isSectionRow(row) || !includeRow(row)) return;
      const rowId = String(row?.id || "").trim();
      if (!rowId) return;
      if (isFieldObservationChapter(chapterId)) {
        itemCount += 1;
        rowMap.set(rowId, `${chapterId}.${itemCount}`);
        return;
      }
      const sectionId = resolveSectionId(row, chapterId);
      if (!sectionMap.has(sectionId)) sectionMap.set(sectionId, sectionMap.size + 1);
      const count = (sectionCounts.get(sectionId) || 0) + 1;
      sectionCounts.set(sectionId, count);
      rowMap.set(rowId, `${chapterId}.${sectionMap.get(sectionId)}.${count}`);
    });

    return { rowMap, sectionMap };
  };

  const orderRowsForChapter = (chapter) => {
    const rows = Array.isArray(chapter?.rows) ? [...chapter.rows] : [];
    const order = Array.isArray(chapter?.meta?.order) ? chapter.meta.order : [];
    if (!order.length) return rows;
    const sections = rows.filter((row) => isSectionRow(row));
    const items = rows.filter((row) => !isSectionRow(row));
    const byId = new Map(items.map((row) => [String(row?.id || ""), row]));
    const orderedItems = [];
    order.forEach((id) => {
      const key = String(id || "");
      const row = byId.get(key);
      if (row && !orderedItems.includes(row)) orderedItems.push(row);
    });
    items.forEach((row) => {
      if (!orderedItems.includes(row)) orderedItems.push(row);
    });
    return [...sections, ...orderedItems];
  };

  const buildChapterRows = (chapter, options = {}) => {
    const chapterId = String(chapter?.id || "");
    const includeRow = typeof options.includeRow === "function" ? options.includeRow : isReportReadyRow;
    const toText = options.toText;
    const rows = Array.isArray(options.rows) ? [...options.rows] : orderRowsForChapter(chapter);
    const titleForFinding = typeof options.titleForFinding === "function"
      ? options.titleForFinding
      : (row, cid) => (cid === "4.8" ? String(row?.titleOverride || "").trim() : "");

    const includedSections = buildIncludedSections(rows, includeRow);
    const { rowMap, sectionMap } = buildRenumberMap(rows, chapterId, includeRow);
    const output = [];

    rows.forEach((row) => {
      if (isSectionRow(row)) {
        const sectionId = String(row?.id || "").trim();
        if (!sectionId || !includedSections.has(sectionId)) return;
        output.push({
          kind: "section",
          id: resolveSectionDisplayId(sectionId, chapterId, sectionMap),
          title: resolveSectionTitle(row),
        });
        return;
      }
      if (!includeRow(row)) return;
      const rowId = String(row?.id || "").trim();
      output.push({
        kind: "finding",
        id: rowMap.get(rowId) || rowId,
        title: titleForFinding(row, chapterId),
        finding: resolveFindingText(row, toText),
        recommendation: resolveRecommendationText(row, toText),
        priority: resolvePriorityText(row),
      });
    });

    return output;
  };

  window.AutoBerichtReportRows = {
    rowToText,
    stripLeadingNumber,
    isSectionRow,
    isReportReadyRow,
    resolveSectionId,
    resolveSectionTitle,
    resolveFindingText,
    resolveRecommendationText,
    resolvePriorityText,
    isFieldObservationChapter,
    resolveSectionDisplayId,
    buildIncludedSections,
    buildRenumberMap,
    orderRowsForChapter,
    buildChapterRows,
  };
})();
