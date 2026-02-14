/*
 * No-VBA Word export orchestrator.
 *
 * Responsibilities:
 * - Read a user-selected `.docx` template and replace placeholder markers.
 * - Build chapter payloads from sidecar rows using report-ready filtering.
 * - Inject generated media (logos + spider chart) and per-chapter thermo tables.
 * - Write final output as `Outputs/YYYY-MM-DD_AutoBericht_NoVBA.docx`.
 *
 * Key dependencies:
 * - `window.AutoBerichtReportRows` for row projection/renumbering.
 * - `window.AutoBerichtWordDocxZip` for ZIP read/write helpers.
 * - `window.AutoBerichtWordDocxXml` for OOXML marker/relation helpers.
 * - `window.AutoBerichtSpiderChart` for shared spider rendering.
 */
(() => {
  const textDecoder = new TextDecoder();
  const textEncoder = new TextEncoder();
  const reportRows = window.AutoBerichtReportRows || {};
  const zipTools = window.AutoBerichtWordDocxZip;
  const xmlTools = window.AutoBerichtWordDocxXml;

  const requireHelper = (scope, toolbox, name) => {
    const fn = toolbox?.[name];
    if (typeof fn !== "function") {
      throw new Error(`${scope}.${name} helper is unavailable.`);
    }
    return fn;
  };

  const unzipAllEntries = requireHelper("AutoBerichtWordDocxZip", zipTools, "unzipAllEntries");
  const buildZipStore = requireHelper("AutoBerichtWordDocxZip", zipTools, "buildZipStore");

  const xmlEscape = requireHelper("AutoBerichtWordDocxXml", xmlTools, "xmlEscape");
  const hasMarker = requireHelper("AutoBerichtWordDocxXml", xmlTools, "hasMarker");
  const replaceParagraphMarker = requireHelper("AutoBerichtWordDocxXml", xmlTools, "replaceParagraphMarker");
  const replaceAllParagraphMarkers = requireHelper("AutoBerichtWordDocxXml", xmlTools, "replaceAllParagraphMarkers");
  const replaceTextMarkers = requireHelper("AutoBerichtWordDocxXml", xmlTools, "replaceTextMarkers");
  const ensurePngContentType = requireHelper("AutoBerichtWordDocxXml", xmlTools, "ensurePngContentType");
  const getNextRelId = requireHelper("AutoBerichtWordDocxXml", xmlTools, "getNextRelId");
  const appendRelationship = requireHelper("AutoBerichtWordDocxXml", xmlTools, "appendRelationship");
  const emuFromCm = requireHelper("AutoBerichtWordDocxXml", xmlTools, "emuFromCm");
  const drawingXml = requireHelper("AutoBerichtWordDocxXml", xmlTools, "drawingXml");
  const ensureUpdateFieldsOnOpen = requireHelper("AutoBerichtWordDocxXml", xmlTools, "ensureUpdateFieldsOnOpen");

  const stripLeadingNumber = typeof reportRows.stripLeadingNumber === "function"
    ? reportRows.stripLeadingNumber
    : (value) => String(value || "").replace(/^\s*\d+(?:\.\d+)*(?:\s|[.:-]\s*)?/, "").trim();

  const rowToText = typeof reportRows.rowToText === "function"
    ? reportRows.rowToText
    : (value, toText) => {
      if (typeof toText === "function") return toText(value);
      if (Array.isArray(value)) return value.join("\n");
      if (value == null) return "";
      return String(value);
    };

  const isSectionRow = typeof reportRows.isSectionRow === "function"
    ? reportRows.isSectionRow
    : (row) => String(row?.kind || "").toLowerCase() === "section";

  const isIncludedRow = typeof reportRows.isReportReadyRow === "function"
    ? reportRows.isReportReadyRow
    : (row) => {
      const ws = row?.workstate;
      if (!ws || ws.includeFinding == null) return false;
      return ws.includeFinding === true && ws.done === true;
    };

  const resolveSectionId = typeof reportRows.resolveSectionId === "function"
    ? reportRows.resolveSectionId
    : (row, chapterId) => {
      const sectionId = String(row?.sectionId || "").trim();
      if (sectionId) return sectionId;
      const parts = String(row?.id || "").split(".");
      if (parts.length >= 2) return `${parts[0]}.${parts[1]}`;
      return `${chapterId}.1`;
    };

  const resolveSectionTitle = typeof reportRows.resolveSectionTitle === "function"
    ? reportRows.resolveSectionTitle
    : (row) => {
      const rawTitle = String(row?.title || row?.id || "");
      const cleaned = stripLeadingNumber(rawTitle);
      return cleaned || rawTitle;
    };

  const resolveFindingText = typeof reportRows.resolveFindingText === "function"
    ? (row, toText) => reportRows.resolveFindingText(row, toText)
    : (row, toText) => {
      const ws = row?.workstate;
      if (ws && Object.prototype.hasOwnProperty.call(ws, "findingText")) {
        return rowToText(ws.findingText, toText);
      }
      return rowToText(row?.master?.finding, toText);
    };

  const resolveRecommendationText = typeof reportRows.resolveRecommendationText === "function"
    ? (row, toText) => reportRows.resolveRecommendationText(row, toText)
    : (row, toText) => {
      const ws = row?.workstate || {};
      if (ws.includeRecommendation === false) return "";
      if (Object.prototype.hasOwnProperty.call(ws, "recommendationText")) {
        return rowToText(ws.recommendationText, toText);
      }
      return rowToText(row?.master?.recommendation, toText);
    };

  const resolvePriorityText = typeof reportRows.resolvePriorityText === "function"
    ? reportRows.resolvePriorityText
    : (row) => {
      const ws = row?.workstate || {};
      const raw = Number(ws.priority);
      if (!Number.isFinite(raw)) return "";
      const value = Math.round(raw);
      if (value < 1 || value > 4) return "";
      return String(value);
    };

  const buildAddressLine = (meta = {}) => {
    const address = String(meta?.address || "").trim();
    const postal = String(meta?.plz || "").trim();
    const city = String(meta?.city || "").trim();
    const cityLine = [postal, city].filter(Boolean).join(" ");
    return [address, cityLine].filter(Boolean).join(", ");
  };

  const roundToNearestTen = (value) => {
    const raw = Number(value);
    if (!Number.isFinite(raw)) return 0;
    return Math.max(0, Math.min(100, Math.round(raw / 10) * 10));
  };

  const thermoTextByLocale = (locale, companyName) => {
    const normalized = String(locale || "de-CH").toLowerCase();
    const company = String(companyName || "").trim() || "Company";
    if (normalized.startsWith("fr")) {
      return {
        companyLabel: `Autoevaluation de ${company}`,
        consultantLabel: "Evaluation par Suva",
      };
    }
    if (normalized.startsWith("it")) {
      return {
        companyLabel: `Autovalutazione di ${company}`,
        consultantLabel: "Valutazione da parte di Suva",
      };
    }
    return {
      companyLabel: `Selbstbeurteilung von ${company}`,
      consultantLabel: "Beurteilung durch Suva",
    };
  };

  const thermoCellXml = ({
    width = 320,
    widthType = "dxa",
    text = "",
    fill = "",
    align = "center",
    bold = false,
    showGrid = false,
  }) => [
    "<w:tc><w:tcPr>",
    `<w:tcW w:w="${width}" w:type="${widthType}"/>`,
    showGrid ? [
      "<w:tcBorders>",
      "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "</w:tcBorders>",
    ].join("") : "",
    fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : "",
    "</w:tcPr>",
    paragraphXml(text, { align, bold }),
    "</w:tc>",
  ].join("");

  const buildThermoTableXml = ({
    companyValue = 0,
    consultantValue = 0,
    locale = "de-CH",
    companyName = "Company",
  }) => {
    const segments = 10;
    const labelWidth = 4680; // 50% of 9360 twips
    const segmentWidth = 468; // each 5% of 9360 twips
    const companyRounded = roundToNearestTen(companyValue);
    const consultantRounded = roundToNearestTen(consultantValue);
    const companyFilled = Math.max(0, Math.min(segments, Math.round(companyRounded / 10)));
    const consultantFilled = Math.max(0, Math.min(segments, Math.round(consultantRounded / 10)));
    const labels = thermoTextByLocale(locale, companyName);

    const scaleCells = Array.from({ length: segments }, (_, index) => {
      const marker = index === 0 ? "-" : index === segments - 1 ? "+" : "";
      return thermoCellXml({
        width: segmentWidth,
        widthType: "dxa",
        text: marker,
        align: "center",
        bold: !!marker,
        showGrid: false,
      });
    }).join("");

    const buildFillCells = (filled, fillColor) => (
      Array.from({ length: segments }, (_, index) => thermoCellXml({
        width: segmentWidth,
        widthType: "dxa",
        text: "",
        fill: index < filled ? fillColor : "",
        align: "center",
        showGrid: true,
      })).join("")
    );

    const scaleRow = [
      "<w:tr>",
      thermoCellXml({ width: labelWidth, widthType: "dxa", text: "", align: "left" }),
      scaleCells,
      "</w:tr>",
    ].join("");

    const companyRow = [
      "<w:tr>",
      thermoCellXml({ width: labelWidth, widthType: "dxa", text: labels.companyLabel, align: "left" }),
      buildFillCells(companyFilled, "BDD7EE"),
      "</w:tr>",
    ].join("");

    const consultantRow = [
      "<w:tr>",
      thermoCellXml({ width: labelWidth, widthType: "dxa", text: labels.consultantLabel, align: "left" }),
      buildFillCells(consultantFilled, "F8CBAD"),
      "</w:tr>",
    ].join("");

    const gridCols = [
      `<w:gridCol w:w="${labelWidth}"/>`,
      ...Array.from({ length: segments }, () => `<w:gridCol w:w="${segmentWidth}"/>`),
    ].join("");

    return [
      "<w:tbl>",
      "<w:tblPr>",
      "<w:tblpPr w:vertAnchor=\"text\" w:horzAnchor=\"text\" w:leftFromText=\"141\" w:rightFromText=\"141\" w:tblpX=\"105\" w:tblpY=\"1\"/>",
      "<w:tblOverlap w:val=\"never\"/>",
      "<w:tblW w:w=\"5000\" w:type=\"pct\"/>",
      "<w:jc w:val=\"start\"/>",
      "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>",
      "<w:tblLayout w:type=\"fixed\"/>",
      "<w:tblCellMar>",
      "<w:top w:w=\"0\" w:type=\"dxa\"/>",
      "<w:start w:w=\"108\" w:type=\"dxa\"/>",
      "<w:bottom w:w=\"0\" w:type=\"dxa\"/>",
      "<w:end w:w=\"108\" w:type=\"dxa\"/>",
      "</w:tblCellMar>",
      "<w:tblBorders>",
      "<w:top w:val=\"nil\"/>",
      "<w:left w:val=\"nil\"/>",
      "<w:bottom w:val=\"nil\"/>",
      "<w:right w:val=\"nil\"/>",
      "<w:insideH w:val=\"nil\"/>",
      "<w:insideV w:val=\"nil\"/>",
      "</w:tblBorders>",
      "<w:tblLook w:val=\"0000\" w:noHBand=\"0\" w:noVBand=\"0\" w:firstColumn=\"0\" w:lastRow=\"0\" w:lastColumn=\"0\" w:firstRow=\"0\"/>",
      "</w:tblPr>",
      "<w:tblGrid>",
      gridCols,
      "</w:tblGrid>",
      scaleRow,
      companyRow,
      consultantRow,
      "</w:tbl>",
    ].join("");
  };

  const buildThermoScoreMap = (spiderData) => {
    const map = new Map();
    const addRows = (rows) => {
      (rows || []).forEach((row) => {
        const id = String(row?.id || "").trim();
        if (!id) return;
        map.set(id, {
          company: Number(row?.company || 0),
          consultant: Number(row?.consultant || 0),
        });
      });
    };
    addRows(spiderData?.effective?.chapters_1_14);
    addRows(spiderData?.effective?.chapters_1_11);
    return map;
  };

  const resolveThermoScore = (scoreMap, chapterId) => {
    const id = String(chapterId || "").trim();
    const value = scoreMap.get(id);
    if (value) return value;
    return { company: 0, consultant: 0 };
  };

  const thermoMarkersForChapter = (chapterId) => {
    const id = String(chapterId || "").trim();
    if (!id) return [];
    const variants = new Set([
      id,
      id.replace(/\./g, "_"),
      id.replace(/[^0-9A-Za-z.]/g, ""),
      id.replace(/[^0-9A-Za-z]/g, ""),
    ]);
    return Array.from(variants).filter(Boolean).map((token) => `THERMO${token}$$`);
  };

  const isFieldObservationChapter = typeof reportRows.isFieldObservationChapter === "function"
    ? reportRows.isFieldObservationChapter
    : (chapterId) => String(chapterId || "").includes(".");

  const resolveSectionDisplayId = typeof reportRows.resolveSectionDisplayId === "function"
    ? reportRows.resolveSectionDisplayId
    : (sectionId, chapterId, sectionMap) => {
      const key = String(sectionId || "").trim();
      if (!key) return "";
      if (sectionMap && sectionMap.has(key)) return `${chapterId}.${sectionMap.get(key)}`;
      return key;
    };

  const buildIncludedSections = typeof reportRows.buildIncludedSections === "function"
    ? (rows) => reportRows.buildIncludedSections(rows, isIncludedRow)
    : (rows) => {
      const included = new Set();
      rows.forEach((row) => {
        if (isSectionRow(row) || !isIncludedRow(row)) return;
        const sectionId = String(row?.sectionId || "").trim();
        if (sectionId) included.add(sectionId);
      });
      return included;
    };

  const buildRenumberMap = typeof reportRows.buildRenumberMap === "function"
    ? (rows, chapterId) => reportRows.buildRenumberMap(rows, chapterId, isIncludedRow)
    : (rows, chapterId) => {
      const rowMap = new Map();
      const sectionMap = new Map();
      const sectionCounts = new Map();
      let itemCount = 0;
      rows.forEach((row) => {
        if (isSectionRow(row) || !isIncludedRow(row)) return;
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

  const orderRowsForChapter = typeof reportRows.orderRowsForChapter === "function"
    ? reportRows.orderRowsForChapter
    : (chapter) => Array.isArray(chapter?.rows) ? [...chapter.rows] : [];

  const buildChapterRows = (chapter, toText) => {
    if (typeof reportRows.buildChapterRows === "function") {
      return reportRows.buildChapterRows(chapter, {
        toText,
        includeRow: isIncludedRow,
        titleForFinding: (row, chapterId) => (
          chapterId === "4.8" ? String(row?.titleOverride || "").trim() : ""
        ),
      });
    }
    const rows = orderRowsForChapter(chapter);
    const includedSections = buildIncludedSections(rows);
    const chapterId = String(chapter?.id || "");
    const { rowMap, sectionMap } = buildRenumberMap(rows, chapterId);
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
      if (!isIncludedRow(row)) return;
      const rowId = String(row?.id || "").trim();
      output.push({
        kind: "finding",
        id: rowMap.get(rowId) || rowId,
        title: chapter?.id === "4.8" ? String(row?.titleOverride || "").trim() : "",
        finding: resolveFindingText(row, toText),
        recommendation: resolveRecommendationText(row, toText),
        priority: resolvePriorityText(row),
      });
    });
    return output;
  };

  const paragraphXml = (text, options = {}) => {
    const safe = xmlEscape(text || "");
    const boldStart = options.bold ? "<w:rPr><w:b/></w:rPr>" : "";
    const pPrParts = [];
    if (options.styleId) {
      pPrParts.push(`<w:pStyle w:val="${xmlEscape(options.styleId)}"/>`);
    }
    if (options.align) {
      pPrParts.push(`<w:jc w:val="${xmlEscape(options.align)}"/>`);
    }
    const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";
    return `<w:p>${pPr}<w:r>${boldStart}<w:t xml:space=\"preserve\">${safe}</w:t></w:r></w:p>`;
  };

  const normalizeMarkdownInline = (value) => String(value || "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/_([^_]+)_/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/^\s*>\s?/g, "");

  const splitTextByUrls = (text) => {
    const source = String(text || "");
    const out = [];
    const urlRegex = /(https?:\/\/[^\s<>"']+)/g;
    let last = 0;
    let match = urlRegex.exec(source);
    while (match) {
      const start = match.index;
      let rawUrl = match[1];
      const rawEnd = start + rawUrl.length;
      let end = rawEnd;
      let trim = rawUrl;
      while (/[.,;!?]$/.test(trim)) {
        trim = trim.slice(0, -1);
        end -= 1;
      }
      if (start > last) out.push({ type: "text", text: source.slice(last, start) });
      if (trim) out.push({ type: "link", text: trim, url: trim });
      if (rawEnd > end) {
        out.push({ type: "text", text: source.slice(end, rawEnd) });
      }
      last = end;
      urlRegex.lastIndex = end;
      match = urlRegex.exec(source);
    }
    if (last < source.length) out.push({ type: "text", text: source.slice(last) });
    return out;
  };

  const parseInlineMarkdownSegments = (text) => {
    const source = String(text || "");
    const parts = [];
    const markdownLinkRegex = /\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g;
    let cursor = 0;
    let match = markdownLinkRegex.exec(source);
    while (match) {
      if (match.index > cursor) {
        parts.push({ type: "text", text: source.slice(cursor, match.index) });
      }
      parts.push({
        type: "link",
        text: normalizeMarkdownInline(match[1]),
        url: match[2],
      });
      cursor = match.index + match[0].length;
      match = markdownLinkRegex.exec(source);
    }
    if (cursor < source.length) {
      parts.push({ type: "text", text: source.slice(cursor) });
    }

    const expanded = [];
    parts.forEach((part) => {
      if (part.type === "link") {
        expanded.push(part);
        return;
      }
      splitTextByUrls(normalizeMarkdownInline(part.text)).forEach((token) => expanded.push(token));
    });
    return expanded;
  };

  const inlineSegmentsToXml = (segments, options = {}) => {
    const bold = !!options.bold;
    const xml = (segments || []).map((segment) => {
      if (!segment) return "";
      if (segment.type === "link") {
        const url = String(segment.url || "").trim();
        const label = String(segment.text || url).trim() || url;
        if (!url) return "";
        const instr = xmlEscape(`HYPERLINK "${url}"`);
        return [
          `<w:fldSimple w:instr="${instr}">`,
          "<w:r>",
          "<w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>",
          `<w:t xml:space=\"preserve\">${xmlEscape(label)}</w:t>`,
          "</w:r>",
          "</w:fldSimple>",
        ].join("");
      }
      const value = String(segment.text || "");
      if (!value) return "";
      return [
        "<w:r>",
        bold ? "<w:rPr><w:b/></w:rPr>" : "",
        `<w:t xml:space=\"preserve\">${xmlEscape(value)}</w:t>`,
        "</w:r>",
      ].join("");
    }).join("");
    return xml || "<w:r><w:t xml:space=\"preserve\"></w:t></w:r>";
  };

  const richParagraphXml = (text, options = {}) => {
    const pPrParts = [];
    if (options.styleId) {
      pPrParts.push(`<w:pStyle w:val="${xmlEscape(options.styleId)}"/>`);
    }
    if (options.align) {
      pPrParts.push(`<w:jc w:val="${xmlEscape(options.align)}"/>`);
    }
    const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";
    const segments = parseInlineMarkdownSegments(String(text || ""));
    if (options.bullet) {
      segments.unshift({ type: "text", text: "• " });
    }
    return `<w:p>${pPr}${inlineSegmentsToXml(segments, options)}</w:p>`;
  };

  const letteredParagraphXml = (label, text, options = {}) => {
    const safeLabel = xmlEscape(label || "");
    const safeText = xmlEscape(text || "");
    const numId = String(options?.numId || "").trim();
    if (numId) {
      return [
        "<w:p>",
        "<w:pPr>",
        "<w:pStyle w:val=\"ListParagraph\"/>",
        "<w:numPr>",
        "<w:ilvl w:val=\"0\"/>",
        `<w:numId w:val="${xmlEscape(numId)}"/>`,
        "</w:numPr>",
        "</w:pPr>",
        `<w:r><w:t xml:space=\"preserve\">${safeText}</w:t></w:r>`,
        "</w:p>",
      ].join("");
    }
    return [
      "<w:p>",
      "<w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr>",
      `<w:r><w:t xml:space=\"preserve\">${safeLabel}.</w:t></w:r>`,
      "<w:r><w:tab/></w:r>",
      `<w:r><w:t xml:space=\"preserve\">${safeText}</w:t></w:r>`,
      "</w:p>",
    ].join("");
  };

  const multiParagraphXml = (text) => {
    const lines = String(text || "").split(/\r?\n/);
    const nonEmpty = lines.length ? lines : [""];
    return nonEmpty.map((line) => {
      const trimmed = line.trim();
      const isBullet = trimmed.startsWith("- ") || trimmed.startsWith("* ");
      const value = isBullet ? trimmed.slice(2) : line;
      return richParagraphXml(value, { bullet: isBullet });
    }).join("");
  };

  const alphaLabel = (index) => {
    let n = Number(index) + 1;
    if (!Number.isFinite(n) || n < 1) return "";
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  };

  const collapseSingleLine = (text) => String(text || "")
    .replace(/\r?\n+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const buildChapter0Xml = (chapter, toText, options = {}) => {
    const rows = buildChapterRows(chapter, toText).filter((entry) => entry.kind === "finding");
    if (!rows.length) return paragraphXml("(No included findings)");
    const numId = String(options?.numId || "").trim();
    return rows.map((entry, index) => {
      const label = alphaLabel(index);
      const line = collapseSingleLine(entry.recommendation || "");
      return letteredParagraphXml(label, line, { numId });
    }).join("");
  };

  const findUpperLetterNumId = (numberingXml) => {
    const xml = String(numberingXml || "");
    if (!xml) return "";
    const abstractWithUpperLetter = new Set();
    const abstractRegex = /<w:abstractNum\b[^>]*w:abstractNumId="([^"]+)"[^>]*>[\s\S]*?<\/w:abstractNum>/g;
    let abstractMatch = abstractRegex.exec(xml);
    while (abstractMatch) {
      const abstractId = String(abstractMatch[1] || "").trim();
      const block = abstractMatch[0];
      const hasUpperLetter = /<w:lvl\b[^>]*w:ilvl="0"[^>]*>[\s\S]*?<w:numFmt\b[^>]*w:val="upperLetter"/.test(block);
      if (abstractId && hasUpperLetter) abstractWithUpperLetter.add(abstractId);
      abstractMatch = abstractRegex.exec(xml);
    }
    if (!abstractWithUpperLetter.size) return "";

    const numRegex = /<w:num\b[^>]*w:numId="([^"]+)"[^>]*>[\s\S]*?<w:abstractNumId\b[^>]*w:val="([^"]+)"[^>]*\/>[\s\S]*?<\/w:num>/g;
    let numMatch = numRegex.exec(xml);
    while (numMatch) {
      const numId = String(numMatch[1] || "").trim();
      const abstractId = String(numMatch[2] || "").trim();
      if (numId && abstractWithUpperLetter.has(abstractId)) return numId;
      numMatch = numRegex.exec(xml);
    }
    return "";
  };

  const buildChapterTableXml = (chapter, toText) => {
    const rows = buildChapterRows(chapter, toText);
    const chapterMeta = chapter?.meta || {};
    const positivesText = chapterMeta.positivesInclude === true && chapterMeta.positivesDone === true
      ? String(chapterMeta.positivesText || "").trim()
      : "";
    if (!rows.length && !positivesText) return paragraphXml("(No included findings)");
    const chapterId = String(chapter?.id || "");
    const widthCol1 = 3150;
    const widthCol2 = 5220;
    const widthCol3 = 630;
    const widthLeftBlock = widthCol1 + widthCol2;
    const widthTotal = widthLeftBlock + widthCol3;

    const tableRows = [];
    tableRows.push(
      [
        "<w:tr>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthLeftBlock}" w:type="dxa"/>`,
        "<w:gridSpan w:val=\"2\"/>",
        "<w:tcBorders><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tcBorders>",
        "</w:tcPr>",
        positivesText ? multiParagraphXml(positivesText) : paragraphXml(""),
        "</w:tc>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthCol3}" w:type="dxa"/>`,
        "<w:tcBorders><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tcBorders>",
        "</w:tcPr>",
        paragraphXml("✓", { bold: true, align: "center" }),
        "</w:tc>",
        "</w:tr>",
      ].join(""),
    );

    tableRows.push(
      [
        "<w:tr>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthLeftBlock}" w:type="dxa"/>`,
        "<w:gridSpan w:val=\"2\"/>",
        "</w:tcPr>",
        paragraphXml("Systempunkte mit Verbesserungspotenzial", { bold: true }),
        "</w:tc>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthCol3}" w:type="dxa"/>`,
        "</w:tcPr>",
        paragraphXml(""),
        "</w:tc>",
        "</w:tr>",
      ].join(""),
    );

    tableRows.push(
      [
        "<w:tr>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthCol1}" w:type="dxa"/>`,
        "</w:tcPr>",
        paragraphXml("Ist-Zustand", { bold: true }),
        "</w:tc>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthCol2}" w:type="dxa"/>`,
        "</w:tcPr>",
        paragraphXml("Lösungsansätze", { bold: true }),
        "</w:tc>",
        "<w:tc><w:tcPr>",
        `<w:tcW w:w="${widthCol3}" w:type="dxa"/>`,
        "</w:tcPr>",
        paragraphXml("Prio", { bold: true, align: "center" }),
        "</w:tc>",
        "</w:tr>",
      ].join(""),
    );

    rows.forEach((entry) => {
      if (entry.kind === "section") {
        const sectionHeading = `${entry.id || ""}${entry.title ? ` ${entry.title}` : ""}`.trim();
        const sectionText = String(entry.title || "").trim() || sectionHeading;
        tableRows.push(
          [
            "<w:tr>",
            "<w:tc><w:tcPr>",
            `<w:tcW w:w="${widthTotal}" w:type="dxa"/>`,
            "<w:gridSpan w:val=\"3\"/>",
            "</w:tcPr>",
            paragraphXml(sectionText, { styleId: "Heading2" }),
            "</w:tc>",
            "</w:tr>",
          ].join(""),
        );
        return;
      }

      const heading = `${entry.id || ""}${entry.title ? ` ${entry.title}` : ""}`.trim();
      let findingCell = "";
      if (chapterId === "4.8") {
        const headingText = stripLeadingNumber(String(entry.title || "").trim() || heading) || heading;
        findingCell = `${paragraphXml(headingText, { styleId: "Heading3" })}${multiParagraphXml(entry.finding || "")}`;
      } else {
        const findingLines = String(entry.finding || "").split(/\r?\n/);
        const headingTextRaw = String(findingLines.shift() || "").trim()
          || String(entry.title || "").trim()
          || String(entry.id || "").trim();
        const headingText = stripLeadingNumber(headingTextRaw) || headingTextRaw;
        findingCell = paragraphXml(headingText, { styleId: "Heading3" });
        if (findingLines.length) {
          findingCell += findingLines.map((line) => richParagraphXml(line)).join("");
        }
      }
      const recommendationCell = multiParagraphXml(entry.recommendation || "");
      const priorityCell = paragraphXml(entry.priority || "", { bold: true, align: "center" });

      tableRows.push(
        [
          "<w:tr>",
          "<w:tc><w:tcPr>",
          `<w:tcW w:w="${widthCol1}" w:type="dxa"/>`,
          "</w:tcPr>",
          findingCell,
          "</w:tc>",
          "<w:tc><w:tcPr>",
          `<w:tcW w:w="${widthCol2}" w:type="dxa"/>`,
          "</w:tcPr>",
          recommendationCell,
          "</w:tc>",
          "<w:tc><w:tcPr>",
          `<w:tcW w:w="${widthCol3}" w:type="dxa"/>`,
          "</w:tcPr>",
          priorityCell,
          "</w:tc>",
          "</w:tr>",
        ].join(""),
      );
    });

    return [
      "<w:tbl>",
      "<w:tblPr>",
      "<w:tblpPr w:vertAnchor=\"text\" w:horzAnchor=\"text\" w:leftFromText=\"141\" w:rightFromText=\"141\" w:tblpX=\"105\" w:tblpY=\"1\"/>",
      "<w:tblOverlap w:val=\"never\"/>",
      "<w:tblW w:w=\"5000\" w:type=\"pct\"/>",
      "<w:jc w:val=\"start\"/>",
      "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>",
      "<w:tblLayout w:type=\"fixed\"/>",
      "<w:tblCellMar>",
      "<w:top w:w=\"0\" w:type=\"dxa\"/>",
      "<w:start w:w=\"108\" w:type=\"dxa\"/>",
      "<w:bottom w:w=\"0\" w:type=\"dxa\"/>",
      "<w:end w:w=\"108\" w:type=\"dxa\"/>",
      "</w:tblCellMar>",
      "<w:tblBorders>",
      "<w:top w:val=\"nil\"/>",
      "<w:left w:val=\"nil\"/>",
      "<w:bottom w:val=\"nil\"/>",
      "<w:right w:val=\"nil\"/>",
      "<w:insideH w:val=\"nil\"/>",
      "<w:insideV w:val=\"nil\"/>",
      "</w:tblBorders>",
      "<w:tblLook w:val=\"0000\" w:noHBand=\"0\" w:noVBand=\"0\" w:firstColumn=\"0\" w:lastRow=\"0\" w:lastColumn=\"0\" w:firstRow=\"0\"/>",
      "</w:tblPr>",
      "<w:tblGrid>",
      `<w:gridCol w:w="${widthCol1}"/>`,
      `<w:gridCol w:w="${widthCol2}"/>`,
      `<w:gridCol w:w="${widthCol3}"/>`,
      "</w:tblGrid>",
      tableRows.join(""),
      "</w:tbl>",
    ].join("");
  };

  const getNestedDirectory = async (root, parts, options = { create: false }) => {
    let current = root;
    for (const rawPart of parts) {
      const part = String(rawPart || "").trim();
      if (!part) continue;
      try {
        current = await current.getDirectoryHandle(part, options);
      } catch (err) {
        if (options.create) throw err;
        let found = null;
        if (current?.entries) {
          // eslint-disable-next-line no-await-in-loop
          for await (const [name, handle] of current.entries()) {
            if (handle.kind !== "directory") continue;
            if (name.toLowerCase() === part.toLowerCase()) {
              found = handle;
              break;
            }
          }
        }
        if (!found) throw err;
        current = found;
      }
    }
    return current;
  };

  const getOutputsDirectory = async (projectHandle) => {
    try {
      return await getNestedDirectory(projectHandle, ["Outputs"], { create: false });
    } catch (err) {
      try {
        return await getNestedDirectory(projectHandle, ["outputs"], { create: false });
      } catch (err2) {
        return getNestedDirectory(projectHandle, ["Outputs"], { create: true });
      }
    }
  };

  const getFileHandleFromPath = async (projectHandle, path) => {
    const parts = String(path || "").split("/").map((part) => part.trim()).filter(Boolean);
    if (!parts.length) return null;
    let dir = projectHandle;
    for (let i = 0; i < parts.length - 1; i += 1) {
      // eslint-disable-next-line no-await-in-loop
      dir = await getNestedDirectory(dir, [parts[i]], { create: false });
    }
    const fileName = parts[parts.length - 1];
    try {
      return await dir.getFileHandle(fileName);
    } catch (err) {
      if (dir?.entries) {
        for await (const [name, handle] of dir.entries()) {
          if (handle.kind !== "file") continue;
          if (name.toLowerCase() === fileName.toLowerCase()) return handle;
        }
      }
      throw err;
    }
  };

  const tryReadProjectFile = async (projectHandle, relativePath) => {
    try {
      const handle = await getFileHandleFromPath(projectHandle, relativePath);
      if (!handle) return null;
      return await handle.getFile();
    } catch (err) {
      return null;
    }
  };

  const blobToUint8 = async (blob) => new Uint8Array(await blob.arrayBuffer());

  const getImageDimensions = async (blob) => {
    const bitmap = await createImageBitmap(blob);
    const dimensions = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return dimensions;
  };

  const resizeImage = async (blob, maxLongSide) => {
    const bitmap = await createImageBitmap(blob);
    const longSide = Math.max(bitmap.width, bitmap.height);
    const scale = longSide > maxLongSide ? (maxLongSide / longSide) : 1;
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(bitmap, 0, 0, width, height);
    bitmap.close();
    const outputBlob = await new Promise((resolve) => {
      canvas.toBlob((result) => resolve(result), "image/png", 0.92);
    });
    if (!outputBlob) throw new Error("Failed to resize image");
    return outputBlob;
  };

  const writeFileHandle = async (dirHandle, name, data) => {
    const handle = await dirHandle.getFileHandle(name, { create: true });
    const writable = await handle.createWritable();
    await writable.write(data);
    await writable.close();
    return handle;
  };

  const pickLogoFile = async () => {
    if (!window.showOpenFilePicker) throw new Error("File picker unavailable in this browser.");
    const handles = await window.showOpenFilePicker({
      multiple: false,
      excludeAcceptAllOption: false,
      types: [
        {
          description: "Images",
          accept: {
            "image/*": [".png", ".jpg", ".jpeg", ".webp", ".bmp"],
          },
        },
      ],
    });
    if (!handles || !handles.length) throw new Error("No file selected.");
    return handles[0].getFile();
  };

  const prepareLogosForProject = async ({ projectHandle, meta = {} }) => {
    if (!projectHandle) throw new Error("Project folder not selected.");
    const source = await pickLogoFile();
    const largeBlob = await resizeImage(source, 1600);
    const smallBlob = await resizeImage(source, 480);

    const outputs = await getOutputsDirectory(projectHandle);
    await writeFileHandle(outputs, "logo-large.png", largeBlob);
    await writeFileHandle(outputs, "logo-small.png", smallBlob);

    meta.logoLargePath = "Outputs/logo-large.png";
    meta.logoSmallPath = "Outputs/logo-small.png";

    return {
      logoLargePath: meta.logoLargePath,
      logoSmallPath: meta.logoSmallPath,
    };
  };

  const drawSpiderPng = async (spiderData, companyLabel = "Company", project = null) => {
    const stateHelpers = window.AutoBerichtState || {};
    const spiderChart = window.AutoBerichtSpiderChart || {};
    if (typeof spiderChart.drawToBlob !== "function") {
      throw new Error("Spider chart helper unavailable.");
    }
    const formatChapterLabel = typeof stateHelpers.formatChapterLabel === "function"
      ? stateHelpers.formatChapterLabel
      : null;
    const locale = String(project?.meta?.locale || "de-CH");
    const chapters = Array.isArray(project?.chapters) ? project.chapters : [];
    const chapterById = new Map(chapters.map((chapter) => [String(chapter?.id || ""), chapter]));
    const rows = Array.isArray(spiderData?.effective?.chapters_1_11) ? spiderData.effective.chapters_1_11 : [];
    const displayRows = rows.map((row) => {
      const id = String(row?.id || "");
      const chapter = chapterById.get(id);
      const displayLabel = chapter && formatChapterLabel
        ? formatChapterLabel(chapter, locale)
        : String(row?.label || row?.id || "");
      return {
        ...row,
        displayLabel,
      };
    });
    return spiderChart.drawToBlob(displayRows, {
      width: 760,
      height: 500,
      dpr: 2,
      companyLabel,
      suvaLabel: "Suva",
      type: "image/png",
      quality: 0.95,
    });
  };

  const formatDate = (iso) => {
    const raw = String(iso || "");
    if (!raw) {
      const now = new Date();
      return `${String(now.getDate()).padStart(2, "0")}.${String(now.getMonth() + 1).padStart(2, "0")}.${now.getFullYear()}`;
    }
    const y = raw.slice(0, 4);
    const m = raw.slice(5, 7);
    const d = raw.slice(8, 10);
    if (!/^\d{4}$/.test(y) || !/^\d{2}$/.test(m) || !/^\d{2}$/.test(d)) return raw;
    return `${d}.${m}.${y}`;
  };

  const exportReportDocx = async ({
    project,
    projectHandle,
    spiderOverrides = {},
    computeSpider,
    compareIdSegments,
    toText,
  }) => {
    if (!projectHandle) throw new Error("Project folder not selected.");
    if (!window.showOpenFilePicker) throw new Error("File picker unavailable in this browser.");

    const picks = await window.showOpenFilePicker({
      multiple: false,
      excludeAcceptAllOption: false,
      types: [
        {
          description: "Word Template",
          accept: {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
          },
        },
      ],
    });
    if (!picks || !picks.length) throw new Error("No template selected.");
    const templateFile = await picks[0].getFile();
    const entries = await unzipAllEntries(await templateFile.arrayBuffer());
    const map = new Map(entries.map((entry) => [entry.name, entry]));

    const getText = (name) => {
      const entry = map.get(name);
      return entry ? textDecoder.decode(entry.data) : "";
    };

    const setText = (name, xml) => {
      map.set(name, {
        name,
        data: textEncoder.encode(xml),
        flags: map.get(name)?.flags || 0,
      });
    };

    let documentXml = getText("word/document.xml");
    if (!documentXml) throw new Error("Template missing word/document.xml");

    const markerMap = {
      "NAME$$": project?.meta?.projectName || project?.meta?.company || "",
      "COMPANY$$": project?.meta?.company || "",
      "COMPANY_ID$$": project?.meta?.companyId || "",
      "ADDRESS$$": buildAddressLine(project?.meta || {}),
      "AUTHOR$$": project?.meta?.author || project?.meta?.moderator || "",
      "MOD$$": project?.meta?.moderator || "",
      "CO$$": project?.meta?.coModerator || "",
      "DATE$$": formatDate(project?.meta?.createdAt),
    };

    documentXml = replaceTextMarkers(documentXml, markerMap);

    const headerParts = ["word/header1.xml", "word/header2.xml"];
    headerParts.forEach((part) => {
      const xml = getText(part);
      if (!xml) return;
      setText(part, replaceTextMarkers(xml, markerMap));
    });

    const chapters = [...(project?.chapters || [])].sort((a, b) => {
      if (typeof compareIdSegments === "function") return compareIdSegments(a?.id, b?.id);
      return String(a?.id || "").localeCompare(String(b?.id || ""), "de", { numeric: true });
    });
    const chapter0ListNumId = findUpperLetterNumId(getText("word/numbering.xml"));

    chapters.forEach((chapter) => {
      const marker = `CHAPTER${chapter.id}$$`;
      const replacement = String(chapter.id) === "0"
        ? buildChapter0Xml(chapter, toText, { numId: chapter0ListNumId })
        : buildChapterTableXml(chapter, toText);
      const patched = replaceParagraphMarker(documentXml, marker, replacement);
      documentXml = patched.xml;
    });

    let contentTypes = getText("[Content_Types].xml");

    const relPartFor = (xmlPart) => {
      if (xmlPart === "word/document.xml") return "word/_rels/document.xml.rels";
      if (xmlPart === "word/header1.xml") return "word/_rels/header1.xml.rels";
      if (xmlPart === "word/header2.xml") return "word/_rels/header2.xml.rels";
      return "";
    };

    const insertImageAtMarker = async ({
      xmlPart,
      marker,
      imageFile,
      mediaName,
      cmHeight,
      align = "",
    }) => {
      if (!imageFile) return false;
      const xml = xmlPart === "word/document.xml" ? documentXml : getText(xmlPart);
      if (!xml || !hasMarker(xml, marker)) return false;

      const relPart = relPartFor(xmlPart);
      let relsXml = getText(relPart);
      if (!relsXml) return false;

      const imageBytes = await blobToUint8(imageFile);
      map.set(`word/media/${mediaName}`, {
        name: `word/media/${mediaName}`,
        data: imageBytes,
        flags: 0,
      });

      const dimensions = await getImageDimensions(imageFile);
      const targetHeight = emuFromCm(cmHeight);
      const targetWidth = Math.round((targetHeight * dimensions.width) / Math.max(1, dimensions.height));
      const relId = getNextRelId(relsXml);
      relsXml = appendRelationship(relsXml, relId, `media/${mediaName}`);
      setText(relPart, relsXml);

      const drawing = drawingXml(relId, mediaName, targetWidth, targetHeight, align);
      const patched = replaceParagraphMarker(xml, marker, drawing);
      if (!patched.replaced) return false;
      if (xmlPart === "word/document.xml") {
        documentXml = patched.xml;
      } else {
        setText(xmlPart, patched.xml);
      }
      return true;
    };

  const insertChapterThermos = (spiderData) => {
      const scoreMap = buildThermoScoreMap(spiderData);
      if (!scoreMap.size) return;
      const chapterIds = new Set();
      [...(project?.chapters || [])]
        .map((chapter) => String(chapter?.id || "").trim())
        .forEach((id) => {
          if (!/^\d+$/.test(id)) return;
          if (id === "0") return;
          chapterIds.add(id);
        });
      scoreMap.forEach((_, id) => {
        const chapterId = String(id || "").trim();
        if (!/^\d+$/.test(chapterId)) return;
        if (chapterId === "0") return;
        chapterIds.add(chapterId);
      });
      const locale = String(project?.meta?.locale || "de-CH");
      const company = String(project?.meta?.company || "").trim() || "Company";

      chapterIds.forEach((chapterId) => {
        const score = resolveThermoScore(scoreMap, chapterId);
        const thermoXml = buildThermoTableXml({
          companyValue: score.company,
          consultantValue: score.consultant,
          locale,
          companyName: company,
        });
        thermoMarkersForChapter(chapterId).forEach((marker) => {
          if (!hasMarker(documentXml, marker)) return;
          const patched = replaceAllParagraphMarkers(documentXml, marker, thermoXml);
          if (patched.count > 0) {
            documentXml = patched.xml;
          }
        });
      });
    };

    const logoLargeFile = await tryReadProjectFile(projectHandle, project?.meta?.logoLargePath || "Outputs/logo-large.png");
    const logoSmallFile = await tryReadProjectFile(projectHandle, project?.meta?.logoSmallPath || "Outputs/logo-small.png");

    if (logoLargeFile) {
      await insertImageAtMarker({
        xmlPart: "word/document.xml",
        marker: "LOGO_BIG$$",
        imageFile: logoLargeFile,
        mediaName: "autobericht_logo_large.png",
        cmHeight: 3,
        align: "center",
      });
    }

    if (logoSmallFile) {
      await insertImageAtMarker({
        xmlPart: "word/header1.xml",
        marker: "LOGO_SMALL$$",
        imageFile: logoSmallFile,
        mediaName: "autobericht_logo_small.png",
        cmHeight: 0.8,
      });
      await insertImageAtMarker({
        xmlPart: "word/header2.xml",
        marker: "LOGO_SMALL$$",
        imageFile: logoSmallFile,
        mediaName: "autobericht_logo_small.png",
        cmHeight: 0.8,
      });
    }

    if (typeof computeSpider === "function") {
      try {
        const spiderData = await computeSpider({
          project,
          overrides: spiderOverrides || {},
          dirHandle: projectHandle,
        });
        try {
          const spiderBlob = await drawSpiderPng(
            spiderData,
            String(project?.meta?.company || "").trim() || "Company",
            project,
          );
          await insertImageAtMarker({
            xmlPart: "word/document.xml",
            marker: "SPIDER$$",
            imageFile: spiderBlob,
            mediaName: "autobericht_spider.png",
            cmHeight: 10.0,
            align: "center",
          });
        } catch (err) {
          // Keep export running even if spider image generation fails.
        }
        insertChapterThermos(spiderData);
      } catch (err) {
        // Keep export running even if spider computation fails.
      }
    }

    setText("word/document.xml", documentXml);
    const settingsXml = getText("word/settings.xml");
    if (settingsXml) {
      setText("word/settings.xml", ensureUpdateFieldsOnOpen(settingsXml));
    }
    if (contentTypes) {
      contentTypes = ensurePngContentType(contentTypes);
      setText("[Content_Types].xml", contentTypes);
    }

    const outputBytes = buildZipStore(Array.from(map.values()));
    const outputs = await getOutputsDirectory(projectHandle);
    const now = new Date();
    const stamp = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
    const outputName = `${stamp}_AutoBericht_NoVBA.docx`;
    await writeFileHandle(outputs, outputName, outputBytes);

    return {
      savedAs: `Outputs/${outputName}`,
    };
  };

  window.AutoBerichtWordExport = {
    prepareLogosForProject,
    exportReportDocx,
  };
})();
