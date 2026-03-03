/*
 * No-VBA PowerPoint export orchestrator.
 *
 * Responsibilities:
 * - Read a user-selected `.pptx` template and validate required layouts/placeholders.
 * - Build report and training slide plans from sidecar/project data.
 * - Render chapter "page snapshot" images (header/footer/logos/thermo/table) for report chapters.
 * - Write final output as .pptx into outputs/.
 */
(() => {
  const textDecoder = new TextDecoder();
  const textEncoder = new TextEncoder();
  const zipTools = window.AutoBerichtWordDocxZip;
  const reportRows = window.AutoBerichtReportRows || {};

  const unzipAllEntries = zipTools?.unzipAllEntries;
  const buildZipStore = zipTools?.buildZipStore;

  const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";

  const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
  const REL_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
  const REL_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
  const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

  const REPORT_LAYOUTS = {
    cover: "ab_title",
    chapterSeparator: "ab_chapterorange",
    chapterSnapshot: "ab_titleandpicture",
    sectionText: "ab_textandpicture",
    sectionPhotoLow: "ab_4pictures",
    sectionPhotoHigh: "ab_6pictures",
    observationSeparator: "ab_chapterorange",
    observationTextPhotoLow: "ab_textandpicture",
    observationTextPhotoHigh: "ab_6pictures",
    summaryText: "ab_titleandtext",
  };

  const REPORT_OPTIONAL_LAYOUTS = {
    sectionPhotoTwo: "ab_2pictures",
    sectionPhotoThree: "ab_3pictures",
  };

  const TRAINING_TAG_ORDER = [
    "unterlassen",
    "dulden",
    "handeln",
    "vorbild",
    "iceberg",
    "pyramide",
    "stop",
    "sos",
    "verhindern",
    "audit",
    "risikobeurteilung",
    "aviva",
  ];

  const TRAINING_LAYOUT_BY_TAG_BASE = {
    unterlassen: "ab_unterlassen",
    dulden: "ab_dulden",
    handeln: "ab_handeln",
    vorbild: "ab_vorbild",
    verhindern: "ab_verhindern",
    audit: "ab_audit",
    risikobeurteilung: "ab_risikobeurteilung",
    aviva: "ab_aviva",
  };

  const TRAINING_LAYOUTS = {
    introBySuffix: {
      d: "ab_title",
      f: "ab_title",
      i: "ab_title",
    },
    sectionSeparator: "ab_chapterorange",
    defaultPhoto: "ab_picture",
  };

  const SNAPSHOT = {
    // Compact thermo snapshot canvas for chapter intro slides.
    width: 1400,
    height: 560,
    margin: 52,
    headerH: 120,
    thermoH: 110,
  };

  const xmlEscape = (value) => String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");

  const normalizePartName = (name) => String(name || "")
    .replace(/\\/g, "/")
    .replace(/^\/+/, "")
    .replace(/^\.\//, "");

  const relsPartNameForPart = (partName) => {
    const normalized = normalizePartName(partName);
    const parts = normalized.split("/").filter(Boolean);
    const file = parts.pop() || "";
    const dir = parts.join("/");
    return dir ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
  };

  const resolveRelativePartName = (sourcePartName, target) => {
    const rawTarget = String(target || "").trim();
    if (!rawTarget) return "";
    if (rawTarget.startsWith("/")) return normalizePartName(rawTarget);
    const sourceParts = normalizePartName(sourcePartName).split("/").filter(Boolean);
    sourceParts.pop();
    rawTarget.replace(/\\/g, "/").split("/").forEach((segment) => {
      const token = String(segment || "").trim();
      if (!token || token === ".") return;
      if (token === "..") {
        sourceParts.pop();
        return;
      }
      sourceParts.push(token);
    });
    return sourceParts.join("/");
  };

  const ensureZipHelpers = () => {
    if (typeof unzipAllEntries !== "function" || typeof buildZipStore !== "function") {
      throw new Error("ZIP helpers are not available (AutoBerichtWordDocxZip).");
    }
  };

  const parseXml = (xmlText, partName) => {
    const doc = new DOMParser().parseFromString(String(xmlText || ""), "application/xml");
    const parseError = doc.getElementsByTagName("parsererror")[0];
    if (parseError) {
      throw new Error(`XML parse failed for ${partName}`);
    }
    return doc;
  };

  const serializeXml = (doc) => {
    const body = new XMLSerializer().serializeToString(doc);
    // Browser XMLSerializer may already include a declaration for some docs.
    // Emit exactly one declaration to keep package XML valid.
    if (/^\s*<\?xml\b/i.test(body)) return body;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
  };

  const byLocalName = (node, localName) => {
    if (!node) return [];
    const out = [];
    const walker = (current) => {
      if (!current) return;
      if (current.nodeType === 1 && current.localName === localName) out.push(current);
      const children = current.childNodes || [];
      for (let i = 0; i < children.length; i += 1) walker(children[i]);
    };
    walker(node);
    return out;
  };

  const firstByLocalName = (node, localName) => byLocalName(node, localName)[0] || null;

  const getAttr = (el, name) => {
    if (!el) return "";
    return String(el.getAttribute(name) || el.getAttribute(`a:${name}`) || el.getAttribute(`p:${name}`) || "");
  };

  const toInt = (value, fallback = 0) => {
    const n = Number(String(value || "").trim());
    return Number.isFinite(n) ? Math.round(n) : fallback;
  };

  const formatDateIso = (date = new Date()) => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  };

  const formatDateLabel = (isoLike) => {
    const raw = String(isoLike || "");
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

  const toFileSafeSlug = (value, fallback = "Company") => {
    const raw = String(value || "").trim();
    const hyphenated = raw.replace(/\s+/g, "-");
    const cleaned = hyphenated
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
      .replace(/-+/g, "-")
      .replace(/^[.\-\s]+|[.\-\s]+$/g, "");
    return cleaned || fallback;
  };

  const rowToText = (value, toText) => {
    if (typeof reportRows.rowToText === "function") return reportRows.rowToText(value, toText);
    if (typeof toText === "function") return toText(value);
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  };

  const localeBase = (locale) => String(locale || "").trim().toLowerCase().split("-")[0] || "de";

  const resolveLocalizedText = (value, locale = "de-CH") => {
    if (value == null) return "";
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
      return String(value);
    }
    if (Array.isArray(value)) {
      return value.map((item) => resolveLocalizedText(item, locale)).filter(Boolean).join("\n");
    }
    if (typeof value === "object") {
      const exact = String(locale || "").trim();
      const base = localeBase(locale);
      const candidates = [
        exact, exact.toLowerCase(), exact.toUpperCase(),
        base, base.toLowerCase(), base.toUpperCase(),
        "de-CH", "de", "fr-CH", "fr", "it-CH", "it", "en-GB", "en-US", "en",
        "title", "label", "name", "text", "value",
      ];
      for (let i = 0; i < candidates.length; i += 1) {
        const key = candidates[i];
        if (!key || !Object.prototype.hasOwnProperty.call(value, key)) continue;
        const text = resolveLocalizedText(value[key], locale);
        if (text && text !== "[object Object]") return text;
      }
      const keys = Object.keys(value);
      for (let i = 0; i < keys.length; i += 1) {
        const text = resolveLocalizedText(value[keys[i]], locale);
        if (text && text !== "[object Object]") return text;
      }
      return "";
    }
    return String(value);
  };

  const stripLeadingNumber = (value) => {
    if (typeof reportRows.stripLeadingNumber === "function") return reportRows.stripLeadingNumber(value);
    return String(value || "").replace(/^\s*\d+(?:\.\d+)*(?:\s|[.:-]\s*)?/, "").trim();
  };

  const isSectionRow = (row) => {
    if (typeof reportRows.isSectionRow === "function") return reportRows.isSectionRow(row);
    return String(row?.kind || "").toLowerCase() === "section";
  };

  const isIncludedRow = (row) => {
    if (typeof reportRows.isReportReadyRow === "function") return reportRows.isReportReadyRow(row);
    const ws = row?.workstate;
    if (!ws || ws.includeFinding == null) return false;
    return ws.includeFinding === true && ws.done === true;
  };

  const resolveSectionId = (row, chapterId) => {
    if (typeof reportRows.resolveSectionId === "function") return reportRows.resolveSectionId(row, chapterId);
    const sectionId = String(row?.sectionId || "").trim();
    if (sectionId) return sectionId;
    const parts = String(row?.id || "").split(".");
    if (parts.length >= 2) return `${parts[0]}.${parts[1]}`;
    return `${chapterId}.1`;
  };

  const resolveSectionTitle = (row, locale = "de-CH") => {
    if (typeof reportRows.resolveSectionTitle === "function") {
      const fromHelper = reportRows.resolveSectionTitle(row);
      const helperText = resolveLocalizedText(fromHelper, locale);
      if (helperText) return stripLeadingNumber(helperText) || helperText;
    }
    const rawTitle = resolveLocalizedText(row?.title, locale) || String(row?.id || "");
    const cleaned = stripLeadingNumber(rawTitle);
    return cleaned || rawTitle;
  };

  const resolveFindingText = (row, toText) => {
    if (typeof reportRows.resolveFindingText === "function") return reportRows.resolveFindingText(row, toText);
    const ws = row?.workstate;
    if (ws && Object.prototype.hasOwnProperty.call(ws, "findingText")) {
      return rowToText(ws.findingText, toText);
    }
    return rowToText(row?.master?.finding, toText);
  };

  const resolveRecommendationText = (row, toText) => {
    if (typeof reportRows.resolveRecommendationText === "function") return reportRows.resolveRecommendationText(row, toText);
    const ws = row?.workstate || {};
    if (ws.includeRecommendation === false) return "";
    if (Object.prototype.hasOwnProperty.call(ws, "recommendationText")) {
      return rowToText(ws.recommendationText, toText);
    }
    return rowToText(row?.master?.recommendation, toText);
  };

  const resolvePriorityText = (row) => {
    if (typeof reportRows.resolvePriorityText === "function") return reportRows.resolvePriorityText(row);
    const raw = Number(row?.workstate?.priority);
    if (!Number.isFinite(raw)) return "";
    const value = Math.round(raw);
    if (value < 1 || value > 4) return "";
    return String(value);
  };

  const isFieldObservationChapter = (chapterId) => {
    if (typeof reportRows.isFieldObservationChapter === "function") return reportRows.isFieldObservationChapter(chapterId);
    return String(chapterId || "").includes(".");
  };

  const resolveSectionDisplayId = (sectionId, chapterId, sectionMap) => {
    if (typeof reportRows.resolveSectionDisplayId === "function") {
      return reportRows.resolveSectionDisplayId(sectionId, chapterId, sectionMap);
    }
    const key = String(sectionId || "").trim();
    if (!key) return "";
    if (sectionMap && sectionMap.has(key)) return `${chapterId}.${sectionMap.get(key)}`;
    return key;
  };

  const buildRenumberMap = (rows, chapterId) => {
    if (typeof reportRows.buildRenumberMap === "function") {
      return reportRows.buildRenumberMap(rows, chapterId, isIncludedRow);
    }
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

  const orderRowsForChapter = (chapter) => {
    if (typeof reportRows.orderRowsForChapter === "function") return reportRows.orderRowsForChapter(chapter);
    return Array.isArray(chapter?.rows) ? [...chapter.rows] : [];
  };

  const getNestedDirectory = async (root, parts, options = { create: false }) => {
    let current = root;
    for (const rawPart of parts) {
      const part = String(rawPart || "").trim();
      if (!part) continue;
      try {
        // eslint-disable-next-line no-await-in-loop
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
      return await getNestedDirectory(projectHandle, ["outputs"], { create: false });
    } catch (err) {
      return getNestedDirectory(projectHandle, ["outputs"], { create: true });
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

  const writeFileHandle = async (dirHandle, name, data) => {
    const handle = await dirHandle.getFileHandle(name, { create: true });
    const writable = await handle.createWritable();
    await writable.write(data);
    await writable.close();
    return handle;
  };

  const pickPptTemplate = async () => {
    if (!window.showOpenFilePicker) {
      throw new Error("File picker unavailable in this browser.");
    }
    const picks = await window.showOpenFilePicker({
      multiple: false,
      excludeAcceptAllOption: false,
      types: [
        {
          description: "PowerPoint Template",
          accept: {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"],
          },
        },
      ],
    });
    if (!picks?.length) throw new Error("No template selected.");
    return picks[0].getFile();
  };

  const blobToUint8 = async (blob) => new Uint8Array(await blob.arrayBuffer());

  const imageDimensions = async (blob) => {
    const bmp = await createImageBitmap(blob);
    const out = { width: bmp.width, height: bmp.height };
    bmp.close();
    return out;
  };

  const maybeLoadLogo = async (projectHandle, path) => {
    if (!path) return null;
    const file = await tryReadProjectFile(projectHandle, path);
    if (!file) return null;
    const dims = await imageDimensions(file);
    return { file, width: dims.width, height: dims.height };
  };

  const buildPhotoFileMap = (sidecarDoc) => {
    const reportMap = new Map();
    const trainingMap = new Map();
    const photos = sidecarDoc?.photos?.photos || {};

    const addToMap = (map, tag, relPath) => {
      const key = String(tag || "").trim();
      if (!key) return;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(relPath);
    };

    Object.entries(photos).forEach(([relPath, photo]) => {
      const tags = photo?.tags || {};
      (tags.report || []).forEach((tag) => addToMap(reportMap, tag, relPath));
      (tags.observations || []).forEach((tag) => addToMap(reportMap, tag, relPath));
      (tags.training || []).forEach((tag) => addToMap(trainingMap, tag, relPath));
    });

    const dedupe = (map) => {
      map.forEach((arr, key) => {
        const seen = new Set();
        const out = [];
        arr.forEach((item) => {
          const token = String(item || "");
          if (!token || seen.has(token)) return;
          seen.add(token);
          out.push(token);
        });
        map.set(key, out);
      });
    };

    dedupe(reportMap);
    dedupe(trainingMap);

    return { reportMap, trainingMap };
  };

  const normalizeTag = (value) => String(value || "").trim().toLowerCase();

  const compareAlphaNumeric = (a, b) => String(a || "").localeCompare(String(b || ""), "de", { numeric: true });

  const ensureMapArray = (map, key) => map.get(String(key || "").trim()) || [];

  const resolveObservationTag = (row, locale = "de-CH") => {
    const first = resolveLocalizedText(row?.tag, locale).trim();
    if (first) return first;
    const second = resolveLocalizedText(row?.titleOverride, locale).trim();
    if (second) return second;
    const third = resolveLocalizedText(row?.title, locale).trim();
    if (third) return third;
    return String(row?.id || "").trim();
  };

  const resolveObservationTitle = (row, fallback, locale = "de-CH") => {
    const first = resolveLocalizedText(row?.titleOverride, locale).trim();
    if (first) return first;
    const second = resolveLocalizedText(row?.title, locale).trim();
    if (second) return second;
    return resolveLocalizedText(fallback, locale).trim() || String(fallback || "").trim();
  };

  const resolveSpecial48DisplaySectionId = (chapters) => {
    const chapter4 = (chapters || []).find((chapter) => String(chapter?.id || "") === "4");
    if (!chapter4) return "4.8";
    const rows = orderRowsForChapter(chapter4);
    const { sectionMap } = buildRenumberMap(rows, "4");
    return `4.${sectionMap?.size ? sectionMap.size + 1 : 1}`;
  };

  const orderObservationRows = (chapter) => {
    const rows = (chapter?.rows || []).filter((row) => !isSectionRow(row));
    const order = Array.isArray(chapter?.meta?.order) ? chapter.meta.order.map((x) => String(x || "")) : [];
    if (!order.length) return rows;
    const byId = new Map(rows.map((row) => [String(row?.id || ""), row]));
    const out = [];
    const added = new Set();
    order.forEach((id) => {
      const match = byId.get(id);
      if (!match) return;
      out.push(match);
      added.add(id);
    });
    rows.forEach((row) => {
      const id = String(row?.id || "");
      if (!added.has(id)) out.push(row);
    });
    return out;
  };

  const buildSectionBlocks = (chapter, chapterId, locale = "de-CH") => {
    const rows = orderRowsForChapter(chapter);
    const { rowMap, sectionMap } = buildRenumberMap(rows, chapterId);
    const sections = [];
    let current = null;

    rows.forEach((row) => {
      if (isSectionRow(row)) {
        const sectionId = String(row?.id || "").trim();
        current = {
          rawId: sectionId,
          displayId: resolveSectionDisplayId(sectionId, chapterId, sectionMap),
          title: resolveSectionTitle(row, locale),
          rows: [],
        };
        sections.push(current);
        return;
      }
      if (!isIncludedRow(row)) return;
      if (!current) {
        const sid = resolveSectionId(row, chapterId);
        current = {
          rawId: sid,
          displayId: resolveSectionDisplayId(sid, chapterId, sectionMap),
          title: resolveLocalizedText(row?.sectionLabel, locale).trim() || String(sid || "").trim(),
          rows: [],
        };
        sections.push(current);
      }
      const rowId = String(row?.id || "").trim();
      current.rows.push({
        row,
        displayId: rowMap.get(rowId) || rowId,
      });
    });

    return sections.filter((section) => section.rows.length > 0);
  };

  const buildSummaryRecommendationLines = (chapter, toText) => {
    const rows = orderRowsForChapter(chapter);
    const lines = [];
    rows.forEach((row) => {
      if (isSectionRow(row) || !isIncludedRow(row)) return;
      const text = String(resolveRecommendationText(row, toText) || "").trim();
      if (!text) return;
      lines.push(text.replace(/\s+/g, " ").trim());
    });
    const alphaLabel = (index) => {
      let n = Math.max(1, Number(index) || 1);
      let out = "";
      while (n > 0) {
        const rem = (n - 1) % 26;
        out = String.fromCharCode(65 + rem) + out;
        n = Math.floor((n - 1) / 26);
      }
      return out;
    };
    return lines.map((line, index) => `${alphaLabel(index + 1)}. ${line}`);
  };

  const buildSectionFindingLines = (section, toText) => {
    const lines = [];
    section.rows.forEach(({ row, displayId }) => {
      const finding = String(resolveFindingText(row, toText) || "").trim();
      if (!finding) return;
      const head = stripLeadingNumber(finding.replace(/\s+/g, " ").trim()) || finding;
      lines.push(`${displayId} ${head}`.trim());
    });
    return lines;
  };

  const chunkArray = (arr, size) => {
    const out = [];
    if (!Array.isArray(arr) || size < 1) return out;
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  };

  const thermoRows = (locale, companyName) => {
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

  const snapshotLabels = (locale) => {
    const normalized = String(locale || "de-CH").toLowerCase();
    if (normalized.startsWith("fr")) {
      return {
        date: "Date",
        moderator: "Moderateur",
        potential: "Points systeme avec potentiel d'amelioration",
        finding: "Etat actuel",
        recommendation: "Pistes de solution",
        priority: "Prio",
        reportType: "Etat des lieux",
      };
    }
    if (normalized.startsWith("it")) {
      return {
        date: "Data",
        moderator: "Moderatore",
        potential: "Punti di sistema con potenziale di miglioramento",
        finding: "Stato attuale",
        recommendation: "Possibili soluzioni",
        priority: "Prio",
        reportType: "Rilevazione stato attuale",
      };
    }
    return {
      date: "Datum",
      moderator: "Moderator",
      potential: "Systempunkte mit Verbesserungspotenzial",
      finding: "Ist-Zustand",
      recommendation: "Loesungsansaetze",
      priority: "Prio",
      reportType: "Ist-Aufnahme",
    };
  };

  const coverReportTitle = (locale) => {
    const normalized = String(locale || "de-CH").toLowerCase();
    if (normalized.startsWith("fr")) return "Etat des lieux: Presentation du rapport";
    if (normalized.startsWith("it")) return "Rilevazione stato attuale: Presentazione del rapporto";
    return "Ist-Aufnahme: Bericht Besprechung";
  };

  const assessmentSummaryTitle = (locale) => {
    const normalized = String(locale || "de-CH").toLowerCase();
    if (normalized.startsWith("fr")) return "Synthese des evaluations";
    if (normalized.startsWith("it")) return "Sintesi delle valutazioni";
    return "Zusammenfassung der Beurteilungen";
  };

  const roundToNearestTen = (value) => {
    const raw = Number(value);
    if (!Number.isFinite(raw)) return 0;
    return Math.max(0, Math.min(100, Math.round(raw / 10) * 10));
  };

  const buildSpiderScoreMap = (spiderData) => {
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

  const drawThermoBars = (ctx, x, y, width, score, locale, companyName) => {
    const rows = thermoRows(locale, companyName);
    const labelsW = Math.round(width * 0.42);
    const barW = width - labelsW;
    const segCount = 10;
    const segGap = 4;
    const segW = Math.floor((barW - (segCount - 1) * segGap) / segCount);
    const segH = 18;
    const startX = x + labelsW;
    const companyFilled = Math.max(0, Math.min(segCount, Math.round(roundToNearestTen(score.company) / 10)));
    const consultantFilled = Math.max(0, Math.min(segCount, Math.round(roundToNearestTen(score.consultant) / 10)));

    ctx.save();
    ctx.fillStyle = "#23303d";
    ctx.font = "600 19px Arial";
    ctx.fillText(rows.companyLabel, x, y + 26);
    ctx.fillText(rows.consultantLabel, x, y + 58);

    const drawRow = (rowY, filled, fill) => {
      for (let i = 0; i < segCount; i += 1) {
        const segX = startX + i * (segW + segGap);
        ctx.fillStyle = i < filled ? fill : "#ffffff";
        ctx.strokeStyle = "#667788";
        ctx.lineWidth = 1;
        ctx.fillRect(segX, rowY, segW, segH);
        ctx.strokeRect(segX, rowY, segW, segH);
      }
    };

    drawRow(y + 10, companyFilled, "#bdd7ee");
    drawRow(y + 42, consultantFilled, "#f8cbad");

    ctx.fillStyle = "#6f7b86";
    ctx.font = "600 14px Arial";
    ctx.fillText("-", startX - 14, y + 28);
    ctx.fillText("+", startX + barW + 4, y + 28);
    ctx.restore();
  };

  const drawChapterSnapshot = async ({
    chapterLabel,
    score,
    locale,
    company,
    moderator,
    dateLabel,
    logoSmall,
    logoLarge,
  }) => {
    const canvas = document.createElement("canvas");
    canvas.width = SNAPSHOT.width;
    canvas.height = SNAPSHOT.height;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas context unavailable for chapter snapshot.");
    const labels = snapshotLabels(locale);

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, SNAPSHOT.width, SNAPSHOT.height);

    const headY = 0;
    const headH = SNAPSHOT.headerH;
    ctx.fillStyle = "#f8f5ef";
    ctx.fillRect(0, headY, SNAPSHOT.width, headH);
    ctx.strokeStyle = "#d7d0c7";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, headH + 0.5);
    ctx.lineTo(SNAPSHOT.width, headH + 0.5);
    ctx.stroke();

    if (logoSmall?.file) {
      const bmp = await createImageBitmap(logoSmall.file);
      const targetH = 56;
      const targetW = Math.round((targetH * bmp.width) / Math.max(1, bmp.height));
      ctx.drawImage(bmp, SNAPSHOT.margin, 24, targetW, targetH);
      bmp.close();
    } else if (logoLarge?.file) {
      const bmp = await createImageBitmap(logoLarge.file);
      const targetH = 56;
      const targetW = Math.round((targetH * bmp.width) / Math.max(1, bmp.height));
      ctx.drawImage(bmp, SNAPSHOT.margin, 24, targetW, targetH);
      bmp.close();
    }

    ctx.fillStyle = "#1f2933";
    ctx.font = "700 31px Arial";
    ctx.fillText(chapterLabel, SNAPSHOT.margin + 180, 50);
    ctx.font = "500 18px Arial";
    ctx.fillStyle = "#4b5967";
    ctx.fillText(String(company || "").trim() || "Company", SNAPSHOT.margin + 180, 82);

    ctx.textAlign = "right";
    ctx.fillStyle = "#4b5967";
    ctx.font = "600 16px Arial";
    ctx.fillText(`${labels.date}: ${String(dateLabel || "").trim()}`, SNAPSHOT.width - SNAPSHOT.margin, 45);
    ctx.fillText(`${labels.moderator}: ${String(moderator || "").trim() || "-"}`, SNAPSHOT.width - SNAPSHOT.margin, 73);
    ctx.textAlign = "left";

    const thermoX = SNAPSHOT.margin;
    const thermoY = SNAPSHOT.headerH + 86;
    const thermoW = SNAPSHOT.width - SNAPSHOT.margin * 2;
    drawThermoBars(ctx, thermoX, thermoY, thermoW, score, locale, company);
    const rulerY = Math.min(SNAPSHOT.height - 24, thermoY + SNAPSHOT.thermoH + 34);
    ctx.strokeStyle = "#d7d0c7";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(SNAPSHOT.margin, rulerY + 0.5);
    ctx.lineTo(SNAPSHOT.width - SNAPSHOT.margin, rulerY + 0.5);
    ctx.stroke();

    const blob = await new Promise((resolve) => {
      canvas.toBlob((result) => resolve(result), "image/png", 0.94);
    });
    if (!blob) throw new Error("Failed to render chapter snapshot image.");
    return blob;
  };

  const getTemplateMap = async (templateFile) => {
    ensureZipHelpers();
    const entries = await unzipAllEntries(await templateFile.arrayBuffer());
    return new Map(entries.map((entry) => [entry.name, entry]));
  };

  const getEntryText = (map, name) => {
    const entry = map.get(name);
    return entry ? textDecoder.decode(entry.data) : "";
  };

  const setEntryText = (map, name, xml) => {
    const previous = map.get(name);
    map.set(name, {
      name,
      data: textEncoder.encode(String(xml || "")),
      flags: previous?.flags || 0,
    });
  };

  const setEntryBytes = (map, name, bytes, flags = 0) => {
    map.set(name, {
      name,
      data: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes),
      flags,
    });
  };

  const extractPlaceholdersFromDoc = (doc) => {
    const cSld = firstByLocalName(doc, "cSld");
    const spTree = firstByLocalName(cSld, "spTree");
    const children = spTree ? Array.from(spTree.childNodes || []).filter((node) => node.nodeType === 1) : [];
    const placeholders = [];
    children.forEach((shapeEl) => {
      const ph = firstByLocalName(shapeEl, "ph");
      if (!ph) return;
      let xfrm = null;
      if (shapeEl.localName === "graphicFrame") {
        xfrm = firstByLocalName(shapeEl, "xfrm");
      } else {
        const spPr = firstByLocalName(shapeEl, "spPr");
        xfrm = firstByLocalName(spPr, "xfrm");
      }
      const off = firstByLocalName(xfrm, "off");
      const ext = firstByLocalName(xfrm, "ext");
      const typeRaw = String(getAttr(ph, "type") || "").trim();
      const type = typeRaw || "body";
      const idxRaw = String(getAttr(ph, "idx") || "").trim();
      const idxKey = /^\d+$/.test(idxRaw) ? idxRaw : "";
      placeholders.push({
        type,
        idxKey,
        idxSort: idxKey ? Number(idxKey) : Number.MAX_SAFE_INTEGER,
        bounds: (off && ext)
          ? {
            x: toInt(getAttr(off, "x"), 0),
            y: toInt(getAttr(off, "y"), 0),
            cx: Math.max(1, toInt(getAttr(ext, "cx"), 1)),
            cy: Math.max(1, toInt(getAttr(ext, "cy"), 1)),
          }
          : null,
      });
    });
    placeholders.sort((a, b) => (a.idxSort - b.idxSort) || compareAlphaNumeric(a.type, b.type));
    return placeholders;
  };

  const getLayoutInfos = (templateMap) => {
    const keyByNormalized = new Map();
    Array.from(templateMap.keys()).forEach((key) => {
      keyByNormalized.set(normalizePartName(key).toLowerCase(), key);
    });

    const getEntryTextByNormalized = (normalizedPartName) => {
      const actual = keyByNormalized.get(normalizePartName(normalizedPartName).toLowerCase());
      if (!actual) return "";
      return getEntryText(templateMap, actual);
    };

    const masterPlaceholderCache = new Map();
    const resolveMasterPlaceholderMap = (layoutPartName) => {
      const relsPart = relsPartNameForPart(layoutPartName);
      const relsXml = getEntryTextByNormalized(relsPart);
      if (!relsXml) return null;
      const relsDoc = parseXml(relsXml, relsPart);
      const relNodes = Array.from(relsDoc.getElementsByTagNameNS(NS_REL, "Relationship"));
      const masterRel = relNodes.find((rel) => String(rel.getAttribute("Type") || "").trim() === REL_SLIDE_MASTER);
      if (!masterRel) return null;
      const masterPart = resolveRelativePartName(layoutPartName, masterRel.getAttribute("Target") || "");
      if (!masterPart) return null;
      if (masterPlaceholderCache.has(masterPart)) return masterPlaceholderCache.get(masterPart);
      const masterXml = getEntryTextByNormalized(masterPart);
      if (!masterXml) {
        masterPlaceholderCache.set(masterPart, null);
        return null;
      }
      const masterDoc = parseXml(masterXml, masterPart);
      const masterMap = new Map();
      extractPlaceholdersFromDoc(masterDoc).forEach((ph) => {
        if (!ph.bounds) return;
        const key = `${String(ph.type || "").toLowerCase()}|${ph.idxKey}`;
        if (!masterMap.has(key)) {
          masterMap.set(key, ph.bounds);
        }
      });
      masterPlaceholderCache.set(masterPart, masterMap);
      return masterMap;
    };

    const infos = new Map();
    Array.from(templateMap.keys())
      .filter((name) => /^ppt\/slidelayouts\/slidelayout\d+\.xml$/i.test(normalizePartName(name)))
      .forEach((partName) => {
        const xml = getEntryText(templateMap, partName);
        const doc = parseXml(xml, partName);
        const sldLayout = firstByLocalName(doc, "sldLayout");
        const cSld = firstByLocalName(doc, "cSld");
        // Different PPTX producers persist the user-visible layout name
        // either on p:cSld@name or on p:sldLayout@matchingName / @name.
        const layoutName = String(
          getAttr(cSld, "name")
          || getAttr(sldLayout, "matchingName")
          || getAttr(sldLayout, "name")
          || "",
        ).trim();
        if (!layoutName) return;

        const placeholders = extractPlaceholdersFromDoc(doc);
        const masterMap = resolveMasterPlaceholderMap(partName);
        if (masterMap && masterMap.size) {
          placeholders.forEach((ph) => {
            if (ph.bounds) return;
            const key = `${String(ph.type || "").toLowerCase()}|${ph.idxKey}`;
            const inherited = masterMap.get(key);
            if (inherited) ph.bounds = { ...inherited };
          });
        }

        infos.set(layoutName.toLowerCase(), {
          name: layoutName,
          partName,
          placeholders,
        });
      });

    return infos;
  };

  const getLayoutInfo = (layoutInfos, layoutName) => layoutInfos.get(String(layoutName || "").trim().toLowerCase()) || null;

  const pickPlaceholderBounds = (layoutInfo, kinds) => {
    if (!layoutInfo) return null;
    const want = Array.isArray(kinds) ? kinds : [kinds];
    for (let i = 0; i < want.length; i += 1) {
      const kind = String(want[i] || "").trim();
      const match = layoutInfo.placeholders.find(
        (ph) => ph.type.toLowerCase() === kind.toLowerCase() && !!ph.bounds,
      );
      if (match) return match.bounds;
    }
    return null;
  };

  const pickPlaceholderSlot = (layoutInfo, kinds, options = {}) => {
    if (!layoutInfo) return null;
    const want = Array.isArray(kinds) ? kinds : [kinds];
    const requireBounds = options.requireBounds === true;
    for (let i = 0; i < want.length; i += 1) {
      const kind = String(want[i] || "").trim().toLowerCase();
      const match = (layoutInfo.placeholders || []).find((ph) => {
        if (String(ph?.type || "").toLowerCase() !== kind) return false;
        if (!requireBounds) return true;
        return !!ph?.bounds;
      });
      if (match) return match;
    }
    return null;
  };

  const listPictureSlots = (layoutInfo) => {
    if (!layoutInfo) return [];
    return layoutInfo.placeholders
      .filter((ph) => ph.type.toLowerCase() === "pic")
      .filter((ph) => !!ph.bounds || !!ph.idxKey);
  };

  const listPictureBounds = (layoutInfo) => {
    return listPictureSlots(layoutInfo)
      .filter((ph) => !!ph.bounds)
      .map((ph) => ph.bounds);
  };

  const requireLayouts = (layoutInfos, layoutNames) => {
    const missing = [];
    (layoutNames || []).forEach((name) => {
      if (!name) return;
      if (!getLayoutInfo(layoutInfos, name)) missing.push(name);
    });
    if (missing.length) {
      const available = Array.from(layoutInfos.values())
        .map((info) => String(info?.name || "").trim())
        .filter(Boolean)
        .sort(compareAlphaNumeric);
      const preview = available.slice(0, 40);
      const suffix = available.length > preview.length ? ", ..." : "";
      throw new Error(
        `Template is missing required layout(s): ${missing.join(", ")}. `
        + `Detected layouts (${available.length}): ${preview.join(", ")}${suffix}`,
      );
    }
  };

  const requirePlaceholder = (layoutInfos, layoutName, kindSet, label) => {
    const info = getLayoutInfo(layoutInfos, layoutName);
    if (!info) throw new Error(`Template layout not found: ${layoutName}`);
    const want = Array.isArray(kindSet) ? kindSet : [kindSet];
    const matches = (info.placeholders || []).filter((ph) => {
      const type = String(ph?.type || "").toLowerCase();
      return want.some((kind) => type === String(kind || "").toLowerCase());
    });
    if (!matches.length) {
      const kinds = Array.isArray(kindSet) ? kindSet.join("/") : String(kindSet || "");
      throw new Error(`Layout ${layoutName} is missing required ${kinds} placeholder for ${label}.`);
    }
    const wantsPic = want.some((kind) => String(kind || "").toLowerCase() === "pic");
    if (wantsPic) {
      if (matches.some((ph) => !!ph.bounds || !!ph.idxKey)) return;
      const kinds = Array.isArray(kindSet) ? kindSet.join("/") : String(kindSet || "");
      throw new Error(`Layout ${layoutName} has ${kinds} placeholder(s), but no geometry and no idx for ${label}.`);
    }
    if (matches.some((ph) => !!ph.bounds)) return;
    const kinds = Array.isArray(kindSet) ? kindSet.join("/") : String(kindSet || "");
    throw new Error(`Layout ${layoutName} has ${kinds} placeholder(s), but no geometry (x/y/cx/cy) for ${label}.`);
  };

  const validateReportTemplate = (layoutInfos) => {
    const layouts = Object.values(REPORT_LAYOUTS);
    requireLayouts(layoutInfos, layouts);
    requirePlaceholder(layoutInfos, REPORT_LAYOUTS.cover, ["title", "ctrTitle"], "cover title");
    requirePlaceholder(layoutInfos, REPORT_LAYOUTS.chapterSeparator, ["title", "ctrTitle"], "chapter separator title");
    requirePlaceholder(layoutInfos, REPORT_LAYOUTS.chapterSnapshot, ["title", "ctrTitle"], "chapter snapshot title");
    requirePlaceholder(layoutInfos, REPORT_LAYOUTS.chapterSnapshot, ["pic"], "chapter snapshot image");
    requirePlaceholder(layoutInfos, REPORT_LAYOUTS.sectionText, ["title", "ctrTitle"], "section text title");
    requirePlaceholder(layoutInfos, REPORT_LAYOUTS.sectionText, ["body", "subTitle"], "section text body");
    [
      REPORT_LAYOUTS.sectionPhotoLow,
      REPORT_LAYOUTS.sectionPhotoHigh,
      REPORT_LAYOUTS.observationTextPhotoLow,
      REPORT_LAYOUTS.observationTextPhotoHigh,
    ].forEach((layoutName) => {
      const count = listPictureSlots(getLayoutInfo(layoutInfos, layoutName)).length;
      if (count < 1) throw new Error(`Layout ${layoutName} has no picture placeholders.`);
    });
  };

  const localeSuffix = (locale) => {
    const base = String(locale || "de-CH").toLowerCase().split("-")[0];
    if (base === "fr") return "f";
    if (base === "it") return "i";
    return "d";
  };

  const trainingLayoutForTag = (tag, suffix) => {
    const normalized = normalizeTag(tag);
    const base = TRAINING_LAYOUT_BY_TAG_BASE[normalized];
    if (!base) return TRAINING_LAYOUTS.defaultPhoto;
    return `${base}_${suffix}`;
  };

  const validateTrainingTemplate = (layoutInfos, locale, tagList) => {
    const suffix = localeSuffix(locale);
    const introLayout = TRAINING_LAYOUTS.introBySuffix[suffix];
    if (!introLayout) {
      throw new Error(`Unsupported locale suffix for training template validation: ${suffix}`);
    }
    const required = new Set([introLayout, TRAINING_LAYOUTS.sectionSeparator, TRAINING_LAYOUTS.defaultPhoto]);
    (tagList || []).forEach((tag) => required.add(trainingLayoutForTag(tag, suffix)));
    requireLayouts(layoutInfos, Array.from(required));

    const chapterPics = listPictureSlots(getLayoutInfo(layoutInfos, TRAINING_LAYOUTS.sectionSeparator)).length;
    if (chapterPics > 0) {
      throw new Error(`Layout ${TRAINING_LAYOUTS.sectionSeparator} should not contain picture placeholders.`);
    }

    [TRAINING_LAYOUTS.defaultPhoto].forEach((name) => {
      const count = listPictureSlots(getLayoutInfo(layoutInfos, name)).length;
      if (count < 1) throw new Error(`Layout ${name} has no picture placeholders.`);
    });

    (tagList || []).forEach((tag) => {
      const layoutName = trainingLayoutForTag(tag, suffix);
      const count = listPictureSlots(getLayoutInfo(layoutInfos, layoutName)).length;
      if (count < 1) throw new Error(`Layout ${layoutName} has no picture placeholders.`);
    });
  };

  const getPresentationDocs = (templateMap) => {
    const presentationXml = getEntryText(templateMap, "ppt/presentation.xml");
    if (!presentationXml) throw new Error("Template missing ppt/presentation.xml");
    const presentationRelsXml = getEntryText(templateMap, "ppt/_rels/presentation.xml.rels");
    if (!presentationRelsXml) throw new Error("Template missing ppt/_rels/presentation.xml.rels");
    const contentTypesXml = getEntryText(templateMap, "[Content_Types].xml");
    if (!contentTypesXml) throw new Error("Template missing [Content_Types].xml");
    return {
      presentation: parseXml(presentationXml, "ppt/presentation.xml"),
      presentationRels: parseXml(presentationRelsXml, "ppt/_rels/presentation.xml.rels"),
      contentTypes: parseXml(contentTypesXml, "[Content_Types].xml"),
    };
  };

  const nextSlideIndex = (templateMap) => {
    const nums = Array.from(templateMap.keys())
      .map((name) => {
        const m = name.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        return m ? Number(m[1]) : 0;
      })
      .filter((n) => Number.isFinite(n) && n > 0);
    return (nums.length ? Math.max(...nums) : 0) + 1;
  };

  const nextMediaIndex = (templateMap) => {
    const nums = Array.from(templateMap.keys())
      .map((name) => {
        const m = name.match(/^ppt\/media\/image(\d+)\.(png|jpe?g)$/i);
        return m ? Number(m[1]) : 0;
      })
      .filter((n) => Number.isFinite(n) && n > 0);
    return (nums.length ? Math.max(...nums) : 0) + 1;
  };

  const nextRelNumeric = (relsDoc) => {
    const rels = Array.from(relsDoc.getElementsByTagNameNS(NS_REL, "Relationship"));
    const nums = rels
      .map((rel) => String(rel.getAttribute("Id") || ""))
      .map((id) => {
        const m = id.match(/^rId(\d+)$/);
        return m ? Number(m[1]) : 0;
      })
      .filter((n) => Number.isFinite(n) && n > 0);
    return (nums.length ? Math.max(...nums) : 0) + 1;
  };

  const clearExistingSlideLinks = (presentationDoc, relsDoc) => {
    const slideRelType = REL_SLIDE;
    Array.from(relsDoc.getElementsByTagNameNS(NS_REL, "Relationship"))
      .filter((rel) => String(rel.getAttribute("Type") || "") === slideRelType)
      .forEach((rel) => rel.parentNode.removeChild(rel));

    let sldIdLst = presentationDoc.getElementsByTagNameNS(NS_P, "sldIdLst")[0] || null;
    if (!sldIdLst) {
      const root = presentationDoc.getElementsByTagNameNS(NS_P, "presentation")[0] || presentationDoc.documentElement;
      if (!root) throw new Error("Template presentation.xml is invalid: missing p:presentation root.");
      sldIdLst = presentationDoc.createElementNS(NS_P, "p:sldIdLst");
      // Empty templates may not contain p:sldIdLst yet. Insert it in a stable position.
      const children = Array.from(root.childNodes || []).filter((node) => node.nodeType === 1);
      const masterList = children.find((node) => node.localName === "sldMasterIdLst") || null;
      if (masterList && masterList.nextSibling) {
        root.insertBefore(sldIdLst, masterList.nextSibling);
      } else if (masterList) {
        root.appendChild(sldIdLst);
      } else if (children.length) {
        root.insertBefore(sldIdLst, children[0]);
      } else {
        root.appendChild(sldIdLst);
      }
    }
    while (sldIdLst.firstChild) sldIdLst.removeChild(sldIdLst.firstChild);
    return sldIdLst;
  };

  const ensureContentTypeOverride = (contentTypesDoc, partName, contentType) => {
    const types = contentTypesDoc.getElementsByTagNameNS(NS_CT, "Types")[0];
    const overrides = Array.from(contentTypesDoc.getElementsByTagNameNS(NS_CT, "Override"));
    const exists = overrides.some((el) => String(el.getAttribute("PartName") || "") === partName);
    if (exists) return;
    const node = contentTypesDoc.createElementNS(NS_CT, "Override");
    node.setAttribute("PartName", partName);
    node.setAttribute("ContentType", contentType);
    types.appendChild(node);
  };

  const ensureContentTypeDefault = (contentTypesDoc, extension, contentType) => {
    const ext = String(extension || "").replace(/^\./, "").trim().toLowerCase();
    if (!ext) return;
    const types = contentTypesDoc.getElementsByTagNameNS(NS_CT, "Types")[0];
    const defaults = Array.from(contentTypesDoc.getElementsByTagNameNS(NS_CT, "Default"));
    const exists = defaults.some((el) => String(el.getAttribute("Extension") || "").toLowerCase() === ext);
    if (exists) return;
    const node = contentTypesDoc.createElementNS(NS_CT, "Default");
    node.setAttribute("Extension", ext);
    node.setAttribute("ContentType", contentType);
    types.appendChild(node);
  };

  const createSlideRelDoc = (layoutPartName, imageTargets) => {
    const relsDoc = parseXml(`<Relationships xmlns="${NS_REL}"/>`, "slide rels");
    const root = relsDoc.documentElement;

    const layoutRel = relsDoc.createElementNS(NS_REL, "Relationship");
    layoutRel.setAttribute("Id", "rId1");
    layoutRel.setAttribute("Type", REL_SLIDE_LAYOUT);
    layoutRel.setAttribute("Target", `../slideLayouts/${layoutPartName.split("/").pop()}`);
    root.appendChild(layoutRel);

    (imageTargets || []).forEach((target, index) => {
      const rel = relsDoc.createElementNS(NS_REL, "Relationship");
      rel.setAttribute("Id", `rId${index + 2}`);
      rel.setAttribute("Type", REL_IMAGE);
      rel.setAttribute("Target", `../media/${target}`);
      root.appendChild(rel);
    });

    return serializeXml(relsDoc);
  };

  const paragraphRunsXml = (line, size = 1800, bold = false) => {
    const value = xmlEscape(line || "");
    const rPr = [`<a:rPr lang=\"en-US\" sz=\"${size}\"`];
    if (bold) rPr.push(" b=\"1\"");
    rPr.push("/>");
    return `<a:r>${rPr.join("")}<a:t>${value}</a:t></a:r>`;
  };

  const textBodyXml = (text, options = {}) => {
    const size = Number(options.size || 1800);
    const bold = options.bold === true;
    const lines = String(text || "").split(/\r?\n/);
    const paragraphs = (lines.length ? lines : [""]).map((line) => {
      if (!line.trim()) return "<a:p><a:endParaRPr lang=\"en-US\"/></a:p>";
      return `<a:p>${paragraphRunsXml(line, size, bold)}<a:endParaRPr lang=\"en-US\"/></a:p>`;
    }).join("");
    return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody>`;
  };

  const textBodyXmlInherit = (text) => {
    const lines = String(text || "").split(/\r?\n/);
    const paragraphs = (lines.length ? lines : [""]).map((line) => {
      if (!line.trim()) return "<a:p><a:endParaRPr/></a:p>";
      return `<a:p><a:r><a:t>${xmlEscape(line)}</a:t></a:r><a:endParaRPr/></a:p>`;
    }).join("");
    return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody>`;
  };

  const textShapeXml = ({ id, name, bounds, text, size = 1800, bold = false }) => {
    const x = toInt(bounds?.x, 0);
    const y = toInt(bounds?.y, 0);
    const cx = Math.max(1, toInt(bounds?.cx, 1000000));
    const cy = Math.max(1, toInt(bounds?.cy, 300000));
    return [
      "<p:sp>",
      "<p:nvSpPr>",
      `<p:cNvPr id=\"${id}\" name=\"${xmlEscape(name)}\"/>`,
      "<p:cNvSpPr txBox=\"1\"/>",
      "<p:nvPr/>",
      "</p:nvSpPr>",
      "<p:spPr>",
      `<a:xfrm><a:off x=\"${x}\" y=\"${y}\"/><a:ext cx=\"${cx}\" cy=\"${cy}\"/></a:xfrm>`,
      "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
      "<a:noFill/>",
      "</p:spPr>",
      textBodyXml(text, { size, bold }),
      "</p:sp>",
    ].join("");
  };

  const textPlaceholderShapeXml = ({ id, name, text, phType = "body", idxKey = "" }) => {
    const idxAttr = idxKey ? ` idx=\"${xmlEscape(idxKey)}\"` : "";
    return [
      "<p:sp>",
      "<p:nvSpPr>",
      `<p:cNvPr id=\"${id}\" name=\"${xmlEscape(name)}\"/>`,
      "<p:cNvSpPr/>",
      `<p:nvPr><p:ph type=\"${xmlEscape(phType)}\"${idxAttr}/></p:nvPr>`,
      "</p:nvSpPr>",
      "<p:spPr/>",
      textBodyXmlInherit(text),
      "</p:sp>",
    ].join("");
  };

  const fitContain = (imageW, imageH, box) => {
    const bw = Math.max(1, Number(box?.cx || 1));
    const bh = Math.max(1, Number(box?.cy || 1));
    const iw = Math.max(1, Number(imageW || 1));
    const ih = Math.max(1, Number(imageH || 1));
    const scale = Math.min(bw / iw, bh / ih);
    const w = Math.max(1, Math.round(iw * scale));
    const h = Math.max(1, Math.round(ih * scale));
    return {
      x: Math.round(Number(box.x || 0) + (bw - w) / 2),
      y: Math.round(Number(box.y || 0) + (bh - h) / 2),
      cx: w,
      cy: h,
    };
  };

  const pictureShapeXml = ({ id, name, relId, bounds, imageW, imageH }) => {
    const fit = fitContain(imageW, imageH, bounds);
    return [
      "<p:pic>",
      "<p:nvPicPr>",
      `<p:cNvPr id=\"${id}\" name=\"${xmlEscape(name)}\"/>`,
      "<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>",
      "<p:nvPr/>",
      "</p:nvPicPr>",
      "<p:blipFill>",
      `<a:blip r:embed=\"${xmlEscape(relId)}\"/>`,
      "<a:stretch><a:fillRect/></a:stretch>",
      "</p:blipFill>",
      "<p:spPr>",
      `<a:xfrm><a:off x=\"${fit.x}\" y=\"${fit.y}\"/><a:ext cx=\"${fit.cx}\" cy=\"${fit.cy}\"/></a:xfrm>`,
      "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
      "</p:spPr>",
      "</p:pic>",
    ].join("");
  };

  const fallbackPictureBounds = (layoutInfo) => {
    const body = pickPlaceholderBounds(layoutInfo, ["body", "subTitle"]);
    if (body) return body;
    const title = pickPlaceholderBounds(layoutInfo, ["title", "ctrTitle"]);
    if (title) {
      const slideCx = 12192000;
      const slideCy = 6858000;
      const marginX = 551384;
      const y = Math.max(0, Number(title.y || 0) + Number(title.cy || 0) + 180000);
      const cy = Math.max(900000, slideCy - y - 240000);
      return {
        x: marginX,
        y,
        cx: Math.max(2000000, slideCx - marginX * 2),
        cy,
      };
    }
    return { x: 551384, y: 1557338, cx: 11089232, cy: 4608512 };
  };

  const picturePlaceholderShapeXml = ({ id, name, relId, idxKey }) => [
    "<p:pic>",
    "<p:nvPicPr>",
    `<p:cNvPr id=\"${id}\" name=\"${xmlEscape(name)}\"/>`,
    "<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>",
    `<p:nvPr><p:ph type=\"pic\" idx=\"${xmlEscape(idxKey)}\"/></p:nvPr>`,
    "</p:nvPicPr>",
    "<p:blipFill rotWithShape=\"1\">",
    `<a:blip r:embed=\"${xmlEscape(relId)}\"/>`,
    "<a:stretch/>",
    "</p:blipFill>",
    "<p:spPr/>",
    "</p:pic>",
  ].join("");

  const emptyPicturePlaceholderShapeXml = ({ id, name, idxKey }) => [
    "<p:sp>",
    "<p:nvSpPr>",
    `<p:cNvPr id=\"${id}\" name=\"${xmlEscape(name)}\"/>`,
    "<p:cNvSpPr/>",
    `<p:nvPr><p:ph type=\"pic\" idx=\"${xmlEscape(idxKey)}\"/></p:nvPr>`,
    "</p:nvSpPr>",
    "<p:spPr/>",
    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>",
    "</p:sp>",
  ].join("");

  const buildSlideXml = ({ layoutInfo, title, body, images, preservePicPlaceholders = false }) => {
    let nextShapeId = 2;
    const shapes = [];

    const titleSlot = pickPlaceholderSlot(layoutInfo, ["title", "ctrTitle"]);
    const bodySlot = pickPlaceholderSlot(layoutInfo, ["body", "subTitle"]);

    if (String(title || "").trim()) {
      if (!titleSlot) {
        throw new Error(`Layout ${layoutInfo?.name || "(unknown)"} has no title/ctrTitle placeholder.`);
      }
      if (titleSlot.type) {
        shapes.push(textPlaceholderShapeXml({
          id: nextShapeId,
          name: `Title ${nextShapeId}`,
          text: String(title || ""),
          phType: titleSlot.type,
          idxKey: titleSlot.idxKey || "",
        }));
      } else if (titleSlot.bounds) {
        shapes.push(textShapeXml({
          id: nextShapeId,
          name: `Title ${nextShapeId}`,
          bounds: titleSlot.bounds,
          text: String(title || ""),
          size: 2400,
          bold: true,
        }));
      } else {
        throw new Error(`Layout ${layoutInfo?.name || "(unknown)"} title placeholder is not renderable.`);
      }
      nextShapeId += 1;
    }

    if (String(body || "").trim()) {
      if (!bodySlot) {
        throw new Error(`Layout ${layoutInfo?.name || "(unknown)"} has no body/subTitle placeholder.`);
      }
      if (bodySlot.type) {
        shapes.push(textPlaceholderShapeXml({
          id: nextShapeId,
          name: `Body ${nextShapeId}`,
          text: String(body || ""),
          phType: bodySlot.type,
          idxKey: bodySlot.idxKey || "",
        }));
      } else if (bodySlot.bounds) {
        shapes.push(textShapeXml({
          id: nextShapeId,
          name: `Body ${nextShapeId}`,
          bounds: bodySlot.bounds,
          text: String(body || ""),
          size: 1800,
          bold: false,
        }));
      } else {
        throw new Error(`Layout ${layoutInfo?.name || "(unknown)"} body placeholder is not renderable.`);
      }
      nextShapeId += 1;
    }

    const picSlots = listPictureSlots(layoutInfo);
    (images || []).forEach((img, index) => {
      const slot = picSlots[index] || picSlots[0];
      if (!slot) {
        throw new Error(`Layout ${layoutInfo?.name || "(unknown)"} has no picture placeholders.`);
      }
      const relId = `rId${index + 2}`;
      const bounds = slot.bounds || fallbackPictureBounds(layoutInfo);
      if (bounds) {
        shapes.push(pictureShapeXml({
          id: nextShapeId,
          name: `Picture ${nextShapeId}`,
          relId,
          bounds,
          imageW: img.width,
          imageH: img.height,
        }));
      } else if (slot.idxKey) {
        // Bind image to the placeholder identity; PowerPoint resolves geometry.
        shapes.push(picturePlaceholderShapeXml({
          id: nextShapeId,
          name: `Picture Placeholder ${nextShapeId}`,
          relId,
          idxKey: slot.idxKey,
        }));
      } else {
        throw new Error(`Layout ${layoutInfo?.name || "(unknown)"} picture placeholder is not renderable (missing bounds and idx).`);
      }
      nextShapeId += 1;
    });

    if (preservePicPlaceholders === true && (!images || !images.length)) {
      picSlots.forEach((slot) => {
        if (!slot?.idxKey) return;
        shapes.push(emptyPicturePlaceholderShapeXml({
          id: nextShapeId,
          name: `Picture Placeholder ${nextShapeId}`,
          idxKey: slot.idxKey,
        }));
        nextShapeId += 1;
      });
    }

    return [
      `<p:sld xmlns:a=\"${NS_A}\" xmlns:r=\"${NS_R}\" xmlns:p=\"${NS_P}\">`,
      "<p:cSld><p:spTree>",
      "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>",
      "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>",
      shapes.join(""),
      "</p:spTree></p:cSld>",
      "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>",
      "</p:sld>",
    ].join("");
  };

  const loadProjectImage = async (projectHandle, relPath, cache) => {
    const key = String(relPath || "").trim();
    if (!key) return null;
    if (cache.has(key)) return cache.get(key);
    const file = await tryReadProjectFile(projectHandle, key);
    if (!file) throw new Error(`Photo not found: ${key}`);
    const dims = await imageDimensions(file);
    const bytes = await blobToUint8(file);
    const extRaw = String(file.name || key).split(".").pop() || "jpg";
    const ext = extRaw.toLowerCase() === "jpeg" ? "jpg" : extRaw.toLowerCase();
    const out = {
      sourcePath: key,
      file,
      bytes,
      width: dims.width,
      height: dims.height,
      ext: ["png", "jpg", "jpeg"].includes(ext) ? (ext === "jpeg" ? "jpg" : ext) : "jpg",
    };
    cache.set(key, out);
    return out;
  };

  const reportTitleForChapter = (chapter, locale = "de-CH") => {
    const cid = String(chapter?.id || "").trim();
    const title = stripLeadingNumber(resolveLocalizedText(chapter?.title, locale)).trim();
    if (cid === "0") return title || "Management Summary";
    if (!cid) return title;
    if (!title) return cid;
    if (cid.includes(".")) return `${cid} ${title}`;
    return `${cid}. ${title}`;
  };

  const pickSectionPhotoLayout = (layoutInfos, count) => {
    const c = Math.max(1, Number(count) || 1);
    if (c <= 1) return REPORT_LAYOUTS.chapterSnapshot;
    if (c === 2 && getLayoutInfo(layoutInfos, REPORT_OPTIONAL_LAYOUTS.sectionPhotoTwo)) {
      return REPORT_OPTIONAL_LAYOUTS.sectionPhotoTwo;
    }
    if (c === 3 && getLayoutInfo(layoutInfos, REPORT_OPTIONAL_LAYOUTS.sectionPhotoThree)) {
      return REPORT_OPTIONAL_LAYOUTS.sectionPhotoThree;
    }
    if (c <= 4) return REPORT_LAYOUTS.sectionPhotoLow;
    return REPORT_LAYOUTS.sectionPhotoHigh;
  };

  const buildReportSlidePlan = async ({
    project,
    sidecarDoc,
    layoutInfos,
    toText,
    compareIdSegments,
    spiderScoreMap,
    spiderImage,
    logoSmall,
    logoLarge,
  }) => {
    const chapters = [...(project?.chapters || [])].sort((a, b) => {
      if (typeof compareIdSegments === "function") return compareIdSegments(a?.id, b?.id);
      return compareAlphaNumeric(a?.id, b?.id);
    });

    const { reportMap } = buildPhotoFileMap(sidecarDoc);
    const slides = [];

    const dateLabel = formatDateLabel(project?.meta?.createdAt);
    const company = String(project?.meta?.company || "").trim() || "Company";
    const moderator = String(project?.meta?.moderator || "").trim();
    const locale = String(project?.meta?.locale || "de-CH");

    slides.push({
      layout: REPORT_LAYOUTS.cover,
      title: coverReportTitle(locale),
      body: `${dateLabel}\n${snapshotLabels(locale).moderator}: ${moderator || "-"}`,
      images: [],
      preservePicPlaceholders: true,
    });

    for (let i = 0; i < chapters.length; i += 1) {
      const chapter = chapters[i];
      const chapterId = String(chapter?.id || "").trim();
      if (!chapterId) continue;
      const chapterTitle = resolveLocalizedText(chapter?.title, locale).trim();

      if (chapterId === "0") {
        const summaryTitle = stripLeadingNumber(chapterTitle || "Management Summary");
        slides.push({
          layout: REPORT_LAYOUTS.chapterSeparator,
          title: summaryTitle,
          body: "",
          images: [],
        });

        const lines = buildSummaryRecommendationLines(chapter, toText);
        chunkArray(lines, 3).forEach((chunk) => {
          slides.push({
            layout: REPORT_LAYOUTS.summaryText,
            title: summaryTitle,
            body: chunk.join("\n\n"),
            images: [],
          });
        });

        const summaryAssessment = assessmentSummaryTitle(locale);
        if (!(spiderImage instanceof Blob)) {
          throw new Error("Spider chart image is missing for report summary section.");
        }
        slides.push({
          layout: REPORT_LAYOUTS.chapterSeparator,
          title: summaryAssessment,
          body: "",
          images: [],
        });
        slides.push({
          layout: REPORT_LAYOUTS.chapterSnapshot,
          title: summaryAssessment,
          body: "",
          images: [spiderImage],
        });
        continue;
      }

      if (chapterId === "4.8") {
        const displaySectionId = resolveSpecial48DisplaySectionId(chapters);
        const separatorTitle = `${displaySectionId} ${chapterTitle || ""}`.trim();
        slides.push({
          layout: REPORT_LAYOUTS.observationSeparator,
          title: separatorTitle,
          body: "",
          images: [],
        });

        const ordered = orderObservationRows(chapter).filter((row) => isIncludedRow(row));
        let item = 0;
        for (let r = 0; r < ordered.length; r += 1) {
          const row = ordered[r];
          item += 1;
          const displayId = `${displaySectionId}.${item}`;
          const obsTag = resolveObservationTag(row, locale);
          const obsTitle = resolveObservationTitle(row, obsTag, locale);
          const fullTitle = `${displayId} ${obsTitle}`.trim();
          const finding = String(resolveFindingText(row, toText) || "").trim();

          const photos = ensureMapArray(reportMap, obsTag);
          // Observation finding text needs a body placeholder.
          // Use the text-capable layout for the first slide, then continue
          // remaining photos on pure photo layouts.
          const firstLayout = REPORT_LAYOUTS.observationTextPhotoLow;
          const firstSlots = photos.length > 0 ? 1 : 0;
          slides.push({
            layout: firstLayout,
            title: fullTitle,
            body: finding,
            images: photos.slice(0, firstSlots),
          });

          const remaining = photos.slice(firstSlots);
          const chunks = chunkArray(remaining, 6);
          chunks.forEach((chunk) => {
            const layout = pickSectionPhotoLayout(layoutInfos, chunk.length);
            slides.push({
              layout,
              title: fullTitle,
              body: "",
              images: chunk,
            });
          });
        }
        continue;
      }

      const chapterLabel = reportTitleForChapter(chapter, locale);
      slides.push({
        layout: REPORT_LAYOUTS.chapterSeparator,
        title: chapterLabel,
        body: "",
        images: [],
      });

      const score = spiderScoreMap.get(chapterId) || { company: 0, consultant: 0 };
      const snapshot = await drawChapterSnapshot({
        chapterLabel,
        score,
        locale,
        company,
        moderator,
        dateLabel,
        logoSmall,
        logoLarge,
      });

      slides.push({
        layout: REPORT_LAYOUTS.chapterSnapshot,
        title: chapterLabel,
        body: "",
        images: [snapshot],
      });

      const sections = buildSectionBlocks(chapter, chapterId, locale);
      for (let s = 0; s < sections.length; s += 1) {
        const section = sections[s];
        const sectionTitle = `${section.displayId || section.rawId || ""} ${section.title || ""}`.trim();
        const lines = buildSectionFindingLines(section, toText);
        chunkArray(lines, 3).forEach((chunk) => {
          slides.push({
            layout: REPORT_LAYOUTS.sectionText,
            title: sectionTitle,
            body: chunk.join("\n\n"),
            images: [],
          });
        });

        const photos = ensureMapArray(reportMap, section.rawId);
        chunkArray(photos, 6).forEach((chunk) => {
          const layout = pickSectionPhotoLayout(layoutInfos, chunk.length);
          slides.push({
            layout,
            title: sectionTitle,
            body: "",
            images: chunk,
          });
        });
      }
    }

    return slides;
  };

  const buildTrainingSlidePlan = ({ project, sidecarDoc }) => {
    const locale = String(project?.meta?.locale || "de-CH");
    const suffix = localeSuffix(locale);
    const introLayout = TRAINING_LAYOUTS.introBySuffix[suffix];
    if (!introLayout) {
      throw new Error(`Unsupported locale for training export: ${locale}`);
    }
    const { trainingMap } = buildPhotoFileMap(sidecarDoc);

    const orderedTags = [];
    const seen = new Set();
    TRAINING_TAG_ORDER.forEach((tag) => {
      const existing = Array.from(trainingMap.keys()).find((key) => normalizeTag(key) === tag);
      if (!existing) return;
      orderedTags.push(existing);
      seen.add(existing);
    });

    Array.from(trainingMap.keys())
      .filter((key) => !seen.has(key))
      .sort(compareAlphaNumeric)
      .forEach((key) => orderedTags.push(key));

    const slides = [];
    slides.push({
      layout: introLayout,
      title: "Seminar",
      body: "",
      images: [],
    });

    orderedTags.forEach((tag) => {
      const photos = ensureMapArray(trainingMap, tag);
      if (!photos.length) return;
      slides.push({
        layout: TRAINING_LAYOUTS.sectionSeparator,
        title: String(tag || "").trim(),
        body: "",
        images: [],
      });
      const layout = trainingLayoutForTag(tag, suffix);
      slides.push({
        layout,
        title: "",
        body: "",
        images: photos,
      });
    });

    return {
      slides,
      tagsUsed: orderedTags,
      suffix,
    };
  };

  const flattenSlidesByPictureCapacity = (slides, layoutInfos) => {
    const out = [];
    (slides || []).forEach((slide) => {
      const layoutInfo = getLayoutInfo(layoutInfos, slide.layout);
      if (!layoutInfo) {
        throw new Error(`Template layout not found: ${slide.layout}`);
      }
      const photos = Array.isArray(slide.images) ? slide.images : [];
      if (!photos.length) {
        out.push({ ...slide, images: [] });
        return;
      }
      const slotCapacity = Math.max(1, listPictureSlots(layoutInfo).length);
      chunkArray(photos, slotCapacity).forEach((chunk) => {
        out.push({ ...slide, images: chunk });
      });
    });
    return out;
  };

  const materializeSlideImages = async ({ projectHandle, slides, binaryCache }) => {
    const out = [];
    for (let i = 0; i < slides.length; i += 1) {
      const slide = slides[i];
      const materialized = [];
      const images = Array.isArray(slide.images) ? slide.images : [];
      for (let j = 0; j < images.length; j += 1) {
        const source = images[j];
        if (source instanceof Blob) {
          const dims = await imageDimensions(source);
          materialized.push({
            sourcePath: `__blob__${i}_${j}`,
            file: source,
            bytes: await blobToUint8(source),
            width: dims.width,
            height: dims.height,
            ext: "png",
          });
          continue;
        }
        // eslint-disable-next-line no-await-in-loop
        const loaded = await loadProjectImage(projectHandle, String(source || ""), binaryCache);
        materialized.push(loaded);
      }
      out.push({
        ...slide,
        images: materialized,
      });
    }
    return out;
  };

  const writeSlidesToTemplate = ({ templateMap, layoutInfos, slides }) => {
    const docs = getPresentationDocs(templateMap);
    // Some templates only declare jpeg. We emit png/jpg as needed.
    ensureContentTypeDefault(docs.contentTypes, "png", "image/png");
    ensureContentTypeDefault(docs.contentTypes, "jpg", "image/jpeg");
    ensureContentTypeDefault(docs.contentTypes, "jpeg", "image/jpeg");
    const sldIdLst = clearExistingSlideLinks(docs.presentation, docs.presentationRels);
    let nextRel = nextRelNumeric(docs.presentationRels);
    let nextSlide = nextSlideIndex(templateMap);
    let nextMedia = nextMediaIndex(templateMap);
    let nextSlideId = 256;

    slides.forEach((slide) => {
      const layoutInfo = getLayoutInfo(layoutInfos, slide.layout);
      if (!layoutInfo) throw new Error(`Template layout not found: ${slide.layout}`);

      const mediaTargets = [];
      (slide.images || []).forEach((img) => {
        const ext = String(img.ext || "png").toLowerCase() === "jpg" ? "jpg" : "png";
        const mediaName = `image${nextMedia}.${ext}`;
        nextMedia += 1;
        setEntryBytes(templateMap, `ppt/media/${mediaName}`, img.bytes);
        mediaTargets.push(mediaName);
      });

      const slideXml = buildSlideXml({
        layoutInfo,
        title: slide.title,
        body: slide.body,
        images: slide.images,
        preservePicPlaceholders: slide.preservePicPlaceholders === true,
      });
      const slideRelXml = createSlideRelDoc(layoutInfo.partName, mediaTargets);

      const slidePart = `ppt/slides/slide${nextSlide}.xml`;
      const slideRelPart = `ppt/slides/_rels/slide${nextSlide}.xml.rels`;
      nextSlide += 1;

      setEntryText(templateMap, slidePart, slideXml);
      setEntryText(templateMap, slideRelPart, slideRelXml);

      ensureContentTypeOverride(
        docs.contentTypes,
        `/${slidePart}`,
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
      );

      const relId = `rId${nextRel}`;
      nextRel += 1;
      const relNode = docs.presentationRels.createElementNS(NS_REL, "Relationship");
      relNode.setAttribute("Id", relId);
      relNode.setAttribute("Type", REL_SLIDE);
      relNode.setAttribute("Target", `slides/${slidePart.split("/").pop()}`);
      docs.presentationRels.documentElement.appendChild(relNode);

      const sldNode = docs.presentation.createElementNS(NS_P, "p:sldId");
      sldNode.setAttribute("id", String(nextSlideId));
      nextSlideId += 1;
      sldNode.setAttributeNS(NS_R, "r:id", relId);
      sldIdLst.appendChild(sldNode);
    });

    setEntryText(templateMap, "ppt/presentation.xml", serializeXml(docs.presentation));
    setEntryText(templateMap, "ppt/_rels/presentation.xml.rels", serializeXml(docs.presentationRels));
    setEntryText(templateMap, "[Content_Types].xml", serializeXml(docs.contentTypes));
  };

  const exportPptx = async ({
    project,
    sidecarDoc,
    projectHandle,
    mode,
    compareIdSegments,
    toText,
    spiderOverrides,
    computeSpider,
  }) => {
    if (!projectHandle) throw new Error("Project folder not selected.");
    if (!project || !Array.isArray(project.chapters)) {
      throw new Error("Project content is missing or invalid.");
    }
    if (!sidecarDoc || typeof sidecarDoc !== "object") {
      throw new Error("Sidecar data is missing.");
    }

    const templateFile = await pickPptTemplate();
    const templateMap = await getTemplateMap(templateFile);
    const layoutInfos = getLayoutInfos(templateMap);
    if (!layoutInfos.size) {
      const allEntries = Array.from(templateMap.keys());
      const xmlEntries = allEntries
        .map((name) => normalizePartName(name))
        .filter((name) => name.toLowerCase().endsWith(".xml"));
      const layoutLikeEntries = xmlEntries.filter((name) => name.toLowerCase().includes("slidelayout"));
      throw new Error(
        `No slide layouts detected in template. `
        + `Entries: ${allEntries.length}, XML entries: ${xmlEntries.length}, `
        + `layout-like XML entries: ${layoutLikeEntries.length}. `
        + `First layout-like entries: ${layoutLikeEntries.slice(0, 10).join(", ") || "(none)"}`,
      );
    }

    const locale = String(project?.meta?.locale || "de-CH");

    const logoSmall = await maybeLoadLogo(projectHandle, project?.meta?.logoSmallPath || "outputs/logo-small.png");
    const logoLarge = await maybeLoadLogo(projectHandle, project?.meta?.logoLargePath || "outputs/logo-large.png");

    let plannedSlides = [];

    if (mode === "report") {
      validateReportTemplate(layoutInfos);
      if (typeof computeSpider !== "function") {
        throw new Error("Spider calculation helper is unavailable.");
      }
      const spiderData = await computeSpider({
        project,
        overrides: spiderOverrides || {},
        dirHandle: projectHandle,
      });
      const spiderImage = await drawSpiderPng(
        spiderData,
        String(project?.meta?.company || "").trim() || "Company",
        project,
      );
      const spiderScoreMap = buildSpiderScoreMap(spiderData);
      plannedSlides = await buildReportSlidePlan({
        project,
        sidecarDoc,
        layoutInfos,
        toText,
        compareIdSegments,
        spiderScoreMap,
        spiderImage,
        logoSmall,
        logoLarge,
      });
    } else if (mode === "training") {
      const training = buildTrainingSlidePlan({ project, sidecarDoc });
      validateTrainingTemplate(layoutInfos, locale, training.tagsUsed);
      plannedSlides = training.slides;
    } else {
      throw new Error(`Unsupported export mode: ${mode}`);
    }

    if (!plannedSlides.length) {
      throw new Error(`No slides generated for ${mode} export.`);
    }

    const expandedPlan = flattenSlidesByPictureCapacity(plannedSlides, layoutInfos);
    const binaryCache = new Map();
    const materializedSlides = await materializeSlideImages({
      projectHandle,
      slides: expandedPlan,
      binaryCache,
    });

    writeSlidesToTemplate({
      templateMap,
      layoutInfos,
      slides: materializedSlides,
    });

    const outputBytes = buildZipStore(Array.from(templateMap.values()));
    const outputs = await getOutputsDirectory(projectHandle);
    const stamp = formatDateIso(new Date());
    const companySlug = toFileSafeSlug(project?.meta?.company || project?.meta?.projectName || "", "Company");

    const outName = mode === "report"
      ? `${stamp}-${companySlug}-Bericht-Besprechung.pptx`
      : `${stamp}-${companySlug}-Seminar-Slides.pptx`;

    await writeFileHandle(outputs, outName, outputBytes);

    return {
      savedAs: `outputs/${outName}`,
      slideCount: materializedSlides.length,
    };
  };

  const exportReportPptx = async (args) => exportPptx({ ...args, mode: "report" });
  const exportTrainingPptx = async (args) => exportPptx({ ...args, mode: "training" });

  window.AutoBerichtPptxExport = {
    exportReportPptx,
    exportTrainingPptx,
  };
})();
