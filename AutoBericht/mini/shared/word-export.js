(() => {
  const textDecoder = new TextDecoder();
  const textEncoder = new TextEncoder();

  const ZIP_LOCAL_FILE_SIG = 0x04034b50;
  const ZIP_CENTRAL_FILE_SIG = 0x02014b50;
  const ZIP_EOCD_SIG = 0x06054b50;

  const xmlEscape = (value) => String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");

  const escapeRegex = (value) => String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let j = 0; j < 8; j += 1) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c >>> 0;
    }
    return table;
  })();

  const crc32 = (bytes) => {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      c = crcTable[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  };

  const concatUint8 = (chunks) => {
    const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    chunks.forEach((chunk) => {
      out.set(chunk, offset);
      offset += chunk.length;
    });
    return out;
  };

  const readU16 = (view, offset) => view.getUint16(offset, true);
  const readU32 = (view, offset) => view.getUint32(offset, true);

  const writeU16 = (arr, offset, value) => {
    arr[offset] = value & 0xff;
    arr[offset + 1] = (value >>> 8) & 0xff;
  };

  const writeU32 = (arr, offset, value) => {
    arr[offset] = value & 0xff;
    arr[offset + 1] = (value >>> 8) & 0xff;
    arr[offset + 2] = (value >>> 16) & 0xff;
    arr[offset + 3] = (value >>> 24) & 0xff;
  };

  const findEocdOffset = (bytes) => {
    const minEocd = 22;
    const maxComment = 0xffff;
    const start = Math.max(0, bytes.length - minEocd - maxComment);
    for (let i = bytes.length - minEocd; i >= start; i -= 1) {
      if (
        bytes[i] === 0x50
        && bytes[i + 1] === 0x4b
        && bytes[i + 2] === 0x05
        && bytes[i + 3] === 0x06
      ) {
        return i;
      }
    }
    return -1;
  };

  const inflateRaw = async (compressed) => {
    const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    const buffer = await new Response(stream).arrayBuffer();
    return new Uint8Array(buffer);
  };

  const unzipAllEntries = async (arrayBuffer) => {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const eocd = findEocdOffset(bytes);
    if (eocd < 0) throw new Error("Invalid zip: EOCD not found");

    const centralDirSize = readU32(view, eocd + 12);
    const centralDirOffset = readU32(view, eocd + 16);
    const centralDirEnd = centralDirOffset + centralDirSize;

    const entries = [];
    let cursor = centralDirOffset;
    while (cursor < centralDirEnd) {
      const sig = readU32(view, cursor);
      if (sig !== ZIP_CENTRAL_FILE_SIG) throw new Error("Invalid zip: central directory signature mismatch");

      const flags = readU16(view, cursor + 8);
      const method = readU16(view, cursor + 10);
      const compressedSize = readU32(view, cursor + 20);
      const nameLen = readU16(view, cursor + 28);
      const extraLen = readU16(view, cursor + 30);
      const commentLen = readU16(view, cursor + 32);
      const localHeaderOffset = readU32(view, cursor + 42);

      const nameBytes = bytes.slice(cursor + 46, cursor + 46 + nameLen);
      const name = textDecoder.decode(nameBytes);

      const localSig = readU32(view, localHeaderOffset);
      if (localSig !== ZIP_LOCAL_FILE_SIG) throw new Error(`Invalid zip: local header mismatch for ${name}`);
      const localNameLen = readU16(view, localHeaderOffset + 26);
      const localExtraLen = readU16(view, localHeaderOffset + 28);
      const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
      const compressed = bytes.slice(dataStart, dataStart + compressedSize);

      let data;
      if (method === 0) {
        data = compressed;
      } else if (method === 8) {
        data = await inflateRaw(compressed);
      } else {
        throw new Error(`Unsupported zip method ${method} for ${name}`);
      }

      entries.push({ name, data, flags });
      cursor += 46 + nameLen + extraLen + commentLen;
    }

    return entries;
  };

  const buildZipStore = (entries) => {
    const localParts = [];
    const centralParts = [];
    let localOffset = 0;

    const sorted = [...entries].sort((a, b) => a.name.localeCompare(b.name, "en", { numeric: true }));

    sorted.forEach((entry) => {
      const nameBytes = textEncoder.encode(entry.name);
      const data = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
      const crc = crc32(data);

      const localHeader = new Uint8Array(30 + nameBytes.length);
      writeU32(localHeader, 0, ZIP_LOCAL_FILE_SIG);
      writeU16(localHeader, 4, 20);
      writeU16(localHeader, 6, (entry.flags || 0) & 0x0800); // keep UTF-8 bit only
      writeU16(localHeader, 8, 0); // store
      writeU16(localHeader, 10, 0);
      writeU16(localHeader, 12, 0);
      writeU32(localHeader, 14, crc);
      writeU32(localHeader, 18, data.length);
      writeU32(localHeader, 22, data.length);
      writeU16(localHeader, 26, nameBytes.length);
      writeU16(localHeader, 28, 0);
      localHeader.set(nameBytes, 30);

      localParts.push(localHeader, data);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      writeU32(centralHeader, 0, ZIP_CENTRAL_FILE_SIG);
      writeU16(centralHeader, 4, 20);
      writeU16(centralHeader, 6, 20);
      writeU16(centralHeader, 8, (entry.flags || 0) & 0x0800);
      writeU16(centralHeader, 10, 0);
      writeU16(centralHeader, 12, 0);
      writeU16(centralHeader, 14, 0);
      writeU32(centralHeader, 16, crc);
      writeU32(centralHeader, 20, data.length);
      writeU32(centralHeader, 24, data.length);
      writeU16(centralHeader, 28, nameBytes.length);
      writeU16(centralHeader, 30, 0);
      writeU16(centralHeader, 32, 0);
      writeU16(centralHeader, 34, 0);
      writeU16(centralHeader, 36, 0);
      writeU32(centralHeader, 38, 0);
      writeU32(centralHeader, 42, localOffset);
      centralHeader.set(nameBytes, 46);
      centralParts.push(centralHeader);

      localOffset += localHeader.length + data.length;
    });

    const localBlob = concatUint8(localParts);
    const centralBlob = concatUint8(centralParts);

    const eocd = new Uint8Array(22);
    writeU32(eocd, 0, ZIP_EOCD_SIG);
    writeU16(eocd, 4, 0);
    writeU16(eocd, 6, 0);
    writeU16(eocd, 8, sorted.length);
    writeU16(eocd, 10, sorted.length);
    writeU32(eocd, 12, centralBlob.length);
    writeU32(eocd, 16, localBlob.length);
    writeU16(eocd, 20, 0);

    return concatUint8([localBlob, centralBlob, eocd]);
  };

  const markdownToPlain = (value) => {
    let text = String(value || "");
    text = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1 $2");
    text = text.replace(/[*_`>#]/g, "");
    return text;
  };

  const stripLeadingNumber = (value) => String(value || "")
    .replace(/^\s*\d+(?:\.\d+)*(?:\s|[.:-]\s*)?/, "")
    .trim();

  const rowToText = (value, toText) => {
    if (typeof toText === "function") return toText(value);
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  };

  const isSectionRow = (row) => String(row?.kind || "").toLowerCase() === "section";

  const isIncludedRow = (row) => {
    const ws = row?.workstate;
    if (!ws || ws.includeFinding == null) return false;
    return ws.includeFinding === true;
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

  const buildIncludedSections = (rows) => {
    const included = new Set();
    rows.forEach((row) => {
      if (isSectionRow(row) || !isIncludedRow(row)) return;
      const sectionId = String(row?.sectionId || "").trim();
      if (sectionId) included.add(sectionId);
    });
    return included;
  };

  const buildRenumberMap = (rows, chapterId) => {
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

    return { rowMap };
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

  const buildChapterRows = (chapter, toText) => {
    const rows = orderRowsForChapter(chapter);
    const includedSections = buildIncludedSections(rows);
    const { rowMap } = buildRenumberMap(rows, chapter?.id || "");
    const output = [];

    rows.forEach((row) => {
      if (isSectionRow(row)) {
        const sectionId = String(row?.id || "").trim();
        if (!sectionId || !includedSections.has(sectionId)) return;
        output.push({ kind: "section", title: resolveSectionTitle(row) });
        return;
      }
      if (!isIncludedRow(row)) return;
      const rowId = String(row?.id || "").trim();
      output.push({
        kind: "finding",
        id: rowMap.get(rowId) || rowId,
        title: chapter?.id === "4.8" ? String(row?.titleOverride || "").trim() : "",
        finding: markdownToPlain(resolveFindingText(row, toText)),
        recommendation: markdownToPlain(resolveRecommendationText(row, toText)),
        priority: resolvePriorityText(row),
      });
    });

    return output;
  };

  const paragraphXml = (text, options = {}) => {
    const safe = xmlEscape(text || "");
    const boldStart = options.bold ? "<w:rPr><w:b/></w:rPr>" : "";
    return `<w:p><w:r>${boldStart}<w:t xml:space=\"preserve\">${safe}</w:t></w:r></w:p>`;
  };

  const multiParagraphXml = (text) => {
    const lines = String(text || "").split(/\r?\n/);
    const nonEmpty = lines.length ? lines : [""];
    return nonEmpty.map((line) => {
      const value = line.trim().startsWith("- ") ? `• ${line.trim().slice(2)}` : line;
      return paragraphXml(value);
    }).join("");
  };

  const buildChapter0Xml = (chapter, toText) => {
    const rows = buildChapterRows(chapter, toText).filter((entry) => entry.kind === "finding");
    if (!rows.length) return paragraphXml("(No included findings)");
    return rows.map((entry) => {
      const heading = paragraphXml(`${entry.id}${entry.title ? ` ${entry.title}` : ""}`.trim(), { bold: true });
      const recommendation = multiParagraphXml(entry.recommendation || "");
      return `${heading}${recommendation}${paragraphXml("")}`;
    }).join("");
  };

  const buildChapterTableXml = (chapter, toText) => {
    const rows = buildChapterRows(chapter, toText);
    if (!rows.length) return paragraphXml("(No included findings)");

    const tableRows = [];
    tableRows.push(
      [
        "<w:tr>",
        "<w:tc><w:tcPr><w:tcW w:w=\"8200\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr>",
        paragraphXml(""),
        "</w:tc>",
        "<w:tc><w:tcPr><w:tcW w:w=\"500\" w:type=\"dxa\"/></w:tcPr>",
        paragraphXml("✓", { bold: true }),
        "</w:tc>",
        "</w:tr>",
      ].join(""),
    );

    tableRows.push(
      [
        "<w:tr>",
        "<w:tc><w:tcPr><w:tcW w:w=\"8200\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr>",
        paragraphXml("Systempunkte mit Verbesserungspotenzial", { bold: true }),
        "</w:tc>",
        "<w:tc><w:tcPr><w:tcW w:w=\"500\" w:type=\"dxa\"/></w:tcPr>",
        paragraphXml(""),
        "</w:tc>",
        "</w:tr>",
      ].join(""),
    );

    tableRows.push(
      [
        "<w:tr>",
        "<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>",
        paragraphXml("Ist-Zustand", { bold: true }),
        "</w:tc>",
        "<w:tc><w:tcPr><w:tcW w:w=\"5200\" w:type=\"dxa\"/></w:tcPr>",
        paragraphXml("Lösungsansätze", { bold: true }),
        "</w:tc>",
        "<w:tc><w:tcPr><w:tcW w:w=\"500\" w:type=\"dxa\"/></w:tcPr>",
        paragraphXml("Prio", { bold: true }),
        "</w:tc>",
        "</w:tr>",
      ].join(""),
    );

    rows.forEach((entry) => {
      if (entry.kind === "section") {
        tableRows.push(
          [
            "<w:tr>",
            "<w:tc><w:tcPr><w:tcW w:w=\"8700\" w:type=\"dxa\"/><w:gridSpan w:val=\"3\"/></w:tcPr>",
            paragraphXml(entry.title || "", { bold: true }),
            "</w:tc>",
            "</w:tr>",
          ].join(""),
        );
        return;
      }

      const heading = `${entry.id || ""}${entry.title ? ` ${entry.title}` : ""}`.trim();
      const findingCell = `${paragraphXml(heading, { bold: true })}${multiParagraphXml(entry.finding || "")}`;
      const recommendationCell = multiParagraphXml(entry.recommendation || "");
      const priorityCell = paragraphXml(entry.priority || "", { bold: true });

      tableRows.push(
        [
          "<w:tr>",
          "<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>",
          findingCell,
          "</w:tc>",
          "<w:tc><w:tcPr><w:tcW w:w=\"5200\" w:type=\"dxa\"/></w:tcPr>",
          recommendationCell,
          "</w:tc>",
          "<w:tc><w:tcPr><w:tcW w:w=\"500\" w:type=\"dxa\"/></w:tcPr>",
          priorityCell,
          "</w:tc>",
          "</w:tr>",
        ].join(""),
      );
    });

    return [
      "<w:tbl>",
      "<w:tblPr>",
      "<w:tblW w:w=\"5000\" w:type=\"pct\"/>",
      "<w:tblLayout w:type=\"fixed\"/>",
      "<w:tblBorders>",
      "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>",
      "</w:tblBorders>",
      "</w:tblPr>",
      "<w:tblGrid>",
      "<w:gridCol w:w=\"3000\"/>",
      "<w:gridCol w:w=\"5200\"/>",
      "<w:gridCol w:w=\"500\"/>",
      "</w:tblGrid>",
      tableRows.join(""),
      "</w:tbl>",
    ].join("");
  };

  const replaceParagraphMarker = (xml, marker, replacementXml) => {
    const regex = new RegExp(`<w:p[^>]*>[\\s\\S]*?${escapeRegex(marker)}[\\s\\S]*?<\\/w:p>`);
    if (!regex.test(xml)) return { xml, replaced: false };
    return {
      xml: xml.replace(regex, replacementXml),
      replaced: true,
    };
  };

  const replaceTextMarkers = (xml, markerMap) => {
    let out = xml;
    Object.entries(markerMap || {}).forEach(([marker, value]) => {
      const safe = xmlEscape(value || "");
      out = out.split(marker).join(safe);
    });
    return out;
  };

  const ensurePngContentType = (xml) => {
    if (/Extension="png"/i.test(xml)) return xml;
    const insert = '<Default Extension="png" ContentType="image/png"/>';
    return xml.replace("</Types>", `${insert}</Types>`);
  };

  const getNextRelId = (relsXml) => {
    const matches = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
    const max = matches.length ? Math.max(...matches) : 0;
    return `rId${max + 1}`;
  };

  const appendRelationship = (relsXml, relId, target) => {
    const rel = `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
    return relsXml.replace("</Relationships>", `${rel}</Relationships>`);
  };

  const emuFromCm = (cm) => Math.round(Number(cm || 0) * 360000);

  const drawingXml = (relId, name, widthEmu, heightEmu) => [
    "<w:p><w:r><w:drawing>",
    "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">",
    `<wp:extent cx=\"${widthEmu}\" cy=\"${heightEmu}\"/>`,
    "<wp:docPr id=\"1\" name=\"Picture\"/>",
    "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
    "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
    "<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"",
    xmlEscape(name),
    "\"/><pic:cNvPicPr/></pic:nvPicPr>",
    "<pic:blipFill>",
    `<a:blip r:embed=\"${relId}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>`,
    "<a:stretch><a:fillRect/></a:stretch>",
    "</pic:blipFill>",
    "<pic:spPr>",
    "<a:xfrm><a:off x=\"0\" y=\"0\"/>",
    `<a:ext cx=\"${widthEmu}\" cy=\"${heightEmu}\"/></a:xfrm>`,
    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
    "</pic:spPr>",
    "</pic:pic>",
    "</a:graphicData>",
    "</a:graphic>",
    "</wp:inline>",
    "</w:drawing></w:r></w:p>",
  ].join("");

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

  const drawSpiderPng = async (spiderData, companyLabel = "Company") => {
    const rows = spiderData?.effective?.chapters_1_11 || [];
    const labels = rows.map((row) => String(row?.id || ""));
    const companyValues = rows.map((row) => Number(row?.company || 0));
    const consultantValues = rows.map((row) => Number(row?.consultant || 0));

    const canvas = document.createElement("canvas");
    canvas.width = 1400;
    canvas.height = 980;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas context unavailable.");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const cx = canvas.width * 0.5;
    const cy = canvas.height * 0.56;
    const radius = Math.min(canvas.width, canvas.height) * 0.32;
    const count = labels.length || 1;

    ctx.strokeStyle = "#dad6cf";
    for (let ring = 1; ring <= 5; ring += 1) {
      const r = (radius * ring) / 5;
      ctx.beginPath();
      for (let i = 0; i < count; i += 1) {
        const angle = (-Math.PI / 2) + ((Math.PI * 2 * i) / count);
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.stroke();
    }

    labels.forEach((label, i) => {
      const angle = (-Math.PI / 2) + ((Math.PI * 2 * i) / count);
      const x = cx + Math.cos(angle) * radius;
      const y = cy + Math.sin(angle) * radius;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.lineTo(x, y);
      ctx.strokeStyle = "#c8c2b8";
      ctx.stroke();

      ctx.font = "22px Segoe UI, Arial, sans-serif";
      ctx.fillStyle = "#374151";
      ctx.textAlign = Math.cos(angle) > 0.2 ? "left" : Math.cos(angle) < -0.2 ? "right" : "center";
      ctx.textBaseline = Math.sin(angle) > 0.2 ? "top" : Math.sin(angle) < -0.2 ? "bottom" : "middle";
      ctx.fillText(label, cx + Math.cos(angle) * (radius + 38), cy + Math.sin(angle) * (radius + 38));
    });

    const drawSeries = (values, stroke, fill) => {
      ctx.beginPath();
      values.forEach((value, i) => {
        const pct = Math.max(0, Math.min(100, Number(value || 0))) / 100;
        const angle = (-Math.PI / 2) + ((Math.PI * 2 * i) / count);
        const x = cx + Math.cos(angle) * radius * pct;
        const y = cy + Math.sin(angle) * radius * pct;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.strokeStyle = stroke;
      ctx.lineWidth = 4;
      ctx.fill();
      ctx.stroke();
    };

    drawSeries(companyValues, "#2563eb", "rgba(37,99,235,0.18)");
    drawSeries(consultantValues, "#dc2626", "rgba(220,38,38,0.14)");

    ctx.font = "24px Segoe UI, Arial, sans-serif";
    const leftX = Math.round(canvas.width * 0.5) - 300;
    const rightX = Math.round(canvas.width * 0.5) + 12;
    ctx.fillStyle = "#2563eb";
    ctx.fillRect(leftX, canvas.height - 58, 24, 4);
    ctx.fillText(companyLabel || "Company", leftX + 34, canvas.height - 52);
    ctx.fillStyle = "#dc2626";
    ctx.fillRect(rightX, canvas.height - 58, 24, 4);
    ctx.fillText("Suva", rightX + 34, canvas.height - 52);

    const blob = await new Promise((resolve) => {
      canvas.toBlob((result) => resolve(result), "image/png", 0.95);
    });
    if (!blob) throw new Error("Failed to render spider image");
    return blob;
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

    chapters.forEach((chapter) => {
      const marker = `CHAPTER${chapter.id}$$`;
      const replacement = String(chapter.id) === "0"
        ? buildChapter0Xml(chapter, toText)
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

    const insertImageAtMarker = async ({ xmlPart, marker, imageFile, mediaName, cmHeight }) => {
      if (!imageFile) return false;
      const xml = xmlPart === "word/document.xml" ? documentXml : getText(xmlPart);
      if (!xml || !xml.includes(marker)) return false;

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

      const drawing = drawingXml(relId, mediaName, targetWidth, targetHeight);
      const patched = replaceParagraphMarker(xml, marker, drawing);
      if (!patched.replaced) return false;
      if (xmlPart === "word/document.xml") {
        documentXml = patched.xml;
      } else {
        setText(xmlPart, patched.xml);
      }
      return true;
    };

    const logoLargeFile = await tryReadProjectFile(projectHandle, project?.meta?.logoLargePath || "Outputs/logo-large.png");
    const logoSmallFile = await tryReadProjectFile(projectHandle, project?.meta?.logoSmallPath || "Outputs/logo-small.png");

    if (logoLargeFile) {
      await insertImageAtMarker({
        xmlPart: "word/document.xml",
        marker: "LOGO_BIG$$",
        imageFile: logoLargeFile,
        mediaName: "autobericht_logo_large.png",
        cmHeight: 2,
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
        const spiderBlob = await drawSpiderPng(
          spiderData,
          String(project?.meta?.company || "").trim() || "Company",
        );
        await insertImageAtMarker({
          xmlPart: "word/document.xml",
          marker: "SPIDER$$",
          imageFile: spiderBlob,
          mediaName: "autobericht_spider.png",
          cmHeight: 8.2,
        });
      } catch (err) {
        // Keep export running even if spider generation fails.
      }
    }

    setText("word/document.xml", documentXml);
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
