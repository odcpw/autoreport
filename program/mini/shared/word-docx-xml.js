/*
 * OOXML helper utilities for placeholder-safe DOCX mutation.
 *
 * Exposes `window.AutoBerichtWordDocxXml` helpers for:
 * - XML escaping
 * - paragraph-aware marker lookup/replacement
 * - text marker replacement
 * - content-type and relationship updates
 * - inline image drawing XML generation
 *
 * Marker replacement is paragraph-aware to tolerate split runs (`w:t`) in Word XML.
 */
(() => {
  const xmlEscape = (value) => String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");

  const xmlUnescape = (value) => String(value || "")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, "\"")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");

  const paragraphHasMarker = (paragraphXml, marker) => {
    const paragraph = String(paragraphXml || "");
    const needle = String(marker || "");
    if (!needle) return false;
    if (paragraph.includes(needle)) return true;

    const textMatches = [...paragraph.matchAll(/<(?:w:t|w:delText)\b[^>]*>([\s\S]*?)<\/(?:w:t|w:delText)>/g)];
    if (textMatches.length) {
      const merged = textMatches.map((match) => xmlUnescape(match[1] || "")).join("");
      if (merged.includes(needle)) return true;
    }

    const stripped = xmlUnescape(paragraph.replace(/<[^>]+>/g, ""));
    return stripped.includes(needle);
  };

  const hasMarker = (xml, marker) => {
    const source = String(xml || "");
    const needle = String(marker || "");
    if (!needle) return false;
    if (source.includes(needle)) return true;
    const paragraphRegex = /<w:p\b[\s\S]*?<\/w:p>/g;
    let match = paragraphRegex.exec(source);
    while (match) {
      if (paragraphHasMarker(match[0], needle)) return true;
      match = paragraphRegex.exec(source);
    }
    return false;
  };

  const replaceParagraphMarker = (xml, marker, replacementXml) => {
    const source = String(xml || "");
    const needle = String(marker || "");
    if (!needle) return { xml: source, replaced: false };

    let replaced = false;
    const paragraphRegex = /<w:p\b[\s\S]*?<\/w:p>/g;
    const out = source.replace(paragraphRegex, (paragraph) => {
      if (replaced) return paragraph;
      if (!paragraphHasMarker(paragraph, needle)) return paragraph;
      replaced = true;
      return String(replacementXml || "");
    });

    if (replaced) return { xml: out, replaced: true };

    const directIndex = source.indexOf(needle);
    if (directIndex < 0) return { xml: source, replaced: false };
    return {
      xml: source.replace(needle, String(replacementXml || "")),
      replaced: true,
    };
  };

  const replaceAllParagraphMarkers = (xml, marker, replacementXml) => {
    let out = String(xml || "");
    let count = 0;
    for (;;) {
      const patched = replaceParagraphMarker(out, marker, replacementXml);
      if (!patched.replaced) break;
      out = patched.xml;
      count += 1;
    }
    return { xml: out, count };
  };

  const replaceTextMarkers = (xml, markerMap) => {
    let out = String(xml || "");
    Object.entries(markerMap || {}).forEach(([marker, value]) => {
      out = out.split(marker).join(xmlEscape(value || ""));
    });
    return out;
  };

  const ensurePngContentType = (xml) => {
    const source = String(xml || "");
    if (/Extension="png"/i.test(source)) return source;
    return source.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>');
  };

  const getNextRelId = (relsXml) => {
    const matches = [...String(relsXml || "").matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
    const max = matches.length ? Math.max(...matches) : 0;
    return `rId${max + 1}`;
  };

  const appendRelationship = (relsXml, relId, target) => {
    const rel = `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
    return String(relsXml || "").replace("</Relationships>", `${rel}</Relationships>`);
  };

  const emuFromCm = (cm) => Math.round(Number(cm || 0) * 360000);

  const drawingXml = (relId, name, widthEmu, heightEmu, align = "") => [
    "<w:p>",
    align ? `<w:pPr><w:jc w:val="${xmlEscape(align)}"/></w:pPr>` : "",
    "<w:r><w:drawing>",
    "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">",
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>`,
    "<wp:docPr id=\"1\" name=\"Picture\"/>",
    "<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>",
    "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
    "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
    "<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"",
    xmlEscape(name),
    "\"/><pic:cNvPicPr><a:picLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></pic:cNvPicPr></pic:nvPicPr>",
    "<pic:blipFill>",
    `<a:blip r:embed="${relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
    "<a:stretch><a:fillRect/></a:stretch>",
    "</pic:blipFill>",
    "<pic:spPr>",
    "<a:xfrm><a:off x=\"0\" y=\"0\"/>",
    `<a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>`,
    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
    "</pic:spPr>",
    "</pic:pic>",
    "</a:graphicData>",
    "</a:graphic>",
    "</wp:inline>",
    "</w:drawing></w:r></w:p>",
  ].join("");

  const ensureUpdateFieldsOnOpen = (settingsXml) => {
    const source = String(settingsXml || "");
    if (!source.trim()) return source;
    const selfClosing = /<w:updateFields\b[^>]*\/>/;
    const openClose = /<w:updateFields\b[^>]*>[\s\S]*?<\/w:updateFields>/;
    if (selfClosing.test(source)) {
      return source.replace(selfClosing, '<w:updateFields w:val="true"/>');
    }
    if (openClose.test(source)) {
      return source.replace(openClose, '<w:updateFields w:val="true"/>');
    }
    if (!source.includes("</w:settings>")) return source;
    return source.replace("</w:settings>", '<w:updateFields w:val="true"/></w:settings>');
  };

  window.AutoBerichtWordDocxXml = {
    xmlEscape,
    hasMarker,
    replaceParagraphMarker,
    replaceAllParagraphMarkers,
    replaceTextMarkers,
    ensurePngContentType,
    getNextRelId,
    appendRelationship,
    emuFromCm,
    drawingXml,
    ensureUpdateFieldsOnOpen,
  };
})();
