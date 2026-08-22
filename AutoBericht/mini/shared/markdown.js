(() => {
  const escapeHtml = (value) => value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");

  const appendSegment = (segments, segment) => {
    if (!segment?.text) return;
    const previous = segments[segments.length - 1];
    if (
      previous
      && previous.type === segment.type
      && previous.url === segment.url
      && previous.bold === segment.bold
      && previous.italic === segment.italic
    ) {
      previous.text += segment.text;
      return;
    }
    segments.push(segment);
  };

  const normalizeHttpUrl = (value) => {
    const raw = String(value || "").trim();
    if (!/^https?:\/\//i.test(raw)) return "";
    try {
      const parsed = new URL(raw);
      if (!/^https?:$/.test(parsed.protocol)) return "";
      return parsed.href;
    } catch (err) {
      return "";
    }
  };

  const parseEmphasisSegments = (value) => {
    const source = String(value || "");
    const segments = [];
    const pattern = /\*\*\*([^*\n]+?)\*\*\*|\*\*([^*\n]+?)\*\*|\*([^*\n]+?)\*/g;
    let cursor = 0;
    let match = pattern.exec(source);
    while (match) {
      if (match.index > cursor) {
        appendSegment(segments, { type: "text", text: source.slice(cursor, match.index), bold: false, italic: false });
      }
      const both = match[1] != null;
      appendSegment(segments, {
        type: "text",
        text: both ? match[1] : (match[2] ?? match[3] ?? ""),
        bold: both || match[2] != null,
        italic: both || match[3] != null,
      });
      cursor = match.index + match[0].length;
      match = pattern.exec(source);
    }
    if (cursor < source.length) {
      appendSegment(segments, { type: "text", text: source.slice(cursor), bold: false, italic: false });
    }
    return segments;
  };

  const expandBareUrls = (segment) => {
    const source = String(segment?.text || "");
    const expanded = [];
    const pattern = /(https?:\/\/[^\s<>"']+)/g;
    let cursor = 0;
    let match = pattern.exec(source);
    while (match) {
      let rawUrl = match[1];
      let suffix = "";
      while (/[.,;!?]$/.test(rawUrl)) {
        suffix = rawUrl.slice(-1) + suffix;
        rawUrl = rawUrl.slice(0, -1);
      }
      if (match.index > cursor) {
        appendSegment(expanded, { ...segment, text: source.slice(cursor, match.index) });
      }
      const url = normalizeHttpUrl(rawUrl);
      if (url) appendSegment(expanded, { ...segment, type: "link", text: rawUrl, url });
      else appendSegment(expanded, { ...segment, text: rawUrl });
      if (suffix) appendSegment(expanded, { ...segment, text: suffix });
      cursor = match.index + match[1].length;
      match = pattern.exec(source);
    }
    if (cursor < source.length) appendSegment(expanded, { ...segment, text: source.slice(cursor) });
    return expanded;
  };

  const parseInlineMarkdownSegments = (value) => {
    const source = String(value || "");
    const parts = [];
    const linkPattern = /\[([^\]\n]+)\]\(([^()\n]*(?:\([^()\n]*\)[^()\n]*)*)\)/g;
    let cursor = 0;
    let match = linkPattern.exec(source);
    const appendPlain = (text) => {
      parseEmphasisSegments(text).forEach((segment) => {
        expandBareUrls(segment).forEach((expanded) => appendSegment(parts, expanded));
      });
    };
    while (match) {
      if (match.index > cursor) appendPlain(source.slice(cursor, match.index));
      const url = normalizeHttpUrl(match[2]);
      if (!url) {
        appendPlain(match[1]);
      } else {
        parseEmphasisSegments(match[1]).forEach((segment) => {
          appendSegment(parts, { ...segment, type: "link", url });
        });
      }
      cursor = match.index + match[0].length;
      match = linkPattern.exec(source);
    }
    if (cursor < source.length) appendPlain(source.slice(cursor));
    return parts;
  };

  const formatInlineMarkdown = (value) => parseInlineMarkdownSegments(value)
    .map((segment) => {
      let content = escapeHtml(String(segment.text || ""));
      if (segment.italic) content = `<em>${content}</em>`;
      if (segment.bold) content = `<strong>${content}</strong>`;
      if (segment.type === "link" && segment.url) {
        const href = escapeHtml(segment.url);
        content = `<a href="${href}" target="_blank" rel="noopener noreferrer">${content}</a>`;
      }
      return content;
    })
    .join("");

  const markdownToHtml = (text) => {
    const lines = String(text || "").split(/\r?\n/);
    const parts = [];
    let inList = false;
    let paragraphLines = [];

    const flushParagraph = () => {
      if (!paragraphLines.length) return;
      const safe = paragraphLines.map((line) => formatInlineMarkdown(line));
      parts.push(`<p>${safe.join("<br>")}</p>`);
      paragraphLines = [];
    };

    lines.forEach((line) => {
      const trimmed = line.trim();
      if (trimmed.startsWith("- ")) {
        flushParagraph();
        if (!inList) {
          parts.push("<ul>");
          inList = true;
        }
        const item = trimmed.slice(2);
        parts.push(`<li>${formatInlineMarkdown(item)}</li>`);
        return;
      }
      if (trimmed === "") {
        flushParagraph();
        if (inList) {
          parts.push("</ul>");
          inList = false;
        }
        return;
      }
      if (inList) {
        parts.push("</ul>");
        inList = false;
      }
      paragraphLines.push(line);
    });
    if (inList) parts.push("</ul>");
    flushParagraph();
    return parts.join("");
  };

  window.AutoBerichtMarkdown = {
    escapeHtml,
    formatInlineMarkdown,
    markdownToHtml,
    normalizeHttpUrl,
    parseInlineMarkdownSegments,
  };
})();
