(() => {
  const escapeHtml = (value) => value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");

  const formatInlineMarkdown = (value) => {
    let out = value;
    out = out.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
    out = out.replace(/\*(.+?)\*/g, "<em>$1</em>");
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (match, label, rawUrl) => {
      const url = String(rawUrl || "").trim();
      if (!/^https?:\/\//i.test(url)) return label;
      return `<a href="${url}" target="_blank" rel="noopener noreferrer">${label}</a>`;
    });
    return out;
  };

  const markdownToHtml = (text) => {
    const lines = String(text || "").split(/\r?\n/);
    const parts = [];
    let inList = false;
    let paragraphLines = [];

    const flushParagraph = () => {
      if (!paragraphLines.length) return;
      const safe = paragraphLines.map((line) => escapeHtml(line)).map((line) => formatInlineMarkdown(line));
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
        parts.push(`<li>${formatInlineMarkdown(escapeHtml(item))}</li>`);
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
  };
})();
