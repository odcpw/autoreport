let markdownInstance = null;

function getMarkdownIt() {
  if (markdownInstance) return markdownInstance;
  if (typeof window !== 'undefined' && typeof window.markdownit === 'function') {
    markdownInstance = window.markdownit({
      html: false,
      linkify: true,
      breaks: true,
    });
  }
  return markdownInstance;
}

export function renderMarkdown(text) {
  const md = getMarkdownIt();
  if (!md) {
    return `<p>${escapeHtml(text || '')}</p>`;
  }
  return md.render(text || '');
}

export function renderMarkdownInline(text) {
  const md = getMarkdownIt();
  if (!md) {
    return escapeHtml((text || '').replace(/\n+/g, ' '));
  }
  return md.renderInline(text || '');
}

function escapeHtml(text) {
  return (text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
