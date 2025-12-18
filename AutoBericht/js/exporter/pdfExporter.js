import { renderMarkdown } from '../utils/markdownRenderer.js';

export function exportReportPdf(snapshot) {
  if (!snapshot?.project) {
    throw new Error('No project data available');
  }
  const baseHref = window.location.href.replace(/[^/]+$/, '');
  const html = buildDocumentHtml(snapshot, baseHref);
  const printWindow = window.open('', '_blank');
  if (!printWindow) {
    throw new Error('Unable to open preview window. Disable popup blockers and try again.');
  }
  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
}

function buildDocumentHtml(snapshot, baseHref) {
  const { project } = snapshot;
  const now = new Date();
  const locale = snapshot.config?.locale || project.meta?.locale || 'en';
  const displayDate = formatLocaleDate(now, locale);
  const margin = normalizeMargin(snapshot.config?.pdfMargin);
  const showHeaderFooter = snapshot.config?.headerFooter !== 'hide';
  const includedRows = collectIncludedRows(project);

  const tocHtml = includedRows
    .map(
      (entry, index) => `
        <li>
          <span class="toc-index">${index + 1}</span>
          <span class="toc-title">${entry.id || ''} — ${entry.title || ''}</span>
        </li>
      `,
    )
    .join('\n');

  const chaptersHtml = includedRows
    .map((entry, index) => {
      const findingText = entry.finding?.useAdjusted
        ? entry.finding.adjustedText || entry.finding.masterText
        : entry.finding?.masterText;
      const recommendations = Object.keys(entry.recommendations || {})
        .map((key) => {
          const rec = entry.recommendations[key];
          const text = rec?.useAdjusted ? rec.adjusted || rec.master : rec?.master;
          return text ? `<li>${renderMarkdown(text)}</li>` : '';
        })
        .filter(Boolean)
        .join('');
      return `
        <section class="chapter" data-index="${index + 1}">
          <header>
            <h2>${entry.id || ''} — ${entry.title || ''}</h2>
            <div class="meta">Level ${entry.level ?? ''}</div>
          </header>
          <div class="finding">${renderMarkdown(findingText || '')}</div>
          ${recommendations ? `<ul class="recommendations">${recommendations}</ul>` : ''}
        </section>
      `;
    })
    .join('\n');

  const headerSection = showHeaderFooter
    ? coverSection(project, displayDate, tocHtml)
    : simpleTocSection(project, displayDate, tocHtml);
  const footerText = showHeaderFooter
    ? formatFooterText(snapshot.config?.footerText, project, displayDate, now, locale)
    : '';

  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <base href="${baseHref}" />
    <title>${project.meta?.company || 'AutoBericht'} Report</title>
    <style>
      @page { margin: ${margin}mm; }
      body { font-family: Arial, sans-serif; margin: 0; color: #222; }
      header.report-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20mm; }
      header.report-header .title-block { flex: 1; text-align: center; }
      header.report-header h1 { margin: 0; font-size: 24px; }
      header.report-header p { margin: 2px 0 0; color: #555; }
      header.report-header img { height: 40px; object-fit: contain; }
      section.cover-page { page-break-after: always; }
      section.cover-page .report-header { margin-bottom: 10mm; }
      section.toc { margin-bottom: 20mm; }
      section.toc h2 { font-size: 18px; margin-bottom: 8px; }
      section.toc ul { list-style: none; padding: 0; margin: 0; }
      section.toc li { display: flex; gap: 8px; margin-bottom: 4px; font-size: 14px; }
      section.toc .toc-index { font-weight: bold; color: #555; min-width: 1.5em; }
      section.toc .toc-title { flex: 1; }
      section.toc.toc-simple { margin-bottom: 10mm; }
      section.toc.toc-simple h2 { font-size: 20px; margin-bottom: 10px; }
      section.chapter { page-break-inside: avoid; margin-bottom: 15mm; }
      section.chapter h2 { margin: 0 0 6px; font-size: 18px; }
      section.chapter .meta { font-size: 12px; color: #666; margin-bottom: 8px; }
      section.chapter .finding { font-size: 14px; }
      ul.recommendations { margin: 10px 0 0 18px; }
      ul.recommendations li { margin-bottom: 4px; }
      .page-footer {
        position: fixed;
        bottom: ${Math.max(margin - 5, 5)}mm;
        left: ${margin}mm;
        right: ${margin}mm;
        font-size: 11px;
        color: #777;
        text-align: right;
      }
    </style>
    <script src="libs/pagedjs/paged.polyfill.js"></script>
    <script>
      document.addEventListener('DOMContentLoaded', () => {
        if (window.PagedPolyfill && window.PagedPolyfill.preview) {
          window.PagedPolyfill.preview().then(() => {
            setTimeout(() => window.print(), 300);
          });
        } else {
          setTimeout(() => window.print(), 300);
        }
      });
    </script>
  </head>
  <body>
    ${headerSection}
    ${chaptersHtml || '<p>No findings selected for report.</p>'}
    ${showHeaderFooter ? `<div class="page-footer">${footerText}</div>` : ''}
  </body>
</html>`;
}

function collectIncludedRows(project) {
  if (!project) return [];
  const entries = [];
  (project.chapters || []).forEach((chapter) => {
    (chapter?.rows || []).forEach((row) => {
      const include = row?.workstate?.includeFinding ?? true;
      if (!include) return;
      entries.push(unifiedRowToEntry(row));
    });
  });
  return entries;
}

function unifiedRowToEntry(row) {
  const ws = row.workstate || {};
  const levels = row.master?.levels || {};
  return {
    id: row.id,
    title: row.titleOverride || row.id,
    level: ws.selectedLevel ?? '',
    finding: {
      masterText: row.master?.finding || '',
      adjustedText: ws.findingOverride || '',
      useAdjusted: Boolean(ws.useFindingOverride),
    },
    recommendations: {
      '1': {
        master: levels['1'] || '',
        adjusted: ws.levelOverrides?.['1'] || '',
        useAdjusted: Boolean(ws.useLevelOverride?.['1']),
      },
      '2': {
        master: levels['2'] || '',
        adjusted: ws.levelOverrides?.['2'] || '',
        useAdjusted: Boolean(ws.useLevelOverride?.['2']),
      },
      '3': {
        master: levels['3'] || '',
        adjusted: ws.levelOverrides?.['3'] || '',
        useAdjusted: Boolean(ws.useLevelOverride?.['3']),
      },
      '4': {
        master: levels['4'] || '',
        adjusted: ws.levelOverrides?.['4'] || '',
        useAdjusted: Boolean(ws.useLevelOverride?.['4']),
      },
    },
  };
}

function renderLogo(src, alt) {
  if (!src) return '<div style="width:60px;height:40px"></div>';
  return `<img src="${src}" alt="${alt || ''} logo" />`;
}

function coverSection(project, displayDate, tocHtml) {
  return `
    <section class="cover-page">
      <header class="report-header">
        ${renderLogo(project.branding?.left, 'left')}
        <div class="title-block">
          <h1>${project.meta?.company || 'Report'}</h1>
          <p>${displayDate}</p>
        </div>
        ${renderLogo(project.branding?.right, 'right')}
      </header>
      <section class="toc">
        <h2>Table of Contents</h2>
        <ul>${tocHtml || '<li>No chapters selected.</li>'}</ul>
      </section>
    </section>
  `;
}

function simpleTocSection(project, displayDate, tocHtml) {
  return `
    <section class="toc toc-simple">
      <h2>${project.meta?.company || 'Report'} — ${displayDate}</h2>
      <ul>${tocHtml || '<li>No chapters selected.</li>'}</ul>
    </section>
  `;
}

function normalizeMargin(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 5 || numeric > 40) {
    return 20;
  }
  return numeric;
}

function formatLocaleDate(date, locale) {
  try {
    return new Intl.DateTimeFormat(locale || undefined, {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    }).format(date);
  } catch (error) {
    return date.toISOString().slice(0, 10);
  }
}

function formatFooterText(template, project, displayDate, dateObj, locale) {
  const company = project.meta?.company || 'Report';
  const time = formatLocaleTime(dateObj, locale);
  const fallback = `${company} — ${displayDate}`;
  if (!template) return fallback;
  return template
    .replace(/[\u0000-\u001f]/g, '')
    .replace(/{company}/gi, company)
    .replace(/{date}/gi, displayDate)
    .replace(/{time}/gi, time);
}
function formatLocaleTime(date, locale) {
  try {
    return new Intl.DateTimeFormat(locale || undefined, {
      hour: '2-digit',
      minute: '2-digit',
    }).format(date);
  } catch (error) {
    return date.toTimeString().slice(0, 5);
  }
}
