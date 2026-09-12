#!/usr/bin/env python3
"""Render included draft rows verbatim; emit review diagnostics outside the preview.

No prose generation, style score, photo rendering, or app-export equivalence claim.
"""
import argparse
import json
import re
import sys
from pathlib import Path
import sidecar_tool as sc

LABELS = {
    'fr': ('Rapport', 'Constat', 'Recommandation'),
    'de': ('Bericht', 'Feststellung', 'Empfehlung'),
    'it': ('Rapporto', 'Constatazione', 'Raccomandazione'),
    'en': ('Report', 'Finding', 'Recommendation'),
}
# Review candidates, not a universal list of forbidden words. Never rewrite text
# automatically: explicit third-person reporting or a literal quotation can be valid.
PATTERNS = {
    'working_annotation': r'rep[eè]re\s+de\s+revue|source\s+de\s+la\s+dict[eé]e|review\s+(?:note|marker)|Pr[uü]fvermerk|nota\s+di\s+revisione',
    'processing_narration': r'selon\s+la\s+dict[eé]e|[eé]changes\s+dict[eé]s|according\s+to\s+the\s+transcript|laut\s+(?:dem\s+)?Transkript|secondo\s+la\s+trascrizione',
    'outside_narrator': r'le\s+consultant\s+(?:recommande|d[eé]crit|rel[eè]ve)|the\s+consultant\s+(?:recommends|describes)|der\s+Berater\s+empfiehlt|il\s+consulente\s+raccomanda',
}


def review_candidates(text, chapter_id, row_id, field):
    return [dict(chapterId=chapter_id, rowId=row_id, field=field,
                 kind=kind, excerpt=match.group(0))
            for kind, pattern in PATTERNS.items()
            for match in re.finditer(pattern, text, re.IGNORECASE)]


def heading(value, locale):
    if isinstance(value, dict):
        value = value.get(locale) or value.get(locale.split('-')[0]) or next(iter(value.values()), '')
    return ' '.join(str(value or '').split())


def ordered_rows(chapter):
    rows = chapter['rows']
    if str(chapter['id']) not in ('0', '4.8'):
        return rows
    by_id = {str(r.get('id')): r for r in rows if r.get('kind') != 'section'}
    result = []
    seen = set()
    for rid in chapter.get('meta', {}).get('order', []):
        rid = str(rid)
        if rid in by_id and rid not in seen:
            result.append(by_id[rid])
            seen.add(rid)
    result.extend(r for r in rows if str(r.get('id')) not in seen)
    return result


def render(doc):
    project = sc.project(doc)
    sc.row_map(doc)  # Reject ambiguous identities before writing a preview.
    locale = str(project.get('meta', {}).get('locale', ''))
    language = locale.split('-')[0].lower()
    if language not in LABELS:
        raise ValueError('Unsupported/missing report locale; render with the author\'s headings explicitly')
    title, finding_label, recommendation_label = LABELS[language]
    parts = ['# ' + title, '']
    report = {'rowsRendered': [], 'reviewCandidates': [], 'chapterTextNotRendered': [],
              'validationLevel': 'Row preview and review candidates only; semantic/style review required',
              'photosRendered': False, 'appExportTested': False}
    for chapter in project['chapters']:
        cid = str(chapter['id'])
        content = []
        for field in ('positivesText', 'frontMatterText'):
            text = chapter.get('meta', {}).get(field, '')
            if text:
                report['chapterTextNotRendered'].append({'chapterId': cid, 'field': field})
                report['reviewCandidates'].extend(review_candidates(str(text), cid, None, field))
        for row in ordered_rows(chapter):
            if row.get('kind') == 'section':
                continue
            sc.check_ws(row)
            ws = row.get('workstate', {})
            if ws.get('includeFinding') is not True:
                continue
            rid = str(row['id'])
            # Do not silently replace absent case text with a generic library finding.
            if not isinstance(ws.get('findingText'), str):
                raise ValueError(f'Missing case findingText: {cid}/{rid}')
            finding = ws['findingText']
            recommendation = ''
            if ws.get('includeRecommendation', True):
                if not isinstance(ws.get('recommendationText'), str):
                    raise ValueError(f'Missing case recommendationText: {cid}/{rid}')
                recommendation = ws['recommendationText']
            if not finding.strip() and not recommendation.strip():
                raise ValueError(f'Included row has no report text: {cid}/{rid}')
            label = heading(row.get('titleOverride') or row.get('sectionLabel') or row.get('tag'), locale)
            content.extend(['### ' + rid + (' — ' + label if label else ''), ''])
            for field, label, text in [('findingText', finding_label, finding),
                                       ('recommendationText', recommendation_label, recommendation)]:
                if text.strip():
                    content.extend(['**' + label + '**', '', text, ''])
                    report['reviewCandidates'].extend(review_candidates(text, cid, rid, field))
            report['rowsRendered'].append({'chapterId': cid, 'rowId': rid})
        if content:
            label = heading(chapter.get('title'), locale)
            parts.extend(['## ' + cid + (' — ' + label if label else ''), ''] + content)
    return '\n'.join(parts), report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('sidecar')
    parser.add_argument('preview')
    args = parser.parse_args()
    doc, _ = sc.read(args.sidecar)
    text, report = render(doc)
    # Copy-only: never overwrite the Sidecar, an earlier preview or another file.
    with Path(args.preview).open('x', encoding='utf-8', newline='\n') as output:
        output.write(text)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, TypeError, KeyError) as error:
        print('ERROR: ' + str(error), file=sys.stderr)
        sys.exit(1)
