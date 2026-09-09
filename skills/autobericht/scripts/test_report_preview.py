"""Check draft-field fidelity, output separation, selection and copy-only behaviour."""
import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import report_preview as preview
from test_helpers import fixture


class ReportPreview(unittest.TestCase):
    def test_unreviewed_included_text_is_preserved_without_private_context(self):
        source = fixture()
        row = source['report']['project']['chapters'][0]['rows'][0]
        row['workstate'].update(done=False, findingText='Selon nos discussions, les délais ne sont pas suivis.',
                               recommendationText='Nous recommandons de reprendre les actions ouvertes.\n\nPréserver ce détail utile.',
                               includeRecommendation=True)
        row['master']['finding'] = 'Le consultant recommande une formule à ne pas importer.'
        row['customer']['remark'] = 'Repère de revue : conserver dans les données sources.'
        source['photos']['photos']['photos/a.jpg']['notes'] = 'Unrelated photo note'
        snapshot = copy.deepcopy(source)
        text, report = preview.render(source)
        self.assertIn(row['workstate']['findingText'], text)
        self.assertIn(row['workstate']['recommendationText'], text)
        self.assertNotIn(row['master']['finding'], text)
        self.assertNotIn(row['customer']['remark'], text)
        self.assertNotIn('Unrelated photo note', text)
        self.assertEqual(report['reviewCandidates'], [])
        self.assertEqual(source, snapshot)
        self.assertFalse(row['workstate']['done'])

    def test_leakage_is_reported_separately_without_silent_rewriting(self):
        source = fixture()
        ws = source['report']['project']['chapters'][0]['rows'][0]['workstate']
        ws['findingText'] = 'Les échanges dictés montrent un écart.\n\nRepère de revue : 02:00.'
        ws['recommendationText'] = 'Le consultant recommande un contrôle.'
        text, report = preview.render(source)
        self.assertIn(ws['findingText'], text)
        self.assertIn(ws['recommendationText'], text)
        self.assertEqual({x['kind'] for x in report['reviewCandidates']},
                         {'processing_narration', 'working_annotation', 'outside_narrator'})
        self.assertNotIn('reviewCandidates', text)
        self.assertNotIn('validationLevel', text)

    def test_disabled_recommendation_and_excluded_rows_do_not_leak(self):
        source = fixture()
        ws = source['report']['project']['chapters'][0]['rows'][0]['workstate']
        ws.update(includeRecommendation=False, recommendationText='Repère de revue : hidden')
        hidden = source['report']['project']['chapters'][1]['rows'][0]
        hidden['workstate'].update(findingText='Do not include this observation', includeFinding=False)
        text, report = preview.render(source)
        self.assertNotIn('hidden', text)
        self.assertNotIn('Do not include this observation', text)
        self.assertEqual(len(report['rowsRendered']), 1)
        self.assertEqual(report['reviewCandidates'], [])

    def test_observation_order_and_unrendered_chapter_content_are_explicit(self):
        source = fixture()
        chapter = source['report']['project']['chapters'][1]
        first = chapter['rows'][0]
        first['workstate'].update(includeFinding=True, findingText='Premier constat.', includeRecommendation=False)
        second = copy.deepcopy(first)
        second['id'] = '4.8.2'
        second['workstate']['findingText'] = 'Second constat.'
        chapter['rows'].append(second)
        chapter['meta'] = {'order': ['4.8.2', '4.8.1'], 'positivesText': 'Texte de chapitre existant.'}
        text, report = preview.render(source)
        self.assertLess(text.index('Second constat.'), text.index('Premier constat.'))
        self.assertEqual(report['chapterTextNotRendered'], [{'chapterId': '4.8', 'field': 'positivesText'}])

    def test_missing_case_text_fails_instead_of_falling_back_to_seed(self):
        source = fixture()
        del source['report']['project']['chapters'][0]['rows'][0]['workstate']['findingText']
        with self.assertRaisesRegex(ValueError, 'Missing case findingText'):
            preview.render(source)

    def test_cli_writes_preview_and_separate_diagnostics_without_overwrite(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / 'sidecar.json'
            target = Path(temp) / 'report.md'
            source.write_text(json.dumps(fixture()))
            before = source.read_bytes()
            cmd = [sys.executable, str(Path(preview.__file__)), str(source), str(target)]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            self.assertEqual(len(json.loads(result.stdout)['rowsRendered']), 1)
            self.assertNotIn('validationLevel', target.read_text())
            self.assertEqual(source.read_bytes(), before)
            existing = target.read_bytes()
            repeat = subprocess.run(cmd, capture_output=True, text=True)
            self.assertNotEqual(repeat.returncode, 0)
            self.assertEqual(target.read_bytes(), existing)


if __name__ == '__main__':
    unittest.main(verbosity=2)
