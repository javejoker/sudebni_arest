"""Contract tests. Provider names below are simulations, not live LLM tests."""
from __future__ import annotations
import copy
import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile
from test_filler import ROOT, filler, synthetic_case, plain, xml_text, W

FIXTURE = (
    'СИНТЕТИЧЕСКИЙ ТЕКСТ. Дата выдачи исполнительного листа: 05.08.2026\n'
    'Рассмотрев гражданское дело в упрощенном письменном производстве 10.06.2026 г.\n'
    'Дата вступления судебного акта в законную силу: 24.07.2026\n'
)
APP = '«___» _________ 20___ года от частного судебного исполнителя'
MOTION = '«___» ____________ 20___ года'


class V11Tests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.out = self.base/'out'
        self.case = synthetic_case()

    def run_pair(self):
        report = filler.generate_pair(self.case, self.out)
        files = [self.out/x['file'] for x in report['files']]
        return report, files

    def test_33_exact_manual_fragments(self):
        report, files = self.run_pair()
        for path in files:
            expected = APP if 'Заявление' in path.name else MOTION
            self.assertIn(expected, plain(path))
            self.assertNotIn('15 сентября 2026', plain(path))
        self.assertTrue(report['manual_completion_required'])

    def test_34_input_cannot_override_locked_slots(self):
        self.case['data']['manual_notification_date'] = 'ATTACK'
        self.case['data']['notification_method'] = 'по телефону'
        _, files = self.run_pair()
        for path in files:
            self.assertNotIn('ATTACK', plain(path))
            self.assertNotIn('по телефону', plain(path))

    def test_35_manual_fragments_direct_run_format(self):
        _, files = self.run_pair()
        for path in files:
            target = APP if 'Заявление' in path.name else MOTION
            root = ET.fromstring(xml_text(path))
            p = next(p for p in root.iter(W+'p') if target in ''.join(t.text or '' for t in p.iter(W+'t')))
            text = ''.join(t.text or '' for t in p.iter(W+'t'))
            start, cursor, covered = text.index(target), 0, 0
            for run in p.iter(W+'r'):
                n = sum(len(t.text or '') for t in run.iter(W+'t'))
                overlap = max(0, min(start+len(target), cursor+n)-max(start, cursor))
                if overlap:
                    props = run.find(W+'rPr')
                    self.assertIsNotNone(props.find(W+'i'))
                    self.assertEqual(props.find(W+'u').get(W+'val'), 'single')
                    covered += overlap
                cursor += n
            self.assertEqual(covered, len(target))

    def test_36_manual_date_absence_is_intentional(self):
        self.case['data'].pop('awareness_date')
        report, _ = self.run_pair()
        self.assertNotIn('MISSING:awareness_date', report['notes'])
        self.assertEqual(report['status'], 'FILLED_REQUIRES_MANUAL_COMPLETION')

    def test_37_force_date_is_separate_metadata(self):
        self.case['data']['legal_force_date'] = '2026-09-10'
        self.case['sources']['legal_force_date'] = 'SYNTHETIC explicit force date'
        report, files = self.run_pair()
        self.assertEqual(report['date_roles']['legal_force_date'], '2026-09-10')
        for path in files:
            self.assertNotIn('10 сентября 2026', plain(path))
            self.assertIn('01 сентября 2026', plain(path))

    def test_38_force_date_never_fills_missing_decision(self):
        self.case['data'].pop('decision_date')
        self.case['source_text'] = 'Дата вступления судебного акта в законную силу: 24.07.2026'
        report, files = self.run_pair()
        self.assertIsNone(report['date_roles']['decision_date'])
        self.assertIn('MISSING:decision_date', report['notes'])
        self.assertTrue(all('24 июля 2026' not in plain(p) for p in files))

    def test_39_labeled_dates_not_issue_date(self):
        values, blocked = filler.extract_date_roles(FIXTURE)
        self.assertEqual(values, {'decision_date': '2026-06-10', 'legal_force_date': '2026-07-24'})
        self.assertEqual(blocked, [])

    def test_40_source_dates_fill_correct_slots(self):
        self.case['data'].pop('decision_date')
        self.case['source_text'] = FIXTURE
        report, files = self.run_pair()
        self.assertEqual(report['date_roles']['decision_date'], '2026-06-10')
        for path in files:
            self.assertIn('10', plain(path))
            self.assertIn('июня 2026', plain(path))
            self.assertNotIn('24 июля 2026', plain(path))
            self.assertNotIn('05 августа 2026', plain(path))

    def test_41_date_source_conflict_blocks_not_overwrites(self):
        self.case['source_text'] = FIXTURE
        report, files = self.run_pair()
        self.assertIn('SOURCE_DATE_CONFLICT:decision_date', report['notes'])
        self.assertIsNone(report['date_roles']['decision_date'])
        for path in files:
            self.assertNotIn('01 сентября 2026', plain(path))
            self.assertNotIn('10 июня 2026', plain(path))

    def test_42_multiple_context_dates_ambiguous(self):
        values, blocked = filler.extract_date_roles('Рассмотрев гражданское дело 10.06.2026 г., договор от 01.01.2024 г.')
        self.assertNotIn('decision_date', values)
        self.assertIn('decision_date', blocked)

    def test_43_invalid_date_never_accepted(self):
        values, blocked = filler.extract_date_roles('Дата вынесения решения: 31.02.2026')
        self.assertEqual(values, {})
        self.assertEqual(blocked, ['decision_date'])

    def test_44_case_number_prefix_suffix_preserved(self):
        self.case['data']['case_number'] = '№ 3110-26-00-2/2611-1'
        _, files = self.run_pair()
        for path in files:
            self.assertIn('3110-26-00-2/2611-1', plain(path))
            self.assertNotIn('№№', plain(path))
            self.assertNotIn('№№ ', plain(path))

    def test_45_identical_data_four_provider_labels_same_docx_bytes(self):
        hashes = []
        for label in ('ChatGPT', 'Gemini', 'DeepSeek', 'Claude'):
            # This varies provenance ONLY. It does not call or test a provider.
            case = copy.deepcopy(self.case)
            case['sources'] = {k: label + ': SYNTHETIC SOURCE' for k in case['sources']}
            out = self.base/label
            report = filler.generate_pair(case, out)
            hashes.append([hashlib.sha256((out/f['file']).read_bytes()).hexdigest() for f in report['files']])
        self.assertTrue(all(value == hashes[0] for value in hashes))

    def test_46_plain_text_sources_only_no_example_defaults(self):
        self.case = {'source_text': 'Номер дела не читается. Дата решения отсутствует.'}
        report, files = self.run_pair()
        self.assertEqual(len(files), 2)
        self.assertIsNone(report['date_roles']['decision_date'])
        for path in files:
            self.assertNotIn('2611-1', plain(path))
            self.assertNotIn('10 июня 2026', plain(path))

    def test_47_source_text_wrong_type_rejected(self):
        self.case['source_text'] = ['not a string']
        with self.assertRaisesRegex(ValueError, 'source_text'):
            self.run_pair()
        self.assertFalse(self.out.exists())

    def test_48_identifier_type_conflict_blank(self):
        self.case['data']['plaintiff_bin'] = '000000000009'
        self.case['sources']['plaintiff_bin'] = 'SYNTHETIC'
        report, files = self.run_pair()
        self.assertIn('CONFLICT:plaintiff_identifier_type', report['notes'])
        motion = next(p for p in files if 'Ходатайство' in p.name)
        self.assertNotIn('000000000009', plain(motion))
        self.assertNotIn('000000000002', plain(motion))

    def test_49_force_before_decision_flagged(self):
        self.case['data']['legal_force_date'] = '2026-08-31'
        self.case['sources']['legal_force_date'] = 'SYNTHETIC'
        report, _ = self.run_pair()
        self.assertIn('DATE_ORDER_CONFLICT:legal_force_before_decision', report['notes'])

    def test_50_multiple_roles_one_line(self):
        values, blocked = filler.extract_date_roles(FIXTURE.replace('\n', ' '))
        self.assertEqual(values['decision_date'], '2026-06-10')
        self.assertEqual(values['legal_force_date'], '2026-07-24')
        self.assertFalse(blocked)

    def test_51_cli_without_repository_working_directory(self):
        inp = self.base/'case.json'
        inp.write_text(json.dumps(self.case, ensure_ascii=False), encoding='utf-8')
        proc = subprocess.run([sys.executable, str(ROOT/'scripts/fill_documents.py'), '--data', str(inp), '--out', str(self.out)], cwd=self.base, capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertEqual(len(list(self.out.glob('*.docx'))), 2)

    def test_52_duplicate_same_date_not_ambiguous(self):
        values, blocked = filler.extract_date_roles('Дата вынесения решения: 10.06.2026. Рассмотрев гражданское дело 10.06.2026 г.')
        self.assertEqual(values['decision_date'], '2026-06-10')
        self.assertFalse(blocked)

    def test_53_conflicting_force_dates_blocked(self):
        values, blocked = filler.extract_date_roles('Дата вступления судебного акта в законную силу: 24.07.2026\nДата вступления судебного акта в законную силу: 25.07.2026')
        self.assertNotIn('legal_force_date', values)
        self.assertIn('legal_force_date', blocked)

    def test_54_invalid_blocked_list_rejected_cleanly(self):
        self.case['uncertain_fields'] = None
        with self.assertRaisesRegex(ValueError, 'arrays of strings'):
            self.run_pair()


if __name__ == '__main__':
    unittest.main()
