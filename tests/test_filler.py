from __future__ import annotations
import copy
import importlib.util
import json
import re
import shutil
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sp = importlib.util.spec_from_file_location('filler', ROOT / 'scripts/fill_documents.py')
filler = importlib.util.module_from_spec(sp)
sp.loader.exec_module(filler)
W = filler.W


def synthetic_case():
    data = {
        'court_name': 'Учебный районный суд города Тестова',
        'court_address': 'г. Тестов, ул. Учебная, 1',
        'court_instrumental': 'Учебным районным судом города Тестова',
        'court_genitive': 'Учебного районного суда города Тестова',
        'case_number': 'ТЕСТ-0000-26-2/0001',
        'decision_date': '2026-09-01', 'awareness_date': '2026-09-15',
        'filing_date': '2026-09-22',
        'plaintiff_name': 'Примеров Пример Примерович',
        'plaintiff_genitive': 'Примерова Примера Примеровича',
        'plaintiff_iin': '000000000002', 'plaintiff_address': 'г. Тестов, ул. Примерная, 2',
        'defendant_name': 'Тестов Тест Тестович',
        'defendant_dative': 'Тестову Тесту Тестовичу',
        'defendant_iin': '000000000001', 'defendant_address': 'г. Тестов, ул. Примерная, 3',
        'defendant_phone': '+7 000 000 00 00', 'defendant_email': 'example@example.invalid',
        'amount': '150 000', 'amount_words': 'сто пятьдесят тысяч',
    }
    facts = {f: 'confirmed' for f in filler.FACTS}
    sources = {k: 'СИНТЕТИЧЕСКИЙ ТЕСТ. Не относится к реальному делу.' for k in data}
    sources.update({'facts.' + k: 'СИНТЕТИЧЕСКИЙ ТЕСТ. Не подтверждение по делу.' for k in facts})
    return {'synthetic': True, 'act_type': 'simplified_written_decision', 'data': data,
            'facts': facts, 'sources': sources, 'evidence_attached': True,
            'decision_copy_attached': True, 'uncertain_fields': [], 'conflicted_fields': []}


def xml_text(path):
    with ZipFile(path) as z:
        return z.read('word/document.xml').decode('utf-8')


def plain(path):
    r = ET.fromstring(xml_text(path))
    return '\n'.join(''.join(t.text or '' for t in p.iter(W+'t')) for p in r.iter(W+'p'))


def format_skeleton(xml):
    r = ET.fromstring(xml)
    for t in r.iter(W+'t'):
        t.text = ''
        t.attrib.pop('{http://www.w3.org/XML/1998/namespace}space', None)
    return ET.tostring(r)


class FillerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.out = Path(self.temp.name)/'out'
        self.case = synthetic_case()

    def generate(self, case=None):
        report = filler.generate_pair(self.case if case is None else case, self.out)
        return report, [self.out/i['file'] for i in report['files']]

    def test_01_pair(self):
        report, files = self.generate()
        self.assertEqual(len(files), 2)
        self.assertEqual(report['status'], 'FILLED_REQUIRES_FINAL_REVIEW')
        self.assertFalse(report['legal_correctness_certified'])
        self.assertTrue(all(f.exists() for f in files))

    def test_02_partial_input_still_pair(self):
        r, fs = self.generate({})
        self.assertEqual(len(fs), 2)
        self.assertTrue(all(f.name.startswith('ЧЕРНОВИК_') for f in fs))

    def test_03_leading_zeros(self):
        _, fs = self.generate()
        for f in fs:
            self.assertIn('000000000001', plain(f))

    def test_04_numeric_iin_not_coerced(self):
        self.case['data']['defendant_iin'] = 123456789012
        r, _ = self.generate()
        self.assertIn('INVALID_TYPE:defendant_iin', r['notes'])

    def test_05_bad_date(self):
        self.case['data']['decision_date'] = '2026-02-30'
        r, fs = self.generate()
        self.assertIn('INVALID_DATE:decision_date', r['notes'])
        for f in fs:
            self.assertNotIn('2025', plain(f))

    def test_06_no_old_city_or_year(self):
        _, fs = self.generate()
        for f in fs:
            self.assertNotIn('Алматы', plain(f))
            self.assertNotIn('2025', plain(f))

    def test_07_all_other_package_parts_byte_identical(self):
        r, fs = self.generate()
        for spec, out in zip(json.loads((ROOT/'field_map.json').read_text())['documents'], fs):
            with ZipFile(ROOT/spec['file']) as a, ZipFile(out) as b:
                self.assertEqual(a.namelist(), b.namelist())
                for name in a.namelist():
                    if name != 'word/document.xml':
                        self.assertEqual(a.read(name), b.read(name), name)

    def test_08_all_formatting_identical_when_evidence_present(self):
        _, fs = self.generate()
        for spec, f in zip(json.loads((ROOT/'field_map.json').read_text())['documents'], fs):
            self.assertEqual(format_skeleton(xml_text(ROOT/spec['file'])), format_skeleton(xml_text(f)))

    def test_09_highlights_preserved(self):
        _, fs = self.generate()
        for spec, f in zip(json.loads((ROOT/'field_map.json').read_text())['documents'], fs):
            a = ET.fromstring(xml_text(ROOT/spec['file']))
            b = ET.fromstring(xml_text(f))
            self.assertEqual([ET.tostring(x) for x in a.iter(W+'highlight')], [ET.tostring(x) for x in b.iter(W+'highlight')])

    def test_10_header_table_preserved(self):
        _, fs = self.generate()
        for f in fs:
            if 'Заявление' in f.name:
                self.assertEqual(len(list(ET.fromstring(xml_text(f)).iter(W+'tbl'))), 1)
                self.assertIn('Тестов Тест Тестович', plain(f))

    def test_11_repeated_case_numbers(self):
        _, fs = self.generate()
        self.assertEqual(sum(plain(f).count(self.case['data']['case_number']) for f in fs), 5)

    def test_12_xml_escaping(self):
        self.case['data']['plaintiff_name'] = 'A & B <Тест>'
        _, fs = self.generate()
        self.assertTrue(all('A & B <Тест>' in plain(f) for f in fs))

    def test_13_kazakh_text(self):
        self.case['data']['defendant_name'] = 'Әбдіқадыр Қаныш Өмірұлы'
        _, fs = self.generate()
        self.assertTrue(all('Әбдіқадыр Қаныш Өмірұлы' in plain(f) for f in fs))

    def test_14_unknown_notification_marks_draft(self):
        self.case['facts']['not_properly_notified'] = 'unknown'
        r, fs = self.generate()
        self.assertIn('FACT_UNCONFIRMED:not_properly_notified', r['notes'])
        self.assertTrue(all(f.name.startswith('ЧЕРНОВИК_') for f in fs))

    def test_15_contradiction_not_silently_resolved(self):
        self.case['facts']['learned_from_bailiff'] = 'contradicted'
        r, _ = self.generate()
        self.assertIn('FACT_CONTRADICTED:learned_from_bailiff', r['notes'])

    def test_16_incompatible_act_still_two_drafts(self):
        self.case['act_type'] = 'notarial_inscription'
        r, fs = self.generate()
        self.assertEqual(len(fs), 2)
        self.assertIn('ACT_TYPE_NOT_CONFIRMED', r['notes'])

    def test_17_provenance_missing(self):
        self.case['sources'] = {}
        r, _ = self.generate()
        self.assertIn('SOURCE_NOT_RECORDED:defendant_iin', r['notes'])

    def test_18_hash_tamper(self):
        root = Path(self.temp.name)/'package'
        shutil.copytree(ROOT, root, ignore=shutil.ignore_patterns('__pycache__'))
        with (root/'templates/motion_restore_term.docx').open('ab') as f:
            f.write(b'CHANGED')
        with self.assertRaisesRegex(ValueError, 'hash mismatch'):
            filler.generate_pair(self.case, self.out, root)
        self.assertFalse(self.out.exists())

    def test_19_evidence_present_removes_only_hint(self):
        _, fs = self.generate()
        motion = next(f for f in fs if 'Ходатайство' in f.name)
        self.assertNotIn('если нету', plain(motion))
        self.assertIn('Доказательства уважительности причин', plain(motion))

    def test_20_evidence_absent_keeps_empty_paragraph(self):
        self.case['evidence_attached'] = False
        _, fs = self.generate()
        motion = next(f for f in fs if 'Ходатайство' in f.name)
        r = ET.fromstring(xml_text(motion))
        ps = list(r.iter(W+'p'))
        self.assertEqual(len(ps), 42)
        self.assertFalse(list(ps[37].iter(W+'numPr')))
        self.assertNotIn('Доказательства уважительности причин', plain(motion))

    def test_21_evidence_unknown_preserves_reference_hint(self):
        self.case['evidence_attached'] = None
        r, fs = self.generate()
        motion = next(f for f in fs if 'Ходатайство' in f.name)
        self.assertIn(filler.GUIDANCE, plain(motion))
        self.assertIn('EVIDENCE_ATTACHMENT_UNKNOWN:template_guidance_preserved', r['notes'])

    def test_22_conflict_number_not_propagated(self):
        self.case['conflicted_fields'] = ['case_number']
        r, fs = self.generate()
        self.assertTrue(all(self.case['data']['case_number'] not in plain(f) for f in fs))

    def test_23_unknown_base_invalidates_inflection(self):
        self.case['uncertain_fields'] = ['court_name']
        r, fs = self.generate()
        self.assertTrue(all('города Тестова' not in plain(f) for f in fs))
        self.assertIn('MISSING_BASE_FIELD:court_instrumental', r['notes'])

    def test_24_no_filing_date_invention(self):
        self.case['data'].pop('filing_date')
        r, fs = self.generate()
        motion = next(f for f in fs if 'Ходатайство' in f.name)
        self.assertIn('«___» __________20__', plain(motion))

    def test_25_invalid_control_character_not_written(self):
        self.case['data']['defendant_name'] = 'Test\x00Invalid'
        r, fs = self.generate()
        self.assertIn('INVALID_TEXT:defendant_name', r['notes'])
        for f in fs:
            ET.fromstring(xml_text(f))

    def test_26_nonempty_output_blocked(self):
        self.out.mkdir()
        (self.out/'old.txt').write_text('existing case')
        with self.assertRaisesRegex(ValueError, 'must be empty'):
            filler.generate_pair(self.case, self.out)

    def test_27_filenames_no_client_identifiers(self):
        _, fs = self.generate()
        for f in fs:
            self.assertNotIn(self.case['data']['defendant_iin'], f.name)
            self.assertNotIn('Тестов', f.name)

    def test_28_bin_not_written_under_iin(self):
        self.case['data'].pop('plaintiff_iin')
        self.case['data']['plaintiff_bin'] = '000000000003'
        r, fs = self.generate()
        self.assertIn('TEMPLATE_IIN_LABEL_NOT_BIN:plaintiff_iin', r['notes'])
        self.assertTrue(all('000000000003' not in plain(f) for f in fs))

    def test_29_long_field_keeps_font_properties(self):
        self.case['data']['defendant_address'] = 'Очень длинный адрес ' * 30
        _, fs = self.generate()
        for spec, f in zip(json.loads((ROOT/'field_map.json').read_text())['documents'], fs):
            self.assertEqual(format_skeleton(xml_text(ROOT/spec['file'])), format_skeleton(xml_text(f)))

    def test_30_date_order_conflict(self):
        self.case['data']['awareness_date'] = '2026-08-01'
        r, _ = self.generate()
        self.assertIn('DATE_ORDER_CONFLICT:awareness_before_decision', r['notes'])

    def test_31_malformed_input_rejected(self):
        with self.assertRaises(ValueError):
            filler.generate_pair({'data': []}, self.out)

    def test_32_legal_paragraph_unchanged(self):
        _, fs = self.generate()
        a = ET.fromstring(xml_text(ROOT/'templates/application_cancel_decision.docx'))
        original = list(a.iter(W+'p'))
        f = next(f for f in fs if 'Заявление' in f.name)
        filled = list(ET.fromstring(xml_text(f)).iter(W+'p'))
        for pi in (19,20,21,22,23,25,27,35):
            self.assertEqual(ET.tostring(original[pi]), ET.tostring(filled[pi]))

if __name__ == '__main__':
    unittest.main()
