#!/usr/bin/env python3
"""Fill the two approved DOCX files without rebuilding their formatting.

Python >= 3.10, standard library only. No network, AI API, Word, or pip required.
Extraction from the employee's PDF/text is performed by the host LLM.
The date-role helper is conservative and only handles explicit numeric anchors.
This is a template-filling utility, not a legal-admissibility assessment.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TEXT_RE = re.compile(r'(<w:t(?=[\s>])[^>]*>)(.*?)(</w:t>)', re.S)
MONTHS = ('', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
          'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря')
FIELDS = (
    'court_name', 'court_address', 'court_instrumental', 'court_genitive',
    'case_number', 'decision_date', 'legal_force_date', 'awareness_date', 'filing_date',
    'plaintiff_name', 'plaintiff_genitive', 'plaintiff_iin', 'plaintiff_bin',
    'plaintiff_address', 'defendant_name', 'defendant_dative', 'defendant_iin',
    'defendant_address', 'defendant_phone', 'defendant_email', 'amount', 'amount_words',
)
FACTS = (
    'not_properly_notified', 'did_not_receive_claim',
    'could_not_present_objections', 'learned_from_bailiff',
    'will_provide_relevant_evidence', 'missed_deadline',
)
ALLOWED_FILES = {
    'templates/motion_restore_term.docx': '02_Ходатайство_о_восстановлении_срока.docx',
    'templates/application_cancel_decision.docx': '01_Заявление_об_отмене_решения.docx',
}
GUIDANCE = '(если есть, если нету убрать)с'


def parse_date(value: str) -> date:
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
        raise ValueError('Expected YYYY-MM-DD')
    return date.fromisoformat(value)



DATE_TOKEN = r'(?<![0-9])(?:[0-9]{2}\.[0-9]{2}\.[0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})(?![0-9])'
FORCE_LABEL = r'Дата\s+вступления\s+судебного\s+акта\s+в\s+законную\s+силу'
DECISION_LABEL = r'Дата\s+вынесения\s+(?:решения|судебного\s+акта)'
DECISION_CONTEXT = r'рассмотрев\s+гражданское\s+дело'
DATE_BOUNDARY = r'(?:' + FORCE_LABEL + '|' + DECISION_LABEL + r'|Дата\s+выдачи|вступил[оа]?\s+в\s+законную\s+силу|' + DECISION_CONTEXT + ')'


def extract_date_roles(text: str) -> tuple[dict[str, str], list[str]]:
    """Read only explicit date roles; never choose the first/latest page date.

    This is not PDF/OCR or a universal judicial-act parser. The host supplies
    source_text verbatim from the chosen act, never a whole bundle of cases.
    Ambiguous dates are blocked rather than guessed. Other date spellings must
    be read by the host LLM with source provenance.
    """
    if not isinstance(text, str) or len(text) > 2_000_000:
        raise ValueError('source_text must be a string of at most 2000000 characters')
    candidates = {'decision_date': [], 'legal_force_date': []}
    for key, label in (('decision_date', DECISION_LABEL), ('legal_force_date', FORCE_LABEL)):
        for m in re.finditer(label + r'\s*[:\-–]?\s*(' + DATE_TOKEN + ')', text, re.I):
            candidates[key].append(m.group(1))
    for m in re.finditer(DECISION_CONTEXT, text, re.I):
        window = re.split(DATE_BOUNDARY, text[m.end():m.end()+1200], maxsplit=1, flags=re.I)[0]
        candidates['decision_date'].extend(re.findall(DATE_TOKEN, window))
    found, blocked = {}, []
    for key, items in candidates.items():
        normalized = set()
        invalid = False
        for value in items:
            try:
                if '.' in value:
                    day, month, year = value.split('.')
                    value = f'{year}-{month}-{day}'
                normalized.add(parse_date(value).isoformat())
            except ValueError:
                invalid = True
        if invalid or len(normalized) > 1:
            blocked.append(key)
        elif len(normalized) == 1:
            found[key] = normalized.pop()
    return found, blocked


def prepare_values(case: dict) -> tuple[dict[str, str], list[str]]:
    """Normalize types without guessing absent values, dates, or identifiers."""
    if not isinstance(case, dict):
        raise ValueError('case.json must contain a JSON object')
    data = case.get('data', {})
    if not isinstance(data, dict):
        raise ValueError('data must be an object')
    data = dict(data)
    sources = case.get('sources', {})
    if not isinstance(sources, dict):
        raise ValueError('sources must be an object')
    sources = dict(sources)
    notes: list[str] = []
    blocked = set()
    for list_key in ('uncertain_fields', 'conflicted_fields'):
        items = case.get(list_key, [])
        if not isinstance(items, list) or any(not isinstance(x, str) for x in items):
            raise ValueError('uncertain_fields/conflicted_fields must be arrays of strings')
        blocked.update(items)
    if 'source_text' in case:
        detected, ambiguous = extract_date_roles(case['source_text'])
        blocked.update(ambiguous)
        for key in ambiguous:
            notes.append('AMBIGUOUS_DATE_ROLE:' + key)
        for key, value in detected.items():
            if data.get(key) not in (None, '', value):
                blocked.add(key)
                notes.append('SOURCE_DATE_CONFLICT:' + key)
            else:
                data[key] = value
                sources.setdefault(key, 'source_text:explicit_date_role:' + key)
    for key in sorted(blocked):
        notes.append('UNRESOLVED_FIELD:' + key)
    values: dict[str, str] = {}
    for key in FIELDS:
        v = data.get(key)
        if v is None or v == '' or key in blocked:
            continue
        if not isinstance(v, str):
            notes.append('INVALID_TYPE:' + key)
            continue
        v = v.strip()
        if not v:
            continue
        if len(v) > 2000 or re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]', v):
            notes.append('INVALID_TEXT:' + key)
            continue
        # Multiline PDF fields become a single text field; template paragraph
        # breaks are never changed by the contents of a client field.
        v = re.sub(r'\s+', ' ', v)
        if key.endswith(('_iin', '_bin')) and not re.fullmatch(r'[0-9]{12}', v):
            notes.append('INVALID_IDENTIFIER:' + key)
            continue
        if key == 'case_number':
            v = re.sub(r'^(?:№\s*)+', '', v)
            if not v:
                notes.append('INVALID_TEXT:case_number')
                continue
        if key.endswith('_date'):
            try:
                parse_date(v)
            except ValueError:
                notes.append('INVALID_DATE:' + key)
                continue
        values[key] = v
    for key in values:
        if not isinstance(sources.get(key), str) or not sources[key].strip():
            notes.append('SOURCE_NOT_RECORDED:' + key)
    # Grammar variants are supplied by GPT, derived only from the same source
    # name. If missing, leave the corresponding slot blank, never substitute a
    # different court/person. Do not derive them from conflicting base values.
    dependencies = {
        'court_instrumental': 'court_name', 'court_genitive': 'court_name',
        'plaintiff_genitive': 'plaintiff_name', 'defendant_dative': 'defendant_name',
        'amount_words': 'amount',
    }
    for key, base in dependencies.items():
        if key in values and base not in values:
            values.pop(key)
            notes.append('MISSING_BASE_FIELD:' + key)
    if values.get('plaintiff_bin') and values.get('plaintiff_iin'):
        notes.append('CONFLICT:plaintiff_identifier_type')
    elif values.get('plaintiff_bin'):
        values['plaintiff_identifier'] = values['plaintiff_bin']
        values['plaintiff_identifier_label'] = 'БИН'
    elif values.get('plaintiff_iin'):
        values['plaintiff_identifier'] = values['plaintiff_iin']
        values['plaintiff_identifier_label'] = 'ИИН'
    optional = {'plaintiff_bin', 'legal_force_date', 'awareness_date'}
    if values.get('plaintiff_bin'):
        optional.add('plaintiff_iin')
    for key in FIELDS:
        if key not in values and key not in optional:
            notes.append('MISSING:' + key)
    for prefix in ('decision', 'awareness', 'filing'):
        key = prefix + '_date'
        if key in values:
            d = parse_date(values[key])
            full = f'{d.day:02d} {MONTHS[d.month]} {d.year}'
            values[prefix + '_full'] = full
            values[prefix + '_quoted'] = f'«{d.day:02d}» {MONTHS[d.month]} {d.year}'
    if 'decision_full' in values:
        values['decision_from'] = 'от ' + values['decision_full'] + ' года'
        values['decision_full_space'] = values['decision_full'] + ' '
    transforms = {
        'court_header': ('court_name', lambda x: 'В ' + x),
        'case_with_symbol': ('case_number', lambda x: '№' + x),
        'case_after_word': ('case_number', lambda x: 'дело ' + x),
        'plaintiff_claim': ('plaintiff_genitive', lambda x: 'по иску ' + x),
        'defendant_to': ('defendant_dative', lambda x: 'к ' + x),
        'amount_phrase': ('amount', lambda x: 'в размере ' + x),
        'amount_words_brackets': ('amount_words', lambda x: '(' + x + ')'),
    }
    for out_key, (source_key, fn) in transforms.items():
        if source_key in values:
            values[out_key] = fn(values[source_key])
    if case.get('act_type') != 'simplified_written_decision':
        notes.append('ACT_TYPE_NOT_CONFIRMED')
    fact_values = case.get('facts', {})
    if not isinstance(fact_values, dict):
        raise ValueError('facts must be an object')
    for fact in FACTS:
        state = fact_values.get(fact, 'unknown')
        if state != 'confirmed':
            notes.append(('FACT_CONTRADICTED:' if state == 'contradicted' else 'FACT_UNCONFIRMED:') + fact)
        elif not isinstance(sources.get('facts.' + fact), str) or not sources['facts.' + fact].strip():
            notes.append('SOURCE_NOT_RECORDED:facts.' + fact)
    if case.get('decision_copy_attached') is not True:
        notes.append('DECISION_COPY_NOT_CONFIRMED')
    evidence = case.get('evidence_attached')
    if evidence is not None and type(evidence) is not bool:
        raise ValueError('evidence_attached must be true, false, or null')
    if evidence is None:
        notes.append('EVIDENCE_ATTACHMENT_UNKNOWN:template_guidance_preserved')
    if 'decision_date' in values and 'awareness_date' in values:
        if values['awareness_date'] < values['decision_date']:
            notes.append('DATE_ORDER_CONFLICT:awareness_before_decision')
    if 'filing_date' in values and 'awareness_date' in values:
        if values['filing_date'] < values['awareness_date']:
            notes.append('DATE_ORDER_CONFLICT:filing_before_awareness')
    if 'decision_date' in values and 'legal_force_date' in values:
        if values['legal_force_date'] < values['decision_date']:
            notes.append('DATE_ORDER_CONFLICT:legal_force_before_decision')
    return values, sorted(set(notes))


def replace_range(original_nodes: list[ET.Element], node_index: dict[int, int],
                  current: list[str], start: int, end: int, value: str) -> None:
    """Patch a span across multiple w:t nodes, preserving every original run."""
    cursor = 0
    first = True
    for node in original_nodes:
        original = node.text or ''
        lo, hi = max(start, cursor), min(end, cursor + len(original))
        if lo < hi:
            i = node_index[id(node)]
            a, b = lo - cursor, hi - cursor
            current[i] = current[i][:a] + (value if first else '') + current[i][b:]
            first = False
        cursor += len(original)
    if first:
        raise ValueError('Empty or invalid replacement range')


def fill_one(source: Path, dest: Path, spec: dict, values: dict[str, str],
             evidence_attached: bool | None) -> dict:
    """Copy a ZIP package, modifying only allowlisted word/document.xml spans."""
    if hashlib.sha256(source.read_bytes()).hexdigest() != spec['sha256']:
        raise ValueError('Template hash mismatch: ' + source.name)
    with ZipFile(source) as zin:
        xml = zin.read('word/document.xml').decode('utf-8')
        root = ET.fromstring(xml)
        paragraphs = list(root.iter(W + 'p'))
        if len(paragraphs) != spec['paragraph_count']:
            raise ValueError('Unexpected paragraph count')
        nodes = list(root.iter(W + 't'))
        text_matches = list(TEXT_RE.finditer(xml))
        if len(text_matches) != len(nodes):
            raise ValueError('Unsupported XML text-node structure')
        indexes = {id(node): i for i, node in enumerate(nodes)}
        current = [n.text or '' for n in nodes]
        slots = list(spec['slots'])
        remove_evidence_number = False
        if source.name == 'motion_restore_term.docx' and evidence_attached is not None:
            ptext = ''.join(n.text or '' for n in paragraphs[37].iter(W + 't'))
            if evidence_attached:
                a, b = ptext.index(GUIDANCE), len(ptext)
            else:
                a, b = 0, len(ptext)
                remove_evidence_number = True
            slots.append({'paragraph': 37, 'start': a, 'end': b,
                          'original': ptext[a:b], 'field': '__empty', 'missing': ''})
        last_start: dict[int, int] = {}
        changed_slots = 0
        for slot in sorted(slots, key=lambda s: (s['paragraph'], s['start']), reverse=True):
            pi = slot['paragraph']
            ns = list(paragraphs[pi].iter(W + 't'))
            text = ''.join(n.text or '' for n in ns)
            a, b = slot['start'], slot['end']
            if text[a:b] != slot['original'] or a < 0 or b <= a:
                raise ValueError('Invalid field-map anchor')
            if b > last_start.get(pi, len(text)):
                raise ValueError('Overlapping replacement slots')
            last_start[pi] = a
            # Locked slots are fixed manual-entry text, never client values.
            value = slot['missing'] if slot.get('locked') else values.get(slot['field'], slot['missing'])
            if value != slot['original']:
                replace_range(ns, indexes, current, a, b, value)
                changed_slots += 1
        # Do not reserialize the entire XML. All run properties, paragraph
        # properties, tabs, empty paragraphs, fonts, table XML and namespace
        # declarations retain their original bytes.
        changes = []
        for i, (node, m) in enumerate(zip(nodes, text_matches)):
            if current[i] == (node.text or ''):
                continue
            opening = m.group(1)
            if current[i] and (current[i][0].isspace() or current[i][-1].isspace()) and 'xml:space=' not in opening:
                opening = opening[:-1] + ' xml:space="preserve">'
            changes.append((m.start(), m.end(), opening + escape(current[i]) + m.group(3)))
        for a, b, replacement in reversed(changes):
            xml = xml[:a] + replacement + xml[b:]
        if remove_evidence_number:
            # Honor the template's explicit "remove if absent" instruction.
            # Keep the paragraph itself as an empty line; suppress only its
            # automatic list number. This is the sole conditional layout edit.
            pm = list(re.finditer(r'<w:p(?=[\s>])[^>]*>.*?</w:p>', xml, re.S))
            match = pm[37]
            pxml, count = re.subn(r'<w:numPr>.*?</w:numPr>', '', match.group(0), count=1, flags=re.S)
            if count != 1:
                raise ValueError('Evidence paragraph numbering not found')
            xml = xml[:match.start()] + pxml + xml[match.end():]
        final_root = ET.fromstring(xml)
        final_paragraphs = list(final_root.iter(W + 'p'))
        for check in spec.get('manual_fragments', []):
            paragraph = final_paragraphs[check['paragraph']]
            text = ''.join(t.text or '' for t in paragraph.iter(W + 't'))
            start = text.find(check['text'])
            if start < 0 or text.count(check['text']) != 1:
                raise ValueError('Manual notification fragment changed')
            cursor = 0
            for run in paragraph.iter(W + 'r'):
                length = sum(len(t.text or '') for t in run.iter(W + 't'))
                if max(start, cursor) < min(start + len(check['text']), cursor + length):
                    props = run.find(W + 'rPr')
                    italic = None if props is None else props.find(W + 'i')
                    underline = None if props is None else props.find(W + 'u')
                    if (italic is None or italic.get(W + 'val', 'true') not in ('1', 'true', 'on')
                            or underline is None or underline.get(W + 'val', 'single') != 'single'):
                        raise ValueError('Manual fragment must be italic and underlined')
                cursor += length
        with ZipFile(dest, 'w') as zout:
            zout.comment = zin.comment
            for info in zin.infolist():
                content = xml.encode('utf-8') if info.filename == 'word/document.xml' else zin.read(info.filename)
                zout.writestr(info, content)
    return {'file': dest.name, 'changed_slots': changed_slots,
            'source_sha256': spec['sha256'],
            'output_sha256': hashlib.sha256(dest.read_bytes()).hexdigest(),
            'conditional_number_suppressed': remove_evidence_number}


def generate_pair(case: dict, out_dir: Path, root: Path = ROOT) -> dict:
    values, notes = prepare_values(case)
    field_map = json.loads((root / 'field_map.json').read_text(encoding='utf-8'))
    specs = field_map['documents']
    if len(specs) != 2 or {s['file'] for s in specs} != set(ALLOWED_FILES):
        raise ValueError('Exactly the two approved templates are required')
    # Validate both originals before producing any client artifact.
    for spec in specs:
        f = root / spec['file']
        if hashlib.sha256(f.read_bytes()).hexdigest() != spec['sha256']:
            raise ValueError('Template hash mismatch: ' + f.name)
    out_dir = Path(out_dir).resolve()
    if out_dir.exists() and any(out_dir.iterdir()):
        raise ValueError('Output folder must be empty; avoid mixing different clients')
    out_dir.mkdir(parents=True, exist_ok=True)
    draft = bool(notes)
    report = {
        'version': field_map['version'],
        'status': 'DRAFT_INCOMPLETE_OR_UNCONFIRMED' if draft else 'FILLED_REQUIRES_MANUAL_COMPLETION',
        'legal_correctness_certified': False,
        'manual_completion_required': True,
        'manual_fields': ['notification_date', 'notification_method'],
        'date_roles': {key: values.get(key) for key in ('decision_date', 'legal_force_date')},
        'visual_review_required': True,
        'notes': notes, 'files': [],
    }
    with tempfile.TemporaryDirectory(prefix='sudebni_arest_') as temp:
        td = Path(temp)
        for spec in specs:
            name = ('ЧЕРНОВИК_' if draft else '') + ALLOWED_FILES[spec['file']]
            meta = fill_one(root / spec['file'], td / name, spec, values, case.get('evidence_attached'))
            report['files'].append(meta)
        # Commit the pair only after both DOCX files have been successfully made.
        for item in report['files']:
            (out_dir / item['file']).write_bytes((td / item['file']).read_bytes())
        (out_dir / 'validation.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--data', type=Path, required=True, help='Internal case JSON prepared by GPT')
    ap.add_argument('--out', type=Path, required=True, help='Empty output folder, separate for each client')
    args = ap.parse_args()
    try:
        case = json.loads(args.data.read_text(encoding='utf-8-sig'))
        report = generate_pair(case, args.out)
    except (OSError, ValueError, KeyError, TypeError, ET.ParseError) as exc:
        print('GENERATION_FAILED: ' + str(exc), file=sys.stderr)
        return 2
    # No client identifiers or legal claims in stdout.
    print(json.dumps({'status': report['status'], 'docx_count': len(report['files']),
                      'notes_count': len(report['notes'])}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
