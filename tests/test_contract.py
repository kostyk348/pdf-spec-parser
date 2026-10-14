"""
Unit tests for the output contract layer (pdf_spec_parser.contract).

Run from the repository root:

    python -m unittest discover -s tests -v
"""

import csv
import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_spec_parser import contract                       # noqa: E402
from pdf_spec_parser.parser import CATEGORY_VOCAB, FLAG_VOCAB  # noqa: E402


def item(**overrides):
    """A valid parser item; override fields per test."""
    base = {
        'pos': 1,
        'category': 'Коммутационное оборудование',
        'name': 'Реле промежуточное электромагнитное',
        'code': 'РЕК78/4',
        'params': '4 переключающих контакта (4PDT)',
        'qty': 3.0,
        'unit': 'шт',
        'qty_raw': '3',
        'flags': [],
    }
    base.update(overrides)
    return base


class NormalisationTest(unittest.TestCase):

    def test_records_have_exactly_the_contract_fields(self):
        record = contract.normalize_item(item())
        self.assertEqual(tuple(record), contract.FIELDS)

    def test_missing_optional_fields_get_documented_defaults(self):
        record = contract.normalize_item({'pos': 7, 'name': 'Провод'})
        self.assertEqual(record['code'], '')
        self.assertEqual(record['params'], '')
        self.assertEqual(record['qty_raw'], '')
        self.assertEqual(record['flags'], [])

    def test_unparsed_quantity_becomes_null_and_flagged(self):
        record = contract.normalize_item(
            item(qty='комплект', unit=None, qty_raw='комплект', flags=[]))
        self.assertIsNone(record['qty'])
        self.assertIsNone(record['unit'])
        self.assertIn('UNPARSED_QTY', record['flags'])
        self.assertEqual(record['qty_raw'], 'комплект')

    def test_whole_floats_are_normalised_to_int(self):
        self.assertEqual(contract.normalize_item(item(qty=3.0))['qty'], 3)

    def test_fractional_quantity_stays_float(self):
        self.assertEqual(contract.normalize_item(item(qty=2.5))['qty'], 2.5)

    def test_unit_is_dropped_when_quantity_is_null(self):
        record = contract.normalize_item(
            item(qty=None, unit='шт', qty_raw='', flags=['NO_QTY']))
        self.assertIsNone(record['unit'])

    def test_unknown_flags_are_not_propagated(self):
        record = contract.normalize_item(item(flags=['NOT_A_FLAG']))
        self.assertEqual(record['flags'], [])


class ValidationTest(unittest.TestCase):

    def test_valid_records_pass(self):
        records = contract.to_records([item(), item(pos=2, name='Зажим')])
        status, issues = contract.verdict(records)
        self.assertEqual(status, 'PASS')
        self.assertEqual(issues, [])

    def test_empty_name_is_an_error(self):
        status, issues = contract.verdict([contract.normalize_item(item(name=''))])
        self.assertEqual(status, 'FAIL')
        self.assertIn('E_NAME_EMPTY', [i['code'] for i in issues])

    def test_duplicate_pos_is_an_error(self):
        records = contract.to_records([item(), item(pos=1, name='Второй')])
        codes = [i['code'] for i in contract.validate_records(records)]
        self.assertIn('E_POS_DUPLICATE', codes)

    def test_descending_pos_is_an_error(self):
        records = contract.to_records([item(pos=3), item(pos=4), item(pos=2)])
        codes = [i['code'] for i in contract.validate_records(records)]
        self.assertIn('E_POS_ORDER', codes)

    def test_pos_out_of_range_is_an_error(self):
        records = [contract.normalize_item(item(pos=1000))]
        codes = [i['code'] for i in contract.validate_records(records)]
        self.assertIn('E_POS_RANGE', codes)

    def test_unknown_category_is_an_error(self):
        records = [contract.normalize_item(item(category='Что-то новое'))]
        codes = [i['code'] for i in contract.validate_records(records)]
        self.assertIn('E_CATEGORY_UNKNOWN', codes)

    def test_every_category_value_is_in_the_vocabulary(self):
        self.assertIn('Прочее', CATEGORY_VOCAB)
        self.assertEqual(len(CATEGORY_VOCAB), len(set(CATEGORY_VOCAB)))

    def test_unit_without_quantity_is_an_error(self):
        record = contract.normalize_item(item())
        record['qty'] = None            # bypass normalisation on purpose
        record['unit'] = 'шт'
        codes = [i['code'] for i in contract.validate_records([record])]
        self.assertIn('E_UNIT_WITHOUT_QTY', codes)

    def test_quantity_without_unit_is_an_error(self):
        record = contract.normalize_item(item())
        record['unit'] = None
        codes = [i['code'] for i in contract.validate_records([record])]
        self.assertIn('E_QTY_WITHOUT_UNIT', codes)

    def test_missing_field_is_an_error(self):
        record = contract.normalize_item(item())
        del record['qty_raw']
        codes = [i['code'] for i in contract.validate_records([record])]
        self.assertIn('E_SCHEMA_FIELD_MISSING', codes)

    def test_undocumented_extra_field_is_an_error(self):
        record = contract.normalize_item(item())
        record['comment'] = 'не из контракта'
        codes = [i['code'] for i in contract.validate_records([record])]
        self.assertIn('E_SCHEMA_FIELD_UNKNOWN', codes)

    def test_unknown_flag_is_an_error(self):
        record = contract.normalize_item(item())
        record['flags'] = ['MADE_UP']
        codes = [i['code'] for i in contract.validate_records([record])]
        self.assertIn('E_FLAG_UNKNOWN', codes)

    def test_negative_quantity_is_an_error(self):
        record = contract.normalize_item(item(qty=-1))
        codes = [i['code'] for i in contract.validate_records([record])]
        self.assertIn('E_QTY_NEGATIVE', codes)

    def test_unknown_component_is_a_warning_not_a_failure(self):
        records = contract.to_records([item(flags=['UNKNOWN_COMPONENT'])])
        status, issues = contract.verdict(records)
        self.assertEqual(status, 'PASS')
        self.assertIn('W_UNKNOWN_COMPONENT', [i['code'] for i in issues])

    def test_websearch_source_is_reported_as_a_warning(self):
        records = contract.to_records([item(flags=['WEB_INFO'])])
        status, issues = contract.verdict(records)
        self.assertEqual(status, 'PASS')
        self.assertIn('W_NONDETERMINISTIC', [i['code'] for i in issues])

    def test_every_flag_in_the_vocabulary_is_usable(self):
        for flag in FLAG_VOCAB:
            record = contract.normalize_item(item(flags=[flag]))
            self.assertIn(flag, record['flags'])


class WriterTest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='pdf-spec-contract-')
        self.records = contract.to_records([
            item(),
            item(pos=2, name='Комплект крепежа', code='', qty=None, unit=None,
                 qty_raw='комплект', flags=['UNPARSED_QTY']),
        ])

    def test_csv_header_and_row_count(self):
        path = contract.write_csv(self.records, os.path.join(self.tmp, 'a.csv'))
        with open(path, encoding='utf-8-sig', newline='') as fh:
            rows = list(csv.reader(fh))
        self.assertEqual(rows[0], list(contract.FIELDS))
        self.assertEqual(len(rows), 3)

    def test_csv_is_utf8_with_bom_for_excel(self):
        path = contract.write_csv(self.records, os.path.join(self.tmp, 'b.csv'))
        with open(path, 'rb') as fh:
            self.assertTrue(fh.read(3) == b'\xef\xbb\xbf')

    def test_csv_writes_null_quantity_as_empty_cell(self):
        path = contract.write_csv(self.records, os.path.join(self.tmp, 'c.csv'))
        with open(path, encoding='utf-8-sig', newline='') as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(rows[1]['qty'], '')
        self.assertEqual(rows[1]['unit'], '')
        self.assertEqual(rows[1]['qty_raw'], 'комплект')
        self.assertEqual(rows[1]['flags'], 'UNPARSED_QTY')

    def test_csv_joins_multiple_flags_with_semicolon(self):
        records = contract.to_records([
            item(pos=8, qty=None, unit=None, qty_raw='комплект',
                 flags=['UNPARSED_QTY', 'UNKNOWN_COMPONENT'])])
        path = contract.write_csv(records, os.path.join(self.tmp, 'd.csv'))
        with open(path, encoding='utf-8-sig', newline='') as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(rows[0]['flags'], 'UNPARSED_QTY;UNKNOWN_COMPONENT')

    def test_json_document_shape(self):
        source = os.path.join(self.tmp, 'input.pdf')
        with open(source, 'wb') as fh:
            fh.write(b'%PDF-1.4 fake')
        document = contract.build_document(
            self.records, source, {'pages': 1, 'table_page': 1,
                                   'columns': ['item', 'desc', 'qty']}, [])
        path = contract.write_json(document, os.path.join(self.tmp, 'e.json'))
        with open(path, encoding='utf-8') as fh:
            loaded = json.load(fh)
        self.assertEqual(loaded['schema_version'], contract.SCHEMA_VERSION)
        self.assertEqual(loaded['counts']['records'], 2)
        self.assertEqual(loaded['counts']['flags'], {'UNPARSED_QTY': 1})
        self.assertEqual(loaded['source']['file'], 'input.pdf')
        self.assertEqual(len(loaded['source']['sha256']), 64)
        self.assertEqual(loaded['items'][0]['name'],
                         'Реле промежуточное электромагнитное')
        self.assertEqual(loaded['diagnostics']['issues'], [])

    def test_file_sha256_matches_hashlib(self):
        import hashlib
        source = os.path.join(self.tmp, 'input2.pdf')
        payload = b'%PDF-1.4 other'
        with open(source, 'wb') as fh:
            fh.write(payload)
        self.assertEqual(contract.file_sha256(source),
                         hashlib.sha256(payload).hexdigest())


if __name__ == '__main__':
    unittest.main()
