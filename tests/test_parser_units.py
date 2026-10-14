"""
Unit tests for the geometry/knowledge parts of pdf_spec_parser.parser.

These tests do not need a PDF: they build block lists in the coordinate
contract of the decoders, i.e. (page, x_left, y_top, text), and check header
detection, column projection, row bands and the item formatter.

    python -m unittest discover -s tests -v
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_spec_parser import parser as P                     # noqa: E402


class TextUtilitiesTest(unittest.TestCase):

    def test_clean_text_strips_pdf_escapes_and_brackets(self):
        self.assertEqual(P.clean_text('[Реле \\/ РЕК]'), 'Реле / РЕК')

    def test_clean_text_collapses_whitespace(self):
        self.assertEqual(P.clean_text('  Зажим   винтовой  '), 'Зажим винтовой')

    def test_normalize_num_accepts_comma_decimal(self):
        self.assertEqual(P.normalize_num('2,5'), 2.5)

    def test_normalize_num_returns_none_for_text(self):
        self.assertIsNone(P.normalize_num('комплект'))


class HeaderCellTest(unittest.TestCase):

    def test_russian_headers_are_recognised(self):
        cases = {
            'Поз.': 'item', '№': 'item', 'Наименование': 'desc',
            'Кол.': 'qty', 'Кол-во': 'qty', 'Ед. изм.': 'unit_col',
            'Обозначение': 'code', 'Примечание': 'note',
        }
        for text, expected in cases.items():
            self.assertEqual(P._header_column_id(text), expected, text)

    def test_long_sentences_are_not_headers(self):
        self.assertIsNone(P._header_column_id(
            'Общее количество жил в кабеле указано в ведомости'))

    def test_data_like_short_strings_are_not_headers(self):
        for text in ('Зажим винтовой', 'ЗВИ-4', 'шт', '1', 'Датчик'):
            self.assertIsNone(P._header_column_id(text), text)


def blocks_for(columns, rows, header_top=100, pitch=50, page=1):
    """Build (page, x, y, text) blocks for a table given column positions."""
    blocks = [(page, x, header_top, header) for x, header in columns]
    for index, row in enumerate(rows):
        y = header_top + pitch * (index + 1)
        for x, header in columns:
            value = row.get(header)
            if value in (None, ''):
                continue
            blocks.append((page, x, y, str(value)))
    return blocks


class AutoDetectTest(unittest.TestCase):

    def test_columns_are_projected_between_detected_header_positions(self):
        blocks = blocks_for(
            [(50, 'Поз.'), (100, 'Наименование'), (400, 'Кол.')],
            [{'Поз.': 1, 'Наименование': 'Реле', 'Кол.': 3}])
        col_map, page = P.auto_detect_columns(blocks)
        self.assertEqual(page, 1)
        self.assertEqual(col_map['item'], (0, 75))
        self.assertEqual(col_map['desc'], (75, 250))
        self.assertEqual(col_map['qty'], (250, float('inf')))

    def test_column_order_of_the_drawing_wins(self):
        """A sheet with Кол. left of Наименование still maps correctly."""
        blocks = blocks_for(
            [(60, 'Поз.'), (115, 'Кол.'), (170, 'Ед. изм.'), (240, 'Наименование')],
            [{'Поз.': 1, 'Кол.': 20, 'Ед. изм.': 'шт', 'Наименование': 'Зажим'}])
        col_map, _page = P.auto_detect_columns(blocks)
        self.assertLess(col_map['qty'][0], col_map['desc'][0])
        self.assertLess(col_map['unit_col'][0], col_map['desc'][0])
        self.assertEqual(col_map['item'][1], 87)

    def test_header_row_at_the_top_of_the_page_is_found(self):
        """The old bottom-10% rule missed this and fell back to GOST ranges."""
        blocks = [(1, 45, 12, '№'), (1, 200, 12, 'Наименование'),
                  (1, 470, 12, 'Кол-во'),
                  (1, 45, 40, '1'), (1, 200, 40, 'Реле'), (1, 470, 40, '4'),
                  (1, 45, 810, 'Лист 3')]
        col_map, page = P.auto_detect_columns(blocks)
        self.assertIsNotNone(col_map)
        self.assertEqual(page, 1)
        self.assertIn('desc', col_map)

    def test_no_header_row_falls_back_to_fixed_ranges(self):
        blocks = [(1, 450, 3000, 'мусор'), (1, 455, 3010, 'без'),
                  (1, 460, 3020, 'заголовков')]
        col_map, page = P.auto_detect_columns(blocks)
        self.assertIsNone(col_map)
        self.assertEqual(page, 1)

    def test_single_header_cell_is_not_enough(self):
        blocks = [(1, 50, 100, 'Поз.'), (1, 50, 150, '1')]
        col_map, _page = P.auto_detect_columns(blocks)
        self.assertIsNone(col_map)

    def test_scattered_keywords_are_not_one_header_row(self):
        """Candidates on different baselines must not be merged into a row."""
        blocks = [(1, 50, 100, 'Поз.'), (1, 400, 200, 'Кол.'),
                  (1, 800, 400, 'Ед. изм.')]
        col_map, _page = P.auto_detect_columns(blocks)
        self.assertIsNone(col_map)


class RowBandTest(unittest.TestCase):

    def test_dense_sheet_gets_a_small_band(self):
        self.assertAlmostEqual(P._row_band([100, 122, 144, 166]), 13.2, places=5)

    def test_wide_drawing_gets_a_large_band(self):
        self.assertAlmostEqual(P._row_band([3000, 3120, 3240]), 72.0, places=5)

    def test_single_row_falls_back_to_default(self):
        self.assertEqual(P._row_band([100]), P.Y_TOL_DEFAULT)


class ParseTableTest(unittest.TestCase):

    def test_rows_are_built_from_position_numbers(self):
        blocks = blocks_for(
            [(50, 'Поз.'), (100, 'Наименование'), (400, 'Кол.')],
            [{'Поз.': 1, 'Наименование': 'Реле', 'Кол.': 3},
             {'Поз.': 2, 'Наименование': 'Зажим', 'Кол.': 12}])
        col_map, page = P.auto_detect_columns(blocks)
        rows = P.parse_spec_table(blocks, col_map, page)
        self.assertEqual([r['pos'] for r in rows], [1, 2])
        self.assertEqual(rows[0]['name'], 'Реле')
        self.assertEqual(rows[1]['qty_raw'], '12')

    def test_no_position_numbers_raises_no_items_error(self):
        blocks = [(1, 50, 100, 'Поз.'), (1, 100, 100, 'Наименование'),
                  (1, 100, 150, 'Реле')]
        col_map, page = P.auto_detect_columns(blocks)
        with self.assertRaises(P.NoItemsError) as ctx:
            P.parse_spec_table(blocks, col_map, page)
        self.assertIn('column detection failed', str(ctx.exception))

    def test_error_message_carries_the_diagnostics(self):
        blocks = [(1, 50, 100, 'Поз.'), (1, 100, 100, 'Наименование'),
                  (1, 100, 150, 'Реле')]
        col_map, page = P.auto_detect_columns(blocks)
        with self.assertRaises(P.NoItemsError) as ctx:
            P.parse_spec_table(blocks, col_map, page)
        message = str(ctx.exception)
        self.assertIn('table_page', message)
        self.assertIn('col_map', message)

    def test_two_line_name_reads_top_to_bottom(self):
        blocks = [(1, 50, 100, 'Поз.'), (1, 100, 100, 'Наименование'),
                  (1, 400, 100, 'Кол.'),
                  (1, 50, 150, '5'), (1, 100, 150, 'Датчик температуры'),
                  (1, 100, 170, 'ТСМ-100'), (1, 400, 150, '1')]
        col_map, page = P.auto_detect_columns(blocks)
        rows = P.parse_spec_table(blocks, col_map, page)
        self.assertEqual(rows[0]['name'], 'Датчик температуры ТСМ-100')

    def test_dense_rows_do_not_merge(self):
        blocks = [(1, 45, 12, '№'), (1, 200, 12, 'Наименование'),
                  (1, 470, 12, 'Кол-во')]
        for i in range(1, 6):
            y = 12 + 22 * i
            blocks += [(1, 45, y, str(i)), (1, 200, y, 'Реле %d' % i),
                       (1, 470, y, str(i * 2))]
        col_map, page = P.auto_detect_columns(blocks)
        rows = P.parse_spec_table(blocks, col_map, page)
        self.assertEqual(len(rows), 5)
        self.assertEqual([r['name'] for r in rows],
                         ['Реле 1', 'Реле 2', 'Реле 3', 'Реле 4', 'Реле 5'])

    def test_code_and_unit_columns_are_captured(self):
        blocks = blocks_for(
            [(50, 'Поз.'), (100, 'Наименование'), (400, 'Обозначение'),
             (600, 'Кол.'), (700, 'Ед. изм.')],
            [{'Поз.': 1, 'Наименование': 'Реле', 'Обозначение': 'РЕК78/4',
              'Кол.': 3, 'Ед. изм.': 'шт'}])
        col_map, page = P.auto_detect_columns(blocks)
        rows = P.parse_spec_table(blocks, col_map, page)
        self.assertEqual(rows[0]['code'], 'РЕК78/4')
        self.assertEqual(rows[0]['unit_raw'], 'шт')

    def test_stats_are_filled_when_requested(self):
        blocks = blocks_for(
            [(50, 'Поз.'), (100, 'Наименование'), (400, 'Кол.')],
            [{'Поз.': 1, 'Наименование': 'Реле', 'Кол.': 3}])
        col_map, page = P.auto_detect_columns(blocks)
        stats = {}
        P.parse_spec_table(blocks, col_map, page, stats=stats)
        self.assertEqual(stats['table_page'], page)
        self.assertEqual(stats['column_hits']['item'], 1)
        self.assertIn('row_band', stats)

    def test_missing_headers_use_the_fixed_gost_map(self):
        """Fallback path: GOST-style coordinates without any header row."""
        blocks = [(1, 450, 3000, '1'), (1, 700, 3000, 'Реле'),
                  (1, 3900, 3000, '3')]
        rows = P.parse_spec_table(blocks)
        self.assertEqual(rows[0]['pos'], 1)
        self.assertEqual(rows[0]['name'], 'Реле')
        self.assertEqual(rows[0]['qty_raw'], '3')


class ExtractQtyTest(unittest.TestCase):

    def test_quantity_with_unit(self):
        self.assertEqual(P.extract_qty('12 шт'), (12.0, 'шт'))

    def test_quantity_with_comma_decimal(self):
        self.assertEqual(P.extract_qty('1,5 м'), (1.5, 'м'))

    def test_bare_number_defaults_to_pieces(self):
        self.assertEqual(P.extract_qty('250'), (250.0, 'шт'))

    def test_text_quantity_is_returned_as_is(self):
        self.assertEqual(P.extract_qty('комплект'), ('комплект', 'шт'))

    def test_empty_quantity(self):
        self.assertEqual(P.extract_qty(''), ('', 'шт'))


class ClassifyTest(unittest.TestCase):

    def test_known_components(self):
        self.assertEqual(P.classify_component('реле промежуточное'),
                         'Коммутационное оборудование')
        self.assertEqual(P.classify_component('провод пугв'),
                         'Провода и кабели')

    def test_unknown_component_is_other(self):
        self.assertEqual(P.classify_component('датчик температуры'), 'Прочее')

    def test_every_category_is_in_the_vocabulary(self):
        for name in P.CATEGORY_KEYWORDS:
            self.assertIn(P.classify_component(name), P.CATEGORY_VOCAB)


class FormatItemTest(unittest.TestCase):

    def base_row(self, **overrides):
        row = {'pos': 1, 'name': 'Реле промежуточное ИЕК РЕК 78/4',
               'code': 'РЕК78/4', 'qty_raw': '3', 'unit_raw': 'шт'}
        row.update(overrides)
        return row

    def test_known_component_is_not_flagged(self):
        item = P.format_item(self.base_row(), websearch=False)
        self.assertEqual(item['flags'], [])
        self.assertEqual(item['name'], 'Реле промежуточное электромагнитное')
        self.assertEqual(item['qty'], 3.0)
        self.assertEqual(item['unit'], 'шт')
        self.assertEqual(item['code'], 'РЕК78/4')

    def test_unknown_component_is_flagged_and_left_empty(self):
        item = P.format_item(
            self.base_row(name='Датчик температуры ТСМ-100'), websearch=False)
        self.assertEqual(item['flags'], ['UNKNOWN_COMPONENT'])
        self.assertEqual(item['name'], 'Датчик температуры ТСМ-100')
        self.assertEqual(item['params'], '(требуется уточнение)')

    def test_text_quantity_becomes_null_with_flag(self):
        item = P.format_item(self.base_row(qty_raw='комплект', unit_raw=''),
                             websearch=False)
        self.assertIsNone(item['qty'])
        self.assertIsNone(item['unit'])
        self.assertEqual(item['qty_raw'], 'комплект')
        self.assertIn('UNPARSED_QTY', item['flags'])

    def test_empty_quantity_cell_is_flagged(self):
        item = P.format_item(self.base_row(qty_raw='', unit_raw=''),
                             websearch=False)
        self.assertIsNone(item['qty'])
        self.assertIn('NO_QTY', item['flags'])

    def test_centimetres_are_converted_to_metres(self):
        item = P.format_item(
            self.base_row(name='Провод монтажный ПВ3 1х2,5 белый',
                          qty_raw='250', unit_raw='см'), websearch=False)
        self.assertEqual(item['qty'], 2.5)
        self.assertEqual(item['unit'], 'м')
        self.assertIn('QTY_CONVERTED', item['flags'])

    def test_unit_column_wins_over_the_quantity_cell(self):
        item = P.format_item(self.base_row(qty_raw='12', unit_raw='упак'),
                             websearch=False)
        self.assertEqual(item['unit'], 'упак')

    def test_brand_name_is_removed(self):
        item = P.format_item(self.base_row(name='Сальник MG12 Chint'),
                             websearch=False)
        self.assertEqual(item['name'], 'Ввод кабельный (сальник)')

    def test_item_has_exactly_the_contract_fields(self):
        from pdf_spec_parser import contract
        item = P.format_item(self.base_row(), websearch=False)
        self.assertEqual(tuple(item), contract.FIELDS)


class DecodeTest(unittest.TestCase):

    def test_pdfplumber_blocks_are_words_not_characters(self):
        """A per-character stream would break the whole geometry."""
        path = os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), 'examples',
            'template_b_continuation.pdf')
        if not os.path.exists(path):
            self.skipTest('example PDFs not present')
        blocks = P.decode_with_pdfplumber(path)
        texts = [text for _p, _x, _y, text in blocks]
        self.assertGreater(len(blocks), 20)
        # whole words survive, they are not split into single characters
        self.assertIn('Продолжение', texts)
        self.assertGreater(len([t for t in texts if len(t) > 3]), 10)
        # only a handful of one-letter words are legitimate (units like "м")
        self.assertLess(
            len([t for t in texts if len(t) == 1 and t.isalpha()]), 5)

    def test_missing_file_raises(self):
        with self.assertRaises(Exception):
            P.decode_with_pdfplumber('/nonexistent/nope.pdf')


if __name__ == '__main__':
    unittest.main()
