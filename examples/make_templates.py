#!/usr/bin/env python3
"""
Generate the example PDF templates used by the test-suite and by
`python -m pdf_spec_parser.validate`.

Three templates with genuinely different layout (page format, column set,
column order, row pitch, multi-line cells, noise text):

  template_a_gost_drawing.pdf  — large-format vector drawing (A0-ish, scaled
                                 user space), spec table above the title block,
                                 columns: Поз. / Наименование / Обозначение /
                                 Кол. / Ед. изм., row pitch 120 pt.
  template_b_continuation.pdf  — A4 portrait continuation sheet, table at the
                                 top of the page, dense row pitch (22 pt),
                                 footer text below the table, one quantity
                                 cell without a number ("комплект").
  template_c_board_spec.pdf    — A3 landscape board spec, unusual column
                                 ORDER (Поз. / Кол. / Ед. / Наименование /
                                 Примечание), multi-line names, note column.

The generated PDFs are committed to the repository, so the test-suite and CI
do not need this script (and therefore no font installed on the runner).
Re-generate with:

    python examples/make_templates.py            # writes PDFs + manifest
    python examples/make_templates.py --check    # regenerate, compare to git

The Cyrillic text needs a TrueType font with Cyrillic coverage (reportlab's
built-in Helvetica has none). One of the usual DejaVu/Liberation paths is
used, or $PDF_SPEC_TEMPLATE_FONT.
"""

import argparse
import hashlib
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

FONT_CANDIDATES = [
    os.environ.get('PDF_SPEC_TEMPLATE_FONT', ''),
    '/usr/share/fonts/TTF/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
]

BOLD_CANDIDATES = [
    os.environ.get('PDF_SPEC_TEMPLATE_FONT_BOLD', ''),
    '/usr/share/fonts/TTF/DejaVuSans-Bold.ttf',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    '/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
]


def _first_existing(paths):
    for p in paths:
        if p and os.path.exists(p):
            return p
    return None


# ──────────────────────────────────────────────────────────────────────
# TEMPLATE A — large-format GOST-style drawing (scaled user space)
# ──────────────────────────────────────────────────────────────────────

TEMPLATE_A = {
    'id': 'template_a',
    'file': 'template_a_gost_drawing.pdf',
    'title': 'Блок WA.00.000 — схема электрическая принципиальная',
    'layout': ('large-format drawing (5940×4200 pt), spec table with 5 columns '
               '(Поз. / Наименование / Обозначение / Кол. / Ед. изм.) above the '
               'title block, row pitch 120 pt, one two-line name'),
    'exercises': ['header row + code column + separate unit column',
                  'two-line name merged into one item',
                  'quantity in centimetres converted to metres',
                  'component outside the knowledge base (websearch path)',
                  'drawing annotations ignored'],
    # (x, header, align)
    'columns': [
        ('item', 450, 'Поз.', 'left'),
        ('desc', 700, 'Наименование', 'left'),
        ('code', 2100, 'Обозначение', 'left'),
        ('qty', 3900, 'Кол.', 'left'),
        ('unit', 4200, 'Ед. изм.', 'left'),
    ],
    'header_top': 3000,
    'pitch': 120,
    'font_size': 32,
    'line_gap': 34,
    'rows': [
        {'item': '1', 'desc': ['Реле промежуточное ИЕК РЕК 78/4'],
         'code': 'РЕК78/4', 'qty': '3', 'unit': 'шт'},
        {'item': '2', 'desc': ['Зажим винтовой ЗВИ-6'],
         'code': 'ЗВИ6', 'qty': '12', 'unit': 'шт'},
        {'item': '3', 'desc': ['Провод монтажный ПВ3', '1х2,5 белый'],
         'code': 'ПВ3', 'qty': '250', 'unit': 'см'},
        {'item': '4', 'desc': ['Лампа матричная НД-16 красный'],
         'code': 'НД16', 'qty': '8', 'unit': 'шт'},
        {'item': '5', 'desc': ['Датчик температуры ТСМ-100'],
         'code': 'ТСМ100', 'qty': '2', 'unit': 'шт'},
    ],
    # text outside the table (annotation area + title block below the table)
    'free_text': [
        (300, 200, 'СХЕМА ЭЛЕКТРИЧЕСКАЯ ПРИНЦИПИАЛЬНАЯ', 'bold', 44),
        (300, 280, 'Оборудование шкафа управления насосной станцией', 'font', 36),
        (300, 360, 'Все размеры в миллиметрах', 'font', 36),
        (300, 4180, 'Блок WA.00.000', 'bold', 36),
        (2400, 4180, 'ООО Техмонтаж', 'font', 36),
        (4800, 4180, 'Лист 3', 'font', 36),
    ],
    'expected': [
        {'pos': 1, 'name_contains': 'Реле промежуточное', 'qty': 3.0, 'unit': 'шт',
         'code': 'РЕК78/4', 'category': 'Коммутационное оборудование'},
        {'pos': 2, 'name_contains': 'Блок зажимов', 'qty': 12.0, 'unit': 'шт',
         'code': 'ЗВИ6', 'category': 'Клеммы и монтаж'},
        {'pos': 3, 'name_contains': 'Провод монтажный', 'qty': 2.5, 'unit': 'м',
         'code': 'ПВ3', 'category': 'Провода и кабели'},
        {'pos': 4, 'name_contains': 'Светосигнальный индикатор', 'qty': 8.0,
         'unit': 'шт', 'code': 'НД16', 'category': 'Светосигнальная арматура'},
        {'pos': 5, 'name_contains': 'Датчик температуры', 'qty': 2.0, 'unit': 'шт',
         'code': 'ТСМ100', 'category': 'Прочее'},
    ],
}

# ──────────────────────────────────────────────────────────────────────
# TEMPLATE B — A4 portrait continuation sheet (dense rows, table on top)
# ──────────────────────────────────────────────────────────────────────

TEMPLATE_B = {
    'id': 'template_b',
    'file': 'template_b_continuation.pdf',
    'title': 'Продолжение спецификации оборудования и материалов',
    'layout': ('A4 portrait (595×842 pt), table at the TOP of the page, 5 '
               'columns (№ / Обозначение / Наименование / Кол-во / Ед. изм.), '
               'dense row pitch 22 pt, footer text below the table'),
    'exercises': ['table at the top of the page (headers are not the lowest text)',
                  'dense rows: fixed row band would merge neighbouring items',
                  'quantity cell without a number ("комплект")',
                  'brand name removal (Сальник MG12 Chint)',
                  'footer text below the table ignored'],
    'columns': [
        ('item', 45, '№', 'left'),
        ('code', 95, 'Обозначение', 'left'),
        ('desc', 200, 'Наименование', 'left'),
        ('qty', 470, 'Кол-во', 'left'),
        ('unit', 518, 'Ед. изм.', 'left'),
    ],
    'header_top': 80,
    'pitch': 22,
    'font_size': 9,
    'line_gap': 11,
    'rows': [
        {'item': '1', 'code': 'РП21', 'desc': ['Реле промежуточное РП-21'],
         'qty': '4', 'unit': 'шт'},
        {'item': '2', 'code': 'РП21-Р', 'desc': ['Разъём для реле РП-21'],
         'qty': '4', 'unit': 'шт'},
        {'item': '3', 'code': 'ПВ3', 'desc': ['Провод ПВ3 1х2,5 белый'],
         'qty': '1,5', 'unit': 'м'},
        {'item': '4', 'code': 'КВВГ', 'desc': ['Кабель КВВГ 4х1,0'],
         'qty': '12', 'unit': 'м'},
        {'item': '5', 'code': 'НГИ1', 'desc': ['Наконечник гильзовый НГИ 1,0'],
         'qty': '100', 'unit': 'шт'},
        {'item': '6', 'code': 'MG12', 'desc': ['Сальник MG12 Chint'],
         'qty': '2', 'unit': 'шт'},
        {'item': '7', 'code': 'СТ200', 'desc': ['Стяжка нейлоновая 200'],
         'qty': '50', 'unit': 'шт'},
        {'item': '8', 'code': 'КР', 'desc': ['Комплект крепежа'],
         'qty': 'комплект', 'unit': ''},
    ],
    'free_text': [
        (45, 40, 'Продолжение ведомости, начало см. лист 1', 'font', 9),
        (400, 800, 'Формат A4', 'font', 9),
        (45, 820, 'Лист 3 из 6', 'font', 9),
    ],
    'expected': [
        {'pos': 1, 'name_contains': 'Реле промежуточное', 'qty': 4.0, 'unit': 'шт',
         'code': 'РП21', 'category': 'Коммутационное оборудование'},
        {'pos': 2, 'name_contains': 'Розетка', 'qty': 4.0, 'unit': 'шт',
         'code': 'РП21-Р', 'category': 'Коммутационное оборудование'},
        {'pos': 3, 'name_contains': 'Провод монтажный', 'qty': 1.5, 'unit': 'м',
         'code': 'ПВ3', 'category': 'Провода и кабели'},
        {'pos': 4, 'name_contains': 'Кабель контрольный', 'qty': 12.0, 'unit': 'м',
         'code': 'КВВГ', 'category': 'Провода и кабели'},
        {'pos': 5, 'name_contains': 'Наконечник штыревой', 'qty': 100.0,
         'unit': 'шт', 'code': 'НГИ1', 'category': 'Монтажные материалы'},
        {'pos': 6, 'name_contains': 'Ввод кабельный', 'qty': 2.0, 'unit': 'шт',
         'code': 'MG12', 'category': 'Корпусное оборудование и аксессуары'},
        {'pos': 7, 'name_contains': 'Стяжка нейлоновая', 'qty': 50.0, 'unit': 'шт',
         'code': 'СТ200', 'category': 'Монтажные материалы'},
        {'pos': 8, 'name_contains': 'Комплект крепежа', 'qty': None,
         'unit': None, 'code': 'КР', 'category': 'Прочее'},
    ],
}

# ──────────────────────────────────────────────────────────────────────
# TEMPLATE C — A3 landscape board spec (unusual column order + note column)
# ──────────────────────────────────────────────────────────────────────

TEMPLATE_C = {
    'id': 'template_c',
    'file': 'template_c_board_spec.pdf',
    'title': 'Спецификация навесного оборудования щита ЩРН-12',
    'layout': ('A3 landscape (1191×842 pt), columns in an unusual ORDER '
               '(Поз. / Кол. / Ед. / Наименование / Примечание), row pitch '
               '40 pt, one two-line name, note column'),
    'exercises': ['quantity column left of the name column',
                  'note column present and ignored',
                  'two-line name merged into one item',
                  'component with a known category but no rule for its '
                  'parameters (Автомат ВА 47-29 IEK)'],
    'columns': [
        ('item', 60, 'Поз.', 'left'),
        ('qty', 115, 'Кол.', 'left'),
        ('unit', 170, 'Ед. изм.', 'left'),
        ('desc', 240, 'Наименование', 'left'),
        ('note', 760, 'Примечание', 'left'),
    ],
    'header_top': 560,
    'pitch': 40,
    'font_size': 10,
    'line_gap': 18,
    'rows': [
        {'item': '1', 'qty': '20', 'unit': 'шт', 'desc': ['Зажим винтовой ЗВИ-4'],
         'note': ''},
        {'item': '2', 'qty': '6', 'unit': 'шт',
         'desc': ['Выключатель клавишный ВН-10 250В'], 'note': 'см. лист 1'},
        {'item': '3', 'qty': '3', 'unit': 'шт', 'desc': ['Автомат ВА 47-29 IEK'],
         'note': ''},
        {'item': '4', 'qty': '24', 'unit': 'шт', 'desc': ['Изолятор Dr 16-4-5-6'],
         'note': ''},
        {'item': '5', 'qty': '1', 'unit': 'компл',
         'desc': ['Датчик температуры', 'ТСМ-100'], 'note': 'по проекту'},
    ],
    'free_text': [
        (60, 200, 'Щит навесной ЩРН-12, степень защиты IP31', 'font', 11),
        (60, 230, 'Габаритные размеры корпуса указаны в миллиметрах', 'font', 11),
        (60, 780, 'Лист 1', 'font', 10),
    ],
    'expected': [
        {'pos': 1, 'name_contains': 'Блок зажимов', 'qty': 20.0, 'unit': 'шт',
         'code': '', 'category': 'Клеммы и монтаж'},
        {'pos': 2, 'name_contains': 'Выключатель клавишный', 'qty': 6.0,
         'unit': 'шт', 'code': '', 'category': 'Коммутационное оборудование'},
        {'pos': 3, 'name_contains': 'Автомат', 'qty': 3.0, 'unit': 'шт',
         'code': '', 'category': 'Коммутационное оборудование'},
        {'pos': 4, 'name_contains': 'Изолятор', 'qty': 24.0, 'unit': 'шт',
         'code': '', 'category': 'Монтажные материалы'},
        {'pos': 5, 'name_contains': 'Датчик температуры', 'qty': 1.0,
         'unit': 'компл', 'code': '', 'category': 'Прочее'},
    ],
}

TEMPLATES = [TEMPLATE_A, TEMPLATE_B, TEMPLATE_C]


# ──────────────────────────────────────────────────────────────────────
# RENDERING
# ──────────────────────────────────────────────────────────────────────

def render(tpl, font, font_bold, out_dir):
    """Render one template to PDF, return the output path."""
    from reportlab.lib.pagesizes import A3, A4, landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas

    sizes = {
        'template_a': (5940, 4200),
        'template_b': A4,
        'template_c': landscape(A3),
    }
    width, height = sizes[tpl['id']]

    pdfmetrics.registerFont(TTFont('TplRegular', font))
    pdfmetrics.registerFont(TTFont('TplBold', font_bold))

    path = os.path.join(out_dir, tpl['file'])
    c = canvas.Canvas(path, pagesize=(width, height))
    c.setTitle(tpl['title'])
    c.setAuthor('pdf-spec-parser templates')

    # frame
    c.setLineWidth(1)
    c.rect(20, 20, width - 40, height - 40, stroke=1, fill=0)

    fs = tpl['font_size']
    pitch = tpl['pitch']
    gap = tpl['line_gap']
    header_top = tpl['header_top']

    # header row
    c.setFont('TplBold', fs)
    for _cid, x, header, _align in tpl['columns']:
        c.drawString(x, height - header_top, header)

    # data rows
    for i, row in enumerate(tpl['rows']):
        top = header_top + pitch * (i + 1)
        for cid, x, _header, _align in tpl['columns']:
            value = row.get(cid)
            if value in (None, ''):
                continue
            lines = value if isinstance(value, list) else [value]
            c.setFont('TplRegular', fs)
            for j, line in enumerate(lines):
                c.drawString(x, height - (top + j * gap), line)

    # free (non-table) text
    for x, top, text, style, size in tpl['free_text']:
        c.setFont('TplBold' if style == 'bold' else 'TplRegular', size)
        c.drawString(x, height - top, text)

    c.showPage()
    c.save()
    return path


def build_manifest(out_dir):
    """Build the manifest (ground truth per template)."""
    templates = []
    for tpl in TEMPLATES:
        path = os.path.join(out_dir, tpl['file'])
        with open(path, 'rb') as fh:
            digest = hashlib.sha256(fh.read()).hexdigest()
        templates.append({
            'id': tpl['id'],
            'file': tpl['file'],
            'title': tpl['title'],
            'layout': tpl['layout'],
            'exercises': tpl['exercises'],
            'sha256': digest,
            'expected_items': tpl['expected'],
        })
    return {
        'schema_version': '1.0',
        'generator': 'examples/make_templates.py',
        'note': ('Ground truth for `python -m pdf_spec_parser.validate` and the '
                 'test-suite. `sha256` pins the exact PDF revision the '
                 'expectations were written for.'),
        'templates': templates,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(
        description='Generate the example PDF templates of pdf-spec-parser.')
    parser.add_argument('--out-dir', default=HERE,
                        help='where to write the PDFs (default: examples/)')
    parser.add_argument('--check', action='store_true',
                        help='do not write: report whether the manifest on disk '
                             'matches freshly generated PDFs')
    args = parser.parse_args(argv)

    font = _first_existing(FONT_CANDIDATES)
    font_bold = _first_existing(BOLD_CANDIDATES)
    if not font or not font_bold:
        print('ERROR: no Cyrillic TrueType font found. Set '
              'PDF_SPEC_TEMPLATE_FONT / PDF_SPEC_TEMPLATE_FONT_BOLD.',
              file=sys.stderr)
        return 1

    out_dir = args.out_dir
    if args.check:
        import tempfile
        tmp = tempfile.mkdtemp(prefix='pdf-spec-templates-')
        for tpl in TEMPLATES:
            render(tpl, font, font_bold, tmp)
        fresh = build_manifest(tmp)
        with open(os.path.join(out_dir, 'templates_manifest.json')) as fh:
            current = json.load(fh)
        # The PDF byte hash changes on every render (embedded creation date),
        # so compare the manifest content without it.
        def strip(manifest):
            return [{k: v for k, v in entry.items() if k != 'sha256'}
                    for entry in manifest['templates']]
        ok = json.dumps(strip(fresh), sort_keys=True) == json.dumps(strip(current), sort_keys=True)
        hashes = [e['sha256'] for e in fresh['templates']] == \
                 [e['sha256'] for e in current['templates']]
        print('templates_manifest.json: '
              + ('matches freshly generated PDFs' if ok else 'STALE'))
        if ok and not hashes:
            print('note: PDF bytes differ (regeneration is not byte-stable) — '
                  're-run without --check and commit the manifest')
        return 0 if ok else 1

    for tpl in TEMPLATES:
        path = render(tpl, font, font_bold, out_dir)
        print('wrote %s' % os.path.relpath(path))

    manifest = build_manifest(out_dir)
    manifest_path = os.path.join(out_dir, 'templates_manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
        fh.write('\n')
    print('wrote %s' % os.path.relpath(manifest_path))
    return 0


if __name__ == '__main__':
    sys.exit(main())
