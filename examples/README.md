# Examples

This directory contains **three committed example PDFs** with genuinely
different layouts, plus the generator that produced them and the ground truth
used by the test-suite and by `python -m pdf_spec_parser.validate`.

| file | what it is |
|------|------------|
| `template_a_gost_drawing.pdf` | wide (5940×4200 pt) vector drawing, spec table above the title block, columns Поз. / Наименование / Обозначение / Кол. / Ед. изм., row pitch 120 pt, one two-line name |
| `template_b_continuation.pdf` | A4 portrait continuation sheet, table at the top of the page, dense rows (22 pt), footer text below the table, one quantity cell `комплект` |
| `template_c_board_spec.pdf` | A3 landscape board spec, unusual column order (Поз. / Кол. / Ед. / Наименование / Примечание), note column, two-line name |
| `templates_manifest.json` | ground truth: layout description, SHA-256 of each PDF and the expected items |
| `make_templates.py` | regenerates the PDFs and the manifest (needs a Cyrillic TrueType font) |

Run the validator over all of them:

```bash
python -m pdf_spec_parser.validate --examples --out-dir out/
```

## Command-line usage

```bash
# Parse a PDF and generate obfuscated spec (deterministic, offline)
python -m pdf_spec_parser "drawing.pdf" "output_spec" --verbose --no-websearch

# Contract output only
python -m pdf_spec_parser "drawing.pdf" --format csv,json
```

## Python API usage

```python
from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf, contract

stats = {}
items = parse_pdf("drawing.pdf", verbose=True, stats=stats)

print(f"Parsed {len(items)} items:")
for item in items:
    print(f"  {item['pos']}. [{item['category']}] {item['name']}")
    print(f"       {item['params']}")
    print(f"       qty={item['qty']} {item['unit']}  flags={item['flags']}")

export_xlsx(items, "spec.xlsx", cabinet_name="My Cabinet")
export_pdf(items, "spec.pdf", title_text="My Cabinet")

records = contract.to_records(items)
contract.write_csv(records, "spec.csv")
status, issues = contract.verdict(records)
print(status, issues)
```

## Working with the data

Each formatted item is a dict (see `pdf_spec_parser/contract.py` for the
normative schema):

```python
{
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
```

`qty` is a number or `None` (never a string), `unit` is `None` if and only if
`qty` is `None`, and `flags` tells a consumer which rows need a human look.

## Regenerating the templates

```bash
python examples/make_templates.py            # write PDFs + manifest
python examples/make_templates.py --check    # report manifest drift
```

Font resolution: `$PDF_SPEC_TEMPLATE_FONT` / `$PDF_SPEC_TEMPLATE_FONT_BOLD`
or one of the usual DejaVu/Liberation paths.
