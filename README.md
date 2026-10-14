# pdf-spec-parser

**Universal procurement PDF specification parser & obfuscator.**

Extracts specification tables from GOST-font vector PDF drawings, automatically classifies components by type, removes brand/proprietary information, and exports to clean formatted Excel + PDF — plus a **documented CSV/JSON output contract** for machine consumption.

> **Provenance.** Grown out of the Lost Tech Museum archive record
> [`pdf-spec-parser`](https://kostyk348.github.io/lost-tech-museum/digests/portfolio-review/)
> (digest “Портфель kostyk348: что почти готово и что стоит дореализовать”),
> filed as a `poc` tool with `ready: 2 / value: 4` and the pitch:
> *«Сервис “PDF-спецификация → таблица” для снабжения и проектирования:
> загрузка пачки PDF, нормализация в CSV/JSON, экспорт в смету. Обфускатор
> продаётся отдельно как режим “поделиться спецификацией без NDA”»*.
> The audit recorded three gaps — *no output schema, no validation across
> templates, no CLI error documentation*. Version 1.1.0 closes exactly those
> three: see [Output contract](#output-contract-csv--json),
> [Validation on templates](#validation-on-templates) and
> [Errors and exit codes](#errors-and-exit-codes).

---

## The Problem

Procurement drawings from Russian/Eastern European manufacturers often use:
- GOST fonts with Identity-H CID encoding (not extractable by standard PDF tools)
- Mixed font types in a single document
- Non-standard table layouts
- Proprietary brand names (ИЕК, РЕК, Chint, etc.) that need removal

Most PDF parsers choke on these files. pdf-spec-parser handles them.

## How It Works

```
Input PDF (GOST CID / standard)
        │
        ▼
┌───────────────────┐
│ 1. Auto-detect    │ ← pdfplumber for standard, CID decoder for GOST
│    font type      │
└────────┬──────────┘
         ▼
┌───────────────────┐
│ 2. Find the       │ ← Locates the header row (Поз / Наименование / Кол …)
│    header row     │   on the last pages and projects the columns from the
│                   │   header cell positions
└────────┬──────────┘
         ▼
┌───────────────────┐
│ 3. Parse items    │ ← Groups text blocks into rows using the real row pitch
│    & quantities   │
└────────┬──────────┘
         ▼
┌─────────────────────────┐
│ 4. Classify components  │ ← Domain knowledge + websearch fallback
│    Remove brand names   │
└────────┬────────────────┘
         ▼
┌───────────────────────────────────────────┐
│ 5. Export                                 │
│    .xlsx / .pdf  → humans (flags column)  │
│    .csv  / .json → machines (contract)    │
└───────────────────────────────────────────┘
```

Column detection is derived from the detected header row instead of hard-coded
coordinates, so a wide GOST drawing, an A4 continuation sheet and an A3 sheet
with an unusual column order all work with the same code. The fixed GOST
ranges remain as a fallback for tables whose header row is not recognised.

## Supported Components

| Category | Examples |
|----------|----------|
| Switching equipment | Relays, sockets, contactors, circuit breakers |
| Terminals & mounting | Terminal blocks, busbars |
| Mounting materials | Ferrules, ring/flat terminals, insulators, cable ties |
| Wires & cables | Installation wires, control cables, flexible grounding |
| Signal equipment | Indicator lamps, LEDs, signal towers |
| Enclosures | Distribution boxes, cabinets, cable glands |
| Power equipment | Transformers, power supplies, drivers |

## Installation

```bash
# Clone the repository
git clone https://github.com/kostyk348/pdf-spec-parser.git
cd pdf-spec-parser

# Install (Python 3.11+)
pip install -e .

# Or install dependencies manually
pip install pdfplumber openpyxl reportlab fonttools PyMuPDF requests
```

### GOST Font

For decoding GOST Identity-H CID fonts, you need a `GOSTTypeA_2.ttf` font file:

```bash
# Place it in one of these locations:
./fonts/GOSTTypeA_2.ttf
~/Downloads/fonts/GOSTTypeA_2.ttf
```

Without the font, the tool still works on standard PDF files (pdfplumber path).
**Note:** the CID path is not covered by the test-suite or CI, because the GOST
font is proprietary and is not shipped with the repository.

## Usage

### Command-line

```bash
# Basic usage — output files get the `_anon` suffix
python -m pdf_spec_parser drawing.pdf

# Custom output prefix
python -m pdf_spec_parser drawing.pdf output_prefix

# Verbose mode (shows progress and every parsed item)
python -m pdf_spec_parser drawing.pdf --verbose

# Machine-readable output only, deterministic and offline
python -m pdf_spec_parser drawing.pdf --format csv,json --no-websearch

# Everything into a directory of its own
python -m pdf_spec_parser drawing.pdf --out-dir out/ --format all
```

| Option | Meaning |
|--------|---------|
| `input` | input PDF drawing |
| `output` | output prefix without extension (default: `<input>_anon`) |
| `--verbose`, `-v` | detailed progress logging |
| `--format LIST` | `csv`, `json`, `xlsx`, `pdf` or `all` (default: `all`) |
| `--out-dir DIR` | write the output files into `DIR` |
| `--no-websearch` | never query the web for unknown components (deterministic, offline) |

Without `--no-websearch` the parameters of components outside the knowledge
base come from a DuckDuckGo lookup: those rows are flagged `WEB_INFO` and the
run is not reproducible.

### Python API

```python
from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf
from pdf_spec_parser import contract

stats = {}
items = parse_pdf("drawing.pdf", websearch=False, stats=stats)

# items is a list of dicts — see the contract below for every field
# [
#   {
#     'pos': 1,
#     'category': 'Коммутационное оборудование',
#     'name': 'Реле промежуточное электромагнитное',
#     'code': 'РЕК78/4',
#     'params': '4 переключающих контакта (4PDT)',
#     'qty': 3.0,
#     'unit': 'шт',
#     'qty_raw': '3',
#     'flags': [],
#   },
#   ...
# ]

# Human exports
export_xlsx(items, "output.xlsx", cabinet_name="Блок WA.00.000")
export_pdf(items, "output.pdf", title_text="Блок WA.00.000")

# Machine contract
records = contract.to_records(items)
contract.write_csv(records, "spec.csv")
document = contract.build_document(records, "drawing.pdf", stats)
contract.write_json(document, "spec.json")
status, issues = contract.verdict(records)
```

`stats` receives the run diagnostics (`blocks`, `pages`, `table_page`,
`columns`, `column_hits`, `row_band`, `fallback`, `websearch`); they end up in
`document['source']`.

## Output contract (CSV / JSON)

`schema_version 1.0`. Every record has exactly these nine fields, in this
order — the module docstring of `pdf_spec_parser/contract.py` is the normative
version of this table:

| field | type | rules |
|-------|------|-------|
| `pos` | integer | position number from the drawing, 1..999, unique and ascending |
| `category` | string | one of the 7 classes above plus `Прочее`; never empty |
| `name` | string | de-branded component name; never empty |
| `code` | string | drawing designation (`ВА 47-29`); `''` when the table has no designation column |
| `params` | string | parameters/ratings, comma separated; `''` when nothing is known |
| `qty` | number \| null | quantity as a number; `null` when the cell holds no number |
| `unit` | string \| null | unit of measure; `null` if and only if `qty` is `null` |
| `qty_raw` | string | the quantity cell exactly as read from the PDF |
| `flags` | array[string] | subset of the flag vocabulary below; `[]` for a confident row |

### Flags

| flag | meaning |
|------|---------|
| `UNKNOWN_COMPONENT` | name not resolved by the knowledge base — params need a human look |
| `WEB_INFO` | params came from the websearch fallback (non-deterministic) |
| `QTY_CONVERTED` | quantity converted cm → m |
| `NO_QTY` | the quantity cell was empty |
| `UNPARSED_QTY` | the quantity cell had text but no number (e.g. `комплект`); `qty` is `null` |

### CSV layout

* UTF-8 **with BOM** (`utf-8-sig`) so Excel opens Cyrillic correctly.
* Header row with the nine field names, then one row per record.
* `\n` line terminator, `,` separator, RFC 4180 quoting.
* `flags` joined with `;`.
* `qty` uses a dot decimal separator; `null` values are empty cells (pair them
  with `unit == ''`).

```csv
pos,category,name,code,params,qty,unit,qty_raw,flags
1,Коммутационное оборудование,Реле промежуточное электромагнитное,РЕК78/4,4 переключающих контакта (4PDT),3,шт,3,
3,Провода и кабели,Провод монтажный медный повышенной гибкости (заземление),ПВ3,"Число жил и сечение: 1 × 2,5 мм², цвет: белый",2.5,м,250,QTY_CONVERTED
5,Прочее,Датчик температуры ТСМ-100,ТСМ100,(требуется уточнение),2,шт,2,UNKNOWN_COMPONENT
8,Прочее,Комплект крепежа,КР,(требуется уточнение),,,комплект,UNPARSED_QTY;UNKNOWN_COMPONENT
```

### JSON layout

```json
{
  "schema_version": "1.0",
  "generator": "pdf-spec-parser 1.1.0",
  "source": {
    "file": "template_c_board_spec.pdf",
    "sha256": "4ead5208ff7c02d687f1a5261737b8d9f1b05653e2914f14e57094307043547a",
    "pages": 1, "blocks": 57, "table_page": 1,
    "columns": ["desc", "item", "note", "qty", "unit_col"],
    "column_hits": {"item": 5, "desc": 23, "qty": 9, "code": 0, "unit": 10},
    "row_band": 24.0, "websearch": false
  },
  "counts": {"records": 5, "flags": {"UNKNOWN_COMPONENT": 2}},
  "items": ["…nine fields per record…"],
  "diagnostics": {"issues": ["…validator issues…"]}
}
```

### Unrecognised lines

Nothing is dropped silently — every emitted row carries its own flags and the
run diagnostics carry the aggregate counters:

* Text outside every detected column (annotations, title block, footers) is
  ignored, and the volume of text seen is reported in `source.blocks` /
  `source.column_hits`.
* Records are anchored on position numbers — the one thing every spec table
  guarantees. A stray number inside the position column therefore produces a
  record with an empty name, and that is a contract **error**
  (`E_NAME_EMPTY`), not a silent pass.
* A repeated header row in the middle of a long table is recognised and
  skipped instead of being parsed as data.
* A component the knowledge base does not know is still emitted, with its raw
  name, the flag `UNKNOWN_COMPONENT`, and either websearch parameters
  (`WEB_INFO`) or the literal `(требуется уточнение)`.
* An unparseable quantity cell (`комплект`) keeps its text in `qty_raw`, sets
  `qty`/`unit` to `null` and adds `UNPARSED_QTY`.

## Validation on templates

`python -m pdf_spec_parser.validate` runs the whole pipeline over a set of
PDFs, checks the produced records against the contract **and** against the
ground truth in `examples/templates_manifest.json`, and prints one verdict per
file saying exactly where it broke:

```bash
python -m pdf_spec_parser.validate --examples --out-dir out/
python -m pdf_spec_parser.validate drawing.pdf
```

```
[PASS] examples/template_a_gost_drawing.pdf — 5 records, contract PASS, manifest PASS
       table page 1 | row band 72.0 | columns code, desc, item, qty, unit_col
       column hits: code=8, desc=27, item=5, qty=6, unit=9
       [warning] W_UNKNOWN_COMPONENT: pos 5: component not in the knowledge base, params need a human look
[PASS] examples/template_b_continuation.pdf — 8 records, contract PASS, manifest PASS
[PASS] examples/template_c_board_spec.pdf — 5 records, contract PASS, manifest PASS
summary: 3/3 templates PASS
```

A failing file reports the stage it broke at
(`decode` / `parse` / `contract` / `manifest`), the exception together with its
diagnostics (table page, detected columns, per-column hit counts) or the
manifest mismatch, plus every flag that needs a human look.

### Example templates

`examples/` holds three committed PDFs with genuinely different layout, so the
geometry logic is exercised instead of assumed:

| template | page | layout it exercises |
|----------|------|---------------------|
| `template_a_gost_drawing.pdf` | 5940×4200 drawing | wide scaled user space, 5 columns incl. Обозначение + Ед. изм., row pitch 120 pt, two-line name, quantity in cm |
| `template_b_continuation.pdf` | A4 portrait | table at the **top** of the sheet, dense rows (22 pt), footer text below the table, quantity cell `комплект` |
| `template_c_board_spec.pdf` | A3 landscape | columns in an unusual order (Поз. / Кол. / Ед. / Наименование / Примечание), note column, two-line name |

Each template deliberately contains a component outside the knowledge base, a
row that needs a human look, and non-table text. Regenerate them (needs a
Cyrillic TTF font) with:

```bash
python examples/make_templates.py            # PDFs + manifest
python examples/make_templates.py --check    # detect manifest drift
```

`examples/templates_manifest.json` pins each PDF by SHA-256 together with its
expected items; the test-suite fails if a PDF changes without the manifest.

## Errors and exit codes

| code | meaning | what to do |
|------|---------|------------|
| 0 | success | — |
| 1 | unexpected internal error | a bug — run with `--verbose` and report it |
| 2 | usage error: bad arguments, file not found/directory, not a readable PDF | check the path and `--format` |
| 3 | no text could be extracted | the PDF is a scan without a text layer; OCR it first |
| 4 | text extracted, but no spec table rows recognised | the message lists the table page, the detected columns and the per-column hit counts |
| 5 | produced records violate the output contract | the errors are printed and filed in `diagnostics.issues`; `E_NAME_EMPTY` normally means a stray number in the position column |
| 6 | nothing to write | `--format` selected no output |

Validator exit codes: `0` — every file passed, `4` — a template produced no
records at all, `5` — a file failed (pipeline / contract / manifest),
`2` — usage error.

### Contract error codes

`validate_records()` returns issues with a `code`, a `severity` (`error` makes
the verdict FAIL, `warning` does not) and a message:

| code | severity | trigger |
|------|----------|---------|
| `E_SCHEMA_FIELD_MISSING` / `E_SCHEMA_FIELD_UNKNOWN` | error | a record does not have exactly the nine contract fields |
| `E_CONTRACT_TYPE` | error | records are not a list of objects |
| `E_POS_TYPE` / `E_POS_RANGE` | error | `pos` is not an integer in 1..999 |
| `E_POS_DUPLICATE` / `E_POS_ORDER` | error | the same position twice, or out of ascending order |
| `E_CATEGORY_EMPTY` / `E_CATEGORY_UNKNOWN` | error | category empty or outside the vocabulary |
| `E_NAME_EMPTY` | error | nothing was recognised in the name column |
| `E_FIELD_TYPE` | error | `code` / `params` / `qty_raw` is not a string |
| `E_QTY_TYPE` / `E_QTY_NEGATIVE` | error | `qty` is not a number or is negative |
| `E_UNIT_WITHOUT_QTY` / `E_QTY_WITHOUT_UNIT` | error | `unit` and `qty` disagree |
| `E_FLAGS_TYPE` / `E_FLAG_UNKNOWN` | error | flags are not a list of known flags |
| `W_UNKNOWN_COMPONENT` | warning | a component needs a human look |
| `W_NONDETERMINISTIC` | warning | params came from the websearch fallback |

## Tests & CI

The test-suite is stdlib `unittest` (no extra dependency) and covers the
contract layer, the geometry/knowledge units, the three example templates, the
validator and the CLI exit codes:

```bash
python -m unittest discover -s tests -v
```

`.github/workflows/ci.yml` installs the package (`pip install -e .`) on Python
3.11 and 3.13, runs the test-suite, validates every example template, does a
CLI smoke run per template and uploads the generated `out/` as an artifact.

## Known limitations

* The CID decoder needs the proprietary `GOSTTypeA_2.ttf`, so the CID path is
  not covered by the tests (the pdfplumber path is the tested one). One of its
  font search paths is still Windows-flavoured (`~\Downloads\fonts\...`).
* The column layout is inferred from the header row: a table without a
  recognisable header row falls back to the fixed GOST ranges, and a header row
  whose cells sit on visibly different baselines may not be recognised.
* Components with a known category but no parameter rule (e.g.
  `Автомат ВА 47-29 IEK`) are emitted with the generic
  `(требуется уточнение)` parameter and the flag `UNKNOWN_COMPONENT`.
* For PUGV wires the parameter list can contain a redundant default section
  (`Число жил и сечение: 1 × 2,5 мм², Сечение: 2.5 мм²`).
* `--format pdf` registers Arial when it is present (Windows) and falls back to
  Helvetica otherwise, so Cyrillic in the generated PDF needs a suitable font
  on the host.

## Brand Obfuscation

The tool automatically detects and removes or generalizes brand names:

| Original | Obfuscated |
|----------|------------|
| Реле промежуточное ИЕК РЕК 78/4 | Реле промежуточное электромагнитное |
| Сальник MG12 Chint | Ввод кабельный (сальник) MG12 |
| Автомат ВА 47-29 IEK | (generalized) |

## Dependencies

- **pdfplumber** — standard PDF text extraction
- **PyMuPDF** (fitz) — low-level PDF stream access for CID decoding
- **fonttools** — GOST font CID map parsing
- **openpyxl** — Excel export
- **reportlab** — PDF export with Cyrillic font support
- **requests** — websearch fallback for unknown components

No new dependency was added in 1.1.0: the contract layer, the validator and the
tests use the standard library only. PDF generation for the example templates
uses reportlab, which is already a runtime dependency.

## What changed in 1.1.0

* **Output contract** (CSV + JSON), documented in `pdf_spec_parser/contract.py`:
  field names, types, flags, normalisation rules, unrecognised-line policy.
* **Validation across templates**: three committed example PDFs with different
  layouts, a SHA-256-pinned manifest, and `python -m pdf_spec_parser.validate`
  printing a verdict per template.
* **CLI error documentation**: four output formats, `--out-dir`,
  `--no-websearch`, six documented exit codes, contract error codes.
* Test-suite (`tests/`, stdlib `unittest`) and a GitHub Actions workflow.
* Bug fixes found by the validation run:
  * the pdfplumber decoder returned **one block per character**, which turned
    every name into `Р е л е` and made category detection fail — it now
    returns one block per word;
  * the header row had to sit in the bottom 10 % of the page, so any table with
    its header at the top fell back to hard-coded GOST coordinates — columns
    are now projected from the detected header row;
  * the fixed vertical row band merged dense sheets and split wide drawings —
    the band now follows the real row pitch;
  * the quantity field could be a *string* (`комплект`) presented as a number —
    it is now `null` plus the flag `UNPARSED_QTY`;
  * a Latin-script PDF without Cyrillic text was reported as “no text
    extracted” even though text had been extracted;
  * a two-line component name was reassembled in the wrong order;
  * a repeated header row could be parsed as table data;
  * `python -m pdf_spec_parser` (documented everywhere) failed with
    `No module named pdf_spec_parser.__main__`;
  * `pip install -e .` failed: the build backend
    `setuptools.backends._legacy:_Backend` does not exist;
  * the package metadata claimed MIT while `LICENSE` is Apache-2.0.

## License

Apache 2.0

## See Also

- [SINT-UA project family](https://github.com/kostyk348?tab=repositories) — related tooling for deterministic systems optimization
- [Lost Tech Museum](https://kostyk348.github.io/lost-tech-museum/) — the archive this project is catalogued in
