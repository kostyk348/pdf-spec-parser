# pdf-spec-parser

**Universal procurement PDF specification parser & obfuscator.**

Extracts specification tables from GOST-font vector PDF drawings, automatically classifies components by type, removes brand/proprietary information, and exports to clean formatted Excel + PDF.

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
│ 2. Scan last      │ ← Finds "Поз", "Наименование", "Кол" headers
│    pages for      │
│    spec table     │
└────────┬──────────┘
         ▼
┌───────────────────┐
│ 3. Parse items    │ ← Maps text blocks to columns by X coordinate
│    & quantities   │
└────────┬──────────┘
         ▼
┌─────────────────────────┐
│ 4. Classify components  │ ← Domain knowledge + websearch fallback
│    Remove brand names   │
└────────┬────────────────┘
         ▼
┌───────────────────┐
│ 5. Export         │ ← Formatted Excel (.xlsx) + clean PDF
│    Excel + PDF    │
└───────────────────┘
```

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

# Install with pip
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

## Usage

### Command-line

```bash
# Basic usage — output files get `_anon` suffix
python -m pdf_spec_parser drawing.pdf

# Custom output prefix
python -m pdf_spec_parser drawing.pdf output_prefix

# Verbose mode (shows progress)
python -m pdf_spec_parser drawing.pdf --verbose

# Full example
python -m pdf_spec_parser "Блок WA.00.000.pdf" "Блок_WA_spec" --verbose
```

### Python API

```python
from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf

# Parse a PDF
items = parse_pdf("drawing.pdf", verbose=True)

# items is a list of dicts:
# [
#   {
#     'pos': 1,
#     'category': 'Коммутационное оборудование',
#     'name': 'Реле промежуточное электромагнитное',
#     'params': '4 переключающих контакта (4PDT), Ток контактов: 3 А, ...',
#     'qty': 3.0,
#     'unit': 'шт',
#   },
#   ...
# ]

# Export separately
export_xlsx(items, "output.xlsx", cabinet_name="Блок WA.00.000")
export_pdf(items, "output.pdf", title_text="Блок WA.00.000")
```

## Output Format

Both Excel and PDF produce a 6-column table:

| № | Категория | Наименование (обезличенное) | Параметры / Номиналы | Кол-во | Ед. изм. |
|---|-----------|----------------------------|----------------------|--------|----------|
| 1 | Коммутационное оборудование | Реле промежуточное электромагнитное | 4PDT, 3 А, ~220 В AC | 3 | шт |
| 2 | Провода и кабели | Провод монтажный термостойкий | Сечение: 0,5 мм², цвет: белый | 1.77 | м |

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

## License

Apache 2.0

## See Also

- [SINT-UA project family](https://github.com/kostyk348?tab=repositories) — related tooling for deterministic systems optimization
