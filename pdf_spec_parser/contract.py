"""
pdf_spec_parser.contract — the documented output contract of pdf-spec-parser.

The parser internally produces permissive item dicts. This module is the
*contract layer*: it normalises items into records with fixed field names and
types, writes them as CSV/JSON, and validates them. Anything a consumer reads
automatically (a procurement system, a spreadsheet import, a test) should go
through the contract layer.

──────────────────────────────────────────────────────────────────────
RECORD SCHEMA (schema_version 1.0)
──────────────────────────────────────────────────────────────────────

Every record has exactly these nine fields, in this order:

  field      type            meaning / rules
  ---------  --------------  ------------------------------------------------
  pos        integer         position number from the drawing, 1..999,
                             unique and ascending within one file
  category   string          component class, one of CATEGORY_VOCAB
                             (7 classes + 'Прочее'); never empty
  name       string          de-branded component name; never empty
  code       string          drawing designation (e.g. 'ВА 47-29'); ''
                             when the table has no designation column
  params     string          parameters/ratings, comma separated; '' when
                             nothing is known
  qty        number | null   quantity as a number; null when the quantity
                             cell holds no number (see qty_raw + flags)
  unit       string | null   unit of measure; null if and only if qty is null
  qty_raw    string          the quantity cell exactly as read from the PDF
  flags      array[string]   subset of FLAG_VOCAB; [] for a confident row

Documented normalisation rules applied while building records:

  * qty is never a string. A cell like "комплект" yields qty=null,
    qty_raw="комплект", unit=null and the flag UNPARSED_QTY. An empty cell
    yields qty=null and the flag NO_QTY.
  * The unit column of the drawing wins over the unit guessed from the
    quantity cell. cm → m conversion is reported with QTY_CONVERTED.
  * name/params/code/qty_raw are always strings (never null) — '' means
    "the drawing has nothing here".
  * flags is always a list; unknown flag names are a contract violation.

UNRECOGNISED LINES — what happens to them:

  * Lines outside every detected column (drawing annotations, title block,
    footers) are ignored. They never appear in the output; they are counted
    in diagnostics.skipped_blocks.
  * A table line without a position number produces no record (records are
    built from position numbers, the only anchor the format guarantees).
    The count of such leftover lines is filed as a W_ORPHAN_LINE warning.
  * A component the knowledge base does not know is still emitted: the row
    keeps the raw name, gets the flag UNKNOWN_COMPONENT, and its params come
    either from the websearch fallback (flag WEB_INFO, non-deterministic) or
    are the literal string '(требуется уточнение)'. Run with
    websearch=False (CLI: --no-websearch) for a deterministic result.
  * Nothing is ever dropped silently: every emitted row carries its own
    flags, and the run diagnostics carry the aggregate counters.

──────────────────────────────────────────────────────────────────────
CSV LAYOUT
──────────────────────────────────────────────────────────────────────

  * UTF-8 **with BOM** (`utf-8-sig`) so Excel opens Cyrillic correctly.
  * One header row with the nine field names above, then one row per record.
  * `\n` line terminator, `,` separator, RFC 4180 quoting (csv module).
  * flags are joined with ';' inside their cell.
  * qty is written with a dot decimal separator; null values are empty cells
    (pair that with unit == '').

──────────────────────────────────────────────────────────────────────
JSON LAYOUT
──────────────────────────────────────────────────────────────────────

  {
    "schema_version": "1.0",
    "generator": "pdf-spec-parser 1.0.0",
    "source": {"file": "...", "sha256": "...", "pages": 1, "blocks": 335,
               "table_page": 1, "columns": [...], "row_band": 72.0,
               "fallback": null, "websearch": false},
    "counts": {"records": 5, "flags": {"UNKNOWN_COMPONENT": 1}},
    "items": [ { ...nine fields... } ],
    "diagnostics": {"issues": [ ...validator issues... ]}
  }

──────────────────────────────────────────────────────────────────────
EXIT CODES (CLI)
──────────────────────────────────────────────────────────────────────

  0  success
  1  unexpected internal error (bug — please report)
  2  usage error (bad arguments, input file not found / not a PDF)
  3  no text could be extracted from the PDF
  4  text extracted, but no spec table rows recognised
  5  produced records violate the output contract (see diagnostics.issues)
  6  nothing to write (no output format selected)
"""

import csv
import hashlib
import json
import os

from . import __version__
from .parser import CATEGORY_VOCAB, FLAG_VOCAB, clean_text

SCHEMA_VERSION = '1.0'

FIELDS = ('pos', 'category', 'name', 'code', 'params', 'qty', 'unit',
          'qty_raw', 'flags')

FLAG_SEPARATOR = ';'
CSV_ENCODING = 'utf-8-sig'
JSON_ENCODING = 'utf-8'

# ── exit codes ────────────────────────────────────────────────────────
EXIT_OK = 0
EXIT_INTERNAL = 1
EXIT_USAGE = 2
EXIT_NO_TEXT = 3
EXIT_NO_ITEMS = 4
EXIT_CONTRACT = 5
EXIT_NO_OUTPUT = 6

FORMATS = ('csv', 'json', 'xlsx', 'pdf')


# ──────────────────────────────────────────────────────────────────────
# NORMALISATION
# ──────────────────────────────────────────────────────────────────────

def _as_text(value):
    """Coerce anything to a string; None becomes ''."""
    if value is None:
        return ''
    return clean_text(str(value))


def normalize_item(item: dict) -> dict:
    """
    Coerce a parser item into a contract-conformant record.

    Type coercion only — this never invents data. A quantity cell that holds
    no number becomes qty=None and gains the matching flag, so the reason is
    visible in the output instead of being hidden in a string field.
    """
    flags = [f for f in (item.get('flags') or []) if f in FLAG_VOCAB]
    qty = item.get('qty')
    if isinstance(qty, bool) or not isinstance(qty, (int, float)):
        if qty is not None:
            # A non-numeric quantity the formatter did not flag: flag it here
            # rather than passing a string off as a quantity.
            if 'UNPARSED_QTY' not in flags:
                flags.append('UNPARSED_QTY')
        qty = None
    elif float(qty).is_integer():
        qty = int(qty)
    else:
        qty = float(qty)

    unit = _as_text(item.get('unit'))
    if qty is None:
        unit = None

    pos = item.get('pos')
    if isinstance(pos, str) and pos.strip().isdigit():
        pos = int(pos.strip())

    return {
        'pos': pos,
        'category': _as_text(item.get('category')),
        'name': _as_text(item.get('name')),
        'code': _as_text(item.get('code')),
        'params': _as_text(item.get('params')),
        'qty': qty,
        'unit': unit,
        'qty_raw': _as_text(item.get('qty_raw')),
        'flags': flags,
    }


def to_records(items) -> list:
    """Normalise a list of parser items into contract records."""
    return [normalize_item(item) for item in items]


# ──────────────────────────────────────────────────────────────────────
# VALIDATION
# ──────────────────────────────────────────────────────────────────────

def _issue(code, severity, message, pos=None, field=None):
    issue = {'code': code, 'severity': severity, 'message': message}
    if pos is not None:
        issue['pos'] = pos
    if field is not None:
        issue['field'] = field
    return issue


def validate_records(records) -> list:
    """
    Validate records against the output contract.

    Returns a list of issues, each
    {'code', 'severity' ('error'|'warning'), 'message', 'pos'?, 'field'?}.
    An empty list means the records conform. Only 'error' issues make the
    verdict FAIL — the CLI exits with EXIT_CONTRACT then.
    """
    issues = []
    if not isinstance(records, list):
        return [_issue('E_CONTRACT_TYPE', 'error',
                       'records must be a list, got %s' % type(records).__name__)]

    seen_pos = set()
    previous_pos = None
    for index, record in enumerate(records):
        where = 'record #%d' % (index + 1)
        if not isinstance(record, dict):
            issues.append(_issue('E_CONTRACT_TYPE', 'error',
                                 '%s is not an object' % where))
            continue

        missing = [f for f in FIELDS if f not in record]
        extra = [f for f in record if f not in FIELDS]
        if missing:
            issues.append(_issue('E_SCHEMA_FIELD_MISSING', 'error',
                                 '%s is missing field(s): %s'
                                 % (where, ', '.join(missing))))
        if extra:
            issues.append(_issue('E_SCHEMA_FIELD_UNKNOWN', 'error',
                                 '%s has undocumented field(s): %s'
                                 % (where, ', '.join(extra))))
        if missing:
            continue

        pos = record['pos']
        if isinstance(pos, bool) or not isinstance(pos, int):
            issues.append(_issue('E_POS_TYPE', 'error',
                                 '%s: pos must be an integer, got %r'
                                 % (where, pos), field='pos'))
        elif not 1 <= pos <= 999:
            issues.append(_issue('E_POS_RANGE', 'error',
                                 '%s: pos %r outside 1..999' % (where, pos),
                                 pos=pos, field='pos'))
        else:
            if pos in seen_pos:
                issues.append(_issue('E_POS_DUPLICATE', 'error',
                                     'pos %d appears more than once' % pos,
                                     pos=pos, field='pos'))
            if previous_pos is not None and pos < previous_pos:
                issues.append(_issue('E_POS_ORDER', 'error',
                                     'pos %d breaks ascending order (after %d)'
                                     % (pos, previous_pos), pos=pos,
                                     field='pos'))
            previous_pos = pos
            seen_pos.add(pos)

        category = record['category']
        if not isinstance(category, str) or not category:
            issues.append(_issue('E_CATEGORY_EMPTY', 'error',
                                 '%s: category is empty' % where,
                                 pos=pos if isinstance(pos, int) else None,
                                 field='category'))
        elif category not in CATEGORY_VOCAB:
            issues.append(_issue('E_CATEGORY_UNKNOWN', 'error',
                                 '%s: category %r outside the vocabulary (%s)'
                                 % (where, category, ', '.join(CATEGORY_VOCAB)),
                                 pos=pos if isinstance(pos, int) else None,
                                 field='category'))

        if not isinstance(record['name'], str) or not record['name'].strip():
            issues.append(_issue('E_NAME_EMPTY', 'error',
                                 '%s: name is empty — nothing was recognised '
                                 'in the name column' % where,
                                 pos=pos if isinstance(pos, int) else None,
                                 field='name'))

        for field in ('code', 'params', 'qty_raw'):
            if not isinstance(record[field], str):
                issues.append(_issue('E_FIELD_TYPE', 'error',
                                     '%s: %s must be a string, got %s'
                                     % (where, field,
                                        type(record[field]).__name__),
                                     field=field))

        qty = record['qty']
        if qty is not None:
            if isinstance(qty, bool) or not isinstance(qty, (int, float)):
                issues.append(_issue('E_QTY_TYPE', 'error',
                                     '%s: qty must be a number or null, got %r'
                                     % (where, qty), field='qty'))
            elif qty < 0:
                issues.append(_issue('E_QTY_NEGATIVE', 'error',
                                     '%s: qty %r is negative' % (where, qty),
                                     field='qty'))

        unit = record['unit']
        if qty is None and unit is not None:
            issues.append(_issue('E_UNIT_WITHOUT_QTY', 'error',
                                 '%s: unit %r set while qty is null'
                                 % (where, unit), field='unit'))
        if qty is not None and (not isinstance(unit, str) or not unit.strip()):
            issues.append(_issue('E_QTY_WITHOUT_UNIT', 'error',
                                 '%s: qty %r without a unit of measure'
                                 % (where, qty), field='unit'))

        flags = record['flags']
        if not isinstance(flags, list):
            issues.append(_issue('E_FLAGS_TYPE', 'error',
                                 '%s: flags must be a list, got %s'
                                 % (where, type(flags).__name__), field='flags'))
        else:
            unknown = [f for f in flags if f not in FLAG_VOCAB]
            if unknown:
                issues.append(_issue('E_FLAG_UNKNOWN', 'error',
                                     '%s: undocumented flag(s): %s'
                                     % (where, ', '.join(map(str, unknown))),
                                     field='flags'))
            if 'UNKNOWN_COMPONENT' in flags:
                issues.append(_issue('W_UNKNOWN_COMPONENT', 'warning',
                                     'pos %r: component not in the knowledge '
                                     'base, params need a human look'
                                     % record['pos'],
                                     pos=pos if isinstance(pos, int) else None))
            if 'WEB_INFO' in flags:
                issues.append(_issue('W_NONDETERMINISTIC', 'warning',
                                     'pos %r: params came from the websearch '
                                     'fallback (non-deterministic)'
                                     % record['pos'],
                                     pos=pos if isinstance(pos, int) else None))

    return issues


def verdict(records):
    """Return ('PASS'|'FAIL', issues) for a list of records."""
    issues = validate_records(records)
    failed = any(i['severity'] == 'error' for i in issues)
    return ('FAIL' if failed else 'PASS'), issues


# ──────────────────────────────────────────────────────────────────────
# DOCUMENT BUILDING & WRITING
# ──────────────────────────────────────────────────────────────────────

def file_sha256(path: str) -> str:
    """SHA-256 of a file, streamed (PDFs can be large)."""
    digest = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            digest.update(chunk)
    return digest.hexdigest()


def build_document(records, source_path: str = None, stats: dict = None,
                   issues=None) -> dict:
    """Build the JSON document around a list of records."""
    source = {'file': os.path.basename(source_path) if source_path else None,
              'sha256': file_sha256(source_path) if source_path else None}
    for key in ('pages', 'blocks', 'table_page', 'columns', 'column_hits',
                'row_band', 'fallback', 'websearch'):
        if stats and key in stats:
            source[key] = stats[key]

    flag_counts = {}
    for record in records:
        for flag in record['flags']:
            flag_counts[flag] = flag_counts.get(flag, 0) + 1

    return {
        'schema_version': SCHEMA_VERSION,
        'generator': 'pdf-spec-parser %s' % __version__,
        'source': source,
        'counts': {'records': len(records),
                   'flags': dict(sorted(flag_counts.items()))},
        'items': records,
        'diagnostics': {'issues': list(issues or [])},
    }


def write_csv(records, path: str) -> str:
    """Write records as the contract CSV (UTF-8 with BOM)."""
    with open(path, 'w', encoding=CSV_ENCODING, newline='') as fh:
        writer = csv.writer(fh, lineterminator='\n')
        writer.writerow(FIELDS)
        for record in records:
            row = []
            for field in FIELDS:
                value = record.get(field)
                if field == 'flags':
                    row.append(FLAG_SEPARATOR.join(value or []))
                elif value is None:
                    row.append('')
                else:
                    row.append(value)
            writer.writerow(row)
    return path


def write_json(document, path: str) -> str:
    """Write the JSON contract document."""
    with open(path, 'w', encoding=JSON_ENCODING) as fh:
        json.dump(document, fh, ensure_ascii=False, indent=2, sort_keys=False)
        fh.write('\n')
    return path
