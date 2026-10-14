"""
Validate pdf-spec-parser against a set of PDFs and print a verdict per file.

    python -m pdf_spec_parser.validate examples/template_*.pdf
    python -m pdf_spec_parser.validate --examples      # everything in examples/
    python -m pdf_spec_parser.validate in.pdf --out-dir out/

For every input file the validator runs the full pipeline (the websearch
fallback is OFF by default, so runs are deterministic and offline), turns the
result into contract records, validates them, and — when the manifest
(examples/templates_manifest.json, --manifest to point elsewhere) knows the
file — compares the records against the expected items. It then prints one
verdict per template saying exactly where it broke:

    [PASS] examples/template_b_continuation.pdf — 8 records, contract PASS,
           manifest PASS
           table page 1 | row band 13.2 | columns code,desc,item,qty,unit_col
    [FAIL] examples/broken.pdf — no spec items found
           reason: NoItemsError: No spec items found — column detection failed
           (table_page=1, col_map={'item': (0, 575), ...}, 340 text blocks ...)

Exit code: 0 when every file passes, 5 when at least one file fails (either
the pipeline broke, the contract was violated, or a manifest expectation did
not hold), 4 when a template could not be parsed into records at all, 2 for a
usage error.
"""

import argparse
import json
import os
import sys

from . import contract
from .parser import NoItemsError, NoTextError


def default_manifest_path():
    """examples/templates_manifest.json next to the repository root."""
    here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(os.path.dirname(here), 'examples',
                             'templates_manifest.json')
    return candidate if os.path.exists(candidate) else None


def load_manifest(path):
    """Return {file_basename: entry} for the manifest, or {}."""
    if not path or not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as fh:
        data = json.load(fh)
    entries = {}
    for entry in data.get('templates', []):
        entries[entry['file']] = entry
    return entries


def compare_with_manifest(records, entry):
    """
    Compare records against the manifest expectations.

    Returns a list of problem strings (empty = match).
    """
    problems = []
    expected = entry.get('expected_items', [])
    by_pos = {}
    for record in records:
        by_pos.setdefault(record['pos'], []).append(record)

    if len(records) != len(expected):
        problems.append('expected %d records, got %d'
                        % (len(expected), len(records)))

    for want in expected:
        pos = want['pos']
        matches = by_pos.get(pos)
        if not matches:
            problems.append('pos %s missing (expected name contains %r)'
                            % (pos, want.get('name_contains')))
            continue
        if len(matches) > 1:
            problems.append('pos %s appears %d times' % (pos, len(matches)))
        got = matches[0]
        needle = want.get('name_contains')
        if needle and needle.lower() not in got['name'].lower():
            problems.append('pos %s: name %r does not contain %r'
                            % (pos, got['name'], needle))
        if 'category' in want and got['category'] != want['category']:
            problems.append('pos %s: category %r != expected %r'
                            % (pos, got['category'], want['category']))
        if 'code' in want and got['code'] != want['code']:
            problems.append('pos %s: code %r != expected %r'
                            % (pos, got['code'], want['code']))
        if 'qty' in want:
            want_qty = want['qty']
            got_qty = got['qty']
            same = (want_qty == got_qty if want_qty is None
                    else got_qty is not None and abs(got_qty - want_qty) < 1e-6)
            if not same:
                problems.append('pos %s: qty %r != expected %r'
                                % (pos, got_qty, want_qty))
        if 'unit' in want and got['unit'] != want['unit']:
            problems.append('pos %s: unit %r != expected %r'
                            % (pos, got['unit'], want['unit']))

    extra = sorted(set(by_pos) - {w['pos'] for w in expected})
    if extra:
        problems.append('unexpected pos value(s): %s'
                        % ', '.join(str(p) for p in extra))
    return problems


def validate_one(pdf_path, manifest, out_dir=None, websearch=False,
                 write_outputs=True):
    """Run one file through the pipeline; return a verdict dict."""
    name = os.path.basename(pdf_path)
    entry = manifest.get(name)
    result = {'file': pdf_path, 'name': name, 'manifest': bool(entry),
              'records': 0, 'problems': [], 'issues': [], 'stats': {},
              'status': 'FAIL', 'stage': None, 'error': None,
              'written': []}

    try:
        from .parser import parse_pdf
        items = parse_pdf(pdf_path, websearch=websearch, stats=result['stats'])
    except NoTextError as exc:
        result.update(stage='decode', error='%s: %s' % (type(exc).__name__, exc))
        return result
    except NoItemsError as exc:
        result.update(stage='parse', error='%s: %s' % (type(exc).__name__, exc))
        return result
    except Exception as exc:                      # noqa: BLE001 - report it
        result.update(stage='pipeline',
                      error='%s: %s' % (type(exc).__name__, exc))
        return result

    records = contract.to_records(items)
    result['records'] = len(records)
    result['records_data'] = records
    status, issues = contract.verdict(records)
    result['issues'] = issues

    if status == 'FAIL':
        result.update(stage='contract',
                      error='the produced records violate the output contract')
        return result

    if entry:
        problems = compare_with_manifest(records, entry)
        if problems:
            result.update(stage='manifest',
                          error='does not match examples/templates_manifest.json')
            result['problems'] = problems
            return result

    if write_outputs and out_dir:
        os.makedirs(out_dir, exist_ok=True)
        stem = os.path.splitext(name)[0]
        prefix = os.path.join(out_dir, stem)
        document = contract.build_document(records, pdf_path, result['stats'],
                                           issues)
        result['written'] = [contract.write_csv(records, prefix + '.csv'),
                             contract.write_json(document, prefix + '.json')]

    result['status'] = 'PASS'
    return result


def format_report(result):
    """Human readable verdict block for one file."""
    lines = []
    stats = result['stats']
    if result['status'] == 'PASS':
        detail = ['%d records, contract PASS' % result['records']]
        if result['manifest']:
            detail.append('manifest PASS')
        else:
            detail.append('manifest: no expectation for this file')
        lines.append('[PASS] %s — %s' % (result['file'], ', '.join(detail)))
        lines.append('       table page %s | row band %s | columns %s%s'
                     % (stats.get('table_page'), stats.get('row_band'),
                        ', '.join(stats.get('columns', [])) or 'fallback',
                        '' if stats.get('websearch') is False
                        else ' | websearch ON (non-deterministic)'))
        hits = stats.get('column_hits')
        if hits:
            lines.append('       column hits: %s'
                         % ', '.join('%s=%d' % kv for kv in sorted(hits.items())))
        for issue in result['issues']:
            lines.append('       [%s] %s: %s'
                         % (issue['severity'], issue['code'], issue['message']))
        for path in result['written']:
            lines.append('       wrote %s' % path)
        return lines

    lines.append('[FAIL] %s — broken at stage: %s'
                 % (result['file'], result['stage']))
    if result['error']:
        lines.append('       reason: %s' % result['error'])
    if result['stage'] == 'contract':
        for issue in result['issues']:
            if issue['severity'] == 'error':
                lines.append('       [error] %s: %s'
                             % (issue['code'], issue['message']))
    for problem in result['problems']:
        lines.append('       manifest mismatch: %s' % problem)
    if stats:
        lines.append('       diagnostics: blocks=%s pages=%s table_page=%s '
                     'columns=%s hits=%s row_band=%s fallback=%s'
                     % (stats.get('blocks'), stats.get('pages'),
                        stats.get('table_page'),
                        ','.join(stats.get('columns', [])) or '-',
                        stats.get('column_hits'), stats.get('row_band'),
                        stats.get('fallback')))
    return lines


def main(argv=None):
    parser = argparse.ArgumentParser(
        description='Validate pdf-spec-parser output against the contract '
                    '(and against examples/templates_manifest.json).',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m pdf_spec_parser.validate examples/template_a_gost_drawing.pdf
  python -m pdf_spec_parser.validate --examples
  python -m pdf_spec_parser.validate in.pdf --out-dir out/

Exit codes:
  0 every file passed            4 a file produced no records at all
  2 usage error                  5 a file failed (contract/manifest/pipeline)
        """)
    parser.add_argument('pdf', nargs='*', help='PDF file(s) to validate')
    parser.add_argument('--examples', action='store_true',
                        help='validate every PDF in examples/')
    parser.add_argument('--manifest', default=None,
                        help='path to templates_manifest.json '
                             '(default: examples/templates_manifest.json)')
    parser.add_argument('--out-dir', default=None,
                        help='also write .csv/.json for every passing file')
    parser.add_argument('--websearch', action='store_true',
                        help='allow the (non-deterministic) websearch fallback')
    args = parser.parse_args(argv)

    pdfs = list(args.pdf)
    if args.examples:
        here = os.path.dirname(os.path.abspath(__file__))
        examples_dir = os.path.join(os.path.dirname(here), 'examples')
        if not os.path.isdir(examples_dir):
            print('Error: no examples/ directory next to the package',
                  file=sys.stderr)
            return contract.EXIT_USAGE
        pdfs.extend(sorted(
            os.path.join(examples_dir, f) for f in os.listdir(examples_dir)
            if f.lower().endswith('.pdf')))
    if not pdfs:
        parser.error('no input: pass PDF paths or --examples')

    missing = [p for p in pdfs if not os.path.exists(p)]
    if missing:
        print('Error: file not found: %s' % ', '.join(missing), file=sys.stderr)
        return contract.EXIT_USAGE

    manifest_path = args.manifest or default_manifest_path()
    manifest = load_manifest(manifest_path)
    if not manifest:
        print('note: no manifest found (%s) — checking the contract only'
              % (manifest_path or 'none'), file=sys.stderr)

    results = [validate_one(pdf, manifest, args.out_dir, args.websearch)
               for pdf in pdfs]

    for result in results:
        for line in format_report(result):
            print(line)
        print()

    passed = [r for r in results if r['status'] == 'PASS']
    summary = 'summary: %d/%d templates PASS' % (len(passed), len(results))
    print(summary)
    if len(passed) == len(results):
        return contract.EXIT_OK
    if any(r['stage'] == 'parse' or r['stage'] == 'decode' for r in results):
        return contract.EXIT_NO_ITEMS
    return contract.EXIT_CONTRACT


if __name__ == '__main__':
    sys.exit(main())
