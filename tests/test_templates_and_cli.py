"""
End-to-end tests: the three committed example templates, the CLI contract
(exit codes, output formats) and the validate module.

These tests need the runtime dependencies (pdfplumber, openpyxl, reportlab),
i.e. exactly what `pip install -e .` provides.

    python -m unittest discover -s tests -v
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXAMPLES = os.path.join(ROOT, 'examples')
MANIFEST = os.path.join(EXAMPLES, 'templates_manifest.json')
sys.path.insert(0, ROOT)

from pdf_spec_parser import contract, parse_pdf        # noqa: E402
from pdf_spec_parser import validate as V              # noqa: E402

MANIFEST_DATA = {}
if os.path.exists(MANIFEST):
    with open(MANIFEST, encoding='utf-8') as _fh:
        MANIFEST_DATA = json.load(_fh)
TEMPLATES = MANIFEST_DATA.get('templates', [])
TEMPLATE_IDS = [t['id'] for t in TEMPLATES]


def run_cli(*args, cwd=ROOT):
    """Run the package CLI in a subprocess and return (code, stdout, stderr)."""
    env = dict(os.environ)
    env['PYTHONPATH'] = ROOT + os.pathsep + env.get('PYTHONPATH', '')
    proc = subprocess.run([sys.executable, '-m', 'pdf_spec_parser', *args],
                          cwd=cwd, env=env, capture_output=True, text=True,
                          timeout=300)
    return proc.returncode, proc.stdout, proc.stderr


@unittest.skipUnless(TEMPLATES, 'examples/templates_manifest.json not present')
class TemplateTest(unittest.TestCase):
    """Every expected item of every committed template must be parsed."""

    def test_manifest_pins_the_committed_pdfs(self):
        for entry in TEMPLATES:
            path = os.path.join(EXAMPLES, entry['file'])
            self.assertTrue(os.path.exists(path), entry['file'])
            with open(path, 'rb') as fh:
                digest = hashlib.sha256(fh.read()).hexdigest()
            self.assertEqual(digest, entry['sha256'],
                             '%s changed — re-run examples/make_templates.py '
                             'and commit the manifest' % entry['file'])

    def test_three_templates_with_distinct_layouts_are_present(self):
        self.assertGreaterEqual(len(TEMPLATES), 3)
        layouts = {t['layout'] for t in TEMPLATES}
        self.assertEqual(len(layouts), len(TEMPLATES),
                         'templates must differ in layout')

    def test_every_template_parses_into_contract_records(self):
        for entry in TEMPLATES:
            path = os.path.join(EXAMPLES, entry['file'])
            with self.subTest(template=entry['id']):
                stats = {}
                items = parse_pdf(path, websearch=False, stats=stats)
                records = contract.to_records(items)
                status, issues = contract.verdict(records)
                errors = [i for i in issues if i['severity'] == 'error']
                self.assertEqual(status, 'PASS', errors)
                self.assertEqual(len(records),
                                 len(entry['expected_items']),
                                 'record count')
                self.assertEqual([r['pos'] for r in records],
                                 sorted(r['pos'] for r in records))
                self.assertIn('table_page', stats)
                self.assertIn('columns', stats)

    def test_every_template_matches_the_manifest(self):
        for entry in TEMPLATES:
            path = os.path.join(EXAMPLES, entry['file'])
            with self.subTest(template=entry['id']):
                items = parse_pdf(path, websearch=False)
                records = contract.to_records(items)
                problems = V.compare_with_manifest(records, entry)
                self.assertEqual(problems, [], problems)

    def test_validator_reports_pass_for_all_examples(self):
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-validate-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        code = V.main(['--examples', '--out-dir', out_dir])
        self.assertEqual(code, 0)
        written = sorted(os.listdir(out_dir))
        self.assertEqual(len(written), 2 * len(TEMPLATES))
        self.assertIn('template_b_continuation.csv', written)
        self.assertIn('template_a_gost_drawing.json', written)

    def test_validator_writes_contract_conformant_csv(self):
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-validate-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        V.main(['--examples', '--out-dir', out_dir])
        import csv
        path = os.path.join(out_dir, 'template_b_continuation.csv')
        with open(path, encoding='utf-8-sig', newline='') as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(len(rows), 8)
        self.assertEqual(list(rows[0]), list(contract.FIELDS))
        # pos 8 is the "комплект" row: no number, no unit, flagged.
        last = [r for r in rows if r['pos'] == '8'][0]
        self.assertEqual(last['qty'], '')
        self.assertEqual(last['unit'], '')
        self.assertIn('UNPARSED_QTY', last['flags'])

    def test_validator_reports_a_broken_template_with_a_reason(self):
        """A PDF without a recognisable table must FAIL with a diagnosis."""
        from reportlab.pdfgen import canvas
        tmp = tempfile.mkdtemp(prefix='pdf-spec-broken-')
        self.addCleanup(shutil.rmtree, tmp, True)
        path = os.path.join(tmp, 'no_table.pdf')
        c = canvas.Canvas(path)
        c.drawString(72, 700, 'Annotation only, no spec table here')
        c.showPage()
        c.save()

        result = V.validate_one(path, {}, None)
        self.assertEqual(result['status'], 'FAIL')
        self.assertIn(result['stage'], ('parse', 'contract'))
        report = '\n'.join(V.format_report(result))
        self.assertIn('[FAIL]', report)
        self.assertIn('reason:', report)


class CliTest(unittest.TestCase):

    def test_module_invocation_works(self):
        """`python -m pdf_spec_parser` needs a __main__ module to exist."""
        code, out, err = run_cli('--help')
        self.assertEqual(code, 0, err)
        self.assertIn('--format', out)

    def test_cli_writes_all_four_formats(self):
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-cli-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        pdf = os.path.join(EXAMPLES, 'template_a_gost_drawing.pdf')
        code, out, err = run_cli(pdf, '--out-dir', out_dir, '--no-websearch')
        self.assertEqual(code, 0, err)
        for suffix in ('.csv', '.json', '.xlsx', '.pdf'):
            # default prefix keeps the documented "_anon" suffix
            self.assertTrue(
                os.path.exists(os.path.join(
                    out_dir, 'template_a_gost_drawing_anon' + suffix)),
                suffix)
        self.assertIn('contract: PASS', out)
        self.assertIn('records: 5', out)

    def test_cli_explicit_prefix_wins(self):
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-cli-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        pdf = os.path.join(EXAMPLES, 'template_a_gost_drawing.pdf')
        code, _out, err = run_cli(pdf, os.path.join(out_dir, 'my_spec'),
                                  '--format', 'csv')
        self.assertEqual(code, 0, err)
        self.assertEqual(sorted(os.listdir(out_dir)), ['my_spec.csv'])

    def test_cli_format_selection(self):
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-cli-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        pdf = os.path.join(EXAMPLES, 'template_c_board_spec.pdf')
        code, _out, err = run_cli(pdf, '--out-dir', out_dir,
                                  '--format', 'csv,json', '--no-websearch')
        self.assertEqual(code, 0, err)
        self.assertEqual(sorted(os.listdir(out_dir)),
                         ['template_c_board_spec_anon.csv',
                          'template_c_board_spec_anon.json'])

    def test_cli_json_matches_the_documented_shape(self):
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-cli-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        pdf = os.path.join(EXAMPLES, 'template_b_continuation.pdf')
        code, _out, err = run_cli(pdf, '--out-dir', out_dir, '--no-websearch')
        self.assertEqual(code, 0, err)
        with open(os.path.join(out_dir, 'template_b_continuation_anon.json'),
                  encoding='utf-8') as fh:
            document = json.load(fh)
        self.assertEqual(document['schema_version'], '1.0')
        self.assertTrue(document['generator'].startswith('pdf-spec-parser '))
        self.assertEqual(document['counts']['records'], 8)
        self.assertEqual(len(document['source']['sha256']), 64)
        self.assertFalse(document['source']['websearch'])
        self.assertEqual(list(document['items'][0]), list(contract.FIELDS))

    def test_cli_rejects_an_unknown_format(self):
        pdf = os.path.join(EXAMPLES, 'template_a_gost_drawing.pdf')
        code, _out, err = run_cli(pdf, '--format', 'yaml')
        self.assertEqual(code, contract.EXIT_USAGE)
        self.assertIn('--format', err)

    def test_cli_missing_file_exit_code(self):
        code, _out, err = run_cli('/nonexistent/nope.pdf')
        self.assertEqual(code, contract.EXIT_USAGE)
        self.assertIn('File not found', err)

    def test_cli_no_text_exit_code(self):
        from reportlab.pdfgen import canvas
        tmp = tempfile.mkdtemp(prefix='pdf-spec-empty-')
        self.addCleanup(shutil.rmtree, tmp, True)
        path = os.path.join(tmp, 'empty.pdf')
        canvas.Canvas(path).save()          # a PDF page without any text
        code, _out, err = run_cli(path, '--out-dir', tmp)
        self.assertEqual(code, contract.EXIT_NO_TEXT)
        self.assertIn('no-text', err)

    def test_cli_no_items_exit_code(self):
        """Text, but no table: exit code 4 and a diagnostic message."""
        from reportlab.pdfgen import canvas
        tmp = tempfile.mkdtemp(prefix='pdf-spec-annot-')
        self.addCleanup(shutil.rmtree, tmp, True)
        path = os.path.join(tmp, 'annotations.pdf')
        c = canvas.Canvas(path)
        for i in range(6):
            c.drawString(72, 700 - i * 20, 'Пояснение %d к чертежу' % i)
        c.showPage()
        c.save()
        code, _out, err = run_cli(path, '--out-dir', tmp)
        self.assertEqual(code, contract.EXIT_NO_ITEMS, err)
        self.assertIn('no-items', err)
        self.assertIn('column detection failed', err)

    def test_cli_contract_violation_exit_code(self):
        """A stray number inside the position column must not pass silently."""
        out_dir = tempfile.mkdtemp(prefix='pdf-spec-violation-')
        self.addCleanup(shutil.rmtree, out_dir, True)
        records = [contract.normalize_item(
            {'pos': 1, 'category': 'Прочее', 'name': '', 'qty_raw': ''})]
        status, issues = contract.verdict(records)
        self.assertEqual(status, 'FAIL')
        self.assertEqual(contract.EXIT_CONTRACT, 5)
        self.assertIn('E_NAME_EMPTY', [i['code'] for i in issues])


class PackageTest(unittest.TestCase):

    def test_public_api_is_importable(self):
        import pdf_spec_parser as pkg
        for name in ('parse_pdf', 'export_xlsx', 'export_pdf', 'contract',
                     'to_records', 'verdict', 'write_csv', 'write_json'):
            self.assertTrue(hasattr(pkg, name), name)
        self.assertEqual(pkg.__version__, '1.1.0')

    def test_exit_codes_are_the_documented_numbers(self):
        self.assertEqual(contract.EXIT_OK, 0)
        self.assertEqual(contract.EXIT_INTERNAL, 1)
        self.assertEqual(contract.EXIT_USAGE, 2)
        self.assertEqual(contract.EXIT_NO_TEXT, 3)
        self.assertEqual(contract.EXIT_NO_ITEMS, 4)
        self.assertEqual(contract.EXIT_CONTRACT, 5)
        self.assertEqual(contract.EXIT_NO_OUTPUT, 6)

    def test_contract_module_docstring_documents_the_schema(self):
        doc = contract.__doc__ or ''
        for token in ('RECORD SCHEMA', 'CSV LAYOUT', 'JSON LAYOUT',
                      'EXIT CODES', 'UNRECOGNISED LINES'):
            self.assertIn(token, doc, token)
        for field in contract.FIELDS:
            self.assertIn(field, doc, field)


if __name__ == '__main__':
    unittest.main()
