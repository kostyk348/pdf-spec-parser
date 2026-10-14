"""
pdf_spec_parser — Universal procurement PDF spec parser & obfuscator.

Extracts specification tables from GOST-font vector PDFs,
classifies components, removes brand data, and exports to
formatted Excel + PDF, plus a documented CSV/JSON output contract.

Usage:
    from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf
    items = parse_pdf("drawing.pdf")
    export_xlsx(items, "output.xlsx")
    export_pdf(items, "output.pdf")

Machine-readable output (see pdf_spec_parser.contract for the full schema,
field types, flags and exit codes):
    from pdf_spec_parser import contract
    records = contract.to_records(items)
    contract.write_csv(records, "spec.csv")
    status, issues = contract.verdict(records)
"""

__version__ = "1.1.0"

from . import contract                                   # noqa: E402
from .parser import (CATEGORY_VOCAB, FLAG_VOCAB, NoItemsError,  # noqa: E402
                     NoTextError, PdfSpecError, classify_component,
                     export_pdf, export_xlsx, format_item, parse_pdf)
from .contract import (SCHEMA_VERSION, build_document, to_records,  # noqa: E402
                       validate_records, verdict, write_csv, write_json)

__all__ = [
    "parse_pdf", "export_xlsx", "export_pdf", "format_item",
    "classify_component", "CATEGORY_VOCAB", "FLAG_VOCAB",
    "PdfSpecError", "NoTextError", "NoItemsError",
    "contract", "SCHEMA_VERSION", "to_records", "validate_records", "verdict",
    "build_document", "write_csv", "write_json",
]
