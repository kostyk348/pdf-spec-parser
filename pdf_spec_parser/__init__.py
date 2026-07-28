"""
pdf_spec_parser — Universal procurement PDF spec parser & obfuscator.

Extracts specification tables from GOST-font vector PDFs,
classifies components, removes brand data, and exports to
formatted Excel + PDF.

Usage:
    from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf
    items = parse_pdf("drawing.pdf")
    export_xlsx(items, "output.xlsx")
    export_pdf(items, "output.pdf")
"""

from .parser import parse_pdf, export_xlsx, export_pdf

__version__ = "1.0.0"
__all__ = ["parse_pdf", "export_xlsx", "export_pdf"]
