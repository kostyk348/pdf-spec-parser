"""
Allow `python -m pdf_spec_parser ...`.

The README documents the module-level CLI, so the package needs a __main__.
"""

import sys

from .parser import main

if __name__ == '__main__':
    sys.exit(main())
