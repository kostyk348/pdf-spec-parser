#!/usr/bin/env python3
"""
pdf_spec_parser — Universal procurement PDF spec parser & obfuscator.

Auto-detects PDF font type (standard vs GOST CID), extracts spec tables,
classifies components by domain knowledge + websearch fallback,
removes brand names, and exports to formatted Excel + PDF.

CLI:
    python -m pdf_spec_parser input.pdf [output_prefix] [--verbose]

Python API:
    from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf
    items = parse_pdf("input.pdf", verbose=True)
    export_xlsx(items, "output.xlsx")
    export_pdf(items, "output.pdf")
"""

import sys
import os
import re
import json
import logging

log = logging.getLogger(__name__)


# ──────────────────────────────────────────────────────────────────────
# ERRORS
#
# Both exceptions subclass the historical exception types, so existing
# `except ValueError` / `except RuntimeError` code keeps working, while the
# CLI can map each failure to a documented exit code (see contract.py).
# ──────────────────────────────────────────────────────────────────────

class PdfSpecError(Exception):
    """Base class for every expected pdf-spec-parser failure."""


class NoTextError(RuntimeError, PdfSpecError):
    """No text could be extracted from the PDF (scan without a text layer?)."""


class NoItemsError(ValueError, PdfSpecError):
    """Text was extracted, but no spec table rows could be recognised."""


# ──────────────────────────────────────────────────────────────────────
# UTILITIES
# ──────────────────────────────────────────────────────────────────────

def clean_text(text: str) -> str:
    """Clean PDF escaping, brackets, and normalize whitespace."""
    text = text.replace('\\/', '/').replace('\\(', '(').replace('\\)', ')')
    text = text.replace('\\ ', ' ')
    text = re.sub(r'\\(\s|$)', r'\1', text)
    text = re.sub(r'\[([^\]]*)\]', r'\1', text)  # strip decode_all.py brackets
    text = re.sub(r'\s+', ' ', text)
    return text.strip().strip(',;:').strip()


def normalize_num(s: str):
    """Convert comma-decimal string to float."""
    try:
        return float(s.replace(',', '.'))
    except ValueError:
        return None


# ──────────────────────────────────────────────────────────────────────
# 1. PDF DECODER — pdfplumber main, GOST CID fallback
# ──────────────────────────────────────────────────────────────────────

def decode_with_pdfplumber(pdf_path: str):
    """
    Extract text blocks using pdfplumber.

    One block per *word* — a per-character stream (the previous behaviour)
    destroys the table geometry: column ranges are matched per block, so
    characters of one cell scatter over several columns, and every row is
    rebuilt as "Р е л е" (characters joined by spaces). Word blocks keep the
    coordinate contract of the CID decoder: (page, x_left, y_top, text).
    """
    import pdfplumber
    blocks = []
    with pdfplumber.open(pdf_path) as pdf:
        for pn, page in enumerate(pdf.pages, 1):
            for word in page.extract_words(use_text_flow=False,
                                           keep_blank_chars=False):
                text = (word.get('text') or '').strip()
                if text:
                    blocks.append((pn, round(word['x0']), round(word['top']),
                                   text))
    return blocks


def decode_with_cid(pdf_path: str, font_dir: str = None):
    """Extract text using GOST CID decoder for Identity-H fonts."""
    from fontTools.ttLib import TTFont
    import fitz

    if font_dir is None:
        font_dir = os.path.join(os.path.dirname(__file__), '..', 'fonts')
    # Search common locations for the GOST font
    search_paths = [
        os.path.join(font_dir, 'GOSTTypeA_2.ttf'),
        os.path.expanduser(r'~\Downloads\fonts\GOSTTypeA_2.ttf'),
    ]
    font_path = next((p for p in search_paths if os.path.exists(p)), None)
    if not font_path:
        return None

    font = TTFont(font_path)
    glyph_order = font.getGlyphOrder()
    cid_to_chr = {}
    for table in font['cmap'].tables:
        if table.platformID == 3 and table.platEncID == 1:
            for code, name in table.cmap.items():
                if name in glyph_order:
                    gid = glyph_order.index(name)
                    if 0x0400 <= code <= 0x04FF:
                        cid_to_chr[gid] = chr(code)
                    elif gid not in cid_to_chr:
                        cid_to_chr[gid] = chr(code)
    cid_to_chr[0x03] = ' '
    cid_to_chr[0x11] = ''
    cid_to_chr[0x72] = 'ё'
    for c in range(0x20, 0x7F):
        if c not in cid_to_chr:
            cid_to_chr[c] = chr(c)

    def decode_hex_chunk(hex_str):
        parts = []
        for i in range(0, len(hex_str), 4):
            if i + 4 <= len(hex_str):
                cid = int(hex_str[i:i+4], 16)
                parts.append(cid_to_chr.get(cid, '?'))
        return ''.join(parts)

    doc = fitz.open(pdf_path)
    blocks = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        for xref in page.get_contents():
            stream = doc.xref_stream(xref)
            if not stream:
                continue
            try:
                text = stream.decode('latin-1', errors='replace')
            except Exception:
                continue
            for match in re.finditer(r'BT(.*?)ET', text, re.DOTALL):
                bt_block = match.group(1)
                current_x = current_y = 0.0
                for line in bt_block.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    td = re.match(r'([0-9.-]+)\s+([0-9.-]+)\s+Td', line)
                    if td:
                        current_x = float(td.group(1))
                        current_y = float(td.group(2))
                    hm = re.match(r'<([0-9A-Fa-f]+)>\s*Tj', line)
                    if hm:
                        decoded = decode_hex_chunk(hm.group(1))
                        if decoded.strip():
                            blocks.append((page_num + 1, round(current_x),
                                           round(current_y), decoded))
                    pm = re.match(r'\(([^)]*)\)\s*Tj', line)
                    if pm:
                        txt = pm.group(1)
                        if txt.strip():
                            blocks.append((page_num + 1, round(current_x),
                                           round(current_y), txt))
    doc.close()
    return blocks


def decode_blocks(pdf_path: str):
    """Auto-detect PDF type and decode. Returns list of (page, x, y, text)."""
    log.info("Decoding %s", pdf_path)

    plumber_blocks = decode_with_pdfplumber(pdf_path)
    cyrillic_count = sum(1 for _, _, _, t in plumber_blocks
                         if re.search(r'[А-Яа-я]', t))
    if plumber_blocks:
        if cyrillic_count > 10:
            log.info("pdfplumber: %d blocks, %d Cyrillic — OK",
                     len(plumber_blocks), cyrillic_count)
            return plumber_blocks
        log.info("pdfplumber: %d blocks, only %d Cyrillic — trying CID decoder",
                 len(plumber_blocks), cyrillic_count)

    cid_blocks = decode_with_cid(pdf_path)
    if cid_blocks:
        log.info("CID decoder: %d blocks", len(cid_blocks))
        return cid_blocks

    if plumber_blocks:
        # No GOST font available (or not a CID drawing). Text was extracted —
        # discard it only if there is really nothing usable. A Latin-script
        # drawing used to be reported as "no text extracted" here.
        log.warning("CID decoder found nothing — keeping %d pdfplumber blocks "
                    "(%d Cyrillic)", len(plumber_blocks), cyrillic_count)
        return plumber_blocks

    log.error("FAILED: no text extracted")
    return []


# ──────────────────────────────────────────────────────────────────────
# 2. AUTO-DETECT SPEC TABLE COLUMNS
# ──────────────────────────────────────────────────────────────────────

HEADER_KEYWORDS = {
    'поз':      'item',
    'наимен':   'desc',
    'назван':   'desc',
    'обознач':  'code',
    'кол':      'qty',
    'количеств': 'qty',
    'тип':      'type_col',
    'марка':    'type_col',
    'примечан': 'note',
    'категор':  'category_col',
    'ед':       'unit_col',
    'параметр': 'param_col',
    'номинал':  'param_col',
}

HEADER_EXACT = {
    'поз', 'поз.', '№', 'количество', 'кол', 'кол.',
    'наименование', 'наимен.', 'обозначение', 'обознач.',
    'тип', 'марка', 'примечание', 'ед.', 'ед. изм.',
}

# A header row is one visual row: its cells share a baseline (a few points
# apart at most). Candidates further apart than this start a new band.
HEADER_BAND_TOLERANCE = 12

# Header cells are short. Long strings that merely contain a keyword
# ("Количество жил в кабеле") must not be mistaken for a header cell.
MAX_HEADER_WORDS = 3
MAX_HEADER_LEN = 20


def _header_column_id(text: str):
    """
    Return the logical column id if `text` looks like a table header cell.

    Returns one of the HEADER_KEYWORDS values, 'item' for '№', or None.

    Note: there is deliberately no "the header must sit in the bottom 10 % of
    the page" rule here. That rule rejected every table whose header row is
    at the top of the sheet (a normal continuation sheet), which made the
    layout silently fall back to the hard-coded GOST coordinates.
    """
    t = text.strip().lower().rstrip('.:,;')
    if not t or len(t) > MAX_HEADER_LEN or len(t.split()) > MAX_HEADER_WORDS:
        return None
    for kw, col_id in HEADER_KEYWORDS.items():
        if kw in t:
            return col_id
    if t in ('№', 'n', 'no'):
        return 'item'
    return None


def _header_bands(blocks, pages):
    """Group header candidates of `pages` into bands (one band = one row)."""
    candidates = []
    for page, x, y, text in blocks:
        if page not in pages:
            continue
        col_id = _header_column_id(text)
        if col_id:
            candidates.append((page, col_id, x, y, text))
    candidates.sort(key=lambda c: (c[0], c[3]))

    bands = []
    for cand in candidates:
        if (bands and cand[0] == bands[-1][0]
                and abs(cand[3] - bands[-1][1]) <= HEADER_BAND_TOLERANCE):
            bands[-1][2].append(cand)
        else:
            bands.append([cand[0], cand[3], [cand]])
    return bands


def auto_detect_columns(blocks, last_n_pages=2):
    """
    Find the spec table header row and project one x-range per column.

    Returns (col_map, table_page) where col_map maps a column id
    ('item', 'desc', 'qty', 'code', 'unit_col', ...) to a half-open x-range
    (left, right; the right-most column is open-ended), or (None, last_page)
    when no header row was recognised — the caller then falls back to
    FIXED_COL_MAP.

    The ranges are derived from the *detected* header cell positions, so the
    same code handles a wide GOST drawing, an A4 continuation sheet and an
    A3 sheet whose columns are in an unusual order. Column order and the set
    of columns are whatever the drawing has; unknown columns are detected and
    simply not used.
    """
    if not blocks:
        return None, None

    last_page = max(b[0] for b in blocks)
    check_pages = set(range(max(1, last_page - last_n_pages + 1), last_page + 1))

    best = None
    for page, _y, cells in _header_bands(blocks, check_pages):
        by_col = {}
        for _page, col_id, x, y, text in cells:
            if col_id not in by_col or x < by_col[col_id][0]:
                by_col[col_id] = (x, y, text)
        if len(by_col) < 2:
            continue
        # A data row almost never carries both a position number column and a
        # name column header; a real header row usually does.
        score = len(by_col) + (2 if 'item' in by_col and 'desc' in by_col else 0)
        if best is None or score > best[0]:
            best = (score, page, by_col)

    if best is None:
        log.info("No header row recognised on pages %s", sorted(check_pages))
        return None, last_page

    _score, page, by_col = best
    ordered = sorted((x, col_id) for col_id, (x, _y, _t) in by_col.items())

    col_map = {}
    for i, (x, col_id) in enumerate(ordered):
        left = 0 if i == 0 else (ordered[i - 1][0] + x) // 2
        right = ((x + ordered[i + 1][0]) // 2 if i < len(ordered) - 1
                 else float('inf'))
        col_map[col_id] = (left, right)

    log.info("Auto-detected columns on page %d: %s", page,
             ', '.join('%s@%d' % (col_id, x) for x, col_id in ordered))
    return col_map, page


# ──────────────────────────────────────────────────────────────────────
# 3. PARSE SPEC TABLE
# ──────────────────────────────────────────────────────────────────────

FIXED_COL_MAP = {
    'item': (300, 600),
    'desc': (600, 3000),
    'qty': (3800, 5000),
}

# Vertical half-band used to group the blocks of one table row when the row
# pitch is unknown. With >= 2 position numbers the band is derived from the
# real pitch (see _row_band); a dense sheet (22 pt rows) and a wide drawing
# (120 pt rows) need very different bands.
Y_TOL_DEFAULT = 80
Y_TOL_MIN = 8
Y_TOL_MAX = 120


def _row_band(item_ys):
    """
    Half-height of the vertical band that belongs to one table row.

    Derived from the median distance between position numbers, so a dense
    A4 sheet (22 pt pitch) no longer merges neighbouring rows, and a wide
    drawing (120 pt pitch) no longer splits one multi-line row in two.
    """
    ys = sorted(set(item_ys))
    pitches = [b - a for a, b in zip(ys, ys[1:]) if b - a > 1]
    if not pitches:
        return Y_TOL_DEFAULT
    pitches.sort()
    pitch = pitches[len(pitches) // 2]
    return max(Y_TOL_MIN, min(Y_TOL_MAX, pitch * 0.6))


def parse_spec_table(blocks, col_map=None, table_page=None, stats=None):
    """
    Parse spec table rows from decoded blocks.

    Args:
        blocks: list of (page, x, y, text)
        col_map: dict of column ranges, or None for fallback
        table_page: which page to parse, or None for last page
        stats: optional dict to receive diagnostics (row band, column hits)

    Returns list of dicts with keys: pos, name, code, qty_raw, unit_raw
    """
    if col_map is None:
        col_map = FIXED_COL_MAP

    if table_page is None:
        table_page = max(b[0] for b in blocks)

    page_blocks = [b for b in blocks if b[0] == table_page]

    item_nums = []   # (y, num)
    descs = []       # (y, x, text)
    qtys = []        # (y, x, text)
    codes = []       # (y, x, text)
    units = []       # (y, x, text)

    for _, x, y, text in page_blocks:
        text_stripped = text.strip()
        if not text_stripped or not re.search(r'[A-Za-zА-Яа-я0-9]', text_stripped):
            continue
        # A repeated header row (or the header row itself, when it falls into
        # the vertical band of a neighbouring row) is not table data.
        if _header_column_id(text_stripped):
            continue

        col_type = None
        for cid, (x1, x2) in col_map.items():
            if x1 <= x <= x2:
                col_type = cid
                break

        if col_type == 'item':
            num_text = text_stripped.strip('[]').strip().rstrip('.')
            if num_text.isdigit() and 1 <= int(num_text) <= 100:
                item_nums.append((y, int(num_text)))
        elif col_type == 'qty':
            qtys.append((y, x, text_stripped))
        elif col_type == 'desc':
            descs.append((y, x, text_stripped))
        elif col_type == 'code':
            codes.append((y, x, text_stripped))
        elif col_type in ('unit_col', 'unit'):
            units.append((y, x, text_stripped))

    if stats is not None:
        stats['table_page'] = table_page
        stats['columns'] = sorted(col_map)
        stats['column_hits'] = {'item': len(item_nums), 'desc': len(descs),
                                'qty': len(qtys), 'code': len(codes),
                                'unit': len(units)}

    if not item_nums:
        raise NoItemsError(
            "No spec items found — column detection failed "
            "(table_page=%r, col_map=%r, %d text blocks on that page, "
            "rows seen: desc=%d qty=%d code=%d unit=%d)"
            % (table_page, col_map, len(page_blocks), len(descs), len(qtys),
               len(codes), len(units)))

    item_nums.sort(key=lambda t: -t[0])
    descs.sort(key=lambda t: -t[0])
    qtys.sort(key=lambda t: -t[0])

    y_tol = _row_band([y for y, _ in item_nums])
    if stats is not None:
        stats['row_band'] = y_tol

    rows = []
    for pos_y, pos_num in item_nums:
        # Reading order inside one row: top line first, then left to right.
        # (The global sorts above are bottom-up, which used to scramble a
        # name that wraps onto a second line.)
        desc_parts = [(dy, dx, dt) for dy, dx, dt in descs
                      if abs(pos_y - dy) < y_tol]
        qty_parts = [(qy, qx, qt) for qy, qx, qt in qtys
                     if abs(pos_y - qy) < y_tol]
        code_parts = [(cy, cx, ct) for cy, cx, ct in codes
                      if abs(pos_y - cy) < y_tol]
        unit_parts = [(uy, ux, ut) for uy, ux, ut in units
                      if abs(pos_y - uy) < y_tol]

        desc_parts.sort(key=lambda t: (t[0], t[1]))
        seen = set()
        desc_texts = []
        for _, _, t in desc_parts:
            clean = t.strip()
            if clean and clean not in seen:
                desc_texts.append(clean)
                seen.add(clean)

        qty_parts.sort(key=lambda t: (t[0], t[1]))
        seen = set()
        qty_texts = []
        for _, _, t in qty_parts:
            clean = t.strip()
            if clean and clean not in seen:
                qty_texts.append(clean)
                seen.add(clean)

        code_text = ''
        if code_parts:
            code_parts.sort(key=lambda t: (t[0], t[1]))
            code_text = ' '.join(t for _, _, t in code_parts)

        unit_text = ''
        if unit_parts:
            unit_parts.sort(key=lambda t: (t[0], t[1]))
            unit_text = ' '.join(t for _, _, t in unit_parts)

        rows.append({
            'pos': pos_num,
            'name': ' '.join(desc_texts),
            'code': clean_text(code_text),
            'qty_raw': ' '.join(qty_texts),
            'unit_raw': clean_text(unit_text),
        })

    rows.sort(key=lambda r: r['pos'])
    log.info("Parsed %d spec items", len(rows))
    return rows


# ──────────────────────────────────────────────────────────────────────
# 4. COMPONENT KNOWLEDGE BASE
# ──────────────────────────────────────────────────────────────────────

COLOR_MAP = {
    'белый': 'белый', 'жёлтый': 'жёлтый', 'синий': 'синий',
    'красный': 'красный', 'коричневый': 'коричневый', 'чёрный': 'чёрный',
    'зелёный': 'зелёный', 'жёлто': 'жёлто-зелёный',
    'серый': 'серый',
}

CATEGORY_KEYWORDS = {
    'реле': 'Коммутационное оборудование',
    'выключатель': 'Коммутационное оборудование',
    'розетка': 'Коммутационное оборудование',
    'разъём': 'Коммутационное оборудование',
    'рубильник': 'Коммутационное оборудование',
    'контактор': 'Коммутационное оборудование',
    'пускатель': 'Коммутационное оборудование',
    'автомат': 'Коммутационное оборудование',
    'зажим': 'Клеммы и монтаж',
    'клемм': 'Клеммы и монтаж',
    'колодк': 'Клеммы и монтаж',
    'шина': 'Клеммы и монтаж',
    'наконечник': 'Монтажные материалы',
    'изолятор': 'Монтажные материалы',
    'стяжка': 'Монтажные материалы',
    'хомут': 'Монтажные материалы',
    'провод': 'Провода и кабели',
    'кабель': 'Провода и кабели',
    'шнур': 'Провода и кабели',
    'лампа': 'Светосигнальная арматура',
    'светосигнальн': 'Светосигнальная арматура',
    'индикатор': 'Светосигнальная арматура',
    'светильник': 'Светосигнальная арматура',
    'сальник': 'Корпусное оборудование и аксессуары',
    'корпус': 'Корпусное оборудование и аксессуары',
    'щит': 'Корпусное оборудование и аксессуары',
    'бокс': 'Корпусное оборудование и аксессуары',
    'шкаф': 'Корпусное оборудование и аксессуары',
    'din': 'Корпусное оборудование и аксессуары',
    'трансформатор': 'Силовое оборудование',
    'блок питания': 'Силовое оборудование',
    'драйвер': 'Силовое оборудование',
}


def classify_component(name_lower: str) -> str:
    """Classify component name into a category."""
    for kw, cat in CATEGORY_KEYWORDS.items():
        if kw in name_lower:
            return cat
    return 'Прочее'


# Closed category vocabulary — every item['category'] is one of these. The
# contract validator (contract.py) rejects anything outside the list.
CATEGORY_VOCAB = tuple(sorted(set(CATEGORY_KEYWORDS.values()))) + ('Прочее',)

# Flags attached to an item. They are part of the output contract: a consumer
# must be able to tell a confident row from a row that needs a human look.
FLAG_VOCAB = (
    'UNKNOWN_COMPONENT',   # name not resolved by the knowledge base
    'WEB_INFO',            # params filled from the websearch fallback
    'QTY_CONVERTED',       # quantity converted cm → m
    'NO_QTY',              # the quantity cell was empty
    'UNPARSED_QTY',        # the quantity cell had text but no number
)


# ─── Websearch fallback ───────────────────────────────────────────────

SPEC_CACHE = {}


def websearch_component(raw_name: str):
    """Search for component specs via DuckDuckGo HTML."""
    query = f'{" ".join(raw_name.split()[:5])} технические характеристики'
    url = 'https://html.duckduckgo.com/html/'
    try:
        import requests
        resp = requests.post(url, data={'q': query},
                           headers={'User-Agent': 'Mozilla/5.0'},
                           timeout=10)
        if resp.status_code == 200:
            results = re.findall(r'class="result__snippet">(.*?)</a>',
                                 resp.text, re.DOTALL)
            if results:
                snippet = re.sub(r'<[^>]+>', '', results[0]).strip()
                return snippet[:300]
    except Exception:
        pass
    return None


def websearch_with_cache(raw_name: str):
    """Websearch with result caching."""
    key = raw_name[:50].lower()
    if key in SPEC_CACHE:
        return SPEC_CACHE[key]
    result = websearch_component(raw_name)
    SPEC_CACHE[key] = result
    return result


# ──────────────────────────────────────────────────────────────────────
# 5. ITEM FORMATTER
# ──────────────────────────────────────────────────────────────────────

def extract_qty(qty_raw: str):
    """Parse quantity string into (numeric_qty, unit)."""
    qty_raw = clean_text(qty_raw)
    qty_raw = re.sub(r'(\d)\s*c\s+екци', r'\1 секции', qty_raw, flags=re.IGNORECASE)
    qty_raw = re.sub(r'секци[яи]', 'секций', qty_raw, flags=re.IGNORECASE)

    unit_patterns = [
        (r'(\d+[.,]?\d*)\s*см', 'см'),
        (r'(\d+[.,]?\d*)\s*секций', 'секций'),
        (r'(\d+[.,]?\d*)\s*м(?!м)', 'м'),
        (r'(\d+[.,]?\d*)\s*упак', 'упак'),
        (r'(\d+[.,]?\d*)\s*шт', 'шт'),
        (r'(\d+[.,]?\d*)\s*мм', 'мм'),
    ]
    for pat, unit in unit_patterns:
        m = re.search(pat, qty_raw)
        if m:
            try:
                qty = float(m.group(1).replace(',', '.'))
            except ValueError:
                qty = m.group(1)
            return qty, unit

    m = re.match(r'\s*(\d+[.,]?\d*)', qty_raw)
    if m:
        try:
            qty = float(m.group(1).replace(',', '.'))
        except ValueError:
            qty = m.group(1)
        return qty, 'шт'
    return qty_raw, 'шт'


def format_item(row: dict, websearch: bool = True) -> dict:
    """
    Transform a raw parsed row into a formatted, classified item.

    Applies domain knowledge to:
    - Classify component category
    - Extract structured parameters (section, color, ratings, etc.)
    - Remove brand names
    - Fallback to websearch for unknown components (disable with
      websearch=False for a deterministic, offline run)

    Returns dict with keys:
        pos       int          position number from the drawing
        category  str          one of CATEGORY_VOCAB
        name      str          de-branded component name (never empty)
        code      str          drawing designation, '' when the table has none
        params    str          parameters/ratings, '' when nothing is known
        qty       float | None None when the cell holds no number
        unit      str | None   None iff qty is None
        qty_raw   str          the quantity cell exactly as read from the PDF
        flags     list[str]    subset of FLAG_VOCAB
    """
    raw_name = clean_text(row['name'])
    qty_raw = row['qty_raw']
    name_lower = raw_name.lower()
    flags = []

    category = classify_component(name_lower)

    # Quantity: the drawing's own unit column wins over the unit guessed
    # from the quantity cell.
    qty, unit = extract_qty(qty_raw)
    unit_raw = clean_text(row.get('unit_raw') or '')
    if unit_raw:
        unit = unit_raw
    if not isinstance(qty, (int, float)):
        flags.append('UNPARSED_QTY' if clean_text(qty_raw) else 'NO_QTY')
        qty = None
        unit = None

    formatted_name = raw_name
    section_found = False
    parts = []
    cleaned = name_lower.replace('\\/', '/').replace('\\(', '(')

    # --- Relay ---
    if 'реле промежуточное' in name_lower:
        formatted_name = 'Реле промежуточное электромагнитное'
        if '78/4' in name_lower or '78' in cleaned:
            parts.append('4 переключающих контакта (4PDT)')
        m = re.search(r'(\d+[.,]?\d*)\s*[АAаa]', cleaned)
        if m:
            parts.append(f'Ток контактов: {m.group(1)} А')
        m = re.search(r'(\d+)\s*[ВBвb]', cleaned)
        if m:
            parts.append(f'Напряжение катушки: ~{m.group(1)} В AC')
        section_found = True

    # --- Relay socket ---
    elif 'разъём' in name_lower and 'реле' in name_lower:
        formatted_name = 'Розетка (колодка) монтажная для реле'
        if '78/4' in name_lower:
            parts.append('Под 4 переключающих контакта (4PDT)')
        parts.append('установка на DIN-рейку / панель')
        section_found = True

    # --- Terminal block ---
    elif 'зажим' in name_lower:
        formatted_name = 'Блок зажимов винтовых (клеммная колодка)'
        m = re.search(r'(\d+[.,]?\d*)', name_lower.replace('зажим', ''))
        if m:
            parts.append(f'Сечение: до {m.group(1)} мм²')
        section_found = True

    # --- Ferrule (fork) terminal ---
    elif 'наконечник вилочный' in name_lower:
        formatted_name = 'Наконечник кабельный вилочный изолированный'
        parts.append('под винт/болт')
        if 'нви' in name_lower:
            parts.append('Сечение провода: 0.5–1.5 мм²')
        for cw, cn in COLOR_MAP.items():
            if cw in name_lower:
                parts.append(f'цвет изоляции: {cn}')
                break
        section_found = True

    # --- Ferrule (sleeve) terminal ---
    elif 'наконечник гильзовый' in name_lower:
        formatted_name = 'Наконечник штыревой втулочный (гильзовый)'
        m = re.search(r'(\d+[.,]?\d*)\s*мм', name_lower)
        if m:
            parts.append(f'Сечение провода: {m.group(1)} мм²')
        section_found = True

    # --- Ring terminal ---
    elif 'наконечник кольцевой' in name_lower:
        formatted_name = 'Наконечник кольцевой'
        cl = name_lower.replace('\\(', '(').replace('\\)', ')')
        m = re.search(r'(\d+[.,]?\d*)\s*[хx]\s*(\d+[.,]?\d*)\s*[хx]\s*(\d+[.,]?\d*)', cl)
        if m:
            parts.append(f'Размеры: {m.group(1)} × {m.group(2)} × {m.group(3)} мм')
        else:
            m = re.search(r'(\d+[.,]?\d*)\s*[хx]\s*(\d+[.,]?\d*)', cl)
            if m:
                parts.append(f'Размеры: {m.group(1)} × {m.group(2)} мм')
        section_found = True

    # --- Flat/knife terminal ---
    elif 'наконечник ножевой' in name_lower:
        formatted_name = 'Наконечник плоский (ножевой) плоский штекер'
        cl = name_lower.replace('\\(', '(').replace('\\)', ')')
        m = re.search(r'(\d+[.,]?\d*)\s*[хx]\s*(\d+[.,]?\d*)', cl)
        if m:
            parts.append(f'Размер контактов: {m.group(1)} × {m.group(2)} мм')
        section_found = True

    # --- Wire ---
    elif 'провод' in name_lower:
        is_pugv = 'пугв' in name_lower or 'пв3' in cleaned
        if 'мпо' in name_lower:
            formatted_name = 'Провод монтажный термостойкий/повышенной надежности'
        elif is_pugv:
            formatted_name = 'Провод монтажный медный повышенной гибкости (заземление)'
            cleaned = re.sub(r'\bпв\d*\b\s*\(?\s*\d*\s*\)?', '', cleaned)
            cleaned = cleaned.replace('(', '').replace(')', '')
        elif 'пв' in name_lower:
            formatted_name = 'Провод монтажный медный одножильный (установка)'
        else:
            formatted_name = 'Провод монтажный'

        m = re.search(r'(\d+)\s*[хx]\s*(\d+[.,]?\d*)', cleaned)
        if m:
            parts.append(f'Число жил и сечение: {m.group(1)} × {m.group(2)} мм²')
        else:
            stripped = re.sub(
                r'\b(провод|монтажный|мПо|пв\d*|пугв|установка|заземление)\b',
                '', cleaned)
            m = re.search(r'(\d+[.,]?\d*)\s*мм', stripped)
            if m:
                parts.append(f'Сечение: {m.group(1)} мм²')
            else:
                nums = re.findall(r'(\d+[.,]?\d*)', stripped)
                if nums:
                    sec = next((n for n in nums if ',' in n or '.' in n), nums[-1])
                    try:
                        if 0.1 <= float(sec.replace(',', '.')) <= 10:
                            parts.append(f'Сечение: {sec} мм²')
                    except ValueError:
                        pass
        if is_pugv and not any('Сечение' in p for p in parts):
            parts.append('Сечение: 2.5 мм²')
        for cw, cn in COLOR_MAP.items():
            if cw in name_lower:
                parts.append(f'цвет: {cn}')
                break
        section_found = True
        if unit == 'см' and isinstance(qty, (int, float)):
            qty = round(qty / 100, 2)
            unit = 'м'
            flags.append('QTY_CONVERTED')

    # --- Cable ---
    elif 'кабель' in name_lower:
        formatted_name = 'Кабель контрольный гибкий многожильный'
        m = re.search(r'(\d+)\s*[хx]\s*(\d+[.,]?\d*)', cleaned)
        if m:
            parts.append(f'Число жил и сечение: {m.group(1)} × {m.group(2)} мм²')
        section_found = True
        if unit == 'см' and isinstance(qty, (int, float)):
            qty = round(qty / 100, 2)
            unit = 'м'
            flags.append('QTY_CONVERTED')

    # --- Indicator lamp ---
    elif 'лампа' in name_lower:
        formatted_name = 'Светосигнальный индикатор (лампа матричная)'
        m = re.search(r'(\d+)\s*[ВBвb]', cleaned)
        if m:
            parts.append(f'Напряжение: ~{m.group(1)} В AC')
        for cw, cn in COLOR_MAP.items():
            if cw in name_lower:
                parts.append(f'цвет свечения: {cn}')
                break
        if 'nd' in name_lower:
            parts.append('посадочный диаметр: 16 мм')
        section_found = True

    # --- Cable gland ---
    elif 'сальник' in name_lower:
        formatted_name = 'Ввод кабельный (сальник)'
        m = re.search(r'(mg|m)\s*(\d+)', name_lower)
        if m:
            parts.append(f'Резьба: MG{m.group(2)} (M{m.group(2)}×1.5)')
        section_found = True

    # --- Switch / rocker ---
    elif 'выключатель' in name_lower:
        formatted_name = 'Выключатель клавишный (рокерный)'
        cl = cleaned.replace('вн', 'в')
        m = re.search(r'(\d+[.,]?\d*)\s*[АAаa]', cl)
        if m:
            current_tok = f'{m.group(1)} А'
        else:
            current_tok = None
        m = re.search(r'(\d+)\s*[ВBвb](?![А-Яа-яA-Za-z])', cl)
        if m:
            volt = f'{m.group(1)} В AC'
            if current_tok:
                parts.append(f'Ток/Напряжение: {current_tok} / {volt}')
            else:
                parts.append(f'Напряжение: {volt}')
        elif current_tok:
            parts.append(f'Ток: {current_tok}')
        for cw, cn in COLOR_MAP.items():
            if cw in name_lower:
                parts.append(f'цвет: {cn}')
                break
        section_found = True

    # --- Enclosure ---
    elif 'корпус' in name_lower or 'щрн' in name_lower:
        formatted_name = 'Корпус металлический распределительный навесной'
        m = re.search(r'(\d+)\s*[хx]\s*(\d+)\s*[хx]\s*(\d+)', cleaned)
        if m:
            parts.append(f'Габариты (ВхШхГ): {m.group(1)} × {m.group(2)} × {m.group(3)} мм')
        m = re.search(r'ip\s*(\d+)', name_lower)
        if m:
            parts.append(f'степень защиты: IP{m.group(1)}')
        section_found = True

    # --- Insulator ---
    elif 'изолятор' in name_lower:
        formatted_name = 'Изолятор (защитный колпачок) для плоского наконечника'
        m = re.search(r'(?:dr)?\s*\d+[-–](\d+)[-–](\d+)[-–](\d+[.,]?\d*)', name_lower)
        if m:
            parts.append(f'размеры: {m.group(1)}×{m.group(2)}×{m.group(3)} мм')
        if 'н2' in name_lower:
            parts.append('Для ножевого наконечника типа Н2')
        section_found = True

    # ─── Unknown component → websearch ───
    if not section_found:
        flags.append('UNKNOWN_COMPONENT')
        web_info = None
        if websearch:
            log.info("Unknown component: %s — searching web", raw_name[:60])
            web_info = websearch_with_cache(raw_name)
        if web_info:
            formatted_name = raw_name.split('[')[0].strip().rstrip('\\').strip()
            parts.append(f'Характеристики: {web_info[:200]}')
            flags.append('WEB_INFO')
        else:
            formatted_name = raw_name
            parts.append('(требуется уточнение)')

    params = ', '.join(parts)
    formatted_name = formatted_name.strip().strip(',;:').strip()
    params = params.strip().strip(',;:').strip()

    return {
        'pos': row['pos'],
        'category': category,
        'name': formatted_name,
        'code': clean_text(row.get('code') or ''),
        'params': params,
        'qty': qty,
        'unit': unit,
        'qty_raw': clean_text(qty_raw),
        'flags': flags,
    }


# ──────────────────────────────────────────────────────────────────────
# 6. EXPORTERS
# ──────────────────────────────────────────────────────────────────────

def export_xlsx(items: list, output_path: str, cabinet_name: str = None):
    """Export formatted items to styled Excel workbook."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Спецификация'

    title_font = Font(bold=True, size=14)
    subtitle_font = Font(bold=True, size=11)
    header_font = Font(bold=True, size=10, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    cell_font = Font(size=10)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='top')

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 55
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 22

    ws.merge_cells('A1:G1')
    ws['A1'] = 'ОБЕЗЛИЧЕННАЯ СПЕЦИФИКАЦИЯ ОБОРУДОВАНИЯ И МАТЕРИАЛОВ'
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    if cabinet_name:
        ws.merge_cells('A2:G2')
        ws['A2'] = cabinet_name
        ws['A2'].font = subtitle_font
        ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    headers = ['№', 'Категория', 'Наименование компонента (обезличенное)',
               'Основные параметры / Номиналы', 'Кол-во', 'Ед. изм.',
               'Флаги (что проверить вручную)']
    for col, title in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

    for i, item in enumerate(items, 5):
        ws.cell(row=i, column=1, value=item['pos']).font = cell_font
        ws.cell(row=i, column=1).alignment = center_align
        ws.cell(row=i, column=1).border = thin_border

        ws.cell(row=i, column=2, value=item['category']).font = cell_font
        ws.cell(row=i, column=2).alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row=i, column=2).border = thin_border

        ws.cell(row=i, column=3, value=item['name']).font = cell_font
        ws.cell(row=i, column=3).alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row=i, column=3).border = thin_border

        ws.cell(row=i, column=4, value=item['params']).font = cell_font
        ws.cell(row=i, column=4).alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row=i, column=4).border = thin_border

        # qty is a number or None; when None the raw cell text is shown so the
        # row is never silently empty.
        qty_value = item['qty'] if item['qty'] is not None else item.get('qty_raw', '')
        ws.cell(row=i, column=5, value=qty_value).font = cell_font
        ws.cell(row=i, column=5).alignment = center_align
        ws.cell(row=i, column=5).border = thin_border
        if isinstance(item['qty'], float):
            ws.cell(row=i, column=5).number_format = '#,##0.00'

        ws.cell(row=i, column=6, value=item['unit'] or '').font = cell_font
        ws.cell(row=i, column=6).alignment = center_align
        ws.cell(row=i, column=6).border = thin_border

        ws.cell(row=i, column=7, value='; '.join(item.get('flags') or [])).font = cell_font
        ws.cell(row=i, column=7).alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row=i, column=7).border = thin_border
        ws.row_dimensions[i].height = 30

    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    wb.save(output_path)
    log.info("Excel: %s", output_path)


def export_pdf(items: list, output_path: str, title_text: str = None):
    """Export formatted items to clean PDF with Cyrillic support."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    # Register Cyrillic TTF font
    arial_path = 'C:\\Windows\\Fonts\\arial.ttf'
    if os.path.exists(arial_path):
        pdfmetrics.registerFont(TTFont('ArialCyr', arial_path))
        FONT_NAME = 'ArialCyr'
        FONT_BOLD = 'ArialCyr'
    else:
        FONT_NAME = 'Helvetica'
        FONT_BOLD = 'Helvetica-Bold'

    doc = SimpleDocTemplate(
        output_path, pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'T', parent=styles['Title'], fontName=FONT_BOLD,
        fontSize=14, spaceAfter=6*mm)
    cell_style = ParagraphStyle(
        'C', parent=styles['Normal'], fontName=FONT_NAME,
        fontSize=7.5, leading=9, wordWrap='CJK')

    header = ['№', 'Категория', 'Наименование компонента (обезличенное)',
              'Основные параметры / Номиналы', 'Кол-во', 'Ед. изм.', 'Флаги']
    data = [header]
    for item in items:
        if item['qty'] is None:
            qty_str = item.get('qty_raw') or '—'
        elif isinstance(item['qty'], float):
            qty_str = f'{item["qty"]:.2f}'.rstrip('0').rstrip('.')
        else:
            qty_str = str(item['qty'])
        data.append([
            str(item['pos']),
            Paragraph(item['category'], cell_style),
            Paragraph(item['name'], cell_style),
            Paragraph(item['params'], cell_style),
            qty_str, item['unit'] or '',
            Paragraph('; '.join(item.get('flags') or []), cell_style)])

    col_widths = [10*mm, 40*mm, 55*mm, 75*mm, 14*mm, 14*mm, 40*mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (4, 0), (5, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
         [colors.white, colors.HexColor('#F2F2F2')]),
    ]))

    story = [
        Paragraph(title_text or 'ОБЕЗЛИЧЕННАЯ СПЕЦИФИКАЦИЯ ОБОРУДОВАНИЯ И МАТЕРИАЛОВ',
                  title_style),
        Spacer(1, 3*mm),
        table,
    ]
    doc.build(story)
    log.info("PDF: %s", output_path)


# ──────────────────────────────────────────────────────────────────────
# 7. HIGH-LEVEL API
# ──────────────────────────────────────────────────────────────────────

def parse_pdf(pdf_path: str, output_prefix: str = None, verbose: bool = False,
              websearch: bool = True, stats: dict = None):
    """
    Full pipeline: decode → detect → parse → classify → format.

    Args:
        pdf_path: Path to the input PDF drawing.
        output_prefix: Output file path prefix (without extension).
                       If None, derived from input name.
        verbose: Enable detailed logging.
        websearch: Allow the websearch fallback for unknown components. Pass
                   False for a deterministic, offline (and faster) run — the
                   parameters of unknown components are left empty then.
        stats: Optional dict, filled with run diagnostics (blocks, pages,
               table page, detected columns, column hits, row band, whether
               the websearch fallback was enabled).

    Returns:
        List of formatted item dicts with keys:
            pos, category, name, code, params, qty, unit, qty_raw, flags

    Raises:
        NoTextError: no text layer could be extracted.
        NoItemsError: text found, but no spec table rows recognised.
    """
    if verbose:
        logging.basicConfig(level=logging.INFO, format='%(message)s')
    else:
        logging.basicConfig(level=logging.WARNING, format='%(message)s')

    if output_prefix is None:
        output_prefix = os.path.splitext(pdf_path)[0] + '_anon'

    blocks = decode_blocks(pdf_path)
    if not blocks:
        raise NoTextError(
            "No text extracted from PDF (scan without a text layer, or an "
            "unsupported font encoding)")

    if stats is not None:
        stats['blocks'] = len(blocks)
        stats['pages'] = max(b[0] for b in blocks)
        stats['websearch'] = bool(websearch)

    col_map, table_page = auto_detect_columns(blocks)
    if col_map:
        rows = parse_spec_table(blocks, col_map, table_page, stats=stats)
    else:
        log.warning("No header row detected — falling back to the fixed "
                    "GOST column ranges")
        if stats is not None:
            stats['fallback'] = 'FIXED_COL_MAP'
        rows = parse_spec_table(blocks, table_page=table_page, stats=stats)

    formatted = []
    for r in rows:
        item = format_item(r, websearch=websearch)
        formatted.append(item)
        if verbose:
            print(f'  {item["pos"]:>2}. [{item["category"]}] {item["name"]}')
            print(f'       {item["params"]}')
            print(f'       qty={item["qty"]} {item["unit"]}'
                  + (f'  flags={",".join(item["flags"])}' if item['flags'] else ''))

    return formatted


# ──────────────────────────────────────────────────────────────────────
# 8. CLI
# ──────────────────────────────────────────────────────────────────────

def _parse_formats(spec: str):
    """Turn 'all' or 'csv,json' into a list of output formats, or None."""
    from . import contract
    wanted = [part.strip().lower() for part in str(spec).split(',') if part.strip()]
    if not wanted:
        return None
    if 'all' in wanted:
        return list(contract.FORMATS)
    unknown = [w for w in wanted if w not in contract.FORMATS]
    if unknown:
        return None
    ordered = []
    for fmt in contract.FORMATS:          # deterministic file order
        if fmt in wanted:
            ordered.append(fmt)
    return ordered


def main(argv=None):
    """
    Command-line entry point.

    Exit codes are documented in contract.py and in the README:
    0 ok, 1 internal error, 2 usage/input error, 3 no text extracted,
    4 no spec items found, 5 contract violation, 6 nothing to write.
    """
    import argparse
    import traceback
    from . import contract

    parser = argparse.ArgumentParser(
        description="pdf-spec-parser — Universal procurement PDF spec parser & obfuscator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m pdf_spec_parser drawing.pdf
  python -m pdf_spec_parser drawing.pdf output_prefix --verbose
  python -m pdf_spec_parser drawing.pdf --format csv,json --no-websearch
  python -m pdf_spec_parser drawing.pdf --out-dir out/ --format all

Output (--format, default all):
  {prefix}.xlsx  formatted spreadsheet for humans (with flags column)
  {prefix}.pdf   formatted PDF for humans
  {prefix}.csv   contract CSV, UTF-8 with BOM
  {prefix}.json  contract JSON document (schema_version 1.0)

Exit codes:
  0 success            3 no text extracted from the PDF
  1 internal error     4 no spec table rows recognised
  2 usage/input error  5 output contract violated
                       6 no output format selected
        """)

    parser.add_argument('input', help='Path to the input PDF drawing')
    parser.add_argument('output', nargs='?', default=None,
                        help='Output file prefix (without extension)')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Show detailed parsing progress')
    parser.add_argument('--format', default='all', metavar='LIST',
                        help='Comma separated output formats (csv,json,xlsx,pdf) '
                             'or "all" (default: all)')
    parser.add_argument('--out-dir', default=None, metavar='DIR',
                        help='Write the output files into DIR '
                             '(default: next to the input PDF)')
    parser.add_argument('--no-websearch', action='store_true',
                        help='Never query the web for unknown components '
                             '(deterministic, offline run)')

    args = parser.parse_args(argv)

    formats = _parse_formats(args.format)
    if formats is None:
        print("Error: --format must be 'all' or a comma separated list of %s "
              "(got %r)" % ('/'.join(contract.FORMATS), args.format),
              file=sys.stderr)
        return contract.EXIT_USAGE

    if not os.path.exists(args.input) or os.path.isdir(args.input):
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        return contract.EXIT_USAGE

    if args.out_dir:
        os.makedirs(args.out_dir, exist_ok=True)

    stats = {}
    try:
        items = parse_pdf(args.input, args.output, verbose=args.verbose,
                          websearch=not args.no_websearch, stats=stats)
    except NoTextError as exc:
        print(f"Error [no-text]: {exc}", file=sys.stderr)
        return contract.EXIT_NO_TEXT
    except NoItemsError as exc:
        print(f"Error [no-items]: {exc}", file=sys.stderr)
        return contract.EXIT_NO_ITEMS
    except PdfSpecError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return contract.EXIT_INTERNAL
    except Exception as exc:                       # noqa: BLE001 - CLI boundary
        if type(exc).__name__ in ('PDFSyntaxError', 'PDFPasswordIncorrect'):
            print(f"Error: not a readable PDF: {exc}", file=sys.stderr)
            return contract.EXIT_USAGE
        traceback.print_exc()
        return contract.EXIT_INTERNAL

    if not items:
        print("Error: No items parsed", file=sys.stderr)
        return contract.EXIT_NO_ITEMS

    prefix = args.output or (os.path.splitext(args.input)[0] + '_anon')
    if args.out_dir:
        prefix = os.path.join(args.out_dir, os.path.basename(prefix))
    cabinet = os.path.splitext(os.path.basename(args.input))[0]

    records = contract.to_records(items)
    status, issues = contract.verdict(records)
    document = contract.build_document(records, args.input, stats, issues)

    written = []
    if 'csv' in formats:
        written.append(contract.write_csv(records, prefix + '.csv'))
    if 'json' in formats:
        written.append(contract.write_json(document, prefix + '.json'))
    if 'xlsx' in formats:
        export_xlsx(records, prefix + '.xlsx', cabinet)
        written.append(prefix + '.xlsx')
    if 'pdf' in formats:
        export_pdf(records, prefix + '.pdf', cabinet)
        written.append(prefix + '.pdf')

    if not written:
        print("Error: nothing to write (empty --format)", file=sys.stderr)
        return contract.EXIT_NO_OUTPUT

    print(f"records: {len(records)} (schema {contract.SCHEMA_VERSION}, "
          f"table page {stats.get('table_page')}, "
          f"columns: {', '.join(stats.get('columns', [])) or 'fallback'})")
    for path in written:
        print(f"wrote {path}")

    for issue in issues:
        stream = sys.stderr if issue['severity'] == 'error' else sys.stdout
        print(f"  [{issue['severity']}] {issue['code']}: {issue['message']}",
              file=stream)

    if status == 'FAIL':
        print("contract: FAIL — see the errors above", file=sys.stderr)
        return contract.EXIT_CONTRACT

    print("contract: PASS")
    return contract.EXIT_OK
