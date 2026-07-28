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
from collections import defaultdict

log = logging.getLogger(__name__)

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
    """Extract text blocks using pdfplumber."""
    import pdfplumber
    blocks = []
    with pdfplumber.open(pdf_path) as pdf:
        for pn, page in enumerate(pdf.pages, 1):
            for ch in page.chars:
                x, y = round(ch['x0']), round(ch['top'])
                text = ch.get('text', '').strip()
                if text:
                    blocks.append((pn, x, y, text))
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

    blocks = decode_with_pdfplumber(pdf_path)
    if blocks:
        cyrillic_count = sum(1 for _, _, _, t in blocks if re.search(r'[А-Яа-я]', t))
        if cyrillic_count > 10:
            log.info("pdfplumber: %d blocks, %d Cyrillic — OK", len(blocks), cyrillic_count)
            return blocks
        log.info("pdfplumber: %d blocks, only %d Cyrillic — trying CID decoder",
                 len(blocks), cyrillic_count)

    blocks = decode_with_cid(pdf_path)
    if blocks:
        log.info("CID decoder: %d blocks", len(blocks))
        return blocks

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


def _is_likely_header(text: str, y: int, max_y: int) -> bool:
    """Check if text looks like a table header (top of page, short)."""
    t = text.strip().lower().rstrip('.:,')
    if y < max_y * 0.9:
        return False
    words = t.split()
    if len(words) > 3:
        return False
    if t in HEADER_EXACT:
        return True
    for kw in HEADER_KEYWORDS:
        if t == kw or t == kw + '.':
            return True
    if len(t) <= 12 and len(words) <= 2:
        for kw in HEADER_KEYWORDS:
            if kw in t:
                return True
    return False


def auto_detect_columns(blocks, last_n_pages=2):
    """
    Scan last pages for table headers, return (col_map, table_page).

    col_map = {'item': (x_min, x_max), 'desc': ..., 'qty': ...}
    """
    if not blocks:
        return None, None

    last_page = max(b[0] for b in blocks)
    check_pages = set(range(max(1, last_page - last_n_pages + 1), last_page + 1))

    all_y = [y for p, _, y, _ in blocks if p in check_pages]
    max_y = max(all_y) if all_y else 0

    candidates = []
    for page, x, y, text in blocks:
        if page not in check_pages:
            continue
        if not _is_likely_header(text, y, max_y):
            continue
        for kw, col_id in HEADER_KEYWORDS.items():
            if kw in text.lower():
                candidates.append((page, col_id, x, y, text))
                break

    if not candidates:
        return None, last_page

    page_cols = defaultdict(lambda: defaultdict(list))
    for page, col_id, x, y, text in candidates:
        page_cols[page][col_id].append(x)

    best_page = max(page_cols.keys(), key=lambda p: len(page_cols[p]))
    cols_x = page_cols[best_page]

    log.debug("Headers on page %d: %s", best_page,
              {k: v for k, v in cols_x.items()})

    # Project column ranges from known GOST table structure
    qty_x = min(cols_x.get('qty', [4000])) if 'qty' in cols_x else 4000
    col_ranges = {
        'item': (0, 600),
        'desc': (600, qty_x - 100),
        'qty': (qty_x - 100, qty_x + 1000),
    }
    if 'code' in cols_x:
        col_ranges['desc'] = (0, col_ranges['desc'][1])

    log.info("Auto-detected %d columns on page %d, %d items expected",
             len(col_ranges), best_page, ...)
    return col_ranges, best_page


# ──────────────────────────────────────────────────────────────────────
# 3. PARSE SPEC TABLE
# ──────────────────────────────────────────────────────────────────────

FIXED_COL_MAP = {
    'item': (300, 600),
    'desc': (600, 3000),
    'qty': (3800, 5000),
}


def parse_spec_table(blocks, col_map=None, table_page=None):
    """
    Parse spec table rows from decoded blocks.

    Args:
        blocks: list of (page, x, y, text)
        col_map: dict of column ranges, or None for fallback
        table_page: which page to parse, or None for last page

    Returns list of dicts with keys: pos, name, code, qty_raw
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

    for _, x, y, text in page_blocks:
        text_stripped = text.strip()
        if not text_stripped or not re.search(r'[A-Za-zА-Яа-я0-9]', text_stripped):
            continue

        col_type = None
        for cid, (x1, x2) in col_map.items():
            if x1 <= x <= x2:
                col_type = cid
                break

        if col_type == 'item':
            num_text = text_stripped.strip('[]').strip()
            if num_text.isdigit() and 1 <= int(num_text) <= 100:
                item_nums.append((y, int(num_text)))
        elif col_type == 'qty':
            qtys.append((y, x, text_stripped))
        elif col_type == 'desc':
            descs.append((y, x, text_stripped))
        elif col_type == 'code':
            codes.append((y, x, text_stripped))

    if not item_nums:
        raise ValueError("No spec items found — column detection failed.")

    item_nums.sort(key=lambda t: -t[0])
    descs.sort(key=lambda t: -t[0])
    qtys.sort(key=lambda t: -t[0])

    Y_TOL = 80
    rows = []
    for pos_y, pos_num in item_nums:
        desc_parts = [(dx, dt) for dy, dx, dt in descs if abs(pos_y - dy) < Y_TOL]
        qty_parts = [(qx, qt) for qy, qx, qt in qtys if abs(pos_y - qy) < Y_TOL]
        code_parts = [(cx, ct) for cy, cx, ct in codes if abs(pos_y - cy) < Y_TOL]

        desc_parts.sort(key=lambda t: t[0])
        seen = set()
        desc_texts = []
        for _, t in desc_parts:
            clean = t.strip()
            if clean and clean not in seen:
                desc_texts.append(clean)
                seen.add(clean)

        qty_parts.sort(key=lambda t: t[0])
        seen = set()
        qty_texts = []
        for _, t in qty_parts:
            clean = t.strip()
            if clean and clean not in seen:
                qty_texts.append(clean)
                seen.add(clean)

        code_text = ''
        if code_parts:
            code_parts.sort(key=lambda t: t[0])
            code_text = ' '.join(t for _, t in code_parts)

        rows.append({
            'pos': pos_num,
            'name': ' '.join(desc_texts),
            'code': clean_text(code_text),
            'qty_raw': ' '.join(qty_texts),
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


def format_item(row: dict) -> dict:
    """
    Transform a raw parsed row into a formatted, classified item.

    Applies domain knowledge to:
    - Classify component category
    - Extract structured parameters (section, color, ratings, etc.)
    - Remove brand names
    - Fallback to websearch for unknown components

    Returns dict with: pos, category, name, params, qty, unit
    """
    raw_name = clean_text(row['name'])
    qty_raw = row['qty_raw']
    name_lower = raw_name.lower()

    category = classify_component(name_lower)
    qty, unit = extract_qty(qty_raw)

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
        if unit == 'см':
            qty = round(qty / 100, 2)
            unit = 'м'

    # --- Cable ---
    elif 'кабель' in name_lower:
        formatted_name = 'Кабель контрольный гибкий многожильный'
        m = re.search(r'(\d+)\s*[хx]\s*(\d+[.,]?\d*)', cleaned)
        if m:
            parts.append(f'Число жил и сечение: {m.group(1)} × {m.group(2)} мм²')
        section_found = True
        if unit == 'см':
            qty = round(qty / 100, 2)
            unit = 'м'

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
        log.info("Unknown component: %s — searching web", raw_name[:60])
        web_info = websearch_with_cache(raw_name)
        if web_info:
            formatted_name = raw_name.split('[')[0].strip().rstrip('\\').strip()
            parts.append(f'Характеристики: {web_info[:200]}')
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
        'params': params,
        'qty': qty,
        'unit': unit,
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

    ws.merge_cells('A1:F1')
    ws['A1'] = 'ОБЕЗЛИЧЕННАЯ СПЕЦИФИКАЦИЯ ОБОРУДОВАНИЯ И МАТЕРИАЛОВ'
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    if cabinet_name:
        ws.merge_cells('A2:F2')
        ws['A2'] = cabinet_name
        ws['A2'].font = subtitle_font
        ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    headers = ['№', 'Категория', 'Наименование компонента (обезличенное)',
               'Основные параметры / Номиналы', 'Кол-во', 'Ед. изм.']
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

        ws.cell(row=i, column=5, value=item['qty']).font = cell_font
        ws.cell(row=i, column=5).alignment = center_align
        ws.cell(row=i, column=5).border = thin_border
        if isinstance(item['qty'], float):
            ws.cell(row=i, column=5).number_format = '#,##0.00'

        ws.cell(row=i, column=6, value=item['unit']).font = cell_font
        ws.cell(row=i, column=6).alignment = center_align
        ws.cell(row=i, column=6).border = thin_border
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
              'Основные параметры / Номиналы', 'Кол-во', 'Ед. изм.']
    data = [header]
    for item in items:
        qty_str = (f'{item["qty"]:.2f}'.rstrip('0').rstrip('.')
                   if isinstance(item['qty'], float) else str(item['qty']))
        data.append([
            str(item['pos']),
            Paragraph(item['category'], cell_style),
            Paragraph(item['name'], cell_style),
            Paragraph(item['params'], cell_style),
            qty_str, item['unit']])

    col_widths = [10*mm, 45*mm, 65*mm, 85*mm, 15*mm, 15*mm]
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

def parse_pdf(pdf_path: str, output_prefix: str = None, verbose: bool = False):
    """
    Full pipeline: decode → detect → parse → classify → format.

    Args:
        pdf_path: Path to the input PDF drawing.
        output_prefix: Output file path prefix (without extension).
                       If None, derived from input name.
        verbose: Enable detailed logging.

    Returns:
        List of formatted item dicts with keys:
            pos, category, name, params, qty, unit
    """
    if verbose:
        logging.basicConfig(level=logging.INFO, format='%(message)s')
    else:
        logging.basicConfig(level=logging.WARNING, format='%(message)s')

    if output_prefix is None:
        output_prefix = os.path.splitext(pdf_path)[0] + '_anon'

    blocks = decode_blocks(pdf_path)
    if not blocks:
        raise RuntimeError("No text extracted from PDF")

    col_map, table_page = auto_detect_columns(blocks)
    if col_map:
        rows = parse_spec_table(blocks, col_map, table_page)
    else:
        log.warning("Auto-detect failed, using fixed column ranges")
        rows = parse_spec_table(blocks)

    formatted = []
    for r in rows:
        item = format_item(r)
        formatted.append(item)
        if verbose:
            print(f'  {item["pos"]:>2}. [{item["category"]}] {item["name"]}')
            print(f'       {item["params"]}')
            print(f'       qty={item["qty"]} {item["unit"]}')

    return formatted


# ──────────────────────────────────────────────────────────────────────
# 8. CLI
# ──────────────────────────────────────────────────────────────────────

def main():
    """Command-line entry point."""
    import argparse

    parser = argparse.ArgumentParser(
        description="pdf-spec-parser — Universal procurement PDF spec parser & obfuscator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m pdf_spec_parser drawing.pdf
  python -m pdf_spec_parser drawing.pdf output_prefix
  python -m pdf_spec_parser drawing.pdf --verbose

Output:
  Creates {output_prefix}.xlsx and {output_prefix}.pdf
        """)

    parser.add_argument('input', help='Path to the input PDF drawing')
    parser.add_argument('output', nargs='?', default=None,
                        help='Output file prefix (without extension)')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Show detailed parsing progress')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: File not found: {args.input}")
        sys.exit(1)

    try:
        items = parse_pdf(args.input, args.output, verbose=args.verbose)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

    if not items:
        print("Error: No items parsed")
        sys.exit(1)

    prefix = args.output or (os.path.splitext(args.input)[0] + '_anon')
    cabinet = os.path.splitext(os.path.basename(args.input))[0]

    export_xlsx(items, prefix + '.xlsx', cabinet)
    export_pdf(items, prefix + '.pdf', cabinet)

    if args.verbose:
        print(f"\nDone: {prefix}.xlsx, {prefix}.pdf")


if __name__ == '__main__':
    main()
