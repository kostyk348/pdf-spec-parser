# Examples

## Command-line usage

```bash
# Parse a PDF and generate obfuscated spec
python -m pdf_spec_parser "drawing.pdf" "output_spec" --verbose
```

## Python API usage

```python
from pdf_spec_parser import parse_pdf, export_xlsx, export_pdf

items = parse_pdf("drawing.pdf", verbose=True)

print(f"Parsed {len(items)} items:")
for item in items:
    print(f"  {item['pos']}. [{item['category']}] {item['name']}")
    print(f"       {item['params']}")
    print(f"       qty={item['qty']} {item['unit']}")

export_xlsx(items, "spec.xlsx", cabinet_name="My Cabinet")
export_pdf(items, "spec.pdf", title_text="My Cabinet")
```

## Working with the data

Each formatted item is a dict:

```python
{
    'pos': 1,
    'category': 'Коммутационное оборудование',
    'name': 'Реле промежуточное электромагнитное',
    'params': '4 переключающих контакта (4PDT), Ток контактов: 3 А, Напряжение катушки: ~220 В AC',
    'qty': 3.0,
    'unit': 'шт',
}
```
