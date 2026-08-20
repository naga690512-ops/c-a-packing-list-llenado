import re
import pdfplumber

def parse_oc_pdf(path):
    with pdfplumber.open(path) as pdf:
        page0_text = pdf.pages[0].extract_text() or ""
        all_tables = []
        for p in pdf.pages:
            all_tables.extend(p.extract_tables())

    def find(pattern, text, default=None):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else default

    header = {
        'numero_orden': find(r'Numero de Orden:\s*(\S+)', page0_text),
        'modelo_id': find(r'Modelo ID:\s*(\S+)', page0_text),
        'proveedor': find(r'Proveedor:\s*(.+)', page0_text),
        'division': find(r'Division:\s*(.+?)\s+Sub Division:', page0_text),
        'color_generico': find(r'Color Generico:\s*(.+?)\s+PANTONE:', page0_text),
        'piezas_totales': find(r'Piezas Totales:\s*(\S+)', page0_text),
        'descripcion_articulo': find(r'Descripcion del Articulo:\s*(.+)', page0_text),
    }

    # SKU table (talla -> sku, piezas totales del PO)
    sku_table = {}
    sku_order = []
    pack_table_rows = []
    for t in all_tables:
        if not t or not t[0]:
            continue
        if t[0][0] == 'SKU' and t[0][1] == 'Talla':
            for row in t[1:]:
                if not row or row[1] == 'Total' or row[0] is None:
                    continue
                sku_table[row[1]] = {'sku': row[0], 'piezas': int(row[2])}
                sku_order.append(row[1])
        elif t[0][0] == 'Pack' and t[0][1] == 'Pack ID':
            pack_table_rows = t[1:]

    # Group pack table rows into packs (A, B, C...)
    packs = []
    current = None
    for row in pack_table_rows:
        pack_letter, pack_id, tipo, unidades, total_packs, total_unidades, talla, cantidad = row
        if pack_letter:  # new pack group starts
            if current:
                packs.append(current)
            current = {
                'letra': pack_letter,
                'pack_id': pack_id or '',
                'tipo': tipo,
                'unidades_por_pack': int(unidades) if unidades else None,
                'total_packs': int(total_packs) if total_packs else None,
                'total_unidades': int(total_unidades) if total_unidades else None,
                'tallas': {}
            }
        if current and talla:
            current['tallas'][talla] = int(cantidad)
    if current:
        packs.append(current)

    return {
        'header': header,
        'sku_table': sku_table,
        'sku_order': sku_order,
        'packs': packs,
    }

if __name__ == '__main__':
    import json
    for f in ['ProductionOrderIcon_08-20-024526.pdf','ProductionOrderIcon_08-20-024625.pdf','ProductionOrderIcon_08-20-024652.pdf']:
        data = parse_oc_pdf(f'/mnt/user-data/uploads/{f}')
        print('=====', f)
        print(json.dumps(data, indent=2, ensure_ascii=False))
