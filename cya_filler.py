"""
Llenador automático del Packing List C&A (entregas nacionales).

Toma los datos ya extraídos de una OC (ver cya_parser.py) más la capacidad
física de caja (packs por caja / piezas por caja -- esto SIEMPRE lo defines
tú a mano porque depende del volumen de la prenda) y llena PL_ACTUALIZADO.xlsx
respetando exactamente las fórmulas y el formato del template original.
"""
from openpyxl import load_workbook
import copy


def _split_por_capacidad(total, capacidad):
    """Divide un total de packs/piezas en renglones de caja llena + caja resto.

    Devuelve lista de tuplas (cantidad_en_esa_caja, cajas_de_ese_tipo)
    -> en el Excel: (Packs per carton, Number of cartons) o
                    (Pieces per carton, Number of cartons)
    Si no se da capacidad, regresa un solo renglón vacío para llenado manual.
    """
    if not capacidad or capacidad <= 0:
        return [(None, None)]
    llenas = total // capacidad
    resto = total % capacidad
    filas = []
    if llenas > 0:
        filas.append((capacidad, llenas))       # (piezas/packs por caja, num de cajas)
    if resto > 0:
        filas.append((resto, 1))                # caja resto: siempre menor distribución
    if not filas:
        filas.append((0, 0))
    return filas


def fill_packing_list(template_path, output_path, oc_data,
                       capacidad_ratio=None, capacidad_solid=None,
                       shipper_code=None, shipper_name_override=None,
                       gross_weight=None, net_weight=None, cbm=None,
                       invoice_number=None):
    """
    template_path: ruta a PL_ACTUALIZADO.xlsx (plantilla en blanco)
    output_path: ruta del archivo a generar
    oc_data: dict devuelto por cya_parser.parse_oc_pdf
    capacidad_ratio: dict {'A': packs_por_caja, 'B': packs_por_caja} (opcional)
    capacidad_solid: int (piezas por caja, un solo valor para todas las tallas)
                      o dict {talla: piezas_por_caja} si varía por talla
    shipper_code, gross_weight, net_weight, cbm, invoice_number: campos que
      NO vienen en la OC -- captúralos tú; si no se pasan, quedan en blanco
      para que los llenes a mano en el Excel resultante.
    """
    wb = load_workbook(template_path)
    ws = wb['Sheet1']
    header = oc_data['header']
    sku_table = oc_data['sku_table']
    sku_order = oc_data['sku_order']
    packs = oc_data['packs']

    ratio_packs = [p for p in packs if p['tipo'] == 'PACK']
    solid_packs = [p for p in packs if p['tipo'] == 'SKU']

    # --- Encabezado ---
    ws['B2'] = shipper_name_override or header.get('proveedor')
    if shipper_code:
        ws['B3'] = shipper_code
    ws['B4'] = header.get('modelo_id')
    ws['B5'] = header.get('division')
    ws['B6'] = header.get('numero_orden')
    if gross_weight is not None:
        ws['B9'] = gross_weight
    if net_weight is not None:
        ws['B10'] = net_weight
    if cbm is not None:
        ws['B11'] = cbm
    if invoice_number:
        ws['B12'] = invoice_number
    # B7 (TOTAL CARTONS) y B8 (TOTAL PIECES) son fórmulas -> no se tocan

    warnings = []

    # --- Bloques "Ratio Pack" (A: filas 18-19, B: filas 23-24) ---
    ratio_rows = {'A': (18, 19), 'B': (23, 24)}
    for pack in ratio_packs[:2]:  # el template solo soporta A y B
        letra = pack['letra']
        if letra not in ratio_rows:
            warnings.append(f"Pack '{letra}' no tiene bloque Ratio Pack disponible en el template (solo A y B).")
            continue
        row_label, row_qty = ratio_rows[letra]
        if letra == 'A':
            ws.cell(row=row_label, column=2, value=header.get('color_generico'))  # B18
        # tallas -> columnas C..
        col = 3
        for talla in sku_order:
            if talla in pack['tallas']:
                ws.cell(row=row_label, column=col, value=talla)
                cantidad_total = pack['tallas'][talla]
                total_packs = pack['total_packs']
                ratio = cantidad_total / total_packs if total_packs else 0
                if ratio != int(ratio):
                    warnings.append(
                        f"Pack {letra}, talla {talla}: la razón {cantidad_total}/{total_packs} "
                        f"no da un entero exacto ({ratio}). Verifica el prepack."
                    )
                ws.cell(row=row_qty, column=col, value=round(ratio))
                col += 1

    # --- Tabla "PACK No." (filas 27-30) ---
    pack_no_rows = {'A': [27, 28], 'B': [29, 30]}
    for pack in ratio_packs[:2]:
        letra = pack['letra']
        filas_disp = pack_no_rows.get(letra)
        if not filas_disp:
            continue
        cap = (capacidad_ratio or {}).get(letra)
        splits = _split_por_capacidad(pack['total_packs'], cap)
        if len(splits) > len(filas_disp):
            warnings.append(
                f"Pack {letra}: la capacidad dada requiere más de {len(filas_disp)} tipos de caja; "
                f"solo se llenaron las primeras {len(filas_disp)}. Ajusta manualmente."
            )
        for (packs_por_caja, num_cajas), fila in zip(splits, filas_disp):
            ws.cell(row=fila, column=1, value=pack['pack_id'])          # A: PACK No.
            # B ya trae 'A'/'B' precargado en el template
            ws.cell(row=fila, column=3, value=num_cajas)                # C: Number of cartons
            ws.cell(row=fila, column=4, value=packs_por_caja)           # D: Packs per carton
            ws.cell(row=fila, column=5, value=pack['unidades_por_pack'])  # E: Pieces Per Pack
            ws.cell(row=fila, column=6, value=pack['total_packs'])      # F: Packs per PO
            # G ya trae la fórmula =C*D*E

    # --- Bloque "Solid Pack" C (filas 36-37): distribución agregada por talla ---
    # Si no hay packs tipo SKU en la OC (Escenario 1), se usa la tabla SKU completa como "solid".
    if solid_packs:
        agregado_talla = {}
        for p in solid_packs:
            for talla, cant in p['tallas'].items():
                agregado_talla[talla] = agregado_talla.get(talla, 0) + cant
    else:
        agregado_talla = {t: sku_table[t]['piezas'] for t in sku_order}

    col = 3
    for talla in sku_order:
        if talla in agregado_talla:
            ws.cell(row=36, column=col, value=talla)
            ws.cell(row=37, column=col, value=agregado_talla[talla])
            col += 1

    # --- Tabla SKU (Solid Pack), filas 40 en adelante ---
    fila = 40
    max_fila_sku = 55  # filas 56-58 son "MIXED", no aplican a entregas nacionales
    for talla in sku_order:
        if talla not in agregado_talla:
            continue
        cantidad_a_entregar = agregado_talla[talla]
        piezas_po = sku_table[talla]['piezas']
        sku_no = sku_table[talla]['sku']

        cap = capacidad_solid
        if isinstance(capacidad_solid, dict):
            cap = capacidad_solid.get(talla)
        splits = _split_por_capacidad(cantidad_a_entregar, cap)

        for pieces_por_caja, num_cajas in splits:
            if fila > max_fila_sku:
                warnings.append(
                    "Se agotaron los renglones disponibles en la tabla SKU (40-55); "
                    "faltó capturar algunas cajas resto manualmente."
                )
                break
            ws.cell(row=fila, column=1, value=sku_no)              # A: SKU No.
            ws.cell(row=fila, column=2, value='C')                 # B: PACK ('C' por convención)
            ws.cell(row=fila, column=3, value=num_cajas)           # C: Number of cartons
            ws.cell(row=fila, column=4, value=pieces_por_caja)     # D: Pieces per carton
            ws.cell(row=fila, column=5, value=cantidad_a_entregar) # E: Pieces to Delivery
            ws.cell(row=fila, column=6, value=piezas_po)           # F: Pieces per PO
            # G ya trae la fórmula =C*D
            fila += 1

    wb.save(output_path)
    return warnings
