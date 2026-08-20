# Packing List C&A — Llenado automático

## Archivos
- `app.py` — app de Streamlit (súbela a tu repo de GitHub junto con los demás archivos y despliégala en Streamlit Community Cloud como tus otras apps).
- `cya_parser.py` — lee la OC en PDF y extrae: encabezado, tabla SKU (talla/pieza/PO) y tabla Pack (Ratio o Solid).
- `cya_filler.py` — llena PL_ACTUALIZADO.xlsx respetando todas las fórmulas del template original.
- `PL_ACTUALIZADO.xlsx` — la plantilla en blanco. **Debe vivir en la misma carpeta que app.py** (el código la busca ahí).
- `requirements.txt` — para Streamlit Cloud.
- `ejemplo_1_618978.xlsx`, `ejemplo_2_622526.xlsx`, `ejemplo_3_616404.xlsx` — los 3 escenarios que me mandaste, ya llenados y verificados (0 errores de fórmula, total de piezas cuadra contra la OC).

## Qué llena solo vs. qué capturas tú

**Automático desde la OC:**
- Shipper/Exporter Name, Model No, DIVISION, PURCHASE ORDER
- Los 3 bloques de tallas (Ratio Pack A, Ratio Pack B, Solid Pack C) con sus cantidades
- PACK No., PACK, Pieces Per Pack, Packs per PO / Pieces per PO
- SKU No. de cada talla
- TOTAL CARTONS y TOTAL PIECES (fórmulas del template, se recalculan solas)

**Sigues capturando tú (no viene en la OC):**
- **Packs por caja / Piezas por caja** — esto es justo lo que me dijiste: depende del volumen de la prenda, no de la OC. La app te pide este dato por cada Pack y para el remanente (Solid), y calcula sola cuántas cajas llenas + una caja resto necesitas.
- Shipper Code, Invoice Number, Gross/Net Weight, CBM
- Dimensiones y conteo de cajas por tipo (J7:K10) — el template ya trae tus 2 tipos de caja por default; ajústalos si cambian.

## Los 3 escenarios que ya cubre
1. **Solo Solid Pack** (ej. OC 618978): la OC no trae desglose de Pack, todo se manda como caja individual por talla.
2. **1 Ratio Pack + remanente** (ej. OC 622526): Pack A con razón fija por talla, más un sobrante tipo SKU.
3. **2 Ratio Packs + remanente** (ej. OC 616404): Pack A y Pack B con distintas razones, más el sobrante.

## Cómo correrla localmente para probar
```
pip install -r requirements.txt
streamlit run app.py
```
