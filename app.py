import streamlit as st
import io
import os
from cya_parser import parse_oc_pdf
from cya_filler import fill_packing_list

st.set_page_config(page_title="Packing List C&A", layout="centered")
st.title("Packing List C&A — Llenado automático")
st.caption("Sube la Orden de Compra (PDF) de C&A y completa la capacidad de caja para generar el Excel.")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "PL_ACTUALIZADO.xlsx")

pdf_file = st.file_uploader("Orden de Compra (PDF)", type=["pdf"])

if pdf_file:
    tmp_pdf = "/tmp/_oc_subida.pdf"
    with open(tmp_pdf, "wb") as f:
        f.write(pdf_file.read())

    data = parse_oc_pdf(tmp_pdf)
    header = data["header"]
    sku_order = data["sku_order"]
    sku_table = data["sku_table"]
    packs = data["packs"]
    ratio_packs = [p for p in packs if p["tipo"] == "PACK"]
    solid_packs = [p for p in packs if p["tipo"] == "SKU"]

    st.subheader("Datos detectados")
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Orden de compra:** {header.get('numero_orden')}")
        st.write(f"**Modelo:** {header.get('modelo_id')}")
        st.write(f"**Proveedor:** {header.get('proveedor')}")
    with c2:
        st.write(f"**División:** {header.get('division')}")
        st.write(f"**Color:** {header.get('color_generico')}")
        st.write(f"**Piezas totales OC:** {header.get('piezas_totales')}")

    st.write(f"**Descripción:** {header.get('descripcion_articulo')}")

    if not packs:
        st.info("Esta OC no trae desglose de Pack — todo se maneja como Solid Pack (caja individual por talla).")
    else:
        for p in packs:
            st.write(f"- Pack **{p['letra']}** ({p['tipo']}): {p['unidades_por_pack']} u/pack, "
                     f"{p['total_packs']} packs, {p['total_unidades']} piezas — tallas: {p['tallas']}")

    st.divider()
    st.subheader("Capacidad de caja (según el volumen de esta prenda)")
    st.caption("Esto NO viene en la OC — captúralo según lo que realmente quepa en tu caja para este estilo.")

    capacidad_ratio = {}
    oc_num = header.get("numero_orden", "sinOC")
    for p in ratio_packs:
        letra = p["letra"]
        capacidad_ratio[letra] = st.number_input(
            f"Packs por caja — Pack {letra} (unidades/pack: {p['unidades_por_pack']}, total packs: {p['total_packs']})",
            min_value=0, value=0, step=1, key=f"cap_{oc_num}_{letra}"
        ) or None

    capacidad_solid = st.number_input(
        "Piezas por caja — Solid Pack / remanente (mismo valor para todas las tallas)",
        min_value=0, value=0, step=1, key=f"cap_solid_{oc_num}"
    ) or None

    # Cantidad real a entregar por talla en el Solid Pack (remanente).
    # Precargada con lo planeado en la OC; ajústala si hay faltante de fabricación.
    if solid_packs:
        planeado_solid = {}
        for p in solid_packs:
            for talla, cant in p["tallas"].items():
                planeado_solid[talla] = planeado_solid.get(talla, 0) + cant
    else:
        planeado_solid = {t: sku_table[t]["piezas"] for t in sku_order}

    cantidades_reales_solid = {}
    if planeado_solid:
        st.caption("Cantidad real a entregar del Solid Pack — ajusta si hubo faltante de fabricación (por defecto, lo planeado en la OC).")
        cols = st.columns(len(planeado_solid))
        for i, (talla, planeado) in enumerate(planeado_solid.items()):
            with cols[i]:
                cantidades_reales_solid[talla] = st.number_input(
                    f"{talla} (OC: {planeado})", min_value=0, value=planeado, step=1,
                    key=f"real_{oc_num}_{talla}"
                )

    with st.expander("Campos que no vienen en la OC (opcional)"):
        shipper_code = st.text_input("Shipper Code", value="70135")
        invoice_number = st.text_input("Invoice Number")
        gross_weight = st.number_input("Gross Weight per Order", min_value=0.0, value=0.0) or None
        net_weight = st.number_input("Net Weight per Order", min_value=0.0, value=0.0) or None
        cbm = st.number_input("CBM", min_value=0.0, value=0.0) or None

    if st.button("Generar Excel", type="primary"):
        out_path = "/tmp/_packing_list_generado.xlsx"
        warnings = fill_packing_list(
            TEMPLATE_PATH, out_path, data,
            capacidad_ratio=capacidad_ratio or None,
            capacidad_solid=capacidad_solid,
            cantidades_reales_solid=cantidades_reales_solid or None,
            shipper_code=shipper_code or None,
            invoice_number=invoice_number or None,
            gross_weight=gross_weight, net_weight=net_weight, cbm=cbm,
        )
        if warnings:
            st.warning("Revisa lo siguiente antes de mandar el archivo:\n\n" + "\n".join(f"- {w}" for w in warnings))
        with open(out_path, "rb") as f:
            st.download_button(
                "Descargar Packing List llenado",
                data=f.read(),
                file_name=f"PackingList_C&A_{header.get('numero_orden')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.success("Excel generado. Recuerda: TOTAL CARTONS/PIECES, Number of cartons y Packs per carton se recalculan automáticamente al abrir en Excel.")
else:
    st.info("Sube el PDF de la Orden de Compra para empezar.")
