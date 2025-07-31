import streamlit as st
import pandas as pd
import datetime
import os
from fpdf import FPDF
import io

# Archivo Excel y hojas
EXCEL_FILENAME = "LISTADO DE CLIENTES Y COMERCIALES 2025-06-10 (2).xlsx"
HOJA_ENTREGAS = "ENTREGADO"
HOJA_ACTIVACIONES = "ACTIVACIONES"

# Crear archivo con hojas si no existe
if not os.path.exists(EXCEL_FILENAME):
    with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, index=False, sheet_name=HOJA_ENTREGAS)
        pd.DataFrame().to_excel(writer, index=False, sheet_name=HOJA_ACTIVACIONES)

# Cargar datos de entregas y activaciones
df_entregas = pd.read_excel(EXCEL_FILENAME, sheet_name=HOJA_ENTREGAS)
df_activaciones = pd.read_excel(EXCEL_FILENAME, sheet_name=HOJA_ACTIVACIONES)

# Convertir columnas Fecha a datetime
if "Fecha" in df_entregas.columns:
    df_entregas["Fecha"] = pd.to_datetime(df_entregas["Fecha"], errors='coerce')
if "Fecha" in df_activaciones.columns:
    df_activaciones["Fecha"] = pd.to_datetime(df_activaciones["Fecha"], errors='coerce')

# Cargar clientes para dropdown
try:
    df_clientes = pd.read_excel(EXCEL_FILENAME, sheet_name=0)
    clientes_lista = df_clientes["nombre_fiscal"].dropna().unique().tolist()
except:
    clientes_lista = []

st.set_page_config(page_title="Gestión de Entregas y Activaciones", layout="wide")
st.title("🎁 Registro de Entregas Publicitarias")

# --- FORMULARIO DE ENTREGAS ---
with st.form("form_entrega"):
    fecha = st.date_input("Fecha de entrega", value=datetime.date.today())
    nombre_cliente = st.selectbox("Buscar cliente", clientes_lista)

    cliente_seleccionado = df_clientes[df_clientes["nombre_fiscal"] == nombre_cliente] if not df_clientes.empty else pd.DataFrame()
    if not cliente_seleccionado.empty:
        cliente_seleccionado = cliente_seleccionado.iloc[0]
        ruc = cliente_seleccionado.get("identificacion", "")
        provincia = cliente_seleccionado.get("provincia", "")
        ciudad = cliente_seleccionado.get("ciudad", "")
        vendedor = cliente_seleccionado.get("NUEVO COMERCIAL", "")
    else:
        ruc = provincia = ciudad = vendedor = ""

    st.text_input("RUC", value=ruc, disabled=True)
    st.text_input("Provincia", value=provincia, disabled=True)
    st.text_input("Ciudad", value=ciudad, disabled=True)
    st.text_input("Vendedor", value=vendedor, disabled=True)

    productos = [
        "GORRA EXTREME", "GORRA PANTRO", "GORRA VOLTMAX",
        "CAMISETA EXTREME", "CAMISETA PANTRO", "CAMISETA VOLTMAX",
        "LLAVERO EXTREME", "LLAVERO PANTRO", "LLAVERO VOLTMAX",
        "STIKER EXTREME", "STIKER PANTRO", "STIKER VOLTMAX",
        "TOMATO", "AGENDA", "BOLSAS PUBLICITARIAS",
        "CASCOS", "GAFAS", "GUANTES", "SOMBRILLA",
        "CARPA PANTRO", "CARPA EXTREME",
        "ROMPETRAFICOS EXTREME", "ROMPETRAFICO PANTRO", "ROMPETRAFICO VOLTMAX",
        "LETREROS", "VINILOS", "MICROPERFORADOS",
        "MANDIL EXTREME", "PELUCHES PANTRO"
    ]
    producto = st.selectbox("Producto Publicitario", productos)
    cantidad = st.number_input("Cantidad entregada", min_value=1, step=1)
    costo_unitario = st.number_input("Costo Unitario ($)", min_value=0.0, format="%.2f")
    proveedor = st.text_input("Proveedor (dejar en blanco si no aplica)")
    observaciones = st.text_area("Observaciones (opcional)")

    submitted_entrega = st.form_submit_button("Registrar entrega")

    if submitted_entrega:
        costo_total = round(cantidad * costo_unitario, 2)
        nueva_entrega = {
            "Fecha": fecha,
            "Cliente": nombre_cliente,
            "RUC": ruc,
            "Provincia": provincia,
            "Ciudad": ciudad,
            "Vendedor": vendedor,
            "Producto Publicitario": producto,
            "Cantidad": cantidad,
            "Costo Unitario": costo_unitario,
            "Costo Total": costo_total,
            "Proveedor": proveedor,
            "Observaciones": observaciones
        }
        df_entregas = pd.concat([df_entregas, pd.DataFrame([nueva_entrega])], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_entregas.to_excel(writer, sheet_name=HOJA_ENTREGAS, index=False)

        st.success(f"✅ Entrega registrada. Total: ${costo_total:.2f}")

st.markdown("---")
st.title("📢 Registro de Activaciones Publicitarias")

# --- FORMULARIO DE ACTIVACIONES ---
with st.form("form_activacion"):
    tipo_activacion = st.selectbox("Tipo de Activación", ["EXTREME", "PANTRO"])

    try:
        df_clientes = pd.read_excel(EXCEL_FILENAME, sheet_name=0)
        clientes = df_clientes["nombre_fiscal"].dropna().unique()
    except:
        clientes = []

    cliente = st.selectbox("Cliente", clientes)
    fecha_activacion = st.date_input("Fecha de Activación", datetime.datetime.now())
    observaciones_activacion = st.text_area("Observaciones")

    articulos_extreme = [
        "Inflable tipo arco Extrememax", "Inflable de batería VOLTMAX", "Bandera EXTREME",
        "Mesa y mantel EXTREME", "Gorra EXTREME", "Gorra VOLTMAX", "Camiseta EXTREME", "Camiseta VOLTMAX",
        "Hoja QR", "Hoja de registro", "Carpa EXTREME", "Vestido EXTREME", "Llavero EXTREME", "Llavero PANTRO",
        "Parlante", "Micrófono"
    ]
    articulos_pantro = [
        "Inflable tipo llanta", "Banderas PANTRO", "Mesa y mantel PANTRO", "Vestido PANTRO",
        "Carpa PANTRO", "Camiseta PANTRO", "Llavero PANTRO"
    ]

    articulos = articulos_extreme if tipo_activacion == "EXTREME" else articulos_pantro

    cantidades = {}
    st.write("### Artículos a entregar:")
    for item in articulos:
        cantidades[item] = st.number_input(f"{item}", min_value=0, value=0, step=1)

    submitted_activacion = st.form_submit_button("Generar Activación")

    if submitted_activacion:
        try:
            df_act = pd.read_excel(EXCEL_FILENAME, sheet_name=HOJA_ACTIVACIONES)
        except:
            df_act = pd.DataFrame()

        numero = len(df_act) + 1
        nueva_fila = {
            "N°": numero,
            "Cliente": cliente,
            "Fecha": fecha_activacion.strftime("%Y-%m-%d"),
            "Tipo": tipo_activacion,
            "Observaciones": observaciones_activacion
        }
        for art in cantidades:
            nueva_fila[art] = cantidades[art]

        df_act = pd.concat([df_act, pd.DataFrame([nueva_fila])], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILENAME, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_act.to_excel(writer, sheet_name=HOJA_ACTIVACIONES, index=False)

        # Generar PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"ACTIVACIÓN PUBLICITARIA N° {numero}", ln=True, align='C')

        pdf.set_font("Arial", size=12)
        pdf.cell(0, 10, f"Cliente: {cliente}", ln=True)
        pdf.cell(0, 10, f"Fecha: {fecha_activacion.strftime('%Y-%m-%d')}", ln=True)
        pdf.cell(0, 10, f"Tipo de Activación: {tipo_activacion}", ln=True)

        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(100, 10, "Artículo", 1)
        pdf.cell(40, 10, "Cantidad", 1, ln=True)

        pdf.set_font("Arial", size=12)
        for art, cant in cantidades.items():
            if cant > 0:
                pdf.cell(100, 10, art, 1)
                pdf.cell(40, 10, str(cant), 1, ln=True)

        pdf.ln(10)
        pdf.multi_cell(0, 10, f"Observaciones:\n{observaciones_activacion}")

        pdf.ln(20)
        pdf.cell(0, 10, "Aprobado por Gerente General: Paola Villamarín", ln=True)
        pdf.cell(0, 10, "Responsable: Mario Ponce", ln=True)

        pdf_file = f"activacion_{numero}.pdf"
        pdf.output(pdf_file)

        st.success(f"Activación registrada y PDF generado.")
        with open(pdf_file, "rb") as file:
            st.download_button("📄 Descargar PDF", data=file, file_name=pdf_file, mime="application/pdf")

st.markdown("---")
st.title("📑 Reporte por Cliente y Fecha")

# Preparar df combinado para reportes
df_entregas["Tipo"] = "Registro"
df_activaciones["Tipo"] = "Activación"
df_combined = pd.concat([df_entregas, df_activaciones], ignore_index=True)

clientes_disponibles = ["Todos"] + sorted(df_combined["Cliente"].dropna().unique())
cliente_filtro = st.selectbox("Filtrar por cliente", clientes_disponibles)
fecha_inicio = st.date_input("Fecha desde", value=df_combined["Fecha"].min().date() if not df_combined.empty else datetime.date.today())
fecha_fin = st.date_input("Fecha hasta", value=df_combined["Fecha"].max().date() if not df_combined.empty else datetime.date.today())
tipo_reporte = st.multiselect("Tipos de registro", ["Registro", "Activación"], default=["Registro", "Activación"])

df_reporte = df_combined.copy()
if cliente_filtro != "Todos":
    df_reporte = df_reporte[df_reporte["Cliente"] == cliente_filtro]

df_reporte = df_reporte[
    (df_reporte["Fecha"] >= pd.to_datetime(fecha_inicio)) &
    (df_reporte["Fecha"] <= pd.to_datetime(fecha_fin)) &
    (df_reporte["Tipo"].isin(tipo_reporte))
]

st.write(f"Mostrando {len(df_reporte)} registros filtrados:")
st.dataframe(df_reporte)

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df_reporte.to_excel(writer, index=False, sheet_name="Reporte")
buffer.seek(0)

st.download_button(
    label="⬇️ Descargar reporte filtrado",
    data=buffer,
    file_name="reporte_entregas_activaciones.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
