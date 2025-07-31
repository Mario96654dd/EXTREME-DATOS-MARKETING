import streamlit as st
import pandas as pd
import datetime

archivo_excel = "LISTADO DE CLIENTES Y COMERCIALES 2025-06-10 (2).xlsx"

# Leer clientes
hojas = pd.ExcelFile(archivo_excel).sheet_names
nombre_hoja_clientes = hojas[0]
df_clientes = pd.read_excel(archivo_excel, sheet_name=nombre_hoja_clientes)

# Nombre hoja entregas
nombre_hoja_entregas = "ENTREGADO"

# Leer entregas previas o crear DataFrame vacío
try:
    df_entregas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja_entregas)
except:
    df_entregas = pd.DataFrame(columns=[
        "Fecha", "Cliente", "RUC", "Provincia", "Ciudad", "Vendedor",
        "Producto Publicitario", "Cantidad", "Costo Unitario", "Costo Total",
        "Proveedor", "Observaciones"
    ])

st.set_page_config(page_title="Registro de Entregas", layout="wide")
st.title("🎁 Registro de Entregas Publicitarias")

with st.form("form_entrega"):
    fecha = st.date_input("Fecha de entrega", value=datetime.date.today())

    nombre_cliente = st.selectbox("Buscar cliente", df_clientes["nombre_fiscal"].dropna().unique())

    cliente_seleccionado = df_clientes[df_clientes["nombre_fiscal"] == nombre_cliente]
    if not cliente_seleccionado.empty:
        cliente_seleccionado = cliente_seleccionado.iloc[0]
        ruc = cliente_seleccionado["identificacion"]
        provincia = cliente_seleccionado["provincia"]
        ciudad = cliente_seleccionado["ciudad"]
        vendedor = cliente_seleccionado["NUEVO COMERCIAL"]
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

    submitted = st.form_submit_button("Registrar entrega")

    if submitted:
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

        with pd.ExcelWriter(archivo_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_entregas.to_excel(writer, sheet_name=nombre_hoja_entregas, index=False)

        st.success(f"✅ Entrega registrada. Total: ${costo_total:.2f}")


















import streamlit as st
import pandas as pd
import os
from fpdf import FPDF
from datetime import datetime

st.set_page_config(page_title="Activaciones Publicitarias", layout="centered")

EXCEL_FILE = "LISTADO DE CLIENTES Y COMERCIALES 2025-06-10 (2).xlsx"
HOJA_ACTIVACIONES = "ACTIVACIONES"

# Leer clientes desde el archivo
df_clientes = pd.read_excel(EXCEL_FILE)
clientes = df_clientes["nombre_fiscal"].dropna().unique()

st.title("Registro de Activación Publicitaria")

# Selección del tipo de activación
tipo_activacion = st.selectbox("Tipo de Activación", ["EXTREME", "PANTRO"])

# Artículos por tipo
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

# Formulario
with st.form("form_activacion"):
    cliente = st.selectbox("Cliente", clientes)
    fecha = st.date_input("Fecha de Activación", datetime.now())
    observaciones = st.text_area("Observaciones")

    cantidades = {}
    st.write("### Artículos a entregar:")
    for item in articulos:
        cantidades[item] = st.number_input(f"{item}", min_value=0, value=0, step=1)

    enviado = st.form_submit_button("Generar Activación")

# Ruta de guardado de PDF
output_dir = "activaciones_pdf"
os.makedirs(output_dir, exist_ok=True)

# Función para crear el PDF
def generar_pdf(numero, cliente, fecha, tipo, cantidades, observaciones):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"ACTIVACIÓN PUBLICITARIA N° {numero}", ln=True, align='C')

    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, f"Cliente: {cliente}", ln=True)
    pdf.cell(0, 10, f"Fecha: {fecha.strftime('%Y-%m-%d')}", ln=True)
    pdf.cell(0, 10, f"Tipo de Activación: {tipo}", ln=True)

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
    pdf.multi_cell(0, 10, f"Observaciones:\n{observaciones}")

    pdf.ln(20)
    pdf.cell(0, 10, "Aprobado por Gerente General: Paola Villamarín", ln=True)
    pdf.cell(0, 10, "Responsable: Mario Ponce", ln=True)

    filename = f"{output_dir}/activacion_{numero}.pdf"
    pdf.output(filename)
    return filename

# Lógica al enviar formulario
if enviado:
    # Leer hoja ACTIVACIONES
    try:
        df_act = pd.read_excel(EXCEL_FILE, sheet_name=HOJA_ACTIVACIONES)
    except:
        df_act = pd.DataFrame()

    numero = len(df_act) + 1

    nueva_fila = {
        "N°": numero,
        "Cliente": cliente,
        "Fecha": fecha.strftime("%Y-%m-%d"),
        "Tipo": tipo_activacion,
        "Observaciones": observaciones
    }
    for art in cantidades:
        nueva_fila[art] = cantidades[art]

    df_act = pd.concat([df_act, pd.DataFrame([nueva_fila])], ignore_index=True)

    # Guardar en Excel
    with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df_act.to_excel(writer, sheet_name=HOJA_ACTIVACIONES, index=False)

    # Crear PDF
    pdf_file = generar_pdf(numero, cliente, fecha, tipo_activacion, cantidades, observaciones)

    st.success(f"Activación registrada y PDF generado.")
    with open(pdf_file, "rb") as file:
        st.download_button("📄 Descargar PDF", data=file, file_name=os.path.basename(pdf_file), mime="application/pdf")


























import streamlit as st
import pandas as pd
from datetime import datetime
import os
from fpdf import FPDF
import io

# Configurar nombre de archivo base
EXCEL_FILENAME = "LISTADO DE CLIENTES Y COMERCIALES 2025-06-10 (2).xlsx"
EXCEL_SHEET_REGISTRO = "ENTREGADO"
EXCEL_SHEET_ACTIVACION = "ACTIVACIONES"

# Crear archivo si no existe
if not os.path.exists(EXCEL_FILENAME):
    with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, index=False, sheet_name=EXCEL_SHEET_REGISTRO)
        pd.DataFrame().to_excel(writer, index=False, sheet_name=EXCEL_SHEET_ACTIVACION)

# Cargar hojas
df_entregas = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_REGISTRO)
df_activaciones = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_ACTIVACION)

# Convertir fecha si existe
if "Fecha" in df_entregas.columns:
    df_entregas["Fecha"] = pd.to_datetime(df_entregas["Fecha"])
if "Fecha" in df_activaciones.columns:
    df_activaciones["Fecha"] = pd.to_datetime(df_activaciones["Fecha"])

# Combinar para reportes
df_entregas["Tipo"] = "Registro"
df_activaciones["Tipo"] = "Activación"
df_combined = pd.concat([df_entregas, df_activaciones], ignore_index=True)

# Guardar en Excel
with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
    if modo == "Registro de Entrega":
        df_entregas.to_excel(writer, index=False, sheet_name=hoja)
    else:
        df_activaciones.to_excel(writer, index=False, sheet_name=hoja)

# Crear PDF
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)
pdf.cell(200, 10, txt=f"Formulario de {modo}", ln=True, align='C')
pdf.cell(200, 10, txt=f"Cliente: {cliente}", ln=True)
pdf.cell(200, 10, txt=f"Fecha: {fecha}", ln=True)
pdf.cell(200, 10, txt=f"Proveedor: {proveedor}", ln=True)

if modo == "Registro de Entrega":
        pdf.cell(200, 10, txt=f"Artículo: {elementos['Artículo']}, Cantidad: {elementos['Cantidad']}", ln=True)
    else:
        for k, v in cantidades.items():
            if v > 0:
                pdf.cell(200, 10, txt=f"{k}: {v}", ln=True)
        pdf.multi_cell(0, 10, txt=f"Observaciones: {observaciones}")

        pdf.ln(10)
        pdf.cell(200, 10, txt="Autorizado por: Paola Villamarín", ln=True)
        pdf.cell(200, 10, txt="Responsable: Mario Ponce", ln=True)

        pdf_filename = f"{modo.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        pdf.output(pdf_filename)
        st.success(f"Guardado exitosamente en la hoja {hoja} y generado PDF: {pdf_filename}")

st.subheader("📑 Reporte por Cliente y Fecha")

clientes_disponibles = ["Todos"] + sorted(df_combined["Cliente"].dropna().unique())
cliente_filtro = st.selectbox("Filtrar por cliente", clientes_disponibles)
fecha_inicio = st.date_input("Fecha desde", df_combined["Fecha"].min().date())
fecha_fin = st.date_input("Fecha hasta", df_combined["Fecha"].max().date())
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
