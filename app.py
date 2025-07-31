import streamlit as st
import pandas as pd
import os
from fpdf import FPDF
from datetime import datetime
import io

# Configuración inicial
st.set_page_config(page_title="Gestión Publicitaria", layout="wide")
EXCEL_FILENAME = "LISTADO DE CLIENTES Y COMERCIALES 2025-06-10 (2).xlsx"
EXCEL_SHEET_ENTREGAS = "ENTREGADO"
EXCEL_SHEET_ACTIVACION = "ACTIVACIONES"

# Crear archivo si no existe
if not os.path.exists(EXCEL_FILENAME):
    with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, index=False, sheet_name=EXCEL_SHEET_ENTREGAS)
        pd.DataFrame().to_excel(writer, index=False, sheet_name=EXCEL_SHEET_ACTIVACION)

# Cargar hojas
try:
    df_entregas = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_ENTREGAS)
except:
    df_entregas = pd.DataFrame()

try:
    df_activaciones = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_ACTIVACION)
except:
    df_activaciones = pd.DataFrame()

# Convertir fechas
if "Fecha" in df_entregas.columns:
    df_entregas["Fecha"] = pd.to_datetime(df_entregas["Fecha"])
if "Fecha" in df_activaciones.columns:
    df_activaciones["Fecha"] = pd.to_datetime(df_activaciones["Fecha"])

# Combinar para reportes
df_entregas["Tipo"] = "Registro"
df_activaciones["Tipo"] = "Activación"
df_combined = pd.concat([df_entregas, df_activaciones], ignore_index=True)

# ------------------------- Reporte -------------------------
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

# Descargar reporte
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

# ------------------------- PDF Opcional -------------------------

# Puedes usar esta sección para generar PDFs si quieres. Se desactivó la parte con errores.
# Descomenta y define `modo`, `cliente`, `fecha`, `proveedor`, `elementos`, `observaciones`, `cantidades` si la vas a usar.

# modo = "Registro de Entrega"  # o "Activación"
# hoja = EXCEL_SHEET_ENTREGAS if modo == "Registro de Entrega" else EXCEL_SHEET_ACTIVACION
# cliente = "Cliente Ejemplo"
# fecha = datetime.now()
# proveedor = "Proveedor X"
# elementos = {"Artículo": "Gorra EXTREME", "Cantidad": 10}
# observaciones = "Sin observaciones"
# cantidades = {"Carpa EXTREME": 1, "Bandera EXTREME": 2}  # ejemplo

# pdf = FPDF()
# pdf.add_page()
# pdf.set_font("Arial", size=12)
# pdf.cell(200, 10, txt=f"Formulario de {modo}", ln=True, align='C')
# pdf.cell(200, 10, txt=f"Cliente: {cliente}", ln=True)
# pdf.cell(200, 10, txt=f"Fecha: {fecha.strftime('%Y-%m-%d')}", ln=True)
# pdf.cell(200, 10, txt=f"Proveedor: {proveedor}", ln=True)

# if modo == "Registro de Entrega":
#     pdf.cell(200, 10, txt=f"Artículo: {elementos['Artículo']}, Cantidad: {elementos['Cantidad']}", ln=True)
# else:
#     for k, v in cantidades.items():
#         if v > 0:
#             pdf.cell(200, 10, txt=f"{k}: {v}", ln=True)
#     pdf.multi_cell(0, 10, txt=f"Observaciones: {observaciones}")

# pdf.ln(10)
# pdf.cell(200, 10, txt="Autorizado por: Paola Villamarín", ln=True)
# pdf.cell(200, 10, txt="Responsable: Mario Ponce", ln=True)

# pdf_filename = f"{modo.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
# pdf.output(pdf_filename)
# st.success(f"Guardado exitosamente en la hoja {hoja} y generado PDF: {pdf_filename}")
