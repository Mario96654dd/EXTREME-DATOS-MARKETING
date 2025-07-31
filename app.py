import streamlit as st
import pandas as pd
from datetime import datetime
import os
from fpdf import FPDF
import io

# Archivo Excel y hojas
EXCEL_FILENAME = "LISTADO DE CLIENTES Y COMERCIALES 2025-06-10 (2).xlsx"
EXCEL_SHEET_REGISTRO = "ENTREGADO"
EXCEL_SHEET_ACTIVACION = "ACTIVACIONES"

# Crear archivo y hojas si no existen
if not os.path.exists(EXCEL_FILENAME):
    with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, index=False, sheet_name=EXCEL_SHEET_REGISTRO)
        pd.DataFrame().to_excel(writer, index=False, sheet_name=EXCEL_SHEET_ACTIVACION)

# Cargar hojas
df_entregas = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_REGISTRO)
df_activaciones = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_ACTIVACION)

# Convertir columna Fecha a datetime si existe
if "Fecha" in df_entregas.columns:
    df_entregas["Fecha"] = pd.to_datetime(df_entregas["Fecha"])
if "Fecha" in df_activaciones.columns:
    df_activaciones["Fecha"] = pd.to_datetime(df_activaciones["Fecha"])

# Cargar clientes para selección
try:
    df_clientes = pd.read_excel(EXCEL_FILENAME, sheet_name=0)
    clientes_lista = df_clientes["nombre_fiscal"].dropna().unique().tolist()
except:
    clientes_lista = []

st.title("🎁 Registro de Entregas Publicitarias")

with st.form("form_entrega"):
    fecha = st.date_input("Fecha de entrega", value=datetime.today().date())

    nombre_cliente = st.selectbox("Buscar cliente", clientes_lista)

    cliente_seleccionado = df_clientes[df_clientes["nombre_fiscal"] == nombre_cliente]
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

        with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_entregas.to_excel(writer, sheet_name=EXCEL_SHEET_REGISTRO, index=False)

        st.success(f"✅ Entrega registrada. Total: ${costo_total:.2f}")
