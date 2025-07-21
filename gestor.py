import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------ FUNCIÓN PARA CONVERTIR A EXCEL -------------------
def convertir_a_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Informe_Limpio')
    output.seek(0)
    return output

# ------------------ CONFIGURACIÓN DE LA PÁGINA -----------------------
st.set_page_config(page_title="Gestor de Informes", layout="wide")

# Colores de Bancamía
rojo_bancamia = "#E30613"
amarillo_bancamia = "#F7B733"

st.markdown(f"<h1 style='color:{rojo_bancamia};'>Gestor de Informes</h1>", unsafe_allow_html=True)
st.markdown("Cargue un archivo Excel, filtre datos, elimine duplicados, reordene columnas y exporte el resultado.")

# ------------------ SUBIR ARCHIVO -------------------
archivo = st.file_uploader("📂 Sube tu archivo Excel", type=["xlsx"])

if archivo:
    xls = pd.ExcelFile(archivo)
    hojas = xls.sheet_names
    hoja_seleccionada = st.selectbox("🗂 Selecciona la hoja a cargar", hojas)
    df = pd.read_excel(xls, sheet_name=hoja_seleccionada)
    df.columns = df.columns.str.strip()  # Limpia espacios en nombres de columnas
    st.success(f"✅ Hoja '{hoja_seleccionada}' cargada exitosamente.")
    st.dataframe(df.head())

    columnas = df.columns.tolist()

    # ------------------ FILTRAR COLUMNAS -------------------
    st.markdown("### 🔎 Filtros por columna")
    filtros = {}
    with st.expander("Aplicar filtros específicos por columna"):
        for col in columnas:
            valores = df[col].dropna().unique()
            if len(valores) > 0 and len(valores) <= 50:
                seleccion = st.multiselect(f"Filtrar '{col}'", valores)
                if seleccion:
                    filtros[col] = seleccion

    for col, valores in filtros.items():
        df = df[df[col].isin(valores)]
    if filtros:
        st.info(f"🔍 Se aplicaron filtros a: {', '.join(filtros.keys())}")

    # ------------------ ELIMINAR DUPLICADOS POR COLUMNA -------------------
    st.markdown("### 🧽 Eliminar duplicados")
    activar_duplicados = st.checkbox("🗑 Activar limpieza de duplicados por columna")
    if activar_duplicados:
        col_dup = st.selectbox("Selecciona la columna para eliminar duplicados", columnas)
        antes = len(df)
        df = df.drop_duplicates(subset=[col_dup])
        despues = len(df)
        st.success(f"✅ {antes - despues} duplicados eliminados usando la columna '{col_dup}'.")

    # ------------------ ORDENAR Y SELECCIONAR COLUMNAS -------------------
    st.markdown("### 🧩 Estructura final del archivo")
    mostrar_seleccion = st.checkbox("✏️ Seleccionar y ordenar columnas")
    if mostrar_seleccion:
        columnas_seleccionadas = st.multiselect(
            "Selecciona y ordena las columnas que deseas mantener",
            columnas,
            default=columnas
        )
        if columnas_seleccionadas:
            df = df[columnas_seleccionadas]
            st.info(f"🧮 Se reordenaron {len(columnas_seleccionadas)} columnas seleccionadas.")

    # ------------------ EXPORTAR RESULTADO -------------------
    st.markdown("### 📤 Exportar archivo")
    if st.button("Generar archivo limpio"):
        excel_bytes = convertir_a_excel(df)
        st.download_button(
            "📥 Descargar archivo Excel",
            data=excel_bytes,
            file_name="informe_limpio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
