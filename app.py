import streamlit as st
import pandas as pd
from io import BytesIO

def transformar_extracto_bbva(df):
    df = df.copy()

    # Fecha en formato dd/mm/yyyy
    df['Fecha'] = pd.to_datetime(df.iloc[:, 0]).dt.strftime("%d/%m/%Y")

    df['Concepto'] = df.iloc[:, 1].astype(str) + " " + df.iloc[:, 2].astype(str) + " " + df.iloc[:, 3].astype(str)
    
    # Asegurarse de que las columnas de valores son numéricas
    
    df['Valor'] = pd.to_numeric(df.iloc[:, 4],errors="coerce").fillna(0) - pd.to_numeric(df.iloc[:, 5],errors="coerce").fillna(0)
    df_final = df[['Fecha', 'Concepto', 'Valor']]
    return df_final

st.title("Transformador de Extracto BBVA")

# Subir archivo
archivo = st.file_uploader("📂 Carga el archivo Excel del extracto", type=["xlsx", "xls"])

if archivo is not None: #Asegurarse de que se ha cargado un archivo
    try:
        df = pd.read_excel(archivo, engine="openpyxl")  # Usa siempre openpyxl
        df = df.iloc[2:].reset_index(drop=True)  # elimina filas 0 y 1 y reinicia el índice
        st.success("Archivo cargado correctamente ✅")
        
        df_transformado = transformar_extracto_bbva(df)
        
        st.subheader("Vista previa del archivo transformado:")
        st.dataframe(df_transformado)


# ---- Descargar en Excel ----
        buffer = BytesIO()
        df_transformado.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar en Excel",
            data=buffer,
            file_name="extracto_BBVA_transformado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e: 
        st.error(f"❌ Error al procesar el archivo: {e}"
)
