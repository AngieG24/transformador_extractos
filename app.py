import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

def transformar_extracto_bbva(df):
    df = df.copy()

    # Fecha en formato dd/mm/yyyy
    df['Fecha'] = pd.to_datetime(df.iloc[:, 0]).dt.strftime("%d/%m/%Y")

    df['Concepto'] = df.iloc[:, 1].astype(str) + " " + df.iloc[:, 2].astype(str) + " " + df.iloc[:, 3].astype(str)
    
    # Asegurarse de que las columnas de valores son numéricas
    df['Valor'] = pd.to_numeric(df.iloc[:, 5],errors="coerce").fillna(0) - pd.to_numeric(df.iloc[:, 4],errors="coerce").fillna(0)
    
    df_final = df[['Fecha', 'Concepto', 'Valor']]
    return df_final

st.title("Transformador de Extractos")

st.subheader("🏦 BBVA")

# Subir archivo
archivo = st.file_uploader("📂 Carga el archivo Excel del extracto", type=["xlsx", "xls"])

if archivo is not None: #Asegurarse de que se ha cargado un archivo
    try:
        df = pd.read_excel(archivo, engine="openpyxl")  # Usa siempre openpyxl para leer
        df = df.iloc[1:].reset_index(drop=True)  # elimina filas 0 y reinicia el índice
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

# -------------------------- Función Banco de Bogotá --------------------------

codigos_dict = {
    "304": "Pago Tarjeta de Credito Banco Bogota Internet o Banca Movil",
    "569": "Pago por Internet Corporativo",
    "715": "Pago impuesto Nal x Int. Corporativo",
    "717": "Recarga Tarjeta Efectiva",
    "853": "Abono compra divisas",
    "670": "Cargo Pago Electronico Planilla Unica No. 000000000000000000000047664212 Nit 0009998600669427",
    "159": "Comision dispersion de pago de proveedores-Otros",
    "183": "Comision transaccion AVAL",
    "502": "Cargo comision consignacion",
    "595": "Comision dispersion pago de nomina",
    "854": "Cargo internacional 20176321000142 Registro Giros",
    "997": "CargoXpor Comision por Cheque de Gerencia Otras oficinas",
    "GT09": "Gravamen Movimientos Financieros",
    "GT10": "Abono ajuste gravamen mov. financieros",
    "GT26": "Cargo IVA",
    "161": "Abono devolucion dispersion pago de proveedores - otros",
    "594": "Abono devolucion dispersion pago de nomina",
    "12": "Pago cheque canje",
    "20": "Consignacion en oficina en cheque",
    "21": "Consignacion nacional en cheque",
    "26": "Transferencia de cuentas del Banco de Bogota",
    "60": "Abono transferencia por canal electrónico",
    "130": "Pago cheque a terceros Chequera No. 3001010628",
    "139": "Pago cheque en oficina Chequera No. 3001010628",
    "160": "Abono dispersion pago a proveedores",
    "164": "Transf",
    "201": "Cargo Dispersion Pago de Proveedores/Otros",
    "216": "Cargo transferencia por Internet o Banca Movil o Swift",
    "220": "Abono transferencia por internet o banca movil",
    "222": "Consignacion local en cheque",
    "276": "Consignacion en cheque",
    "397": "Abono a cuenta en oficina",
    "591": "Cargo Dispersion Pago de Nomina",
    "593": "Cr Ach",
    "611": "Abono por Deposito en Corresponsal de cliente",
    "659": "Cargo Pago de Cartera",
    "929": "Abono por deposito en cajero automatico",
    "995": "CargoXpor Compra de Cheque de Gerencia",
    "2180": "NA",
    "658": "Pago automatico cuota de credito",
    "524": "Cargo reversion recaudo ACH por RECFON",
    "679": "Retefuente remuneracion especial",
    "678": "Interes remuneracion especial",
    "509": "TIMBRE CHEQUERA",
    "508": "IVA CHEQUERA",
    "110": "COMPRA CHEQUERA",
    "GT01": "Intereses ganados",
    "665": "Carga giro empresarial",
    "219": "Pago de servicio o comparendo por canales electronicos",
    "482": "Pago servicio publico por internet o banca movil",
    "106": "Abono por dispersion de fondos por ATH",
    "91": "Abono por recaudos con comprobante",
    "86": "Cargo comision recaudos con comprobante",
    "GT08": "Cargo IVA",
    "570": "Abono recaudo pago electronico ACH",
    "90": "Pago automatico tarjeta de credito",
    "918": "Reversion comision",
    "GT06": "Retencion en la Fuente sobre intereses",
    "GT22": "Intereses por sobregiro",
    "221": "Abono transferencia AVAL por internet o banca movi",
    "644": "Comision PSE",
    "687": "Abono cancelacion CDT en proceso de prescripcion o prescrito",
    "394": "Cargo a cuenta en oficina",
    "938": "Devolucion IVA por ajuste a una comision",
}

def transformar_extracto_bdb(df):
    
    df = df.copy()

    df['id'] = df.iloc[:, 0]
    df['cuenta'] = df.iloc[:, 1].astype(str)
    
    # Columnas de fechas
    df['fecha_ope'] = pd.to_datetime(
        df.iloc[:, 3].astype(str).str.strip(),
        format="%d.%m.%y",
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    df['fecha'] = pd.to_datetime(
        df.iloc[:, 13].astype(str).str.strip(),
        format="%d.%m.%y",
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    df['día'] = pd.to_datetime(df['fecha_ope'], format="%d/%m/%Y", errors="coerce").dt.day
    
    df['numero'] = df.iloc[:, 6].astype(str).str.lstrip('0')

    df['tipo_transaccion'] = df['numero'].astype(str).map(codigos_dict).fillna('Desconocido')
    df['i'] = ""
    df['descripcion'] = ""
    df['it'] = df.iloc[:, 9].astype(str)
    df['provisional'] = ""

    # Importe como número (float)
    df['importe'] = pd.to_numeric(df.iloc[:, 10],errors="coerce").fillna(0)

    # Transformar la columna 'nit' 
    df['nit'] = (
        df.iloc[:, 16]
        .astype(str)                            # aseguramos string
        .str.upper()                            # estandarizamos mayúsculas
        .str.replace(r"[A-Z]", "", regex=True)  # quitamos cualquier letra
        .str.strip()                            # quitamos espacios en blanco    
        .str.lstrip('0')                        # quitamos ceros a la izquierda 
    )

    # Función para aplicar la regla de los 10 dígitos
    def limpiar_nit(valor):
        if not valor or not isinstance(valor, str):  
            return valor
        if re.fullmatch(r"\d{10}", valor):    # si tiene exactamente 10 dígitos
            if 800 <= int(valor[:3]) <= 999:  # si empieza entre 800 y 999
                return valor[:-1]             # quitamos el último dígito
        return valor  # si no cumple, lo dejamos igual

    df['nit'] = df['nit'].apply(limpiar_nit)

    df['nid'] = df.iloc[:, 18].astype(str).str.lstrip('0')
    df['referencia'] = df.iloc[:, 21].astype(str).str.lstrip('0')

    df_final = df[['id', 'cuenta', 'fecha_ope', 'fecha', 'día', 'numero', 'tipo_transaccion', 'i', 'descripcion', 'it', 'provisional', 'importe','nit','nid', 'referencia']]
    return df_final

# -------------------------- Interfaz Streamlit --------------------------

st.subheader("🏦 Banco de Bogotá")

# Subir archivo
archivo = st.file_uploader("📂 Carga el archivo TXT del extracto", type=["txt"])

if archivo is not None: #Asegurarse de que se ha cargado un archivo
    try:
        df = pd.read_csv(archivo, sep=';', decimal=",", encoding='latin1', header=None)
        st.success("Archivo cargado correctamente ✅")
        
        df_transformado = transformar_extracto_bdb(df)
        
        st.subheader("Vista previa del archivo transformado:")
        st.dataframe(df_transformado)


# ---- Descargar en Excel ----
        buffer = BytesIO()
        df_transformado.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)  


        # Aplicar formato con openpyxl
        wb = load_workbook(buffer)
        ws = wb.active

        # Ubicar la columna "importe" (posición 12 en df_final)
        col_importe = df_transformado.columns.get_loc("importe") + 1

        for row in ws.iter_rows(min_row=2, min_col=col_importe, max_col=col_importe):
            for cell in row:
                cell.number_format = '#,##0.00;[Red]-#,##0.00' # miles con "." y decimales con "," y negativos en rojo        

        # Guardar en nuevo buffer
        buffer_final = BytesIO()
        wb.save(buffer_final)
        buffer_final.seek(0)

        st.download_button(
            label="📥 Descargar en Excel",
            data=buffer_final,
            file_name="extracto_BDB_transformado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e: 
        st.error(f"❌ Error al procesar el archivo: {e}"
)