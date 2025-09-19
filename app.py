import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re



# ------------------------------------------------------------
#               Bancos de México 
#  ------------------------------------------------------------

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

# ------------------------------------------------------------
#               Bancos de Colombia 
#  ------------------------------------------------------------

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

cuentas_bancos = {
    "291252245": 1,
    "223589391": 2,
    "040-000016-02": 3,
    "040-000068-06": 4,
    "040-000054-70": 5,
    "040-000038-64": 6,
    "040-000077-86": 7,
    "040-000045-62": 9,
    "4851-0000-3964": 12,
    "4851-6999-6280": 13,
    "IRIS 100598509191": 14,
    "FIC # 8287-1": 15,
    "FIC # 8287-3": 16,
    "FIC # 8287-4": 17,
    "2570": 18,
    "2580": 19,
    "0011-8": 20,
    "0164-0": 21,
    "BBVA 0016": 23,
    "171-2": 25
}

# 1. Diccionario de reglas por banco 

reglas_bancos = {
    "Banco de Bogotá": {
        "columnas": {
            "cuenta": 1,
            "fecha_ope": 3,
            "fecha": 13,
            "numero": 6,
            "it": 9,
            "importe": 10,
            "nit": 16,
            "nid": 18,
            "referencia": 21
        },
        "separador_miles": ".",
        "separador_decimales": ",",
        "codigo_tipo_transaccion": codigos_dict,
        "id": cuentas_bancos
    },

    "Bancolombia": {
        "columnas": {
            "cuenta": 0,
            "fecha_ope": 3,
            "fecha": 3,
            "numero": 6,
            "tipo_transaccion": 7,
            "importe": 5            
        },
        "separador_miles": ".",
        "separador_decimales": ",",
        "id": cuentas_bancos
    }  
}

# 2. Función de limpieza de NIT

# Aplicar la regla de los 10 dígitos
def limpiar_nit(valor):
    if not valor or not isinstance(valor, str):  
        return valor
    if re.fullmatch(r"\d{10}", valor):    # si tiene exactamente 10 dígitos
        if 800 <= int(valor[:3]) <= 999:  # si empieza entre 800 y 999
            return valor[:-1]             # quitamos el último dígito
    return valor                          # si no cumple, lo dejamos igual


# 3. Función para parsear múltiples formatos de fecha

formatos_fecha = [
    "%d.%m.%y",
    "%d.%m.%Y",
    "%d/%m/%Y",
    "%Y%m%d",
    "%Y-%m-%d",
    "%d-%m-%Y"
]

def parsear_fecha_multiple(valor):
    """Intenta convertir un valor a fecha probando varios formatos."""
    if pd.isna(valor):              # Verifica si el valor está vacío o es nulo
        return pd.NaT               # Si el valor es nulo (por ejemplo, está vacío o es NaN), devuelve pd.NaT (que significa "Not a Time", es decir, no hay fecha).
    valor = str(valor).strip()      # Convierte el valor a texto y elimina espacios en blanco al inicio y al final
    for formato in formatos_fecha:  # Intenta convertir el valor a fecha usando varios formatos
        try:
            return pd.to_datetime(valor, format=formato, errors="raise")
        except:
            continue
    return pd.NaT

# 4. Función genérica de transformación para bancos de Colombia

def transformar_extracto(df,banco):
    if banco not in reglas_bancos:
        raise ValueError(f"No hay reglas definidas para el banco '{banco}'")

    reglas = reglas_bancos.get(banco)
    columnas = reglas['columnas']

    df_out = df.copy()
    
    # Mapear columnas según reglas
    
    df_out['cuenta'] = df.iloc[:, columnas['cuenta']].astype(str)
    df_out['id'] = df_out['cuenta'].astype(str).map(cuentas_bancos).fillna('Desconocido')
    df_out['numero'] = df.iloc[:, columnas ['numero']].astype(str).str.lstrip('0')
    
        # Columna tipo_transaccion

    if 'tipo_transaccion' in columnas:
        # Si existe la columna, solo muestra el valor tal cual (como texto)
        df_out['tipo_transaccion'] = df.iloc[:, columnas ['tipo_transaccion']].astype(str)
    else: 
        # Si no existe la columna, busca el código en el diccionario, si no existe muestra 'Desconocido'
        df_out['tipo_transaccion'] = df_out['numero'].astype(str).map(codigos_dict).fillna('Desconocido')

        # Columnas de fechas con función automática
    df_out['fecha_ope'] = (
        df.iloc[:, columnas['fecha_ope']]
        .astype(str)
        .str.strip()
        .apply(parsear_fecha_multiple)
    )
    df_out['fecha_ope'] = df_out['fecha_ope'].dt.strftime("%d/%m/%Y")

    df_out['fecha'] = (
        df.iloc[:, columnas['fecha']]
        .astype(str)
        .str.strip()
        .apply (parsear_fecha_multiple)
    )
    df_out['fecha'] = df_out['fecha'].dt.strftime("%d/%m/%Y")

    df_out['día'] = pd.to_datetime(df_out['fecha_ope'], format="%d/%m/%Y", errors="coerce").dt.day  

        # Importe como número (float)
    df_out['importe'] = pd.to_numeric(df.iloc[:, columnas['importe']],errors="coerce").fillna(0)
   
    df_out['i'] = ""
    df_out['descripcion'] = ""
    df_out['provisional'] = ""

            # Columnas opcionales
    opcionales = {
        'nit': lambda s: 
        s.astype(str)                            # aseguramos string
        .str.upper()                            # estandarizamos mayúsculas
        .str.replace(r"[A-Z]", "", regex=True)  # quitamos cualquier letra
        .str.strip()                            # quitamos espacios en blanco    
        .str.lstrip('0')                        # quitamos ceros a la izquierda 
        .apply(limpiar_nit),                     # aplicamos la regla de los 10 dígitos
        
        'it': lambda s: s.astype(str),
        'nid': lambda s: s.astype(str).str.lstrip('0'),
        'referencia': lambda s: s.astype(str).str.lstrip('0')
    }

    for col, func in opcionales.items():
        if col in columnas:
            df_out[col] = func(df.iloc[:, columnas[col]])
        else:
            df_out[col] = ""
    



    df_final = df_out[['id', 'cuenta', 'fecha_ope', 'fecha', 'día', 'numero', 'tipo_transaccion', 'i', 'descripcion', 'it', 'provisional', 'importe','nit','nid', 'referencia']]
    return df_final

# -------------------------- Interfaz Streamlit --------------------------

st.subheader("🏦 Bancos de Colombia")

#Seleccionar banco
banco_seleccionado = st.selectbox(
    "Selecciona el banco del extracto",
    options=list(reglas_bancos.keys())
)
# Subir archivo
archivos = st.file_uploader(
    "📂 Carga tus extractos",
    type=["txt","csv", "xlsx"],
    accept_multiple_files=True)


# Lista para guardar los resultados de cada archivo transformado

dfs_transformados = []

if archivos is not None: # Verifica si hay archivos cargados
    archivos_cargados = []
    for archivo in archivos:
        try:
            nombre = archivo.name.lower()

            # Detectar tipo de archivo por extensión
            if nombre.endswith(".txt"):
                df = pd.read_csv(archivo, sep=';', decimal=",", encoding='latin1', header=None)
            elif nombre.endswith(".csv"):
                df = pd.read_csv(archivo, sep=",", decimal=".", encoding="latin1", header=None)
            elif nombre.endswith(".xlsx"):
                df = pd.read_excel(archivo, header=None)
            else:
                st.warning(f"Formato no compatible: {archivo.name}")
                continue    

                  
            # Transformar archivo
            df_transformado = transformar_extracto(df, banco=banco_seleccionado)
            dfs_transformados.append(df_transformado)

            archivos_cargados.append(f"✅ {archivo.name}")
        
            #Guadar en la lista
            dfs_transformados.append(df_transformado)

        except Exception as e: 
            archivos_cargados.append(f"❌ {archivo.name} (Error: {e})")     
    
    # Mostrar resumen en un expander
    with st.expander("Ver archivos cargados y estado"):
        for estado in archivos_cargados:
            st.write(estado)

    if dfs_transformados:
        df_transformado = pd.concat(dfs_transformados, ignore_index=True)

        st.success("✅ Archivos procesados y consolidados correctamente") 

        # Ordenar por id de forma ascendente
        df_transformado = df_transformado.sort_values(by="id", ascending=True).reset_index(drop=True)

            # Mostrar vista previa del consolidado
        st.subheader(f"Vista previa:")
        st.dataframe(df_transformado.head(5))   
            
            # Descargar archivo en Excel
        buffer = BytesIO()
        df_transformado.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)

            # Aplicar formato con openpyxl
        wb = load_workbook(buffer)
        ws = wb.active        

            
        if "importe" in df_transformado.columns:
            # Ubicar la columna "importe" (posición 12 en df_final)
            col_importe = df_transformado.columns.get_loc('importe') + 1
            for row in ws.iter_rows(min_row=2, min_col=col_importe, max_col=col_importe):
                for cell in row:
                    cell.number_format = '#,##0.00;[Red]-#,##0.00' # miles con "." y decimales con "," y negativos en rojo        

            # Guardar en nuevo buffer
            buffer_final = BytesIO()
            wb.save(buffer_final)
            buffer_final.seek(0)

            st.download_button(
                label="📥 Descargar extractos consolidados",
                data=buffer_final,
                file_name="extractos_transformados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  






