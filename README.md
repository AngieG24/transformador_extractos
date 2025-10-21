# 🧾 Transformador y Consolidador de Extractos Bancarios

## 📋 Descripción
Este proyecto es una aplicación desarrollada en **Python** utilizando **Streamlit**, **Pandas** y **OpenPyXL**, que automatiza la transformación, limpieza y consolidación de extractos bancarios de diferentes bancos de **México** y **Colombia**.  
Su objetivo es facilitar el trabajo contable, reduciendo tiempos de procesamiento.

## 🚀 Funcionalidades principales
- Carga de archivos `.txt`, `.csv` o `.xlsx` desde la interfaz web.
- Identificación automática del banco y aplicación de reglas personalizadas.
- Limpieza y estandarización de columnas (fechas, importes, NIT, referencias, etc.).
- Consolidación de múltiples extractos en un solo archivo.
- Exportación a Excel con formato contable (miles, decimales, negativos en rojo).
- Compatibilidad con bancos de **México** (BBVA, Banorte, Edenred) y **Colombia** (Banco de Bogotá, Bancolombia, Davivienda).

## ⚙️ Tecnologías utilizadas
- **Python 3.11**
- **Streamlit**
- **Pandas**
- **OpenPyXL**
- **Pathlib**
- **Regex (re)**

## 🧠 Lógica del proyecto
El sistema utiliza un **diccionario de reglas por banco** que mapea la posición de cada columna en el archivo original.  
Cada banco tiene su propio formato, por lo que el transformador identifica automáticamente las columnas relevantes y aplica funciones de:
- Limpieza de texto y números.
- Conversión de fechas a formato estándar (`dd/mm/yyyy`).
- Cálculo del importe final según cargos y abonos.
- Estandarización de nombres de cuenta y referencias.

## 🖥️ Interfaz (Streamlit)
- Selecciona el país y el banco.
- Carga uno o varios extractos.
- Visualiza los primeros registros transformados.
- Descarga el consolidado en formato Excel.

## 📦 Instalación y uso
1. Clona el repositorio:
   ```bash
   git clone https://github.com/tuusuario/transformador-extractos.git
   cd transformador-extractos
Instala las dependencias:

bash
Copiar código
pip install -r requirements.txt
Ejecuta la aplicación:

bash
Copiar código
streamlit run app.py
Abre la interfaz en tu navegador (por defecto: http://localhost:8501)

📊 Ejemplo de salida
cuenta	fecha	fecha_ope	concepto	importe	ref 1	ref 2
BBVA 1234	01/09/2025	01/09/2025	Transferencia recibida	500000	-	-

📁 Estructura del proyecto

├── app_cb.py                 # Código principal del proyecto
├── README.md                 # Documentación del proyecto
├── requirements.txt          # Dependencias del entorno
└── extracto_transformado.xlsx         # Archivo resultante (se genera automáticamente)


La interfaz para el usuaario final la encuentras en el siguiente link: https://transformadorextractos-pwmwnpghg6npw7uvam6kpx.streamlit.app/

👩‍💼 Autora

Angie Galindo
Contadora Pública | Analista de Datos
💡 Apasionada por la automatización de procesos contables mediante Python y herramientas de análisis de datos.
📫 LinkedIn: https://www.linkedin.com/in/angielorenagalindo/
