import pandas as pd
from datetime import datetime
import time
import os
import tkinter as tk
from tkinter import filedialog
import sys


# ========== INICIO ==========
print("🗂️ Selecciona la carpeta que contiene los archivos Excel...")
root = tk.Tk()
root.withdraw()
ruta = filedialog.askdirectory(title="Seleccionar carpeta de trabajo") + os.sep
if not ruta or ruta == os.sep:
    print("❌ No se seleccionó una carpeta. Finalizando.")
    sys.exit()


inicio = time.time()
hoy = datetime.now()
año_actual = hoy.year
mes_actual = hoy.month
mes_anterior = mes_actual - 1 if mes_actual > 1 else 12

meses_es = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
    'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
    'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
}

# Archivos fijos
archivo_base = 'LC Consolidado.xlsx'
archivo_bmw = 'LANDED POR MES DE INGRESO 2025.xlsx'
archivo_ditec = 'LC Ditec.xlsx'

# Buscar el archivo de Inchcape automáticamente
archivo_inchcape = None
for nombre in os.listdir(ruta):
    if "inchcape" in nombre.lower() and nombre.lower().endswith(".xlsx"):
        archivo_inchcape = nombre
        break

if not archivo_inchcape:
    print("❌ No se encontró un archivo de Inchcape en la carpeta seleccionada.")
    sys.exit()

print("🔄 Cargando archivo base...")
base = pd.read_excel(ruta + archivo_base, engine='openpyxl')

# ==== BMW ====
print("🔄 Cargando BMW...")
df_bmw = pd.read_excel(ruta + archivo_bmw, sheet_name='Landed_2025', header=5, engine='openpyxl')
df_bmw['Mes de Ingreso Normalizado'] = df_bmw['Mes de Ingreso'].astype(str).str.lower().str.strip().map(meses_es)
df_bmw_filtrado = df_bmw[df_bmw['Mes de Ingreso Normalizado'] == mes_anterior]

if not df_bmw_filtrado.empty:
    df_bmw_transformado = pd.DataFrame()
    df_bmw_transformado['Texto cab.documento'] = df_bmw_filtrado['BMW Parts invoice No']
    df_bmw_transformado['año_contabilizacion'] = año_actual
    df_bmw_transformado['mes_contabilizacion'] = df_bmw_filtrado['Mes de Ingreso Normalizado']
    df_bmw_transformado['total_general_dinamica'] = df_bmw_filtrado['Costo Puesto en Bodega']
    df_bmw_transformado['total_fob'] = df_bmw_filtrado['Parts value shown on this invoice (LC)']
    df_bmw_transformado['total_gastos'] = df_bmw_transformado['total_general_dinamica'] - df_bmw_transformado['total_fob']
    df_bmw_transformado['lc_final'] = df_bmw_transformado['total_general_dinamica'] / df_bmw_transformado['total_fob']
    df_bmw_transformado['Origen_2'] = 'Alemania'
    df_bmw_transformado['Proveedor'] = 'BMW'
    df_bmw_transformado['MARCA'] = 'BMW'
    df_bmw_transformado['VIA'] = df_bmw_filtrado['Vía']
    for col in base.columns:
        if col not in df_bmw_transformado.columns:
            df_bmw_transformado[col] = None
    df_bmw_transformado = df_bmw_transformado[base.columns]
    base = pd.concat([base, df_bmw_transformado], ignore_index=True)

# ==== Inchcape ====
print(f"🔄 Cargando Inchcape: {archivo_inchcape}...")
df_inch = pd.read_excel(ruta + archivo_inchcape, engine='openpyxl')
df_inch.columns = df_inch.columns.str.replace('\n', ' ').str.strip()

df_inch_transformado = pd.DataFrame()
df_inch_transformado['Texto cab.documento'] = df_inch['Invoice us$']
df_inch_transformado['año_contabilizacion'] = df_inch['Año']
df_inch_transformado['mes_contabilizacion'] = df_inch['Mes Internación'].astype(str).str.lower().str.strip().map(meses_es)
df_inch_transformado['total_fob'] = df_inch['Fob     $USD'] * df_inch['T-C Valorizacion']
df_inch_transformado['total_general_dinamica'] = df_inch['Valor Bodega $USD'] * df_inch['T-C Valorizacion']
df_inch_transformado['total_gastos'] = df_inch_transformado['total_general_dinamica'] - df_inch_transformado['total_fob']
df_inch_transformado['lc_final'] = df_inch_transformado['total_general_dinamica'] / df_inch_transformado['total_fob']
df_inch_transformado['Origen_2'] = df_inch['Proveedor'].map({
    'SOA': 'USA', 'Subaru Corp.': 'JAPON', 'DFSK': 'CHINA', 'ZNA': 'CHINA'
})
df_inch_transformado['MARCA'] = df_inch['Proveedor'].map({
    'SOA': 'SUBARU', 'Subaru Corp.': 'SUBARU', 'DFSK': 'DFSK', 'ZNA': 'DFSK'
})
df_inch_transformado['Proveedor'] = df_inch['Proveedor']
df_inch_transformado['VIA'] = df_inch['Tipo Transporte']
for col in base.columns:
    if col not in df_inch_transformado.columns:
        df_inch_transformado[col] = None
df_inch_transformado = df_inch_transformado[base.columns]
base = pd.concat([base, df_inch_transformado], ignore_index=True)

# ==== Ditec ====
print("🔄 Cargando Ditec...")
df_ditec = pd.read_excel(ruta + archivo_ditec, header=3, engine='openpyxl')
df_ditec.columns = df_ditec.columns.str.replace('\n', ' ').str.strip()
df_ditec['Mes Normalizado'] = df_ditec['Mes'].astype(str).str.lower().str.strip().map(meses_es)
df_ditec_filtrado = df_ditec[(df_ditec['año'] == 2025) & (df_ditec['Mes Normalizado'] == mes_anterior)]

if not df_ditec_filtrado.empty:
    df_ditec_grouped = df_ditec_filtrado.groupby([
        'DocNumFacturaCompra', 'año', 'Mes', 'via', 'Marca'
    ], as_index=False).agg({'Compra': 'sum', 'Precio Almacén': 'sum'})

    df_ditec_transformado = pd.DataFrame()
    df_ditec_transformado['Texto cab.documento'] = df_ditec_grouped['DocNumFacturaCompra']
    df_ditec_transformado['año_contabilizacion'] = df_ditec_grouped['año']
    df_ditec_transformado['mes_contabilizacion'] = df_ditec_grouped['Mes'].astype(str).str.lower().str.strip().map(meses_es)
    df_ditec_transformado['total_fob'] = df_ditec_grouped['Compra']
    df_ditec_transformado['total_general_dinamica'] = df_ditec_grouped['Precio Almacén']
    df_ditec_transformado['lc_final'] = df_ditec_transformado['total_general_dinamica'] / df_ditec_transformado['total_fob']
    df_ditec_transformado['total_gastos'] = df_ditec_transformado['total_general_dinamica'] - df_ditec_transformado['total_fob']
    df_ditec_transformado['Proveedor'] = df_ditec_grouped['Marca']
    df_ditec_transformado['MARCA'] = df_ditec_grouped['Marca']
    df_ditec_transformado['VIA'] = df_ditec_grouped['via']
    df_ditec_transformado['Origen_2'] = df_ditec_transformado['Proveedor'].map({
        'Volvo': 'SUECIA', 'Jaguar': 'REINO UNIDO', 'Land Rover': 'REINO UNIDO', 'Porsche': 'Alemania'
    })
    for col in base.columns:
        if col not in df_ditec_transformado.columns:
            df_ditec_transformado[col] = None
    df_ditec_transformado = df_ditec_transformado[base.columns]
    base = pd.concat([base, df_ditec_transformado], ignore_index=True)

# ==== Limpiar, Normalizar y Exportar ====
print(f"🚫 Eliminando filas con mes actual ({mes_actual})...")
filas_antes = len(base)
base = base[~((base['año_contabilizacion'] == año_actual) & (base['mes_contabilizacion'] == mes_actual))]
print(f"🧾 Filas eliminadas: {filas_antes - len(base)}")

print("🧹 Normalizando VIA...")
via_mapeo = {
    'AEREA': 'Aereo', 'Courrier': 'Courier', 'SEA FULL': 'Maritimo',
    'SEA LCL': 'Maritimo', 'Marítimo': 'Maritimo', 'Terrestre': 'Terrestre'
}
base['VIA'] = base['VIA'].astype(str).str.strip().replace(via_mapeo)

# Guardar
salida = ruta + 'LC Consolidado FINAL.xlsx'
base.to_excel(salida, index=False)
print(f"✅ Archivo exportado correctamente: {salida}")

fin = time.time()
print(f"⏱️ Tiempo total: {int((fin - inicio) // 60)} min {int((fin - inicio) % 60)} seg")

