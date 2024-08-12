# %%
import pandas as pd
import numpy as np
import os
import datetime
import win32com.client
import time
import getpass
from datetime import timedelta

usuario = getpass.getuser()

inicio = time.time()
# %%
import os
from datetime import datetime


# %%
import os
import pandas as pd
from datetime import datetime, timedelta

# Define the directory path
carpeta_fechas = "C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Actualización diaria fechas DT/Actualización Diaria fechas Dts R3"

# Maximum number of days to look back
max_days_back = 30

# Function to try reading the file for a given date
def try_read_file(date_str):
    ruta_arch = f'{date_str} Actualizacion fechas diaria  Dts OEM.xlsx'
    ruta = os.path.join(carpeta_fechas, ruta_arch)
    if os.path.exists(ruta):
        df_fechas = pd.read_excel(ruta, sheet_name='Data')
        print(f"File found: {ruta}")
        return df_fechas
    return None

# Try to find a file for today and up to max_days_back days ago
for days_back in range(max_days_back + 1):
    date_to_try = (datetime.today() - timedelta(days=days_back)).strftime('%d-%m-%Y')
    df_fechas = try_read_file(date_to_try)
    if df_fechas is not None:
        break
else:
    print("No file found within the given date range.")

# Continue with your further processing
if df_fechas is not None:
    # Do something with df_fechas
    pass
else:
    # Handle the case where no file was found
    pass




# %%
df_fechas['Nro. DT'].to_clipboard(index = False, header = False)

# %%


# %%
df_fechas['Nro. DT'].nunique()

# %%
try:
    # Initialize SAP GUI scripting
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    if not application:
        raise Exception("Error obtaining SAP GUI application")

    connection = application.Children(0)
    if not connection:
        raise Exception("Error obtaining SAP GUI connection")

    session = connection.Children(0)
    if not session:
        raise Exception("Error obtaining SAP GUI session")

    # Connect to WScript if available
    try:
        WScript = win32com.client.Dispatch("WScript")
        WScript.ConnectObject(session, "on")
        WScript.ConnectObject(application, "on")
    except Exception as e:
        print(f"Error connecting to WScript: {str(e)}")

    # Maximize the window
    session.findById("wnd[0]").maximize()

    # Execute transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "zmm_seguim_comex_cl"
    session.findById("wnd[0]").sendVKey(0)

    # Press buttons
    session.findById("wnd[0]/usr/btn%_P_TKNUM_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)

    # Export to Excel
    session.findById("wnd[0]/usr/cntlALV_COMEX/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlALV_COMEX/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\lravlic\\Codigos\\automatizacion_gere_comex"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "comex.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close SAP windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

except Exception as e:
    print(f"Error during SAP GUI interaction: {str(e)}")






# %%
# %%
import xlwings as xw
try:
    book = xw.Book("C:/Users/lravlic/Codigos/automatizacion_gere_comex/comex.XLSX")
    book.close()
except Exception as e:
    print(e)


# %%
comex = pd.read_excel("C:/Users/lravlic/Codigos/automatizacion_gere_comex/comex.XLSX")

# %%
dts_a_mir6 = comex[comex['Estatus']=='CERRADO']

# %%
dts_a_mir6.Estatus.value_counts()

# %%
dts_a_mir6['Nro. DT asociado'] = dts_a_mir6['Nro. DT asociado'].astype(str)

# Create the new column 'DT' by concatenating '*' before and after the values
dts_a_mir6['DT'] = '*' + dts_a_mir6['Nro. DT asociado'] + '*'

# %%
dts_a_mir6['DT'].to_clipboard(index=False,header=False)

# %%
import win32com.client

# Initialize SAP GUI scripting
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Maximize the window
session.findById("wnd[0]").maximize()

# Enter transaction code "mir6"
session.findById("wnd[0]/tbar[0]/okcd").text = "mir6"
session.findById("wnd[0]").sendVKey(0)

# Clear and set the focus on the user field
session.findById("wnd[0]/usr/txtSO_USNAM-LOW").text = ""
session.findById("wnd[0]/usr/txtSO_USNAM-LOW").setFocus()
session.findById("wnd[0]/usr/txtSO_USNAM-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey(0)

# Set focus on the text field and perform actions
session.findById("wnd[0]/usr/txtSO_BKTXT-LOW").setFocus()
session.findById("wnd[0]/usr/txtSO_BKTXT-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_SO_BKTXT_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()

# Toggle checkboxes
session.findById("wnd[0]/usr/chkP_IV_BG").setFocus()
session.findById("wnd[0]/usr/chkP_IV_BG").selected = False
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/usr/chkP_IV_BG").selected = True
session.findById("wnd[0]/usr/chkP_IV_IS").setFocus()
session.findById("wnd[0]/usr/chkP_IV_IS").selected = True
session.findById("wnd[0]/usr/chkP_IV_OV").setFocus()
session.findById("wnd[0]/usr/chkP_IV_OV").selected = True
session.findById("wnd[0]/usr/chkP_IV_STO").setFocus()
session.findById("wnd[0]/usr/chkP_IV_STO").selected = True
session.findById("wnd[0]/usr/chkP_IV_EDI").setFocus()
session.findById("wnd[0]/usr/chkP_IV_EDI").selected = True
session.findById("wnd[0]/usr/chkP_IV_RAP").setFocus()
session.findById("wnd[0]/usr/chkP_IV_RAP").selected = True
session.findById("wnd[0]/usr/chkP_IV_BAP").setFocus()
session.findById("wnd[0]/usr/chkP_IV_BAP").selected = True
session.findById("wnd[0]/usr/chkP_IV_PAR").setFocus()
session.findById("wnd[0]/usr/chkP_IV_PAR").selected = True
session.findById("wnd[0]/usr/chkP_IV_ERS").setFocus()
session.findById("wnd[0]/usr/chkP_IV_ERS").selected = True
session.findById("wnd[0]/usr/chkP_IV_SRM").setFocus()
session.findById("wnd[0]/usr/chkP_IV_SRM").selected = True
session.findById("wnd[0]/usr/chkP_IV_TP").setFocus()
session.findById("wnd[0]/usr/chkP_IV_TP").selected = True
session.findById("wnd[0]/usr/chkP_IV_A2A").setFocus()
session.findById("wnd[0]/usr/chkP_IV_A2A").selected = True
session.findById("wnd[0]/usr/chkP_IV_B2B").setFocus()
session.findById("wnd[0]/usr/chkP_IV_B2B").selected = True

# Execute commands and navigate
session.findById("wnd[0]").sendVKey(8)
session.findById("wnd[0]/tbar[0]/btn[86]").press()
session.findById("wnd[0]/tbar[1]/btn[33]").press()
session.findById("wnd[1]/usr/lbl[1,6]").setFocus()
session.findById("wnd[1]/usr/lbl[1,6]").caretPosition = 7
session.findById("wnd[1]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[43]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Set the file path for export
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\lravlic\\Codigos\\land_cost"
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 34
session.findById("wnd[1]/tbar[0]/btn[11]").press()

# Close the windows
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)


# %%

try:
    book = xw.Book("C:/Users/lravlic/Codigos/land_cost/export.XLSX")
    book.close()
except Exception as e:
    print(e)


# %%
df_mir6 = pd.read_excel("C:/Users/lravlic/Codigos/land_cost/export.XLSX")

# %%
df_mir6_filtro_yr = df_mir6[df_mir6['Fe.contabilización'].dt.year >=2024]

# %%
df_mir6_filtro_yr['Texto cab.documento'] = pd.to_numeric(df_mir6_filtro_yr['Texto cab.documento'], errors='coerce')

# Convert the column to integers, dropping NaN values first
df_mir6_filtro_yr = df_mir6_filtro_yr.dropna(subset=['Texto cab.documento'])
df_mir6_filtro_yr['Texto cab.documento'] = df_mir6_filtro_yr['Texto cab.documento'].astype(int)

# Display the DataFrame to verify the changes
print(df_mir6_filtro_yr['Texto cab.documento'])

# %%
df_mir6_filtro_yr['Texto cab.documento'] = df_mir6_filtro_yr['Texto cab.documento'].astype('str')

# %%
filtered_df = df_mir6_filtro_yr[df_mir6_filtro_yr['Texto cab.documento'].str.len() <= 6]

# Print the filtered DataFrame
print(filtered_df['Texto cab.documento'])
filtered_df.to_excel('prueba.xlsx')

# %%
filtered_df.dtypes

# %%


# %%
df_mir6_filtro_yr.to_excel('export_landcost_final.xlsx')

# %%
columnas = ['Nro. DT','Nombre del Embarcador','Vía (Texto)', 'País Origen','Pto. Origen', 'Fe. ATA']
df_fechas_columnas = df_fechas[columnas]

# %%
df_fechas_columnas['Nro. DT'] = df_fechas_columnas['Nro. DT'].astype('str')

# %%
df_cruce_fechas = df_mir6_filtro_yr.merge(df_fechas_columnas, left_on='Texto cab.documento',right_on='Nro. DT', how='left')

# %%


# %%
final = time.time()

# %%
time_difference = final - inicio

# Calculate minutes and seconds
minutes, seconds = divmod(time_difference, 60)

print(f"{minutes} minutos y {seconds} segundos")

# %%



