
import pandas as pd
import datetime
import os
import numpy as np
import getpass
# hoy = datetime.datetime.today() dejar esta linea cuadno se haga el calculo real
hoy = datetime.datetime.today()
#LECTURA DE DFS
from pathlib import Path
usuario = getpass.getuser()

# %%
ruta = f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}'
ruta_repo = Path(ruta)

# %%
import pandas as pd

def excel_to_dataframe(xl_name: str, sh_name: str) -> pd.DataFrame:
    """
    Convert an Excel sheet to a pandas DataFrame.

    :param xl_name: The path to the Excel file.
    :param sh_name: The name of the sheet to be read.
    :return: A pandas DataFrame containing the data from the specified Excel sheet.
    """
    # Load the Excel file
    xls = pd.ExcelFile(xl_name)
    
    # Parse the specified sheet into a DataFrame
    df = xls.parse(sh_name)
    
    return df

# Example usage:



# %%
# Leer el archivo CSV en un DataFrame
print(ruta_repo)
ruta_ddp = ruta_repo.joinpath('DDP.csv')
df_ddp_1 = pd.read_csv(ruta_ddp)

#Archivo Tr
ruta_tubo = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal"
carpetas_tubo = os.listdir(ruta_tubo)
tubo = carpetas_tubo[-2]
ubi_tubo = ruta_tubo + '/' + tubo + '/' + tubo
archivo_tr = ruta_tubo + '/' + tubo + '/' + tubo + ' - TR FINAL R3.xlsx'
archivo_monitor = ruta_tubo + '/' + tubo + '/' + tubo + ' - Monitor R3.xlsx'

df_tr = pd.read_excel(archivo_tr,sheet_name = 'Sheet1', dtype = {'AUX':'str'})
df_mon = pd.read_excel(archivo_monitor, sheet_name = 'Sheet1',dtype = {'Nro. OC SAP':'str','Posición OC SAP':'str'} )
print("Fecha tubo "+ '\n' + tubo)






# %%
df_mon.dtypes

# %%
df_tr= df_tr[df_tr['Grupo de compras'].isin(['RR1','RR3','RR9'])]

# %%
data = [
    '40000139   Mazda Motor Corporation',
    '40000176   GREAT WALL MOTOR COMPANY',
    '40000199   Changan International Co',
    '40000214   MAGYAR SUZUKI CORPORATIO',
    '40000234   Maruti Suzuki India Ltd.',
    '40000236   Renault',
    '40000238   RENAULT DO BRASIL COMERC',
    '40001373   ANHUI JIANGHUAI AUTOMOBI',
    '40002141   Suzuki Motor (Thailand)',
    '40002445   SHENZHEN KOEN ELECTRONIC',
    '40002536   Hefei Winning Auto Parts',
    '40002814   PROMASTER ELECTRONIC CO.',
    '40003055   Renault Samsung Motors C',
    '40003191   SUZUKI MOTOR CORPORATION',
    '40003215   JSN HOLDINGS LLC',
    '40000301   IMCRUZ COMERCIAL S.A.',
    '40003526   PT SUZUKI INDOMOBIL SALE',
    '40003767   CHONGQING SHINDARY INDUS',
    '40003770   BEIJING XUN AN DA TRADE',
    '40003810   GLOBALTRONIC CIA. LTDA.',
    'PU02       DERCO PERU S.A.',
    '40000186   Sociedad de Fabricacion',
    '40000280   MAZDA NORTH AMERICAN OPE',
    '40000460   DERCO COLOMBIA SAS',
    '40003304   CHANGZHOU SUNWOOD INTERN',
    '40003805   Great Wall Motor Middle',
    '40003865   Mobitech Co., Ltd'


]


# %%
df_tr = df_tr[df_tr['Nombre del proveedor'].isin(data)]

# %%
#df_tr = df_tr[df_tr['Cl.documento compras'].isin(['ZSTO', 'ZSPT','ZIPL'])]

# %%
df_tr['Cl.documento compras'].value_counts()

# %%
df_ddp_1.drop_duplicates(subset='Material', inplace=True)

# %%
df_ddp_1 = df_ddp_1[['Material','Familia','Costo UN CLP','Segmentacion','Plan mantención',]]

# %%
columnas= ['Nro_pieza_fabricante_1',	'Cod_Actual_1']
ruta_cod = ruta_repo.joinpath('COD_ACTUAL.csv')

# Leer el archivo CSV en un DataFrame
cadena_de_remplazo = pd.read_csv(ruta_cod)
cadena_de_remplazo = cadena_de_remplazo[columnas]

# %%
df_ddp_1 = df_ddp_1.merge(cadena_de_remplazo, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')
df_ddp_1['Cod_Actual_1'] = df_ddp_1['Cod_Actual_1'].fillna(df_ddp_1['Material'])
df_ddp_1 = df_ddp_1.drop('Nro_pieza_fabricante_1', axis=1)

# %%
df_ddp_1.drop_duplicates(subset='Cod_Actual_1', inplace=True, keep='first')

# %%
df_tr.dtypes

# %%
df_tr= df_tr[df_tr['Cantidad']>0]

# %%
columnas = ['AUX','Status','Material','Texto breve','Cantidad','Fecha','Nombre del proveedor','NomSector_actual','Cl.documento compras']

# %%
df_tr = df_tr[columnas]

# %%
import numpy as np
conditions = [
    df_tr['Status'].str.contains('Facturado')
]

# Define choices
choices = ['Facturado']

# Apply np.select
df_tr['Status_V2'] = np.select(conditions, choices, default='No Facturado')


# %%
df_tr['Status_V2'].value_counts()

# %%
columnas_mon = ['Prefijo OC','Nro. OC SAP','Posición OC SAP','Cód. Mat de prov en OC','Vía (Texto)','Pto. Origen','Pto. Destino','Nro. DT','Fecha Creación OC', 'Fe. ATA','Nro. SOLPED','Cant. UN. OC','Nivel de Urgencia']
df_mon = df_mon[columnas_mon]


# %%
df_mon['Nro. OC SAP'] = df_mon['Nro. OC SAP'].astype('str')
df_mon['Posición OC SAP'] = df_mon['Posición OC SAP'].astype('str')

# %%
df_mon['AUX'] = df_mon['Nro. OC SAP'] + df_mon['Posición OC SAP']

# %%
df_mon.drop_duplicates(subset='AUX', inplace=True)

# %%
df_tr['AUX'] = df_tr['AUX'].astype('str')

# %%
df_mon.nunique()
print(df_mon.shape)

# %%
df_base = df_tr.merge(df_mon, left_on='AUX',right_on='AUX', how='left')


# %%
import datetime
hoy = datetime.datetime.today()
#hoy = datetime.date(2024,8,20)
hoy = pd.to_datetime(hoy)



# %%
df_base['dias_para_oc'] = hoy - df_base['Fecha Creación OC'] 

# %%
df_base.columns

# %%
df_base['dias_para_oc'] =  df_base['dias_para_oc'].dt.days

# %%


# %%


# %%
# Sample DataFrame for demonstration purposes
import pandas as pd
import numpy as np




def evaluate_conditions(row):
    am4 = row['Cl.documento compras']
    x4 = row['Vía (Texto)']
    d4 = row['Status_V2']
    j4 = row['dias_para_oc']
    
    def maritime_or_terrestrial(j4, d4):
        if d4 == "Facturado":
            return d4
        elif j4 > 75:
            return "Si"
        elif 45 <= j4 <= 75:
            return "Alerta 45 dias"
        else:
            return "Dentro de los 45 dias"
    
    def aerial_or_courier(j4, d4, days):
        if d4 == "Facturado":
            return d4
        elif j4 > days:
            return "Si"
        elif days // 2 <= j4 <= days:
            return f"Alerta {days // 2} dias"
        else:
            return f"Dentro de los {days // 2} dias"
    
    if am4 in ["ZVOR", "ZATT"]:
        if x4 in ["Marítimo", "Terrestre"] or pd.isna(x4) or x4 == 0:
            return maritime_or_terrestrial(j4, d4)
        elif x4 in ["Aéreo", "Courier"]:
            return aerial_or_courier(j4, d4, 12)
    else:
        if x4 in ["Marítimo", "Terrestre"] or pd.isna(x4) or x4 == 0:
            return maritime_or_terrestrial(j4, d4)
        # elif x4 in ["Aéreo", "Courier"]:
        #     return aerial_or_courier(j4, d4, 30)
        else:  # Handle the extra condition
            if d4 == "Facturado":
                return d4
            elif j4 > 30:
                return "Si"
            elif 20 <= j4 <= 30:
                return "Alerta 20 dias"
            else:
                return "Dentro de los 20 dias"
    
    return None  # Default case, if any condition is not met

# Apply the function to the DataFrame
df_base['Result'] = df_base.apply(evaluate_conditions, axis=1)




# %%
df_base['Semana'] = df_base['Fecha'].dt.isocalendar().week
df_base['Mes'] = df_base['Fecha'].dt.month
df_base['Año'] = df_base['Fecha'].dt.year

# %%


# %%
df_base = df_base.merge(df_ddp_1, left_on='Material',right_on='Cod_Actual_1', how='left')

# %%


# %%
df_base.shape

# %%
oc_antiguo = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/Bases OC-Producción/OC-Producción Sem26 24-06-2024.xlsx", sheet_name='Consolidado TR OEM', header=2, dtype={'Documento compras':'str'})

# %%
oc_antiguo.columns

# %%
oc_antiguo = oc_antiguo[['Material Actual','Documento compras','Fechas corregidas R3']]

# %%
oc_antiguo['Documento compras'] = oc_antiguo['Documento compras'].astype('str')

# %%
oc_antiguo['AUX_2'] =  oc_antiguo['Documento compras'] + oc_antiguo['Material Actual'] 

# %%
oc_antiguo.drop_duplicates(subset='AUX_2', inplace=True)

# %%
df_base['AUX_2'] = df_base['Nro. OC SAP'] + df_base['Material_x']

# %%
df_base = df_base.merge(oc_antiguo, on='AUX_2', how='left')

# %%
df_base['Fechas corregidas R3'].value_counts()

# %%
df_base['Fechas corregidas R3'].fillna('No', inplace=True )

# %%
df_base['Fecha Creación OC'] = df_base.apply(lambda row: row['Fechas corregidas R3'] if row['Fechas corregidas R3'] != 'No' else row['Fecha Creación OC'], axis=1)

# %%


# %%
df_base

# %%
df_base = df_base.drop(columns=['AUX_2', 'Material_y'])

# %%
df_base.rename(columns={'dias_para_oc':'Hoy - OC (dias)','Result':'Fuera de Estandar'}, inplace = True)

# %%
df_base['Fecha Reporte'] = hoy

# %%
columnas = [ 'Documento compras', 'Material Actual','Cod_Actual_1']
df_base.drop(columns=columnas, inplace=True)

# %%
df_base.to_excel(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/bases python/oc_produccion_semana_{hoy.isocalendar()[1]}.xlsx')
df_base.rename(columns ={'Material_x':'Cod_Actual_1'}, inplace=True)
df_base['AUX'] = df_base['AUX'].astype('str')
#AQUI COMIENZA EVOLUTIVO BO


import os
import pandas as pd
import re
import getpass

# Obtener el nombre de usuario del sistema
# usuario = getpass.getuser()

# # Ruta de la carpeta donde se encuentran los archivos, usando el nombre del usuario dinámicamente
# folder_path = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/bases python/"

# # Filtrar los archivos que contengan "oc_produccion_semana_" en el nombre
# files = [f for f in os.listdir(folder_path) if re.match(r'oc_produccion_semana_\d+\.xlsx', f)]

# # Extraer los números de semana
# week_numbers = [int(re.search(r'(\d+)', f).group()) for f in files]

# # Encontrar el archivo correspondiente a la última semana
# latest_week = max(week_numbers)
# latest_file = f"oc_produccion_semana_{latest_week}.xlsx"

# # Cargar el archivo correspondiente a la última semana
# df_oc = pd.read_excel(os.path.join(folder_path, latest_file))

# # Imprimir el nombre del archivo cargado
# print(f"Archivo cargado: {latest_file}")



# %%
df_oc = df_base

# %%
import pandas as pd
import datetime


# %%
import win32com.client
import getpass
usuario = getpass.getuser()
# Get the current user's username
usuario = getpass.getuser()

# Initialize SAP GUI Scripting
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine

# Establish connection and session
connection = application.Children(0)
session = connection.Children(0)

# Maximize the window
session.findById("wnd[0]").maximize()

# Enter transaction code
session.findById("wnd[0]/tbar[0]/okcd").text = "ME2N"
session.findById("wnd[0]").sendVKey(0)

# Set layout
session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV"

# Set plant codes
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").setFocus()
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtSELPA-LOW").text = "we101"
session.findById("wnd[0]/usr/ctxtSELPA-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1305"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1344"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1335"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]").sendVKey(8)

# Set date range
# session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = "01.01.2023"
# session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").text = "31.12.2024"
# session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").setFocus()
# session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").caretPosition = 10
session.findById("wnd[0]").sendVKey(8)

# Execute the transaction
session.findById("wnd[0]/tbar[1]/btn[33]").press()

# Select a specific row
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 46
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 38
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "46"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

# Export to Excel
session.findById("wnd[0]/tbar[1]/btn[43]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey(4)
session.findById("wnd[2]/usr/ctxtDY_PATH").text = f"C:\\Users\\{usuario}\\Documents\\SAP\\SAP GUI"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "ME2N.XLSX"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[2]/tbar[0]/btn[11]").press()
session.findById("wnd[1]/tbar[0]/btn[11]").press()

# Close the session
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)


# %%
me2n = pd.read_excel(f"C:/Users/{usuario}/Documents/SAP/SAP GUI/ME2N.XLSX", dtype={'Documento compras':'str', 'Centro':'str','Material':'str', 'Posición':'str'})

# %%
me2n = me2n.dropna(subset=['Documento compras'])

# %%
me2n.dtypes

# %%
car_mapping = {
    '1335': "Subaru",
    '1344': "Geely",
    '1305': "DFSK"
 
    # Add more mappings as needed
}

# Apply the mapping to the DataFrame column
me2n['Marca'] = me2n['Centro'].map(car_mapping).fillna("Desconocido")

# %%
import datetime
hoy= datetime.date.today()
hoy = pd.to_datetime(hoy)

# %%
me2n['dias_desde_oc'] = (hoy - me2n['Fecha documento']).dt.days

# %%
me2n

# %%
import pandas as pd

excluir =['100045025',
'100045026',
'100045522',
'100042753']

# Create the dataset
data = {
    'Clase Pedido': ['Z300', 'Z290', 'Z280', 'Z270', 'Z260', 'Z250', 'Z241', 'Z240', 'Z220', 'Z210', 'Z200', 'Z100'],
    'Desc. Clase Pedido': [
        'OC VOR', 'Ped. Inicial Courier', 'Ped. Inicial Aéreo', 'Ped. Inicial Marít.', 'OC Imp. UU.NN. Aéreo',
        'OC Imp. Gtía. Aéreo', 'OC Imp. Rep. Courier', 'OC Imp. Rep. Aérea', 'OC Imp. UU.NN. Marít', 
        'OC Imp. Gtía Marít.', 'OC Imp. Rep. Marít.', 'OC Nacional Repuesto'
    ],
    'Transporte': ['Aereo', 'Aereo', 'Aereo', 'Maritimo', 'Aereo', 'Aereo', 'Aereo', 'Aereo', 'Maritimo', 'Maritimo', 'Maritimo', 'Nacional']
}

# Create DataFrame
clase_pedido = pd.DataFrame(data)

# Display the DataFrame



# %%
me2n['transporte'] = me2n['Documento compras'].apply(lambda x: 'Exportacion' if x in excluir else None)

# Realizar el merge con clase_pedido para los valores no excluidos
me2n = me2n.merge(clase_pedido[['Clase Pedido', 'Transporte']], left_on='Cl.documento compras', right_on='Clase Pedido', how='left')

# Rellenar los valores NaN en la columna 'transporte' con los valores del merge
me2n['transporte'] = me2n.apply(lambda row: row['transporte'] if pd.notna(row['transporte']) else row['Transporte'], axis=1)

# Eliminar la columna extra 'Transporte' creada durante el merge
me2n = me2n.drop(columns=['Transporte'])



# %%
me2n['Documento compras'].drop_duplicates().to_clipboard(index=False, header=False)

# %%
me2n['AUX'] = me2n['Documento compras'] + me2n['Material'] + me2n['Posición']

# %%
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine

# Get the connection and session
connection = application.Children(0)
session = connection.Children(0)

# Maximize the window
session.findById("wnd[0]").maximize()

# Execute the transaction code
session.findById("wnd[0]/tbar[0]/okcd").text = "me80fn"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtSP$00011-LOW").setFocus()
session.findById("wnd[0]/usr/ctxtSP$00011-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_SP$00011_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[16]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]").sendVKey(8)

# Open the context menu and export
session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton("DETAIL_MENU")
session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem("TO_HIST")
session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_HIST/shellcont/shell").pressToolbarButton("&MB_VARIANT")

# Select row and export to Excel
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_HIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_HIST/shellcont/shell").selectContextMenuItem("&XXL")

# Specify export path and filename
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:/Users/{usuario}/Documents/SAP/SAP GUI"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me80fn.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[11]").press()

# Close the windows
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)

# %%
me80fn = pd.read_excel(f"C:/Users/{usuario}/Documents/SAP/SAP GUI/me80fn.XLSX"
                       , dtype={'Material':'str','Documento compras':'str','Posición':'str'}
                       )

# %%
me80fn['AUX'] = me80fn['Documento compras'] + me80fn['Material'] + me80fn['Posición'] 

# %%
me80fn = me80fn.merge(me2n[['AUX','Fecha documento','transporte']], left_on='AUX',right_on='AUX', how='left')

# %%
me80fn['dias_en_factura'] = (me80fn['Fecha de documento'] - me80fn['Fecha documento']).dt.days

# %%
import numpy as np
# Step 1: Add the 'result' column logic
me80fn['result'] = np.where(
    (me80fn['transporte'] == 'Maritimo') & (me80fn['dias_en_factura'] <= 60), 'No',
    np.where((me80fn['transporte'] == 'Aereo') & (me80fn['dias_en_factura'] <= 30), 'No', 'Si')
)


me80fn_dinamica_dias = me80fn[me80fn['Tipo de historial de pedido']=='Q'].groupby('AUX')['dias_en_factura'].mean().reset_index()
# Step 2: Apply the filter and grouping logic
me80fn_dinamica_qty = me80fn[
    (me80fn['Tipo de historial de pedido'] == 'Q') & (me80fn['result'] == 'No')
].groupby('AUX')['Cantidad'].sum().reset_index()

me80fn_dinamica_qty.rename(columns={'Cantidad':'Cantidad facturada'}, inplace=True)


# %%
me2n_cruce = me2n.merge(me80fn_dinamica_dias, left_on='AUX', right_on='AUX', how='left')
me2n_cruce = me2n_cruce.merge(me80fn_dinamica_qty, left_on='AUX', right_on='AUX', how='left')

# %%


# %%

me2n_cruce['fecha_de_entrega'] = me2n_cruce.apply(
    lambda row: row['Fecha documento'] + pd.Timedelta(days=150) if row['Marca'] == "DFSK" and row['transporte'] == "Maritimo"
    else row['Fecha documento'] + pd.Timedelta(days=120) if row['Marca'] in ["Subaru", "Geely"] and row['transporte'] == "Maritimo"
    else row['Fecha documento'] + pd.Timedelta(days=30),
    axis=1
)


# %%
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Assuming the dataframe `me2n_cruce` contains columns 'transporte', 'dias_en_factura', 'fecha_de_entrega', and they are properly formatted

today = pd.Timestamp.today()

# Define conditions
conditions = [
    (me2n_cruce['transporte'].isin(["Nacional", "Exportacion"])),  # OR(transporte="Nacional";transporte="Exportacion")
    ((me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['dias_en_factura'] > 0) & (me2n_cruce['dias_en_factura'] <= 60)),  # AND(transporte="Maritimo";dias_en_factura>0;dias_en_factura<=60)
    ((me2n_cruce['transporte'] == "Aereo") & (me2n_cruce['dias_en_factura'] > 0) & (me2n_cruce['dias_en_factura'] <= 30)),    # AND(transporte="Aereo";dias_en_factura>0;dias_en_factura<=30)
    ((me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['dias_en_factura'] != "") & (me2n_cruce['dias_en_factura'] > 60)),  # AND(transporte="Maritimo";dias_en_factura<>"";dias_en_factura>60)
    ((me2n_cruce['transporte'] == "Aereo") & (me2n_cruce['dias_en_factura'] != "") & (me2n_cruce['dias_en_factura'] > 30)),    # AND(transporte="Aereo";dias_en_factura<>"";dias_en_factura>30)
    ((me2n_cruce['dias_en_factura'].isna()) & (me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['fecha_de_entrega'] > (today + timedelta(days=55)))),  # AND(dias_en_factura="";transporte="Maritimo";fecha_de_entrega>(TODAY()+55))
    ((me2n_cruce['dias_en_factura'].isna()) & (me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['fecha_de_entrega'] < (today + timedelta(days=40)))),  # AND(dias_en_factura="";transporte="Maritimo";fecha_de_entrega<(TODAY()+40))
    ((me2n_cruce['dias_en_factura'].isna()) & (me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['fecha_de_entrega'].between(today + timedelta(days=40), today + timedelta(days=55)))),  # AND(dias_en_factura="";transporte="Maritimo";(TODAY()+40)<=fecha_de_entrega>=(TODAY()+55))
    ((me2n_cruce['dias_en_factura'].isna()) & (me2n_cruce['transporte'] == "Aereo") & ((me2n_cruce['fecha_de_entrega'] - today).dt.days < 14))  # AND(dias_en_factura="";transporte="Aereo";(fecha_de_entrega-TODAY())<14)
]

# Define corresponding results
choices = [
    "No Aplica",  # for Nacional or Exportacion
    "Facturado No Vencido",  # for Maritimo dias_en_factura>0 and dias_en_factura<=60
    "Facturado No Vencido",  # for Aereo dias_en_factura>0 and dias_en_factura<=30
    "Facturado Vencido",     # for Maritimo dias_en_factura>60
    "Facturado Vencido",     # for Aereo dias_en_factura>30
    "Fecha Teórica",         # for Maritimo fecha_de_entrega>(TODAY()+55) and dias_en_factura=""
    "TR Vencido",            # for Maritimo fecha_de_entrega<(TODAY()+40) and dias_en_factura=""
    "Replanificar+45",       # for Maritimo fecha_de_entrega between (TODAY()+40) and (TODAY()+55) and dias_en_factura=""
    "TR Vencido"             # for Aereo (fecha_de_entrega-TODAY())<14 and dias_en_factura=""
]

# Apply the conditions and choices
me2n_cruce['status_v1'] = np.select(conditions, choices, default="Fecha Teórica")
conditions = [
    me2n_cruce['status_v1'].str.contains('Facturado')
]

# Define choices
choices = ['Facturado']

# Apply np.select
me2n_cruce['Status_V2'] = np.select(conditions, choices, default='No Facturado')

# %%
import pandas as pd
import numpy as np

# Assuming the dataframe `me2n_cruce` contains columns 'transporte', 'status_v1', 'dias_desde_oc', 'Ctd.notificada', and 'Cantidad de pedido'

# Define conditions
conditions = [
    (me2n_cruce['transporte'].isin(["Nacional", "Exportacion"])),  # OR(transporte="Nacional";transporte="Exportacion")
    (me2n_cruce['Status_V2'] == "Facturado"),  # status_v1 = "Facturado"
    ((((me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['dias_desde_oc'] > 60)) & (me2n_cruce['Ctd.notificada'] != me2n_cruce['Cantidad de pedido']))|  # AND(transporte="Maritimo";dias_desde_oc>60)
    (((me2n_cruce['transporte'] == "Aereo") & (me2n_cruce['dias_desde_oc'] > 30)) & (me2n_cruce['Ctd.notificada'] != me2n_cruce['Cantidad de pedido']))),  # AND(transporte="Aereo";dias_desde_oc>30)
    ((me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['dias_desde_oc'].between(45, 60))),  # AND(transporte="Maritimo";dias_desde_oc>=45;dias_desde_oc<=60)
    ((me2n_cruce['transporte'] == "Maritimo") & (me2n_cruce['dias_desde_oc'] < 45)),  # AND(transporte="Maritimo";dias_desde_oc<45)
    ((me2n_cruce['transporte'] == "Aereo") & (me2n_cruce['dias_desde_oc'].between(20, 30))),  # AND(transporte="Aereo";dias_desde_oc>=20;dias_desde_oc<=30)
    ((me2n_cruce['transporte'] == "Aereo") & (me2n_cruce['dias_desde_oc'] < 20))  # AND(transporte="Aereo";dias_desde_oc<20)
]

# Define corresponding results
choices = [
    "No Aplica",  # For Nacional or Exportacion
    me2n_cruce['Status_V2'],  # If status_v1 is "Facturado", return status_v1 itself
    "Si",  # For Maritimo dias_desde_oc > 60
     # For Ctd.notificada <> Cantidad de pedido
    "Alerta 45 dias",  # For Maritimo dias_desde_oc between 45 and 60
    "Dentro de los 45 dias",  # For Maritimo dias_desde_oc < 45
    "Alerta 20 dias",  # For Aereo dias_desde_oc between 20 and 30
    "Dentro de los 20 dias"  # For Aereo dias_desde_oc < 20
]

# Apply the conditions and choices
me2n_cruce['fuera_de_estandar'] = np.select(conditions, choices, default="Si")
me2n_cruce['Fecha Reporte'] = hoy

# %%
me2n_cruce.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/HISTORICO_OC.xlsx", index=False)



# %%
df_historico = me2n_cruce

# %%


# %%
df_historico['Prefijo OC'] = ''
df_historico['Cód. Mat de prov en OC'] = ''
df_historico['Vía (Texto)'] = ''
df_historico['Nro. DT'] = ''
df_historico['Pto. Origen'] = ''
df_historico['Pto. Destino'] = ''
df_historico['Fe. ATA'] = ''
df_historico['Nro. SOLPED'] = ''
df_historico['Nivel de Urgencia'] = ''

df_historico['Semana'] = df_historico['fecha_de_entrega'].dt.isocalendar().week
df_historico['Mes'] = df_historico['fecha_de_entrega'].dt.month
df_historico['Año'] = df_historico['fecha_de_entrega'].dt.year
df_historico['Familia'] = ''
df_historico['Costo UN CLP'] = ''
df_historico['Segmentacion'] = ''
df_historico['Plan mantención'] = ''

df_historico['Fechas corregidas R3'] = ''



# %%
df_historico = df_historico[['AUX','status_v1','Material','Texto breve','fecha_de_entrega','Proveedor/Centro suministrador','Cl.documento compras','Marca','Status_V2','Prefijo OC','Documento compras','Posición','Cód. Mat de prov en OC','Vía (Texto)','Pto. Origen','Pto. Destino','Nro. DT','Fecha documento','Fe. ATA','Nro. SOLPED','Cantidad de pedido','Nivel de Urgencia','dias_desde_oc','fuera_de_estandar','Semana','Mes', 'Año','Familia','Costo UN CLP', 'Segmentacion', 'Plan mantención',  'Fechas corregidas R3','Por calcular (cantidad)','Fecha Reporte']]

# %%
df_historico.rename(columns={ 'Por calcular (cantidad)':'Cantidad','Proveedor/Centro suministrador':'Nombre del proveedor','Marca':'NomSector_actual','Documento compras':'Nro. OC SAP','Posición':'Posición OC SAP','Cantidad de pedido':'Cant. UN. OC','dias_desde_oc':'Hoy - OC (dias)','Fecha documento':'Fecha Creación OC','status_v1':'Status','fecha_de_entrega':'Fecha','Status_V2':'Status_V2','fuera_de_estandar':'Fuera de Estandar','Material':'Cod_Actual_1'}, inplace=True)
# %%
# df_oc.drop(columns=['Unnamed: 0'], inplace=True)

# %%

df_final = pd.concat([df_oc, df_historico], axis=0, ignore_index=True)
df_final['AUX'] = df_final['AUX'].astype('str')
df_final.to_excel(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/Consolidado OC Prod Homologado/prueba_consolidado.xlsx')
# %%
df_bo = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/bases python/Seguimiento Producción OEM.xlsm", sheet_name='Seguimiento')

# %%

df_bo = df_bo[['AUX','Input','Fecha de modificación']]
df_bo['AUX'] = df_bo['AUX'].astype('str')
# %%
df_bo.drop_duplicates(subset='AUX', inplace=True)

# %%
df_final = df_final.merge(df_bo, on='AUX', how='left')

columns = [
    "AUX", "Status", "Texto breve", "Cantidad", "Fecha", 
    "Nombre del proveedor", "NomSector_actual", "Cl.documento compras", 
    "Status_V2", "Prefijo OC", "Nro. OC SAP", "Posición OC SAP", 
    "Cód. Mat de prov en OC", "Vía (Texto)", "Pto. Origen", "Pto. Destino", 
    "Nro. DT", "Fecha Creación OC", "Fe. ATA", "Nro. SOLPED", "Cant. UN. OC", 
    "Nivel de Urgencia", "Hoy - OC (dias)", "Fuera de Estandar", "Semana", 
    "Mes", "Año", "Familia", "Costo UN CLP", "Segmentacion", 
    "Plan mantención",  "Cod_Actual_1","Fechas corregidas R3", "Fecha Reporte", "Input", 
    "Fecha de modificación"
]

df_final = df_final[columns]

# %%
df_final.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/OC-Producción OEM/bases python/consolidado BO.xlsx")

# %%


