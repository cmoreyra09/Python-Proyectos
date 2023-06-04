import pandas as pd
from tabulate import tabulate
import os


## Función Leer Archivo
def leerarchivo(filepath, numero):
    data = pd.read_excel(filepath, sheet_name=int(numero))
    return data

def columna(data):
    x = data.columns
    return x

def diccionario(name, columna):
    data = pd.DataFrame({name: columna})
    return data

def lista(data):
    data = data.values.tolist()
    return data

## Variable dinámica para la ruta del archivo correspondiente
## Esto se cambia de acuerdo al archivo

ruta = r"C:\Users\User\OneDrive\Escritorio\Atento\datadelarchivo.xlsx"


# Obtener nombres y dimensiones de las hojas
with pd.ExcelFile(ruta) as xls:
    nombres_hojas = xls.sheet_names ## Sacas los nombres de las hojas.
    num_hojas = len(nombres_hojas) ## Sacas la cantidad de hojas del Excel.

# Declaración de hojas y columnas

dfs = [] ## Aqui estara la data cruda es decir sin columnas o cabezeras
df_columnas = [] ## Aqui estara las cabezeras

for i in range(num_hojas): ## Interamos por la cantidad de numero de hojas
    df = leerarchivo(ruta, i) ## Leyendo el Archivo Excel
    dfs.append(df) ## Agregames a la lista dfs
    df_columnas.append(columna(df)) ## Agregamos las cabeceras a la lista df_columnas

# Declaración de diccionarios finales
df_finales = [] ## Lista de diccionarios Finales
for i in range(num_hojas): ## Se crea por la cantidad de hojas que existe en el excel
    df_final = diccionario(nombres_hojas[i], df_columnas[i]) ## Crea el dataframe de nombres y columnas
    df_finales.append(df_final) ##  Se agrega todo a la variable final

# Formato lista

## Para el printeo del formato tabulate
listas = []
for df_final in df_finales:
    lista_actual = lista(df_final)
    listas.append(lista_actual)

# Ruta
ruta_salida = r'C:\Users\User\OneDrive\Escritorio\Atento'

## Para obtener el nombre del archivo
nombre_data = os.path.splitext(os.path.basename(ruta))[0]

# Para el archivo en TXT
archivo_salida = ruta_salida + '\\'+'Diccionario_'+nombre_data+'.txt'

# Para el arhcivo en EXCEL
archivo_salida_excel = ruta_salida + '\\'+'Diccionario_'+nombre_data+'.xlsx'

for i in range (1,2):
    if i ==1:
        try:
            ## Para TXT
            with open(archivo_salida, 'w') as archivo:
                archivo.write('                                    ----------------------------------------------------\n')
                archivo.write(f'                                      Diccionario de datos Reporte de {nombre_data} \n')
                archivo.write('                                    ----------------------------------------------------\n')
                archivo.write('\n')
                archivo.write('-----------------------------------------------------------------------------------\n')
                archivo.write('\n')

                for i in range(num_hojas):
                    archivo.write(nombres_hojas[i])
                    archivo.write('\n')
                    archivo.write('\n')
                    archivo.write(tabulate(listas[i], headers=df_finales[i].columns, tablefmt='psql'))
                    archivo.write('\n')
                    archivo.write('\n')
                    archivo.write('-----------------------------------------------------------------------------------\n')

            ## Para Exel
            with pd.ExcelWriter(archivo_salida_excel) as writer:
                for i in range(num_hojas):
                    df_finales[i].to_excel(writer, sheet_name=nombres_hojas[i], index=False)

            ## Si esta correcto
            print('Archivo Creado Correctamente')


        ## Si no esta correcto muestra este resultado
        except Exception as e:
            print('Surgió un Error en la salida del archivo de los diccionarios')
