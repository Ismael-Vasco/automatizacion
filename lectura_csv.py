# pip install pandas zipfile36
#D:\\IVASCOG\\Desktop\\python\\aplicacion_powerBi\\datos\\Reporte DepartamentoANTIOQUIA20241007_014406.csv
import pandas as pd
import openpyxl
import os as os

# leer todos los documentos de la carpeta y almacenarlos en un diccionario
def cargar_documentos(ruta):
    documentos_csv = {}
    documenos_xlsx = {}
    
    # Recorre todos los archivos en la carpeta
    for nombre_archivo in os.listdir(ruta):
        if nombre_archivo.endswith('.csv'):
            documentos_csv[nombre_archivo] = pd.read_csv(os.path.join(ruta, nombre_archivo))
        
        if nombre_archivo.endswith('xlsx') or nombre_archivo.endswith('xlx'):
            documenos_xlsx[nombre_archivo] = pd.read_excel(os.path.join(ruta, nombre_archivo))

            # iteración de hojas
            df = openpyxl.load_workbook(os.path.join(ruta, nombre_archivo))
            names = df.sheetnames
            for hoja in names:
                if not hoja.endswith('.xlsx'):
                    documenos_xlsx[hoja] = pd.DataFrame(df[hoja].values)

    return documentos_csv, documenos_xlsx

# eliminar archivos extras
def eliminar_archivos_extra(diccionario: dict):
    for key, value in diccionario.items():
        if key.endswith('.xlsx'):
            eliminado = {}
            eliminado[key] = diccionario[key]
    
    for llave, value in eliminado.items():
        del diccionario[llave]

    return diccionario

def definir_columnas(diccionario: dict):
    for k,v in diccionario.items():
        dataset = diccionario[k]
        # tomar la el primer registro para colocarlo como columna
        columnas = list(dataset.iloc[0])
        # Transponer el DataFrame y luego restablecer los índices
        dataset.columns= columnas[0:] # Establecer la primera fila como nombres de columnas
        # Eliminar la primera fila
        dataset = dataset[1:]
        # Restablecer el índice
        dataset.reset_index(drop=True, inplace=True)

        diccionario[k] = dataset


    return diccionario

# def convertir_a_csv(diccionario: dict):
#     # Crear una nueva carpeta

    
def main():
    # leer todos los documentos de la carpeta y almacenarlos en un diccionario
    #print('presiona la tecla 1 para limpiar carga rlos datos')
    ruta = 'D:\\IVASCOG\\Desktop\\python\\aplicacion_powerBi\\datos'
    # extraer los datos de la carpeta
    documentos_csv, documenos_xlsx = cargar_documentos(ruta)
    # eliminar archivo extra
    documenos_xlsx = eliminar_archivos_extra(documenos_xlsx)
    print('Listo, archivos cargados')

    # limpiar todos los documentos de acuerdo a su necesidad
    documenos_xlsx = definir_columnas(documenos_xlsx)
    print('listo la liempieza')

# convertir todos los documentos en .csv
   
    carpeta_salida = 'archivos_csv'
    os.makedirs(carpeta_salida, exist_ok=True)  # Crea la carpeta si no existe

    # Convertir y guardar los DataFrames en la nueva carpeta
    lista_dicct = [documentos_csv,documenos_xlsx]

    for i in lista_dicct:
        for nombre_tabla, df in i.items():
            df.to_csv(os.path.join(carpeta_salida, f'{nombre_tabla}.csv'), index=False)
    

    # subir los archivos a Power BI

    # crar la automatización semanal por medio de carpetas con nombres especificos


main()