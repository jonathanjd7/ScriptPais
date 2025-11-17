import pandas as pd
import glob
import os

# Define la ruta de la subcarpeta
subfolder_path = 'fichas_a_generar'

# Obtiene una lista de todos los archivos Excel en la subcarpeta
excel_files = glob.glob(os.path.join(subfolder_path, '*.xlsx'))

# Crea una lista vacía para almacenar los datos de todos los archivos
all_data = []

# Itera sobre cada archivo Excel
for file in excel_files:
    # Lee el archivo Excel
    excel_file = pd.ExcelFile(file)

     # Itera sobre cada hoja del archivo Excel
    for sheet_name in excel_file.sheet_names:
        df = excel_file.parse(sheet_name)

         # Selecciona las columnas que te interesan
        df = df[['Codigo postal', 'CCAA', 'PROVINCIA', 'LOCALIDAD', 'Proveedor', 'Links']]

         # Añade los datos de la hoja actual a la lista
        all_data.append(df)

# Concatena todos los datos en un único DataFrame
final_df = pd.concat(all_data)

# Función para corregir caracteres extraños
# def corregir_caracteres(texto):
#     if isinstance(texto, str):
#         texto = texto.encode('utf-8', 'ignore').decode('utf-8')
#         texto = texto.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
#         texto = texto.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
#         texto = texto.replace('ñ', 'n').replace('Ñ', 'N')
#     return texto

# Aplica la función a todas las columnas del DataFrame
# for column in final_df.columns:
#     final_df[column] = final_df[column].apply(corregir_caracteres)

# Guarda los datos en un nuevo archivo CSV
final_df.to_csv('./csvs_codigos_postales_ciudades/csv_codigos_postales_localidades.csv', index=False, encoding='iso-8859-1')
print("Plantilla CSV creada exitosamente.")
    
