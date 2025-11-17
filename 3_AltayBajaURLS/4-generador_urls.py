import pandas as pd
import docx
import csv
import os

# Lee el archivo Excel
excel_file = pd.read_excel('./fichas_a_generar/ficha_a_generar.xlsx')

# Crea un diccionario para almacenar los datos
data = {}

# Itera sobre las filas del archivo Excel
for index, row in excel_file.iterrows():
    # Obtiene el código postal
    codigo_postal = row['Codigo postal']

    # Obtiene la URL
    url = row['Links']

    # Busca el nombre de la plantilla en el archivo Word
    for filename in os.listdir('./plantillas_generadas'):
        if filename.endswith('.docx') and str(codigo_postal) in filename:
            nombre_plantilla = filename
            data[nombre_plantilla] = url
            break

# Escribe los datos en un archivo CSV
with open('./plantillas_generadas/urls.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    #writer.writerow(['nombre_plantilla', 'url'])
    for nombre_plantilla, url in data.items():
        writer.writerow([url])
print("Plantilla CSV creada exitosamente.")