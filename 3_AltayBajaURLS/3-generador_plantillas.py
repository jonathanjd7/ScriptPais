import csv
import os
from docx import Document

def corregir_caracteres(texto):
    if isinstance(texto, str):
        texto = texto.encode('iso-8859-1', 'ignore').decode('iso-8859-1')
        texto = texto.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
        texto = texto.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        texto = texto.replace('ñ', 'n').replace('Ñ', 'N')
    return texto

def crear_plantillas_word(ruta_carpeta_csv, ruta_plantilla):
  """
  Crea plantillas de Word personalizadas para cada código postal en varios archivos CSV.

  Args:
    ruta_carpeta_csv: La ruta de la carpeta que contiene los archivos CSV.
    ruta_plantilla: La ruta de la plantilla de Word.
  """

  # Itera sobre todos los archivos en la carpeta CSV.
  for nombre_archivo in os.listdir(ruta_carpeta_csv):
    # Si el archivo es un archivo CSV.
    if nombre_archivo.endswith(".csv"):
      # Construye la ruta completa del archivo CSV.
      ruta_archivo_csv = os.path.join(ruta_carpeta_csv, nombre_archivo)

      with open(ruta_archivo_csv, 'r', encoding='iso-8859-1') as archivo_csv:
        lector_csv = csv.reader(archivo_csv)
        next(lector_csv)  # Saltar la primera fila si es un encabezado

        for fila in lector_csv:
          codigo_postal = fila[0]
          ccaa = fila[1]
          provincia = fila[2]
          ciudad = fila[3]
          proveedor = fila[4]

          documento = Document(ruta_plantilla)

          # Reemplazar textos en la plantilla
          for parrafo in documento.paragraphs:
            if '{codigopostal}' in parrafo.text:
              parrafo.text = parrafo.text.replace('{codigopostal}', codigo_postal)
            if '{ciudad}' in parrafo.text:
              parrafo.text = parrafo.text.replace('{ciudad}', ciudad)
            if '{provincia}' in parrafo.text:
              parrafo.text = parrafo.text.replace('{provincia}', provincia)
            if '{proveedor}' in parrafo.text:
              parrafo.text = parrafo.text.replace('{proveedor}', proveedor)
            if '{CCAA}' in parrafo.text:
              parrafo.text = parrafo.text.replace('{CCAA}', ccaa)

          # Guardar la nueva plantilla
          nombre_archivo = f'{codigo_postal}_{corregir_caracteres(ciudad)}_plantilla.docx'
          documento.save('./plantillas_generadas/' + nombre_archivo)
          print("Plantilla word para el codigo postal " + codigo_postal + " y la ciudad " + ciudad + " creada exitosamente.")

# Ejemplo de uso
ruta_csv = './csvs_codigos_postales_ciudades'
ruta_plantilla = './plantilla_para_generar.docx'
try:
    crear_plantillas_word(ruta_csv, ruta_plantilla)
except Exception as e:
    print(f"Error al generar el archivo: {e}")