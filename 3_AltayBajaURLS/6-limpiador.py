import os

def borrar_archivos_carpeta(ruta_carpeta):
  """
  Borra todos los archivos de una carpeta específica.

  Args:
    ruta_carpeta: La ruta de la carpeta donde se encuentran los archivos.
  """

  # Itera sobre todos los archivos en la carpeta.
  for nombre_archivo in os.listdir(ruta_carpeta):
    # Construye la ruta completa del archivo.
    ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)

    # Verifica si el elemento es un archivo y no una subcarpeta.
    if os.path.isfile(ruta_archivo):
      # Borra el archivo.
      os.remove(ruta_archivo)

# Ejemplo de uso
ruta_carpeta = "./csvs_codigos_postales_ciudades"
ruta_carpeta2 = "./plantillas_generadas"
#ruta_carpeta3 = "./csv_urls"

borrar_archivos_carpeta(ruta_carpeta)
borrar_archivos_carpeta(ruta_carpeta2)
#borrar_archivos_carpeta(ruta_carpeta3)
print("Archivos limpiados exitosamente.")