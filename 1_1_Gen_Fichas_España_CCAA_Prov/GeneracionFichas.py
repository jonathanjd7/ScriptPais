#CODIGO REALIAZADO POR JONATHAN JD , CON LA AYUDA DE CURSOR
#https://github.com/jonathanjd7

import os
import pandas as pd
from docxtpl import DocxTemplate
import datetime
import locale
import re
import random

# ============================
# 1. CONFIGURACIÓN INICIAL
# ============================

# Configurar la localización para fechas en español (portable)
try:
    # Intentar configurar locale en español (funciona en Linux/Mac)
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        # Intentar alternativa para Windows
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except locale.Error:
        try:
            # Intentar otra alternativa común en Windows
            locale.setlocale(locale.LC_TIME, 'es_ES')
        except locale.Error:
            # Si falla todo, usar locale por defecto del sistema
            print("Advertencia: No se pudo configurar el locale en español. Se usará el locale por defecto del sistema.")

# ============================
# ⚙️ CONTROL DE PRUEBA - EDITA AQUÍ
# ============================
# Cambia NUM_FICHAS_PRUEBA para controlar cuántas fichas generar:
#   None o 0  --> Genera TODAS las fichas del archivo
#   10        --> Genera solo las primeras 10 fichas (PRUEBA)
#   50        --> Genera solo las primeras 50 fichas
NUM_FICHAS_PRUEBA = 0  # <--- CAMBIA ESTE NÚMERO AQUÍ
# ============================

# Variables globales
SECTOR = 'Supermercados y Tiendas de Alimentación' \
''#Verificar que 'SECTOR' coincida con el nombre del sector dentro de 'site_maps_2025_02_12.xlsx'
MES = datetime.datetime.now().strftime("%B")
# BASE_DIR ahora es portable: usa la carpeta donde está el script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_SITEMAPS = os.path.join(BASE_DIR, 'site_maps_2025_02_12.xlsx')
OUTPUT_DIRECTORY_CSV = os.path.join(BASE_DIR, 'csv_creados')
ARCHIVO_EXCEL_GENERADOR = os.path.join(BASE_DIR, 'Generador_csv_ccaa_prov.xlsx')
HOJA1 = 'Comunidades_csv_copiar'
HOJA2 = 'Provincias_csv_copiar'
PLANTILLA_GENERICA = os.path.join(BASE_DIR, 'Plantilla_Generica.docx')
PLANTILLAS_CREADAS_DIR = os.path.join(BASE_DIR, 'Plantillas_creadas')

# Crear directorios necesarios
os.makedirs(OUTPUT_DIRECTORY_CSV, exist_ok=True)
os.makedirs(PLANTILLAS_CREADAS_DIR, exist_ok=True)

# Validación de rutas (BASE_DIR siempre existe porque es la carpeta del script)
if not os.path.exists(EXCEL_FILE_SITEMAPS):
    raise FileNotFoundError(f"El archivo {EXCEL_FILE_SITEMAPS} no existe.")

if not os.path.exists(OUTPUT_DIRECTORY_CSV):
    os.makedirs(OUTPUT_DIRECTORY_CSV)  # Crea el directorio si no existe

if not os.path.exists(ARCHIVO_EXCEL_GENERADOR):
    raise FileNotFoundError(f"El archivo {ARCHIVO_EXCEL_GENERADOR} no existe.")

if not os.path.exists(PLANTILLA_GENERICA):
    raise FileNotFoundError(f"La plantilla genérica {PLANTILLA_GENERICA} no existe.")

if not os.path.exists(PLANTILLAS_CREADAS_DIR):
    os.makedirs(PLANTILLAS_CREADAS_DIR)  # Crea el directorio si no existe

print("Configuración cargada correctamente.")


# ============================
# 2. FUNCIONES
# ============================

# Funciones de utilidad

def eliminar_acentos_slash(text):
    replacements = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ä': 'a', 'ã': 'a', 'å': 'a',
        'Á': 'A', 'À': 'A', 'Â': 'A', 'Ä': 'A', 'Ã': 'A', 'Å': 'A',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'É': 'E', 'È': 'E', 'Ê': 'E', 'Ë': 'E',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'Í': 'I', 'Ì': 'I', 'Î': 'I', 'Ï': 'I',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'ö': 'o', 'õ': 'o', 'ø': 'o',
        'Ó': 'O', 'Ò': 'O', 'Ô': 'O', 'Ö': 'O', 'Õ': 'O', 'Ø': 'O',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
        'Ú': 'U', 'Ù': 'U', 'Û': 'U', 'Ü': 'U',
        'ý': 'y', 'ÿ': 'y', 'Ý': 'Y', 'Ÿ': 'Y',
        'ñ': 'n', 'Ñ': 'n', '/': '_', ' ': '_', '-': '_'
    }
    text = re.sub(r'[áàâäãåéèêëíìîïóòôöõøúùûüýÿñÁÀÂÄÃÅÉÈÊËÍÌÎÏÓÒÔÖÕØÚÙÛÜÝŸÑ/ -]', lambda x: replacements[x.group()], text)
    return re.sub(r'_+', '_', text)

def agregar_caracteres_especiales(text, direcciones_con_acentos, direcciones_con_acentos_correctos):
    if text in direcciones_con_acentos:
        index = direcciones_con_acentos.index(text)
        return direcciones_con_acentos_correctos[index]
    return text

def get_random_urls(EXCEL_FILE_SITEMAPS, num_urls=10):
    df = pd.read_excel(EXCEL_FILE_SITEMAPS)
    blog_urls = df[df['Categoria'] == 'blog']['URL'].tolist()
    if len(blog_urls) < num_urls:
        raise ValueError(f"No hay suficientes URLs de blog. Solo hay {len(blog_urls)} disponibles")
    urls_seleccionadas = random.sample(blog_urls, num_urls)
    return {f'URL_{i+1}': url for i, url in enumerate(urls_seleccionadas)}

def generar_plantilla(context, EXCEL_FILE_SITEMAPS):
    urls_diccionario = get_random_urls(EXCEL_FILE_SITEMAPS, num_urls=10)
    context.update({
        'Localizacion': '{{Localizacion}}',
        'Precio': '{{Precio}}',
        'registros': '{{registros}}',
        'dir': '{{dir}}',
        'phone': '{{phone}}',
        'mail': '{{mail}}',
        'cat': '{{cat}}',
        'web': '{{web}}',
        'Mes' : MES, 
        # Solo modificar estos campos
        'Sector': SECTOR,
        'RegionTipo': context['RegionTipo'],
        'URL_1': urls_diccionario['URL_1'],
        'URL_2':urls_diccionario['URL_2'],
        'URL_3':urls_diccionario['URL_3'],
        'URL_4':urls_diccionario['URL_4'],
        'URL_5':urls_diccionario['URL_5'],
        'URL_6':urls_diccionario['URL_6'],
        'URL_7':urls_diccionario['URL_7'],
        'URL_8':urls_diccionario['URL_8'],
        'URL_9':urls_diccionario['URL_9'],
        'URL_10':urls_diccionario['URL_10']

    })
    doc_generico = DocxTemplate(PLANTILLA_GENERICA)
    tipo = context['RegionTipo'].lower()
    tipo_sin_acentos = eliminar_acentos_slash(tipo)
    nombre_archivo = f'Plantilla_{tipo_sin_acentos}.docx'
    sector_directory = os.path.join(PLANTILLAS_CREADAS_DIR, eliminar_acentos_slash(SECTOR), f'Plantillas_{eliminar_acentos_slash(SECTOR)}')
    os.makedirs(sector_directory, exist_ok=True)
    ruta_archivo = os.path.join(sector_directory, nombre_archivo)
    doc_generico.render(context)
    doc_generico.save(ruta_archivo)
    print(f'Plantilla "{ruta_archivo}" generada con éxito.')
    return DocxTemplate(ruta_archivo)

def generar_documento(fila, context, doc_plantilla, nombre_archivo, sector_directory):
    doc_plantilla.render(context)
    ruta_archivo = os.path.join(sector_directory, nombre_archivo)
    doc_plantilla.save(ruta_archivo)
    print(f'Documento "{ruta_archivo}" generado con éxito.')

def extraer_nombre_doc(nombrefichero):
    nombrefichero = nombrefichero.replace('.docx', '')
    # Eliminar el número al inicio si existe (formato: "numero_nombre")
    partes_nombre = nombrefichero.split('_', 1)
    if len(partes_nombre) > 1 and partes_nombre[0].isdigit():
        nombrefichero = partes_nombre[1]  # Usar solo la parte después del número
    
    if "_Pais_" in nombrefichero and "Vasco" not in nombrefichero:
        partes = nombrefichero.split("_Pais_")
        sector = partes[0]
        region_tipo = "Pais"
        region = partes[1]
    elif "_Comunidad_Autonoma_" in nombrefichero:
        partes = nombrefichero.split("_Comunidad_Autonoma_")
        sector = partes[0]
        region_tipo = "Comunidad_Autonoma"
        region = partes[1]
    elif "_Ciudad_Autonoma_" in nombrefichero:
        partes = nombrefichero.split("_Ciudad_Autonoma_")
        sector = partes[0]
        region_tipo = "Ciudad_Autonoma"
        region = partes[1]
    elif "_Provincia_" in nombrefichero:
        partes = nombrefichero.split("_Provincia_")
        sector = partes[0]
        region_tipo = "Provincia"
        region = partes[1]
    else:
        return None, None, None
    return sector, region_tipo, region
def validar_contexto(context):
       campos_requeridos = ['Localizacion', 'Precio', 'registros', 'dir', 'phone', 'mail', 'cat', 'web', 'Mes', 'Sector', 'RegionTipo']
       for campo in campos_requeridos:
           if campo not in context or context[campo] == '':
               print(f"Advertencia: El campo '{campo}' está vacío o no definido.")
               context[campo] = 'N/A'  # Valor por defecto
       return context  


def generador_csv_con_url(sector_directory_seleccionar, EXCEL_FILE_SITEMAPS):
    try:
        sitemap_df = pd.read_excel(EXCEL_FILE_SITEMAPS)
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return
    columnas = ['Sector', 'España', 'Comunidad_Autonoma', 'Ciudad_Autonoma', 'Provincia', 'Localidad']
    for col in columnas:
        sitemap_df[col] = sitemap_df[col].fillna("").apply(eliminar_acentos_slash)

    ficheros_word = [f for f in os.listdir(sector_directory_seleccionar) if f.endswith('.docx')]
    
    # Ordenar archivos por el número al inicio del nombre (si existe)
    def obtener_numero_archivo(nombre_archivo):
        """Extrae el número del inicio del nombre del archivo para ordenar"""
        # Formato esperado: "numero_nombre.docx" o "nombre.docx"
        partes = nombre_archivo.split('_', 1)
        if partes[0].isdigit():
            return int(partes[0])
        return 999999  # Si no tiene número, ponerlo al final
    
    ficheros_word.sort(key=obtener_numero_archivo)

    combinado_data = []

    for file in ficheros_word:
        sector, region_tipo, region = extraer_nombre_doc(file)
        if sector and region_tipo and region:
            filtrado_fila = None
            if region_tipo == "Comunidad_Autonoma":
                filtrado_fila = sitemap_df[
                    (sitemap_df['Sector'] == sector) &
                    (sitemap_df['Comunidad_Autonoma'] == region)
                ]
            elif region_tipo == "Ciudad_Autonoma":
                filtrado_fila = sitemap_df[
                    (sitemap_df['Sector'] == sector) &
                    (sitemap_df['Ciudad_Autonoma'] == region)
                ]
            elif region_tipo == "Provincia":
                filtrado_fila = sitemap_df[
                    (sitemap_df['Sector'] == sector) &
                    (sitemap_df['Provincia'] == region)
                ]
            elif region_tipo == "Pais":
                filtrado_fila = sitemap_df[
                    (sitemap_df['Sector'] == sector) &
                    (sitemap_df['España'] == region)
                ]

            if filtrado_fila is not None and not filtrado_fila.empty:
                url = filtrado_fila.iloc[0]['URL']
                combinado_data.append([url])
            else:
                print(f"No se encontró URL para el archivo: {file}")

    output_file_url = os.path.join(sector_directory_seleccionar, "urls.csv")
    combinado_df = pd.DataFrame(combinado_data, columns=['URL'])
    combinado_df.to_csv(output_file_url, index=False, header=False, encoding='utf-8')
    print(f"Archivo CSV combinado generado: {output_file_url}")

# Leer datos desde archivos Excel
df_pais_comunidades_ciudades = pd.read_excel(ARCHIVO_EXCEL_GENERADOR, sheet_name=HOJA1)
df_provincias = pd.read_excel(ARCHIVO_EXCEL_GENERADOR, sheet_name=HOJA2)

# Seleccionar datos relevantes
df_seleccionado_pais_comunidades_ciudades = df_pais_comunidades_ciudades.iloc[:21, 1:9]
df_seleccionado_provincias = df_provincias.iloc[:51, 1:9]

# Aplicar límite de prueba si está configurado
total_pais_comunidades_ciudades = len(df_seleccionado_pais_comunidades_ciudades)
total_provincias = len(df_seleccionado_provincias)
total_fichas = total_pais_comunidades_ciudades + total_provincias

if NUM_FICHAS_PRUEBA and NUM_FICHAS_PRUEBA > 0:
    # Aplicar límite proporcionalmente a ambos DataFrames
    # Primero limitar el de paises/comunidades/ciudades
    if NUM_FICHAS_PRUEBA <= total_pais_comunidades_ciudades:
        df_seleccionado_pais_comunidades_ciudades = df_seleccionado_pais_comunidades_ciudades.head(NUM_FICHAS_PRUEBA)
        df_seleccionado_provincias = df_provincias.iloc[:0, 1:9]  # Vaciar provincias
    else:
        # Si el límite es mayor, incluir todas las paises/comunidades/ciudades y parte de provincias
        fichas_restantes = NUM_FICHAS_PRUEBA - total_pais_comunidades_ciudades
        df_seleccionado_provincias = df_seleccionado_provincias.head(fichas_restantes)
    
    print(f"\n[MODO PRUEBA] Generando solo las primeras {NUM_FICHAS_PRUEBA} fichas")
    print(f"   - Países/Comunidades/Ciudades: {len(df_seleccionado_pais_comunidades_ciudades)}")
    print(f"   - Provincias: {len(df_seleccionado_provincias)}")
    print(f"[CONSEJO] Para generar todas las fichas ({total_fichas}), cambia NUM_FICHAS_PRUEBA = None en la línea 38\n")
else:
    print(f"\n[MODO COMPLETO] Generando TODAS las fichas")
    print(f"   - Países/Comunidades/Ciudades: {total_pais_comunidades_ciudades}")
    print(f"   - Provincias: {total_provincias}")
    print(f"   - Total: {total_fichas}\n")

# Guardar datos seleccionados en archivos CSV
df_seleccionado_pais_comunidades_ciudades.to_csv(os.path.join(OUTPUT_DIRECTORY_CSV, 'Generador_csv_ccaa.csv'), index=False)
df_seleccionado_provincias.to_csv(os.path.join(OUTPUT_DIRECTORY_CSV, 'Generador_csv_prov.csv'), index=False)

print("Los archivos CSV se han creado correctamente.")

# Contextos para las plantillas
my_context_pais = {'Mes': MES, 'Sector': SECTOR, 'RegionTipo': 'País'}
my_context_comunidad = {'Mes': MES, 'Sector': SECTOR, 'RegionTipo': 'Comunidad Autónoma'}
my_context_ciudad = {'Mes': MES, 'Sector': SECTOR, 'RegionTipo': 'Ciudad Autónoma'}
my_context_provincia = {'Mes': MES, 'Sector': SECTOR, 'RegionTipo': 'Provincia'}

# Generar las plantillas específicas
context_tipos = [
    my_context_pais,
    my_context_comunidad,
    my_context_provincia,
    my_context_ciudad
]

doc_pais = generar_plantilla(context_tipos[0], EXCEL_FILE_SITEMAPS)
doc_comunidad = generar_plantilla(context_tipos[1], EXCEL_FILE_SITEMAPS)
doc_provincia = generar_plantilla(context_tipos[2], EXCEL_FILE_SITEMAPS)
doc_ciudad = generar_plantilla(context_tipos[3], EXCEL_FILE_SITEMAPS)

# Definir listas de nombres para paises, comunidades, ciudades y provincias
internacional = [
        'Francia', 'Turquia', 'UK', 'Italia', 'Portugal', 'Irlanda', 'Alemania', 'Mexico', 'Belgica', 'Hungría', 
        'Polonia', 'USA', 'Andorra', 'Chile', 'Colombia', 'Suecia', 'Brasil', 'Marruecos', 'Ecuador', 'Albania', 
        'Argentina', 'Armenia', 'Austria', 'Azerbaiyan', 'Belice', 'Bielorrusia', 'Bolivia', 'Bosnia', 'Bulgaria', 
        'Chipre', 'Costa Rica', 'Croacia', 'Cuba', 'Dinamarca', 'El salvador', 'Eslovaquia', 'Eslovenia', 'Estonia', 
        'Finlandia', 'Georgia', 'Gibraltar', 'Grecia', 'Guatemala', 'Guyana', 'Haiti', 'Holanda', 'Honduras', 
        'Hungria', 'Islandia', 'Islas Caiman', 'Kazajistan', 'Letonia', 'Lituania', 'Luxemburgo', 'Macedonia', 
        'Malta', 'Moldavia', 'Monaco', 'Montenegro', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Puerto Rico', 
        'República Checa', 'república Dominicana', 'Rumania', 'Rusia', 'San Marino', 'Serbia', 'Suiza', 'Uruguay', 
        'Uzbekistan', 'Venezuela', 'Republica checa', 'Republica Dominicana']
pais = ['España']

comunidades = ['Andalucía', 'Aragón', 'Asturias', 'Canarias', 'Cantabria', 'Castilla y La Mancha', 'Castilla y León', 
        'Cataluña', 'Comunidad Valenciana', 'Extremadura', 'Galicia', 'Islas Baleares - Illes Balears', 'La Rioja', 
        'Madrid', 'Murcia', 'Navarra', 'País Vasco']
ciudades = ['Ceuta', 'Melilla']

provincias = ['Almería', 'Cádiz', 'Córdoba', 'Granada', 'Huelva', 'Jaén', 'Málaga', 'Sevilla', 'Huesca', 'Teruel', 
        'Zaragoza', 'Asturias','Las Palmas', 'Tenerife','Cantabria', 'Albacete', 'Ciudad Real', 'Cuenca', 'Guadalajara', 'Toledo', 'Ávila', 
        'Burgos', 'León', 'Palencia', 'Salamanca', 'Segovia', 'Soria', 'Valladolid', 'Zamora', 'Barcelona', 
        'Gerona - Girona', 'Lérida - Lleida', 'Tarragona', 'Alicante - Alacant', 'Castellón', 'Valencia', 'Badajoz', 
        'Cáceres', 'A Coruña - La Coruña', 'Lugo', 'Ourense - Orense', 'Pontevedra','Islas Baleares - Illes Balears', 'La Rioja', 'Madrid', 'Murcia', 
        'Navarra', 'Álava - Araba', 'Guipúzcoa - Gipuzkoa', 'Vizcaya - Bizkaia']

# Generar documentos para paises, comunidades y ciudades
direcciones_con_acentos = ['Andalucia', 'Aragon', 'Asturias', 'Canarias', 'Cantabria', 'Castilla y La Mancha', 'Castilla y Leon', 
        'Cataluña', 'Comunidad Valenciana', 'Extremadura', 'Galicia', 'Islas Baleares - Illes Balears', 'La Rioja', 
        'Madrid', 'Murcia', 'Navarra', 'Pais Vasco', 'Almeria', 'Cadiz', 'Cordoba', 'Granada', 'Huelva', 'Jaen', 
        'Malaga', 'Sevilla', 'Huesca', 'Teruel', 'Zaragoza', 'Las Palmas', 'Tenerife', 'Albacete', 'Ciudad Real', 
        'Cuenca', 'Guadalajara', 'Toledo', 'Avila', 'Burgos', 'Leon', 'Palencia', 'Salamanca', 'Segovia', 'Soria', 
        'Valladolid', 'Zamora', 'Barcelona', 'Gerona - Girona', 'Lerida - Lleida', 'Tarragona', 'Alicante - Alacant', 
        'Castellon', 'Valencia', 'Badajoz', 'Caceres', 'A Coruña - La Coruña', 'Lugo', 'Ourense - Orense', 
        'Pontevedra', 'La Rioja', 'Madrid', 'Murcia', 'Navarra', 'Alava - Araba', 'Guipuzcoa - Gipuzkoa', 
        'Vizcaya - Bizkaia', 'Ceuta', 'Melilla']

direcciones_con_acentos_correctos = ['Andalucía', 'Aragón', 'Asturias', 'Canarias', 'Cantabria', 'Castilla y La Mancha', 'Castilla y León', 
        'Cataluña', 'Comunidad Valenciana', 'Extremadura', 'Galicia', 'Islas Baleares - Illes Balears', 'La Rioja', 
        'Madrid', 'Murcia', 'Navarra', 'País Vasco', 'Almería', 'Cádiz', 'Córdoba', 'Granada', 'Huelva', 'Jaén', 
        'Málaga', 'Sevilla', 'Huesca', 'Teruel', 'Zaragoza', 'Las Palmas', 'Tenerife', 'Albacete', 'Ciudad Real', 
        'Cuenca', 'Guadalajara', 'Toledo', 'Ávila', 'Burgos', 'León', 'Palencia', 'Salamanca', 'Segovia', 'Soria', 
        'Valladolid', 'Zamora', 'Barcelona', 'Gerona - Girona', 'Lérida - Lleida', 'Tarragona', 'Alicante - Alacant', 
        'Castellón', 'Valencia', 'Badajoz', 'Cáceres', 'A Coruña - La Coruña', 'Lugo', 'Ourense - Orense', 
        'Pontevedra', 'La Rioja', 'Madrid', 'Murcia', 'Navarra', 'Álava - Araba', 'Guipúzcoa - Gipuzkoa', 
        'Vizcaya - Bizkaia', 'Ceuta', 'Melilla']

sector_directory = os.path.join(PLANTILLAS_CREADAS_DIR, eliminar_acentos_slash(SECTOR))
os.makedirs(sector_directory, exist_ok=True)

# Contador secuencial para numerar los archivos y que coincidan con urls.csv
indice_secuencial = 0

for index, fila in df_seleccionado_pais_comunidades_ciudades.iterrows():
    context = {
        'Localizacion': agregar_caracteres_especiales(fila['nombre'], direcciones_con_acentos, direcciones_con_acentos_correctos),
        'registros': '{:.0f}'.format(fila['registros']) if pd.notna(fila['registros']) else '',
        'dir': fila['dir'] if pd.notna(fila['dir']) else '',
        'phone': '{:.0f}'.format(fila['phone']) if pd.notna(fila['phone']) else '',
        'mail': '{:.0f}'.format(fila['mail']) if pd.notna(fila['mail']) else '',
        'cat': '{:.0f}'.format(fila['cat']) if pd.notna(fila['cat']) else '',
        'web': '{:.0f}'.format(fila['web']) if pd.notna(fila['web']) else '',
        'Precio': '{:.2f}'.format(fila['Precio']) if pd.notna(fila['Precio']) else ''
    }
    
    localizacion_name = eliminar_acentos_slash(fila['nombre'])
    
    # Incrementar índice secuencial antes de generar el archivo (empezará en 1)
    indice_secuencial += 1
    
    if fila['nombre'] in pais:
        nombre_base = f'{eliminar_acentos_slash(SECTOR)}_{eliminar_acentos_slash("Pais")}_{localizacion_name}.docx'
        nombre_archivo = f'{indice_secuencial}_{nombre_base}'
        generar_documento(fila, context, doc_pais, nombre_archivo, sector_directory)
    elif fila['nombre'] in comunidades:
        nombre_base = f'{eliminar_acentos_slash(SECTOR)}_{eliminar_acentos_slash("Comunidad Autonoma")}_{localizacion_name}.docx'
        nombre_archivo = f'{indice_secuencial}_{nombre_base}'
        generar_documento(fila, context, doc_comunidad, nombre_archivo, sector_directory)
    elif fila['nombre'] in ciudades:
        nombre_base = f'{eliminar_acentos_slash(SECTOR)}_{eliminar_acentos_slash("Ciudad Autonoma")}_{localizacion_name}.docx'
        nombre_archivo = f'{indice_secuencial}_{nombre_base}'
        generar_documento(fila, context, doc_ciudad, nombre_archivo, sector_directory)

for index, fila in df_seleccionado_provincias.iterrows():
    context = {
        'Localizacion': agregar_caracteres_especiales(fila['nombre'], direcciones_con_acentos, direcciones_con_acentos_correctos),
        'registros': '{:.0f}'.format(fila['registros']) if pd.notna(fila['registros']) else '',
        'dir': fila['dir'] if pd.notna(fila['dir']) else '',
        'phone': '{:.0f}'.format(fila['phone']) if pd.notna(fila['phone']) else '',
        'mail': '{:.0f}'.format(fila['mail']) if pd.notna(fila['mail']) else '',
        'cat': '{:.0f}'.format(fila['cat']) if pd.notna(fila['cat']) else '',
        'web': '{:.0f}'.format(fila['web']) if pd.notna(fila['web']) else '',
        'Precio': '{:.2f}'.format(fila['Precio']) if pd.notna(fila['Precio']) else ''
    }
    localizacion_name = eliminar_acentos_slash(fila['nombre'])
    
    # Incrementar índice secuencial antes de generar el archivo
    indice_secuencial += 1
    
    if fila['nombre'] in provincias:
        nombre_base = f'{eliminar_acentos_slash(SECTOR)}_{eliminar_acentos_slash("Provincia")}_{localizacion_name}.docx'
        nombre_archivo = f'{indice_secuencial}_{nombre_base}'
        generar_documento(fila, context, doc_provincia, nombre_archivo, sector_directory)

# Ejecutar la generación del CSV con URLs
sector_directory_seleccionar = os.path.join(PLANTILLAS_CREADAS_DIR, eliminar_acentos_slash(SECTOR))
generador_csv_con_url(sector_directory_seleccionar, EXCEL_FILE_SITEMAPS)

'''for file_name in os.listdir(sector_directory):
    if file_name.endswith(".docx"):  # Asegúrate de que sean archivos .docx
        file_path = os.path.join(sector_directory, file_name)
        
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                contenido = file.read()
            
            # Verificar si contiene {{Localizacion}}
            if "{{Localizacion}}" in contenido:
                os.remove(file_path)  # Eliminar el archivo
                print(f"Ficha eliminada: {file_name}")
            else:
                print(f"Ficha conservada: {file_name}")
        
        except Exception as e:
            print(f"Error procesando {file_name}: {e}")'''