

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

# Configurar la localización para fechas en español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except:
        print("[AVISO] No se pudo configurar el idioma español para fechas")

# ============================
# ⚙️ CONTROL DE PRUEBA - EDITA AQUÍ
# ============================
# Cambia NUM_FICHAS_PRUEBA para controlar cuántas fichas generar:
#   None o 0  --> Genera TODAS las fichas
#   10        --> Genera solo las primeras 10 fichas (PRUEBA)
#   50        --> Genera solo las primeras 50 fichas
NUM_FICHAS_PRUEBA = 0  # <--- CAMBIA ESTE NÚMERO AQUÍ
# ============================

# Variables globales
MES = datetime.datetime.now().strftime("%B")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Rutas relativas (PORTABLE - no necesita modificación)
EXCEL_FILE_SITEMAPS = os.path.join(BASE_DIR, 'docsparascript', 'site_maps_2025_02_12.xlsx')
EXCEL_URLS_DATOS_MAESTROS = os.path.join(BASE_DIR, 'datos_maestros', 'datosmaestrosurls.xlsx')
OUTPUT_DIRECTORY_CSV = os.path.join(BASE_DIR, 'csv_creados')
ARCHIVO_EXCEL_PAISES = os.path.join(BASE_DIR, 'docsparascript', 'ITALIA.xlsx')
PLANTILLA_PAISES = os.path.join(BASE_DIR, 'docsparascript', 'plantilla_pais.docx')
PLANTILLAS_CREADAS_DIR = os.path.join(BASE_DIR, 'Plantillas_creadas')

# Crear directorios necesarios
os.makedirs(OUTPUT_DIRECTORY_CSV, exist_ok=True)
os.makedirs(PLANTILLAS_CREADAS_DIR, exist_ok=True)

# Validación de rutas
if not os.path.exists(BASE_DIR):
    raise FileNotFoundError(f"La ruta base {BASE_DIR} no existe.")

# Comentado temporalmente - si no tienes sitemap, comenta esta línea
# if not os.path.exists(EXCEL_FILE_SITEMAPS):
#     raise FileNotFoundError(f"El archivo {EXCEL_FILE_SITEMAPS} no existe.")

if not os.path.exists(ARCHIVO_EXCEL_PAISES):
    raise FileNotFoundError(f"El archivo {ARCHIVO_EXCEL_PAISES} no existe.")

if not os.path.exists(PLANTILLA_PAISES):
    raise FileNotFoundError(f"La plantilla de paises {PLANTILLA_PAISES} no existe.")

print("Configuracion cargada correctamente.")


# ============================
# 2. FUNCIONES
# ============================

def eliminar_acentos_slash(text):
    """Elimina acentos y caracteres especiales para nombres de archivo"""
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
    text = str(text)
    text = re.sub(r'[áàâäãåéèêëíìîïóòôöõøúùûüýÿñÁÀÂÄÃÅÉÈÊËÍÌÎÏÓÒÔÖÕØÚÙÛÜÝŸÑ/ -]', lambda x: replacements[x.group()], text)
    return re.sub(r'_+', '_', text)


def capitalizar_pais(text):
    """Capitaliza correctamente nombres de países"""
    # Palabras que deben ir en minúscula (excepto al inicio)
    palabras_minusculas = ['de', 'del', 'la', 'el', 'los', 'las', 'y', 'le', 'du', 'des']
    
    palabras = text.split()
    resultado = []
    
    for i, palabra in enumerate(palabras):
        if i == 0:  # Primera palabra siempre capitalizada
            resultado.append(palabra.capitalize())
        elif palabra.lower() in palabras_minusculas:
            resultado.append(palabra.lower())
        else:
            resultado.append(palabra.capitalize())
    
    return ' '.join(resultado)


def get_random_urls(EXCEL_URLS_DATOS_MAESTROS, num_urls=10):
    """Obtiene URLs aleatorias desde el archivo de datos maestros"""
    try:
        df = pd.read_excel(EXCEL_URLS_DATOS_MAESTROS)
        # Obtener todas las URLs que no sean nulas
        urls_disponibles = df['URL'].dropna().unique().tolist()
        
        if len(urls_disponibles) == 0:
            print(f"[AVISO] No se encontraron URLs en el archivo de datos maestros. Generando URLs vacias.")
            return {f'URL_{i+1}': 'https://ejemplo.com/blog/' for i in range(num_urls)}
        
        if len(urls_disponibles) < num_urls:
            # Si hay menos URLs disponibles que las necesarias, repetir algunas
            urls_seleccionadas = urls_disponibles * (num_urls // len(urls_disponibles) + 1)
            urls_seleccionadas = urls_seleccionadas[:num_urls]
        else:
            # Seleccionar URLs aleatorias
            urls_seleccionadas = random.sample(urls_disponibles, num_urls)
        
        return {f'URL_{i+1}': url for i, url in enumerate(urls_seleccionadas)}
    except FileNotFoundError:
        print(f"[AVISO] No se encontro el archivo de datos maestros: {EXCEL_URLS_DATOS_MAESTROS}. Generando URLs vacias.")
        return {f'URL_{i+1}': 'https://ejemplo.com/blog/' for i in range(num_urls)}
    except Exception as e:
        print(f"[AVISO] Error al obtener URLs del archivo de datos maestros: {e}. Generando URLs vacias.")
        return {f'URL_{i+1}': 'https://ejemplo.com/blog/' for i in range(num_urls)}


def extraer_nombre_doc_pais(nombrefichero):
    """Extrae información del nombre de archivo de país"""
    nombrefichero = nombrefichero.replace('.docx', '')
    if "_Pais_" in nombrefichero:
        partes = nombrefichero.split("_Pais_")
        sector = partes[0]
        # Extraer país
        pais = partes[1] if len(partes) > 1 else None
        return sector, pais
    return None, None


def generador_csv_con_url_paises(base_directory, EXCEL_FILE_SITEMAPS):
    """Genera CSV con URLs para documentos de países"""
    try:
        sitemap_df = pd.read_excel(EXCEL_FILE_SITEMAPS)
    except FileNotFoundError:
        print(f"[AVISO] No se puede generar CSV de URLs: archivo sitemap no encontrado.")
        return
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return
    
    # Usar la columna correcta del sitemap (Paq_Ps)
    columnas = ['Sector', 'Paq_Ps']
    for col in columnas:
        if col in sitemap_df.columns:
            sitemap_df[col] = sitemap_df[col].fillna("").apply(eliminar_acentos_slash)

    # Buscar todos los archivos .docx en la carpeta base
    ficheros_word = [f for f in os.listdir(base_directory) if f.endswith('.docx') and '_Pais_' in f]

    combinado_data = []

    for file in ficheros_word:
        sector, pais = extraer_nombre_doc_pais(file)
        if sector and pais:
            pais_sin_acentos = eliminar_acentos_slash(pais)
            # Buscar URL en el sitemap
            filtrado_fila = sitemap_df[
                (sitemap_df['Sector'] == eliminar_acentos_slash(sector)) &
                (sitemap_df['Paq_Ps'] == pais_sin_acentos)
            ]

            if not filtrado_fila.empty:
                url = filtrado_fila.iloc[0]['URL']
                combinado_data.append([url])
            else:
                print(f"No se encontro URL para el archivo: {file} (Sector: {sector}, Pais: {pais})")

    output_file_url = os.path.join(base_directory, "urls_paises.csv")
    combinado_df = pd.DataFrame(combinado_data, columns=['URL'])
    combinado_df.to_csv(output_file_url, index=False, header=False, encoding='utf-8')
    print(f"Archivo CSV de URLs generado: {output_file_url}")


# ============================
# 3. PROCESO PRINCIPAL
# ============================

print("\n" + "="*60)
print("GENERADOR DE FICHAS POR PAISES")
print("="*60 + "\n")

# Leer datos del archivo maestro de países
print(f"Leyendo archivo de paises: {ARCHIVO_EXCEL_PAISES}")
df_paises = pd.read_excel(ARCHIVO_EXCEL_PAISES)

# Limpiar y preparar columnas para comparación
df_paises['Sector_clean'] = df_paises['Sector'].fillna('').astype(str).str.strip()
df_paises['Pais_clean'] = df_paises['Pais'].fillna('').astype(str).str.strip()

# Eliminar duplicados de Sector+Pais, manteniendo la PRIMERA ocurrencia (orden del Excel)
# Esto asegura que se mantenga el orden original del Excel
df_paises = df_paises.drop_duplicates(subset=['Sector_clean', 'Pais_clean'], keep='first')

total_paises = len(df_paises)
print(f"Total de combinaciones Sector-Pais únicas en el archivo: {total_paises}")

# Filtrar solo las filas que tienen URL (para mantener el orden con url.csv)
if 'URL' in df_paises.columns:
    # Mantener solo las filas con URL válida, preservando el orden original
    df_paises = df_paises[df_paises['URL'].notna() & (df_paises['URL'].astype(str).str.strip() != '')].copy()
    # Resetear índice para mantener orden secuencial (0, 1, 2, ...)
    df_paises = df_paises.reset_index(drop=True)
    print(f"Filtradas {len(df_paises)} filas con URL válida (manteniendo orden del Excel)")
else:
    print("[AVISO] No se encontro la columna 'URL' en el Excel. Se generaran todas las plantillas.")

# Aplicar límite si NUM_FICHAS_PRUEBA está configurado
if NUM_FICHAS_PRUEBA is not None and NUM_FICHAS_PRUEBA > 0:
    df_paises = df_paises.head(NUM_FICHAS_PRUEBA)
    print(f"[MODO PRUEBA] Generando solo las primeras {NUM_FICHAS_PRUEBA} fichas")
    print(f"[CONSEJO] Para generar todas las fichas ({total_paises}), cambia NUM_FICHAS_PRUEBA = None en la linea 31\n")
else:
    print(f"[MODO COMPLETO] Generando TODAS las {len(df_paises)} fichas\n")

print(f"Columnas disponibles: {df_paises.columns.tolist()}\n")

# Obtener URLs para usar en todas las plantillas
print("Obteniendo URLs desde datos_maestros...")
urls_diccionario = get_random_urls(EXCEL_URLS_DATOS_MAESTROS, num_urls=10)

# Directorio base para documentos (sin sector específico ya que usamos el del Excel)
base_directory = os.path.join(PLANTILLAS_CREADAS_DIR, 'Paises')
os.makedirs(base_directory, exist_ok=True)

# Guardar CSV con los datos procesados
csv_output = os.path.join(OUTPUT_DIRECTORY_CSV, 'Generador_csv_paises.csv')
df_paises.to_csv(csv_output, index=False, encoding='utf-8')
print(f"CSV de paises guardado en: {csv_output}\n")

# Contador de documentos generados
contador = 0
errores = 0
indice_secuencial = 0  # Contador para el número en el nombre del archivo

print("Generando documentos por pais...\n")

# Generar documentos para cada país
for index, fila in df_paises.iterrows():
    try:
        # Obtener y limpiar datos del Excel
        sector_fila = str(fila['Sector']).strip() if pd.notna(fila['Sector']) else ''
        pais_fila = str(fila['Pais']).strip() if pd.notna(fila['Pais']) else None
        
        # Omitir filas sin país válido
        if not pais_fila or pais_fila == 'nan' or pais_fila == '':
            continue
        
        pais = pais_fila
        
        # Incrementar índice secuencial antes de generar el archivo (empezará en 1)
        indice_secuencial += 1
        
        # Capitalizar nombres correctamente
        pais_formateado = capitalizar_pais(pais)
        sector_formateado = sector_fila.strip()
        
        # Preparar contexto con los datos del país
        context = {
            'País': pais_formateado,
            'RegionTipo': 'País',
            'Localizacion': pais_formateado,
            'Localidad': pais_formateado,
            'Provincia': '',
            'Sector': sector_formateado,
            'registros': '{:.0f}'.format(fila['Cuenta de Nombre']) if pd.notna(fila['Cuenta de Nombre']) else '0',
            'dir': '{:.0f}'.format(fila['Cuenta de Direccion']) if pd.notna(fila['Cuenta de Direccion']) else '0',
            'phone': '{:.0f}'.format(fila['Cuenta de Telefono']) if pd.notna(fila['Cuenta de Telefono']) else '0',
            'mail': '{:.0f}'.format(fila['Cuenta de Mail']) if pd.notna(fila['Cuenta de Mail']) else '0',
            'cat': '{:.0f}'.format(fila['Cuenta de Categoria']) if pd.notna(fila['Cuenta de Categoria']) else '0',
            'web': '{:.0f}'.format(fila['Cuenta de Website']) if pd.notna(fila['Cuenta de Website']) else '0',
            'Precio': '{:.2f}'.format(fila['precios']) if pd.notna(fila['precios']) else '0.00',
            'Mes': MES,
            # URLs del blog
            **urls_diccionario
        }
        
        # Nombre del archivo: numero_Sector_Pais_Pais.docx (para mantener orden con url.csv)
        pais_archivo = eliminar_acentos_slash(pais)
        sector_archivo = eliminar_acentos_slash(sector_formateado)
        # Añadir número secuencial al inicio para mantener el orden con url.csv
        nombre_archivo = f'{indice_secuencial}_{sector_archivo}_Pais_{pais_archivo}.docx'
        
        # Generar el documento directamente en la carpeta base
        doc_plantilla = DocxTemplate(PLANTILLA_PAISES)
        doc_plantilla.render(context)
        doc_plantilla.save(os.path.join(base_directory, nombre_archivo))
        
        # Incrementar contador de documentos generados
        contador += 1
        
        # Mostrar progreso
        if contador % 50 == 0 or (len(df_paises) <= 50 and contador % 5 == 0):
            print(f"   [{contador}/{len(df_paises)}] documentos generados...")
        
    except Exception as e:
        print(f"[ERROR] Error procesando fila {index}: {e}")
        errores += 1
        continue

print("\n" + "="*60)
print(f"[COMPLETADO] Proceso completado!")
print(f"   Documentos generados: {contador}")
print(f"   Errores: {errores}")
print(f"   Total procesado: {contador + errores}")
print("="*60 + "\n")

# Generar CSV con URLs del sitemap (si existe)
print("Generando CSV con URLs del sitemap...")
generador_csv_con_url_paises(base_directory, EXCEL_FILE_SITEMAPS)

# Generar CSV con URLs del Excel ITALIA (en el mismo orden que las plantillas generadas)
print("Generando CSV con URLs del Excel ITALIA...")
if 'URL' in df_paises.columns:
    # Obtener URLs en el mismo orden que df_paises (ya filtrado y ordenado)
    urls_italia = df_paises['URL'].tolist()
    urls_csv = os.path.join(base_directory, 'url.csv')
    # Guardar sin encabezado, solo las URLs (en el mismo orden que las plantillas)
    pd.DataFrame(urls_italia, columns=['URL']).to_csv(urls_csv, index=False, header=False, encoding='utf-8')
    print(f"Archivo CSV de URLs (ITALIA) generado: {urls_csv} ({len(urls_italia)} URLs, mismo orden que plantillas)")
else:
    print("[AVISO] No se encontro la columna 'URL' en el Excel ITALIA")

print("\nProceso finalizado con exito!")

