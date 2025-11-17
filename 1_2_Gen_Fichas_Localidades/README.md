# 📝 Generador de Fichas por Localidades

Script automatizado en Python para generar fichas de empresas organizadas por localidades de España, utilizando una plantilla Word personalizable y datos de un archivo Excel maestro.

---

## 🎯 ¿Qué hace este script?

Este script genera automáticamente documentos Word (.docx) personalizados para cada localidad española, incluyendo:

- **Información de la localidad**: nombre, provincia, estadísticas de empresas
- **URLs aleatorias**: 10 URLs del blog extraídas del archivo Excel
- **Datos estadísticos**: registros, direcciones, teléfonos, emails, categorías, webs y precios
- **Formato consistente**: nombres capitalizados correctamente y caracteres especiales manejados

### Ejemplo de salida:
```
Empresas_de_la_localidad_Aguilar_De_La_Frontera_de_Cordoba.docx
```

---

## 📦 Requisitos

### Software necesario
- Python 3.7 o superior
- pip (gestor de paquetes de Python)

### Librerías Python
El script requiere las siguientes librerías (ver `requirements.txt`):
- `pandas` >= 1.3.0
- `openpyxl` >= 3.0.0
- `docxtpl` >= 0.16.0
- `python-docx` >= 0.8.11

### Archivos necesarios
1. **Plantilla Word** (obligatorio):
   ```
   docsparascript/plantilla_localidades.docx
   ```
   - Debe contener placeholders como `{{Localidad}}`, `{{Provincia}}`, `{{URL_1}}`, etc.

2. **Archivo Excel de localidades** (obligatorio):
   ```
   datos_maestros/1_maestro_localidades_nw_20251028_2.xlsx
   ```
   - Debe contener las siguientes columnas:
     - `localidad`: Nombre de la localidad
     - `provincia`: Nombre de la provincia
     - `registros`: Número de registros
     - `Direccion`: Número de direcciones
     - `Telefono`: Número de teléfonos
     - `Mail`: Número de emails
     - `Category`: Número de categorías
     - `Website`: Número de websites
     - `precio`: Precio de la ficha
     - `url`: URLs del blog (una por fila, o múltiples filas con URLs)

---

## 🚀 Instalación

### Opción 1: Usar el script batch (Windows)
1. Ejecuta `1_INSTALAR_DEPENDENCIAS.bat`
2. Espera a que se instalen las dependencias

### Opción 2: Instalación manual
```bash
pip install -r requirements.txt
```

---

## ⚙️ Configuración

### 1. Controlar el número de fichas a generar

Abre `DefinitivoLocalidades.py` y modifica la línea **31**:

```python
NUM_FICHAS_PRUEBA = 10  # Genera 10 fichas
```

**Opciones:**
- `NUM_FICHAS_PRUEBA = 10` → Genera solo las primeras 10 fichas (modo prueba)
- `NUM_FICHAS_PRUEBA = 50` → Genera las primeras 50 fichas
- `NUM_FICHAS_PRUEBA = None` → Genera **TODAS** las fichas del archivo Excel

### 2. Configurar el sector (opcional)

En la línea **35**, puedes cambiar el nombre del sector:

```python
SECTOR = 'Empresas de la localidad '
```

### 3. Modo verbose (opcional)

En la línea **38**, activa mensajes detallados:

```python
VERBOSE = True  # Muestra información detallada durante la ejecución
```

---

## 🏃 Ejecución

### Opción 1: Usar el script batch (Windows)
Doble clic en `EJECUTAR_SCRIPT_LOCALIDADES.bat`

### Opción 2: Ejecutar manualmente
```bash
python DefinitivoLocalidades.py
```

### Opción 3: Desde un IDE
Abre `DefinitivoLocalidades.py` en tu IDE favorito y ejecuta el script.

---

## 📊 Proceso de ejecución

El script realiza los siguientes pasos:

1. **Lectura de datos**: Lee el archivo Excel de localidades
2. **Validación**: Verifica que existan todos los archivos necesarios
3. **Generación de plantilla**: Crea una plantilla base con URLs aleatorias
4. **Generación de fichas**: Para cada localidad:
   - Capitaliza correctamente nombres de localidades y provincias
   - Formatea números y precios
   - Genera un documento Word personalizado
5. **Exportación de datos**:
   - Guarda un CSV con todos los datos procesados
   - Genera un CSV con URLs (mismo número que fichas generadas)

---

## 📁 Estructura de salida

### Directorios creados

```
Script_Gen_Fichas_Localidades_/
│
├── Plantillas_creadas/
│   └── Empresas_de_la_localidad/
│       ├── Plantilla_localidad.docx          # Plantilla base generada
│       ├── Empresas_de_la_localidad_*.docx   # Fichas generadas
│       └── urls_localidades.csv              # CSV con URLs
│
├── csv_creados/
│   └── Generador_csv_localidades.csv         # CSV con todos los datos
│
└── ...
```

### Formato de nombres de archivo

Las fichas se nombran con el siguiente formato:
```
[Sector]_[Localidad]_de_[Provincia].docx
```

Ejemplo:
```
Empresas_de_la_localidad_Aguilar_De_La_Frontera_de_Cordoba.docx
```

---

## 🔧 Funciones principales

### `eliminar_acentos_slash(text)`
Elimina acentos y caracteres especiales de un texto para crear nombres de archivo válidos.
- Convierte: "Aguilar de la Frontera" → "Aguilar_de_la_Frontera"
- Elimina acentos, espacios, guiones y barras

### `capitalizar_localidad(text)`
Capitaliza correctamente nombres de localidades y provincias.
- Respeta palabras en minúscula: "de", "del", "la", "el", "los", "las", "y"
- Ejemplo: "aguilar de la frontera" → "Aguilar de la Frontera"

### `get_random_urls(ARCHIVO_EXCEL_LOCALIDADES, num_urls=10)`
Extrae URLs aleatorias del archivo Excel de localidades.
- Busca la columna 'url' (sin distinguir mayúsculas/minúsculas)
- Filtra URLs vacías
- Retorna un diccionario con `URL_1`, `URL_2`, ..., `URL_10`

### `generar_plantilla_localidad(ARCHIVO_EXCEL_LOCALIDADES)`
Crea una plantilla base con placeholders genéricos y URLs aleatorias.
- Genera 10 URLs aleatorias
- Crea la plantilla con datos genéricos que se reemplazarán después

### `generar_documento_localidad(context, doc_plantilla, nombre_archivo, sector_directory)`
Genera un documento Word individual para una localidad específica.
- Usa la plantilla base
- Reemplaza los placeholders con datos reales de la localidad

### `generador_csv_con_url_localidades(sector_directory, ARCHIVO_EXCEL_LOCALIDADES, num_urls=None)`
Genera un archivo CSV con URLs del archivo de localidades.
- Si `num_urls` es especificado, genera exactamente ese número de URLs (aleatorias)
- Si `num_urls=None`, genera todas las URLs disponibles
- **Nota importante**: El número de URLs generadas coincide con el número de fichas creadas

---

## 📋 Variables en la plantilla Word

La plantilla Word debe contener estos placeholders:

| Placeholder | Descripción | Ejemplo |
|------------|-------------|---------|
| `{{Localidad}}` | Nombre de la localidad | Aguilar de la Frontera |
| `{{Provincia}}` | Nombre de la provincia | Córdoba |
| `{{Precio}}` | Precio formateado | 99.99 |
| `{{registros}}` | Número de registros | 150 |
| `{{dir}}` | Número de direcciones | 120 |
| `{{phone}}` | Número de teléfonos | 85 |
| `{{mail}}` | Número de emails | 90 |
| `{{cat}}` | Número de categorías | 25 |
| `{{web}}` | Número de websites | 75 |
| `{{Mes}}` | Mes actual en español | febrero |
| `{{Sector}}` | Nombre del sector | Empresas de la localidad |
| `{{URL_1}}` a `{{URL_10}}` | URLs aleatorias del blog | https://... |

---

## 🐛 Solución de problemas

### Error: "El archivo de localidades no contiene la columna 'url'"
**Solución**: El script busca una columna llamada 'url' (sin distinguir mayúsculas/minúsculas). Verifica que tu archivo Excel tenga una columna con ese nombre.

### Error: "No hay suficientes URLs"
**Solución**: El archivo Excel no tiene suficientes URLs. Asegúrate de que haya al menos 10 URLs en la columna 'url' del archivo.

### Las fechas no están en español
**Solución**: El script intenta configurar el idioma español automáticamente. Si no funciona, es un aviso menor y no afecta la funcionalidad.

### Los nombres de archivo tienen caracteres extraños
**Solución**: El script elimina automáticamente caracteres especiales. Si persiste, verifica que la función `eliminar_acentos_slash()` esté funcionando correctamente.

---

## 📈 Ejemplo de uso

### Modo prueba (10 fichas)
```python
NUM_FICHAS_PRUEBA = 10
```
**Resultado:**
- Genera 10 fichas Word
- Genera CSV con 10 URLs
- Tiempo estimado: ~5-10 segundos

### Modo completo (todas las fichas)
```python
NUM_FICHAS_PRUEBA = None
```
**Resultado:**
- Genera todas las fichas del archivo Excel (ej: 683 fichas)
- Genera CSV con el mismo número de URLs
- Tiempo estimado: ~2-5 minutos (depende del número de fichas)

---

## 📝 Notas importantes

1. **URLs aleatorias**: Las URLs que aparecen en cada ficha son aleatorias y pueden variar entre ejecuciones.

2. **CSV de URLs**: El archivo `urls_localidades.csv` contiene exactamente el mismo número de URLs que fichas generadas.

3. **Capitalización**: El script capitaliza automáticamente nombres de localidades y provincias según reglas del español.

4. **Caracteres especiales**: Los nombres de archivo se limpian automáticamente para evitar problemas con sistemas de archivos.

5. **Validación de datos**: El script salta automáticamente las filas con localidad o provincia vacías.

---

## 🔄 Actualizaciones recientes

- **Extracción de URLs del archivo de localidades**: Ahora las URLs se extraen directamente del archivo Excel de localidades (`1_maestro_localidades_nw_20251028_2.xlsx`)
- **CSV con número correspondiente**: El CSV de URLs ahora genera exactamente el mismo número de URLs que fichas creadas
- **Búsqueda flexible de columnas**: El script busca la columna 'url' sin distinguir mayúsculas/minúsculas

---

## 📞 Soporte

Si encuentras problemas o tienes preguntas:
1. Verifica que todos los archivos necesarios estén en sus ubicaciones correctas
2. Revisa que el archivo Excel tenga las columnas requeridas
3. Activa `VERBOSE = True` para ver mensajes detallados durante la ejecución

---

## 📄 Licencia

Este script es de uso interno. Modifica según tus necesidades.

---

**Última actualización**: 2025

