# CODIGO REALIAZADO POR JONATHAN JD , CON LA AYUDA DE CURSOR
# https://github.com/jonathanjd7


# Generador de Fichas - España, CCAA y Provincias

Script Python para generar automáticamente documentos Word (fichas) para diferentes sectores y localizaciones geográficas en España (País, Comunidades Autónomas, Ciudades Autónomas y Provincias).

## 📋 Descripción

Este script genera documentos Word personalizados basados en plantillas, utilizando datos de archivos Excel. Crea fichas para diferentes niveles geográficos (País, Comunidades Autónomas, Ciudades Autónomas y Provincias) y genera un archivo CSV con las URLs correspondientes que coinciden con el orden de los documentos generados.

## ✨ Características

- ✅ **Portable**: Funciona en cualquier ordenador sin necesidad de modificar rutas
- ✅ **Control de prueba**: Permite generar un número limitado de fichas para pruebas
- ✅ **Numeración automática**: Los archivos se numeran secuencialmente para mantener el orden
- ✅ **Generación de URLs**: Crea un archivo CSV con URLs que coinciden con el orden de los documentos
- ✅ **Múltiples niveles geográficos**: Soporta País, CCAA, Ciudades Autónomas y Provincias
- ✅ **Plantillas personalizables**: Utiliza plantillas Word con campos dinámicos

## 📦 Requisitos Previos

### Software necesario

- **Python 3.7 o superior**
- **pip** (gestor de paquetes de Python)

### Dependencias de Python

El script requiere las siguientes librerías (instaladas automáticamente con `requirements.txt`):

- `pandas` >= 1.5.0
- `python-docx-template` >= 0.16.0
- `openpyxl` >= 3.0.0

## 📁 Estructura de Archivos Necesarios

Antes de ejecutar el script, asegúrate de tener la siguiente estructura en la carpeta `Gen_Fichas`:

```
Gen_Fichas/
├── GeneracionFichas.py                    # Script principal
├── requirements.txt                  # Dependencias de Python
├── site_maps_2025_02_12.xlsx        # Archivo Excel con URLs y datos de sitemap
├── Generador_csv_ccaa_prov.xlsx     # Archivo Excel con datos de CCAA y provincias
├── Plantilla_Generica.docx          # Plantilla Word base
├── csv_creados/                      # (Se crea automáticamente)
│   ├── Generador_csv_ccaa.csv
│   └── Generador_csv_prov.csv
└── Plantillas_creadas/              # (Se crea automáticamente)
    └── [Sector]/
        └── [Archivos generados]
```

## 🚀 Instalación

### 1. Copiar la carpeta

Copia toda la carpeta `Gen_Fichas` al ordenador donde quieras ejecutar el script. No es necesario modificar ninguna ruta, el script detecta automáticamente su ubicación.

### 2. Instalar dependencias

Abre una terminal en la carpeta `Gen_Fichas` y ejecuta:

```bash
pip install -r requirements.txt
```

### 3. Verificar archivos necesarios

Asegúrate de que todos los archivos requeridos estén presentes:
- `site_maps_2025_02_12.xlsx`
- `Generador_csv_ccaa_prov.xlsx`
- `Plantilla_Generica.docx`

## ⚙️ Configuración

### Configurar el Sector

Edita la línea 42-43 del archivo `GeneracionFichas.py`:

```python
SECTOR = 'Supermercados y Tiendas de Alimentación'
```

**⚠️ IMPORTANTE**: El nombre del sector debe coincidir exactamente con el nombre del sector en el archivo `site_maps_2025_02_12.xlsx`.

### Control de Prueba (Modo de Prueba)

Puedes controlar cuántas fichas generar editando la línea 38 de `GeneracionFichas.py`:

```python
NUM_FICHAS_PRUEBA = 0  # <--- CAMBIA ESTE NÚMERO AQUÍ
```

**Opciones:**
- `None` o `0` → Genera **TODAS** las fichas del archivo
- `10` → Genera solo las primeras **10 fichas** (útil para pruebas)
- `50` → Genera solo las primeras **50 fichas**

**Ejemplo para pruebas:**
```python
NUM_FICHAS_PRUEBA = 10  # Genera solo 10 fichas para probar
```

## 🎯 Uso

### Ejecutar el script

1. Abre una terminal en la carpeta `Gen_Fichas`
2. Ejecuta el script:

```bash
python GeneracionFichas.py
```

### Proceso de ejecución

El script realiza las siguientes acciones en orden:

1. **Validación**: Verifica que todos los archivos necesarios existan
2. **Lectura de datos**: Lee los datos de los archivos Excel
3. **Generación de plantillas**: Crea plantillas específicas para cada tipo de región
4. **Generación de documentos**: Crea los documentos Word numerados secuencialmente
5. **Generación de CSV**: Crea el archivo `urls.csv` con las URLs correspondientes

### Salida esperada

Durante la ejecución verás mensajes como:

```
Configuración cargada correctamente.

[MODO COMPLETO] Generando TODAS las fichas
   - Países/Comunidades/Ciudades: 21
   - Provincias: 51
   - Total: 72

Los archivos CSV se han creado correctamente.
Plantilla "..." generada con éxito.
Documento "..." generado con éxito.
...
Archivo CSV combinado generado: ...
```

## 📂 Estructura de Salida

### Archivos generados

Los documentos se generan en:
```
Plantillas_creadas/[Sector]/
```

### Formato de nombres de archivo

Los archivos se generan con el siguiente formato:
```
{numero}_{Sector}_{Tipo}_{Localizacion}.docx
```

**Ejemplos:**
- `1_Supermercados_y_Tiendas_de_Alimentacion_Ciudad_Autonoma_Ceuta.docx`
- `2_Supermercados_y_Tiendas_de_Alimentacion_Comunidad_Autonoma_Andalucia.docx`
- `3_Supermercados_y_Tiendas_de_Alimentacion_Provincia_Madrid.docx`

### Archivo CSV de URLs

Se genera un archivo `urls.csv` en la misma carpeta que contiene las URLs en el mismo orden que los documentos generados. Cada fila del CSV corresponde con el documento Word del mismo número.

**Ejemplo de `urls.csv`:**
```
https://gsas.es/producto/...
https://gsas.es/producto/...
https://gsas.es/producto/...
```

## 🔧 Funcionalidades Principales

### 1. Generación de Plantillas

El script genera automáticamente plantillas específicas para cada tipo de región:
- Plantilla para País
- Plantilla para Comunidad Autónoma
- Plantilla para Ciudad Autónoma
- Plantilla para Provincia

### 2. Numeración Secuencial

Los archivos se numeran automáticamente (1, 2, 3...) para mantener el orden y facilitar la correspondencia con el archivo `urls.csv`.

### 3. Extracción de URLs

El script busca automáticamente las URLs correspondientes en el archivo `site_maps_2025_02_12.xlsx` basándose en:
- El sector configurado
- El tipo de región (País, CCAA, Ciudad Autónoma, Provincia)
- El nombre de la localización

### 4. Normalización de Nombres

Los nombres de archivo se normalizan automáticamente:
- Eliminación de acentos
- Reemplazo de espacios y caracteres especiales por guiones bajos
- Formato consistente para todos los archivos

## 📊 Datos Requeridos en los Archivos Excel

### `Generador_csv_ccaa_prov.xlsx`

Debe contener dos hojas:

**Hoja 1: `Comunidades_csv_copiar`**
- Columnas: nombre, registros, dir, phone, mail, cat, web, Precio

**Hoja 2: `Provincias_csv_copiar`**
- Columnas: nombre, registros, dir, phone, mail, cat, web, Precio

### `site_maps_2025_02_12.xlsx`

Debe contener las siguientes columnas:
- `Sector`: Nombre del sector (debe coincidir con `SECTOR` en el script)
- `España`: Nombre del país (para fichas de país)
- `Comunidad_Autonoma`: Nombre de la comunidad autónoma
- `Ciudad_Autonoma`: Nombre de la ciudad autónoma
- `Provincia`: Nombre de la provincia
- `Categoria`: Categoría (debe incluir 'blog' para URLs)
- `URL`: URL correspondiente

## 🐛 Solución de Problemas

### Error: "El archivo ... no existe"

**Solución**: Verifica que todos los archivos necesarios estén en la carpeta `Gen_Fichas`:
- `site_maps_2025_02_12.xlsx`
- `Generador_csv_ccaa_prov.xlsx`
- `Plantilla_Generica.docx`

### Error: "No hay suficientes URLs de blog"

**Solución**: Asegúrate de que el archivo `site_maps_2025_02_12.xlsx` contenga al menos 10 filas con `Categoria == 'blog'`.

### Error: "No se encontró URL para el archivo: ..."

**Solución**: Verifica que:
1. El nombre del sector en el script coincida exactamente con el del archivo Excel
2. El nombre de la localización en el Excel coincida con el del archivo generador
3. Los nombres estén escritos correctamente (con o sin acentos según corresponda)

### Los archivos no se ordenan correctamente

**Solución**: El script ordena automáticamente los archivos por el número al inicio. Si hay problemas, verifica que los nombres de archivo sigan el formato `{numero}_{nombre}.docx`.

### Advertencia sobre locale

Si ves el mensaje: "Advertencia: No se pudo configurar el locale en español"

**Solución**: No es crítico, el script funcionará igual. Solo afecta al formato de las fechas. Para solucionarlo:
- **Windows**: Instala el paquete de idioma español
- **Linux/Mac**: Ejecuta `sudo locale-gen es_ES.UTF-8`

## 📝 Notas Importantes

1. **Orden de generación**: Los documentos se generan primero para Países/Comunidades/Ciudades y luego para Provincias, manteniendo el orden del archivo Excel.

2. **Sincronización con URLs**: El archivo `urls.csv` se genera en el mismo orden que los documentos Word, por lo que la fila N del CSV corresponde con el documento número N.

3. **Plantilla genérica**: La plantilla `Plantilla_Generica.docx` debe contener campos de reemplazo como `{{Localizacion}}`, `{{Precio}}`, `{{registros}}`, etc.

4. **Modo portable**: El script detecta automáticamente su ubicación, por lo que puedes copiar toda la carpeta a cualquier ordenador y funcionará sin modificaciones.

## 🔄 Actualizaciones y Mantenimiento

### Cambiar el sector

1. Edita `SECTOR` en la línea 42-43
2. Asegúrate de que el sector exista en `site_maps_2025_02_12.xlsx`
3. Ejecuta el script nuevamente

### Añadir nuevas localizaciones

1. Añade las nuevas filas en `Generador_csv_ccaa_prov.xlsx`
2. Añade las correspondientes URLs en `site_maps_2025_02_12.xlsx`
3. Ejecuta el script

## 📞 Soporte

Si encuentras problemas o tienes preguntas:

1. Verifica que todos los archivos necesarios estén presentes
2. Revisa los mensajes de error en la consola
3. Asegúrate de que los nombres del sector coincidan exactamente
4. Prueba primero con `NUM_FICHAS_PRUEBA = 10` para verificar que todo funciona

## 📄 Licencia

Este script es de uso interno. Asegúrate de tener los permisos necesarios para usar los datos y plantillas incluidos.

---

**Versión**: 1.0  
**Última actualización**: 2025

