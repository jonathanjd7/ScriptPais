## Scripts de Automatización GSAS

Estos scripts automatizan la actualización masiva de fichas de producto en el sitio de GSAS (WooCommerce), tomando el contenido de documentos Word y un archivo CSV con las URLs objetivo.

### Contenido

- `DefinitivoGsas.py`: versión base del proceso de actualización automática.
- `DefinitivoGsasProd..py`: variante que detiene la ejecución si encuentra una página no disponible para permitir una corrección manual antes de continuar.

Ambos comparten la misma lógica principal; la variante *Prod* incorpora la verificación adicional.

### Requisitos previos

- Python 3.x instalado.
- Dependencias: `python-docx`, `selenium`.
- Microsoft Edge instalado junto con `msedgedriver.exe` disponible en el PATH.
- Carpeta `Documentos` en el mismo directorio que los scripts (o en el directorio padre) con:
  - Archivos `.docx` numerados (por ejemplo `1_nombre.docx`, `2_nombre.docx`).
  - Archivo `urls.csv` con una URL por fila, en el orden en que se deben actualizar los productos.

### Flujo de trabajo

1. El script inicia sesión en `https://www.gsas.es/grupo-sanz-acceso/` usando las credenciales incluidas en el código.
2. Busca la carpeta `Documentos`, lista los `.docx` y verifica que coincidan con las URLs del CSV.
3. Para cada par URL/Word:
   - Accede a la página del producto y habilita el modo “Editar producto”.
   - Extrae título, descripciones, tablas, etiquetas, datos SEO y precio del documento Word.
   - Rellena los campos correspondientes en WooCommerce y pulsa “Actualizar”.
   - Espera un tiempo para garantizar el guardado antes de continuar con el siguiente producto.
4. Al final genera un resumen por consola con los elementos procesados, saltos y errores.
5. `DefinitivoGsasProd..py` se detiene inmediatamente si detecta un mensaje de “URL no encontrada” tras abrir una página de producto.

### Cómo usarlo

1. Abre una terminal en `ScriptPais/CodigoGsas`.
2. (Opcional) Crea un entorno virtual e instala dependencias:  
   ```bash
   pip install python-docx selenium
   ```
3. Ejecuta el script deseado:  
   ```bash
   python DefinitivoGsasProd..py
   ```
   o  
   ```bash
   python DefinitivoGsas.py
   ```
4. Observa la terminal para seguir el progreso. Si usas la variante *Prod* y aparece un mensaje de URL no encontrada, corrige la entrada correspondiente antes de relanzar el script.

### Notas

- Ajusta las credenciales en el código si cambian en WordPress.
- Puedes habilitar el modo headless descomentando la línea `options.add_argument("--headless")` si no necesitas ver el navegador.
- Los tiempos de espera (`time.sleep`) son largos para asegurar que los cambios se guarden; ajústalos según el rendimiento del sitio.
- Revisa el resumen final para confirmar qué filas se procesaron o fallaron.

