# Guía de Portabilidad - Generador MIE

Esta guía explica cómo hacer que los scripts sean completamente portables entre diferentes ordenadores.

## ✅ Aspectos que YA son portables

Los scripts ya están diseñados para ser portables en los siguientes aspectos:

1. **Rutas relativas**: Todos los scripts usan rutas relativas (`./`, `os.path.join()`), por lo que funcionan independientemente de dónde esté la carpeta del proyecto.

2. **WebDriver Manager**: Los drivers del navegador se descargan automáticamente, no necesitas instalarlos manualmente.

3. **Estructura de carpetas**: La estructura es autocontenida y puede copiarse a cualquier ubicación.

## ⚠️ Configuración necesaria para portabilidad

Para usar los scripts en otro ordenador, necesitas:

### Paso 1: Copiar la carpeta completa

Copia toda la carpeta `Generador MIE` a cualquier ubicación en el otro ordenador.

### Paso 2: Instalar Python y dependencias

1. **Instalar Python 3.7 o superior** (si no está instalado)
   - Descarga desde: https://www.python.org/downloads/
   - Asegúrate de marcar "Add Python to PATH" durante la instalación

2. **Instalar las dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

### Paso 3: Configurar credenciales

1. **Copia el archivo de ejemplo**:
   ```bash
   # En Windows (PowerShell)
   Copy-Item config.example.py config.py
   
   # En Linux/Mac
   cp config.example.py config.py
   ```

2. **Edita `config.py`** y completa las siguientes credenciales:
   - `WORDPRESS_USERNAME`: Tu usuario de WordPress
   - `WORDPRESS_PASSWORD`: Tu contraseña de WordPress
   - `GOOGLE_MAPS_API_KEY`: Tu clave API de Google Maps (si usas el script 5)

### Paso 4: Verificar estructura de carpetas

Asegúrate de que existan estas carpetas y archivos:

```
Generador MIE/
├── config.py                    ← Debes crearlo desde config.example.py
├── config.example.py
├── requirements.txt
├── 1-duplicador_urls.py
├── 2-generador_csv_codigos_postales.py
├── 3-generador_plantillas.py
├── 4-generador_urls.py
├── 5-mie_subida.py
├── 6-limpiador.py
├── fichas_a_generar/
│   └── ficha_a_generar.xlsx
├── csvs_codigos_postales_ciudades/
├── plantillas_generadas/
└── plantilla_para_generar.docx
```

## 🔧 Configuración avanzada

### Cambiar el navegador

En `config.py`, puedes cambiar el navegador usado:

```python
BROWSER = 'edge'  # Opciones: 'edge', 'chrome', 'firefox'
```

**Requisitos**:
- **Edge**: Viene preinstalado en Windows 10/11
- **Chrome**: Debe estar instalado manualmente
- **Firefox**: Debe estar instalado manualmente

### Modo headless (sin ventana)

Para ejecutar sin mostrar la ventana del navegador:

```python
HEADLESS_MODE = True
```

### Usar variables de entorno (opcional)

Puedes usar variables de entorno en lugar de editar `config.py`:

**Windows (PowerShell)**:
```powershell
$env:WORDPRESS_USERNAME = "tu_usuario"
$env:WORDPRESS_PASSWORD = "tu_contraseña"
$env:GOOGLE_MAPS_API_KEY = "tu_api_key"
```

**Linux/Mac**:
```bash
export WORDPRESS_USERNAME="tu_usuario"
export WORDPRESS_PASSWORD="tu_contraseña"
export GOOGLE_MAPS_API_KEY="tu_api_key"
```

## 📋 Checklist de portabilidad

Antes de usar los scripts en un nuevo ordenador, verifica:

- [ ] Python 3.7+ instalado
- [ ] Dependencias instaladas (`pip install -r requirements.txt`)
- [ ] Archivo `config.py` creado desde `config.example.py`
- [ ] Credenciales completadas en `config.py`
- [ ] Navegador instalado (Edge/Chrome/Firefox según configuración)
- [ ] Estructura de carpetas correcta
- [ ] Archivo Excel `fichas_a_generar/ficha_a_generar.xlsx` presente
- [ ] Plantilla Word `plantilla_para_generar.docx` presente

## 🚨 Problemas comunes

### Error: "No se encontró el archivo config.py"

**Solución**: Copia `config.example.py` como `config.py` y completa las credenciales.

### Error: "Navegador no soportado"

**Solución**: Verifica que el navegador especificado en `config.py` esté instalado, o cambia `BROWSER` a uno disponible.

### Error: "ModuleNotFoundError"

**Solución**: Instala las dependencias con `pip install -r requirements.txt`

### Error: "No se encontró el archivo Excel"

**Solución**: Verifica que el archivo esté en `fichas_a_generar/ficha_a_generar.xlsx` y que la ruta en `config.py` sea correcta.

## 🔒 Seguridad

**IMPORTANTE**: 
- **NO** subas `config.py` a repositorios públicos (ya está en `.gitignore`)
- **NO** compartas `config.py` con credenciales reales
- Usa `config.example.py` como plantilla para compartir

## 📝 Notas adicionales

- Los scripts funcionan en **Windows, Linux y Mac**
- Las rutas usan `/` y `os.path.join()` para compatibilidad multiplataforma
- El tamaño de ventana y otras configuraciones pueden ajustarse en `config.py`
- Los scripts pueden ejecutarse en cualquier orden, pero sigue el flujo recomendado: 1 → 2 → 3 → 4 → 5

