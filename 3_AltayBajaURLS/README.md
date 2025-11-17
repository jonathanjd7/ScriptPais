# Generador MIE - Sistema Automatizado de Generación de Contenido

Sistema automatizado para generar y publicar contenido en WordPress/WooCommerce usando Selenium. Incluye scripts para duplicar productos, generar plantillas Word y subir contenido automáticamente.

## 🚀 Inicio Rápido (Portabilidad)

Los scripts son **completamente portables** entre ordenadores. Para usarlos en un nuevo ordenador:

### Paso 1: Copiar la carpeta completa
Copia toda la carpeta `3_AltayBajaURLS` a cualquier ubicación.

### Paso 2: Instalar dependencias
```bash
pip install -r requirements.txt
```

### Paso 3: Configurar credenciales
```bash
# Copia el archivo de ejemplo
cp config.example.py config.py  # Linux/Mac
# o
Copy-Item config.example.py config.py  # Windows PowerShell

# Edita config.py y completa tus credenciales:
# - WORDPRESS_USERNAME
# - WORDPRESS_PASSWORD  
# - GOOGLE_MAPS_API_KEY (si usas el script 5)
```

### Paso 4: Ejecutar los scripts
```bash
python 1-duplicador_urls.py
python 2-generador_csv_codigos_postales.py
python 3-generador_plantillas.py
python 4-generador_urls.py
python 5-mie_subida.py
```

📖 **Para más detalles, consulta [GUIA_PORTABILIDAD.md](GUIA_PORTABILIDAD.md)**

## 📋 Requisitos Previos

1. **Python 3.7 o superior** instalado
2. **Navegador** instalado (Edge, Chrome o Firefox - configurable en `config.py`)
3. **Archivo Excel** con los datos en `fichas_a_generar/ficha_a_generar.xlsx`
4. **Plantilla Word** en `plantilla_para_generar.docx`

## 📁 Estructura del Proyecto

```
3_AltayBajaURLS/
├── config.py                    # Configuración (crear desde config.example.py)
├── config.example.py            # Plantilla de configuración
├── requirements.txt             # Dependencias de Python
├── README.md                    # Este archivo
├── GUIA_PORTABILIDAD.md         # Guía detallada de portabilidad
├── 1-duplicador_urls.py         # Duplica productos en WordPress
├── 2-generador_csv_codigos_postales.py  # Consolida datos Excel → CSV
├── 3-generador_plantillas.py    # Genera plantillas Word personalizadas
├── 4-generador_urls.py          # Genera CSV con URLs
├── 5-mie_subida.py              # Sube contenido a WordPress
├── 6-limpiador.py               # Limpia archivos generados
├── fichas_a_generar/
│   └── ficha_a_generar.xlsx     # Archivo Excel con datos
├── csvs_codigos_postales_ciudades/
│   └── csv_codigos_postales_localidades.csv  # CSV generado
├── plantillas_generadas/        # Plantillas Word generadas
└── plantilla_para_generar.docx  # Plantilla base Word
```

## 🔄 Flujo de Trabajo

1. **`1-duplicador_urls.py`** → Crea productos duplicados en WordPress
2. **`2-generador_csv_codigos_postales.py`** → Consolida datos de Excel
3. **`3-generador_plantillas.py`** → Genera plantillas Word personalizadas
4. **`4-generador_urls.py`** → Crea CSV con URLs para asociar
5. **`5-mie_subida.py`** → Sube contenido de plantillas a WordPress
6. **`6-limpiador.py`** → (Opcional) Limpia archivos generados

## ⚙️ Configuración

Todas las configuraciones están centralizadas en `config.py`:

- **Credenciales**: WordPress username/password
- **API Keys**: Google Maps API key
- **URLs**: URLs de WordPress
- **Navegador**: Edge, Chrome o Firefox
- **Rutas**: Rutas de archivos y carpetas

Ver `config.example.py` para ver todas las opciones disponibles.

## 🔒 Seguridad

- **NO** subas `config.py` a repositorios públicos (ya está en `.gitignore`)
- **NO** compartas `config.py` con credenciales reales
- Usa `config.example.py` como plantilla para compartir

## 📝 Notas Importantes

- Los scripts usan rutas relativas, funcionan en cualquier ubicación
- WebDriver Manager descarga automáticamente los drivers del navegador
- Compatible con Windows, Linux y Mac
- Requiere conexión a internet para funcionar

## 🆘 Soporte

Para problemas comunes, consulta la sección "Problemas comunes" en [GUIA_PORTABILIDAD.md](GUIA_PORTABILIDAD.md)

