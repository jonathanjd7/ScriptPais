"""
Archivo de ejemplo de configuración para los scripts del Generador MIE.
Copia este archivo como 'config.py' y completa las credenciales.

INSTRUCCIONES:
1. Copia este archivo: cp config.example.py config.py
2. Completa las credenciales en config.py
3. NO subas config.py a repositorios públicos
"""

import os

# ============================================
# CREDENCIALES DE WORDPRESS
# ============================================
WORDPRESS_USERNAME = 'TU_USUARIO_AQUI'
WORDPRESS_PASSWORD = 'TU_CONTRASEÑA_AQUI'

# ============================================
# API KEYS
# ============================================
GOOGLE_MAPS_API_KEY = 'TU_API_KEY_AQUI'

# ============================================
# URLs DE WORDPRESS
# ============================================
WORDPRESS_LOGIN_URL = 'https://mantenimientoinformaticoeconomico.com/mie2025/'
WORDPRESS_BASE_PRODUCT_URL = 'https://mantenimientoinformaticoeconomico.com/producto/reparaciones-informaticas-en-sevilla-codigo-postal-41001/'
WORDPRESS_PRODUCT_DOMAIN = 'https://mantenimientoinformaticoeconomico.com/producto/'

# ============================================
# CONFIGURACIÓN DEL NAVEGADOR
# ============================================
BROWSER = 'edge'  # 'edge', 'chrome', 'firefox'
HEADLESS_MODE = False
WINDOW_SIZE = (800, 700)

# ============================================
# RUTAS DE ARCHIVOS
# ============================================
EXCEL_FILE_PATH = './fichas_a_generar/ficha_a_generar.xlsx'
EXCEL_FOLDER_PATH = './fichas_a_generar'
CSV_FOLDER_PATH = './csvs_codigos_postales_ciudades'
CSV_OUTPUT_FILE = './csvs_codigos_postales_ciudades/csv_codigos_postales_localidades.csv'
WORD_TEMPLATE_PATH = './plantilla_para_generar.docx'
GENERATED_TEMPLATES_FOLDER = './plantillas_generadas'
URLS_CSV_FILE = './plantillas_generadas/urls.csv'

