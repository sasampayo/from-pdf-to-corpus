# ============================================================
# Buscador — From PDF to Corpus
# Sara Sampayo Sande · Centro Crímina, UMH Elche · MIT License
# https://github.com/[URL]
# ============================================================
#
# Extracción automatizada de intervenciones parlamentarias
# individuales a partir de los Diarios de Sesiones del
# Congreso de los Diputados. Diseñada para Google Colaboratory.
#
# Rendimiento validado (n=183 intervenciones):
#   Identificación orador:  κ = 0,984
#   Grupo parlamentario:    κ = 0,979
#   Extracción de texto:    κ = 0,942
#   Tipos de intervención:  κ = 0,701
#   Detección estructural:  F1 = 96,9%
#   Filtrado temático:      F1 = 87,8%
# ============================================================


# ==============================================================
# PARTE 1 — INSTALACIÓN Y CONFIGURACIÓN
# ==============================================================

# !pip install PyPDF2 pdfplumber pandas openpyxl
# !apt-get install -y python3-dev libxml2-dev libxslt1-dev antiword unrtf poppler-utils tesseract-ocr flac ffmpeg lame libmad0 libsox-fmt-mp3 sox libjpeg-dev swig

import os
import re
import pandas as pd
import pdfplumber
from google.colab import drive
from typing import List, Dict, Tuple, Optional
import numpy as np
from datetime import datetime

# ---------------------------------------------------------------
# VARIABLES CONFIGURABLES — editar antes de ejecutar
# ---------------------------------------------------------------

# Ruta a la carpeta de Google Drive con los PDFs a procesar
ENLACE_DRIVE = "/content/drive/MyDrive/RUTA_A_TU_CARPETA"  # ← Cambiar por tu ruta

# Ruta de salida del archivo Excel con los resultados
OUTPUT_PATH = "/content/drive/MyDrive/corpus_parlamentario.xlsx"  # ← Cambiar si se desea

# Palabras clave para el filtrado temático de debates
PALABRAS_CLAVE = [
    # Nombre oficial y referencia legal
    "Ley de Garantía Integral de la Libertad Sexual",
    "Ley del solo sí es sí",
    "Ley 10/2022",
    "garantía integral de la libertad sexual",

    # Conceptos centrales de la ley
    "consentimiento",
    "consentimiento sexual",
    "libertad sexual",
    "indemnidad sexual",
    "libertad e indemnidad sexuales",

    # Términos del debate parlamentario
    "ley de libertad sexual",
    "reforma de la ley de libertad sexual",
    "ley de consentimiento",

    # Términos jurídicos específicos
    "violencia sexual",
    "agresión sexual",
    "delitos sexuales",
    "abusos sexuales",
    "acoso sexual",
    "violación",
    "Protección Integral de la Libertad Sexual",

    # Casos emblemáticos
    "manada",
    "solo sí es sí",
]

# Número de archivos a procesar (None = todos)
NUMERO_ARCHIVOS_A_PROCESAR = None

print("✅ PARTE 1 CARGADA: Configuración básica")


# ==============================================================
# PARTE 1B — CONFIGURACIÓN DE MUESTRA (opcional)
# ==============================================================

import random

# Número de archivos aleatorios a procesar (0 = ninguno aleatorio)
NUMERO_ARCHIVOS_ALEATORIOS = 0

# Archivos específicos a incluir siempre en el procesamiento
# (Los 10 archivos de validación del artículo se incluyen como ejemplo)
archivos_adicionales_especificos = [
    "DSCD-12-PL-182.pdf",
    "DSCD-14-PL-231.pdf",
    "DSCD-12-CO-334.pdf",
    "DSCD-12-PL-79.pdf",
    "DSCD-12-PL-147.pdf",
    "DSCD-14-PL-130.pdf",
    "DSCD-14-PL-227.pdf",
    "DSCD-15-PL-64.PDF",
    "DSCD-15-PL-70.PDF",
    "DSCD-14-PL-204.pdf",
]

# Obtener lista de PDFs disponibles
pdf_files = [f for f in os.listdir(ENLACE_DRIVE) if f.lower().endswith('.pdf')]

# Excluir los específicos de la selección aleatoria para evitar duplicados
pdf_files_filtered = [f for f in pdf_files if f not in archivos_adicionales_especificos]
random.shuffle(pdf_files_filtered)

if NUMERO_ARCHIVOS_ALEATORIOS is None:
    archivos_a_procesar_aleatorios = pdf_files_filtered
else:
    archivos_a_procesar_aleatorios = pdf_files_filtered[:NUMERO_ARCHIVOS_ALEATORIOS]

# Añadir archivos específicos (solo los que existen en el directorio)
archivos_adicionales_existentes = [f for f in archivos_adicionales_especificos if f in pdf_files]
archivos_a_procesar_aleatorios.extend(archivos_adicionales_existentes)
archivos_a_procesar_aleatorios = list(dict.fromkeys(archivos_a_procesar_aleatorios))

print(f"Lista de {len(pdf_files)} archivos PDF encontrados.")
print(f"Se procesarán {len(archivos_a_procesar_aleatorios)} archivos.")


# ==============================================================
# PARTE 2 — FUNCIONES AUXILIARES BÁSICAS
# ==============================================================

def buscar_numero_expediente(texto_completo, titulo_debate=""):
    """Búsqueda contextual del número de expediente."""
    if not texto_completo:
        return ""

    print("🔍 Buscando número de expediente contextual...")

    patron_expediente = r'\(Número de expediente\s+(\d{3}/\d{6})\)'
    todas_coincidencias = []
    for match in re.finditer(patron_expediente, texto_completo):
        numero = match.group(1)
        posicion = match.start()
        todas_coincidencias.append((numero, posicion))
        print(f"      📍 Expediente encontrado: {numero} en posición {posicion}")

    if len(todas_coincidencias) == 1:
        numero = todas_coincidencias[0][0]
        print(f"      ✅ ÚNICO expediente encontrado: {numero}")
        return numero

    elif len(todas_coincidencias) > 1:
        print(f"      ⚠️  MÚLTIPLES expedientes: {[num for num, pos in todas_coincidencias]}")
        lineas = texto_completo.split('\n')
        for i, linea in enumerate(lineas[:20]):
            for numero, posicion in todas_coincidencias:
                if numero in linea:
                    print(f"      ✅ Expediente seleccionado (línea {i}): {numero}")
                    return numero
        primer_numero = todas_coincidencias[0][0]
        print(f"      ✅ Usando primer expediente: {primer_numero}")
        return primer_numero

    patrones_flexibles = [
        r'\(Número de expediente\s+(\d+/\d+)\)',
        r'\(Núm\. de expediente\s+(\d+/\d+)\)',
        r'expediente\s+(\d+/\d+)'
    ]
    for patron in patrones_flexibles:
        matches = re.findall(patron, texto_completo)
        if matches:
            print(f"      ✅ Expediente encontrado (patrón flexible): {matches[0]}")
            return matches[0]

    print(f"      ❌ NO se encontró número de expediente")
    return ""


def extraer_palabras_clave_titulo(titulo):
    """Extrae palabras clave relevantes del título para búsqueda contextual."""
    palabras_comunes = {'de', 'la', 'el', 'y', 'en', 'por', 'para', 'del', 'las', 'los', 'se', 'una', 'un'}
    palabras = re.findall(r'\b[a-zA-Z]{4,}\b', titulo.lower())
    return [p for p in palabras if p not in palabras_comunes][:5]


def extraer_texto_debate_confiable(pdf_path, pagina_inicio, paginas_total=20):
    """Extrae texto confiable de un debate por páginas."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto_completo = ""
            start_page = max(0, pagina_inicio - 1)
            end_page = min(len(pdf.pages), start_page + paginas_total)
            print(f"      📄 Extrayendo páginas {pagina_inicio} a {pagina_inicio + paginas_total}")
            for page_num in range(start_page, end_page):
                pagina = pdf.pages[page_num]
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"
            return texto_completo
    except Exception as e:
        print(f"      ❌ Error extrayendo texto: {e}")
        return ""

print("✅ PARTE 2 CARGADA: Funciones auxiliares básicas")


# ==============================================================
# PARTE 3 — FUNCIONES DE LIMPIEZA Y PREPROCESAMIENTO
# ==============================================================

def limpiar_cabeceras_documento_mejorado(lineas, archivo):
    """Elimina cabeceras de página más efectivamente."""
    lineas_limpias = []
    patrones_cabecera = [
        r'^\d+-\w+-\d+-\w+$',
        r'^:\w+$',
        r'DIARIO DE SESIONES DEL CONGRESO DE LOS DIPUTADOS',
        r'PLENO Y DIPUTACIÓN PERMANENTE',
        r'COMISIONES',
        r'Pág\. \d+',
        r'Núm\. \d+',
        r'Año \d+',
        r'LEGISLATURA'
    ]
    for linea in lineas:
        linea_limpia = linea.strip()
        es_cabecera = any(re.search(patron, linea_limpia) for patron in patrones_cabecera)
        if (len(linea_limpia) < 5 and linea_limpia.isdigit()) or es_cabecera:
            continue
        lineas_limpias.append(linea)
    return lineas_limpias


def extraer_texto_intervencion_sin_cortes(lineas, indice_inicio):
    """Extrae texto de intervención sin cortes por cabeceras."""
    texto = []
    i = indice_inicio
    if i < len(lineas):
        texto.append(lineas[i])
        i += 1
    while i < len(lineas):
        linea_actual = lineas[i].strip()
        if re.match(r'^\d+-\w+-\d+-\w+$', linea_actual) or re.match(r'^:\w+$', linea_actual):
            i += 1
            continue
        if re.match(r'^(?:El señor|La señora)\s+[A-ZÁÉÍÓÚÑ]', linea_actual):
            if len(linea_actual) < 150:
                break
        if (linea_actual.startswith('—') or
            (linea_actual.startswith('(') and any(p in linea_actual.lower() for p in ['aplausos', 'ovación'])) or
            '..........' in linea_actual or
            '................' in linea_actual):
            break
        texto.append(lineas[i])
        i += 1
    resultado = '\n'.join(texto)
    print(f"        📝 Texto extraído: {len(resultado)} caracteres, {i - indice_inicio} líneas")
    return resultado, i


def limpiar_texto_intervencion_final(texto):
    """Limpia el texto final de frases no deseadas."""
    lineas = texto.split('\n')
    lineas_limpias = []
    patrones_final_no_deseado = [
        r'^\d+-\w+-\d+-\w+$',
        r'^:\w+$',
        r'DIARIO DE SESIONES',
        r'CONGRESO DE LOS DIPUTADOS',
        r'PLENO Y DIPUTACIÓN PERMANENTE',
        r'COMISIONES',
        r'Pág\. \d+'
    ]
    for linea in lineas:
        linea_limpia = linea.strip()
        es_no_deseado = any(re.search(patron, linea_limpia) for patron in patrones_final_no_deseado)
        if not es_no_deseado:
            lineas_limpias.append(linea)
        else:
            break
    return '\n'.join(lineas_limpias)

print("✅ PARTE 3 CARGADA: Funciones de limpieza y preprocesamiento")


# ==============================================================
# PARTE 4 — DICCIONARIOS DE ORADORES POR LEGISLATURA
# ==============================================================

# LEGISLATURA CONSTITUYENTE (1977-1979)
DICCIONARIO_CONSTITUYENTE = {
    'BECERRIL BUSTAMANTE': 'UCD',
    'BRABO CASTELLS': 'Comunista',
    'BUSTELO GARCÍA DEL REAL': 'Socialista',
    'CALVET PUIG': 'Comunista',
    'CASTRO GARCÍA': 'Socialista',
    'CRUAÑES MOLINA': 'Socialista',
    'FERNÁNDEZ-ESPAÑA Y FERNÁNDEZ-LATORRE': 'Alianza Popular',
    'GARCÍA BLOISE': 'Socialista',
    'IBÁRRURI GÓMEZ': 'Comunista',
    'IZQUIERDO ROJO': 'Socialista',
    'LAJO PÉREZ': 'Socialistes de Catalunya',
    'MATA I GARRIGA': 'Socialistes de Catalunya',
    'MOLL DE MIGUEL': 'UCD',
    'MORENAS AY DILLO': 'UCD',
    'MORENO GONZÁLEZ': 'UCD',
    'PLA PECHOVERTO': 'Socialista',
    'REVILLA LÓPEZ': 'UCD',
    'RUIZ-TAGLE MORALES': 'Socialista',
    'SABATER LLORENS': 'Socialista',
    'TELLADO ALFONSO': 'UCD',
    'VILARIÑO SALGADO': 'UCD',
}

# I LEGISLATURA (1979-1982)
DICCIONARIO_I_LEGISLATURA = {
    'ARAHUETES PORTERO': 'Centrista',
    'ARCE MOLINA': 'Centrista',
    'BALLETBÓ PUIG': 'Socialistes de Catalunya',
    'BECERRIL BUSTAMANTE': 'Centrista',
    'BRABO CASTELLS': 'Mixto',
    'CRUAÑES MOLINA': 'Socialista',
    'FERNÁNDEZ-ESPAÑA Y FERNÁNDEZ-LATORRE': 'Coalición Democrática',
    'GARCÍA ARIAS': 'Socialista',
    'GARCÍA BLOISE': 'Socialista',
    'GARCÍA-MORENO TEIXEIRA': 'Mixto',
    'IZQUIERDO ROJO': 'Socialista',
    'LAFUENTE ORIVE': 'Centrista',
    'MATA I GARRIGA': 'Socialistes de Catalunya',
    'MORENAS AY DILLO': 'Centrista',
    'MORENO GONZÁLEZ': 'Centrista',
    'PELAYO DUQUE': 'Mixto',
    'PLA PASTOR': 'Socialista',
    'REVILLA LÓPEZ': 'Centrista',
    'RUBIES GARROFE': 'Minoría Catalana',
    'RUIZ-TAGLE MORALES': 'Socialista',
    'SOLANO CARRERAS': 'Mixto',
    'VÁZQUEZ MENÉNDEZ': 'Socialista',
    'VILARIÑO SALGADO': 'Centrista',
    'VINTRO CASTELLS': 'Comunista',
}

# II LEGISLATURA (1982-1986)
DICCIONARIO_II_LEGISLATURA = {
    'ABASCAL Y CALABRIA': 'Socialista',
    'BALLETBÓ PUIG': 'Socialista',
    'BERRUEZO ALBÉNIZ': 'Socialista',
    'CAMPO CASASÚS': 'Socialista',
    'CRUAÑES MOLINA': 'Socialista',
    'CUNILLERA I MESTRES': 'Socialista',
    'FERNÁNDEZ-ESPAÑA Y FERNÁNDEZ-LATORRE': 'Mixto',
    'GARCÍA BLOISE': 'Socialista',
    'GARCÍA-MORENO TEIXEIRA': 'Socialista',
    'GORROÑO ARRIZABALAGA': 'PNV',
    'HERMOSÍN BONO': 'Socialista',
    'LLORCA VILLAPLANA': 'Popular',
    'MADRE ORTEGA': 'Socialista',
    'PELAYO DUQUE': 'Socialista',
    'PINEDO SÁNCHEZ': 'Socialista',
    'PLA PASTOR': 'Socialista',
    'RENAU I MANÉN': 'Socialista',
    'SIMÓN CALVO': 'Socialista',
    'SOLANO CARRERAS': 'Socialista',
    'VALLS BERTRAND': 'Popular',
    'VÁZQUEZ MENÉNDEZ': 'Socialista',
    'VERDÚ ALONSO': 'Socialista',
    'VILLACIÁN PEÑALOSA': 'PNV',
}

# III LEGISLATURA (1986-1989)
DICCIONARIO_III_LEGISLATURA = {
    'AIZPURÚA EGAÑA': 'Mixto',
    'AROZ IBÁÑEZ': 'Socialista',
    'BALLETBÓ PUIG': 'Socialista',
    'BANZO AMAT': 'Coalición Popular',
    'CAMPO CASASÚS': 'Socialista',
    'CRUAÑES MOLINA': 'Socialista',
    'CUENCA I VALERO': 'Minoría Catalana',
    'ESTEVAN BOLEA': 'Coalición Popular',
    'FERNÁNDEZ LABRADOR': 'Coalición Popular',
    'GARCÍA BLOISE': 'Socialista',
    'GARCÍA BOTÍN': 'Coalición Popular',
    'HERMOSÍN BONO': 'Socialista',
    'IZQUIERDO ARIJA': 'Coalición Popular',
    'IZQUIERDO ROJO': 'Socialista',
    'JUAN MILLET': 'Socialista',
    'MADRE ORTEGA': 'Socialista',
    'MORSO PÉREZ': 'CDS',
    'NOVOA CARCACIA': 'Socialista',
    'PELAYO DUQUE': 'Socialista',
    'PINEDO SÁNCHEZ': 'Socialista',
    'PLA PASTOR': 'Socialista',
    'PRIETO MORENO': 'Socialista',
    'RODRÍGUEZ ORTEGA': 'Socialista',
    'RUDI ÚBEDA': 'Coalición Popular',
    'SALARRULLANA DE VERDA': 'Mixto',
    'SÁNCHEZ LÓPEZ': 'Socialista',
    'SANTOS SÁNCHEZ': 'Socialista',
    'SIMÓN CALVO': 'Socialista',
    'TOCINO BISCAROLASAGA': 'Coalición Popular',
    'UGALDE RUIZ DE ASSIN': 'Coalición Popular',
    'VILLALOBOS TALERO': 'Coalición Popular',
    'VISIEDO NIETO': 'Socialista',
    'YABAR STERLING': 'CDS',
}

# IV LEGISLATURA (1989-1993)
DICCIONARIO_IV_LEGISLATURA = {
    'AIZPURÚA EGAÑA': 'Mixto',
    'ALBERDI ALONSO': 'Socialista',
    'ALMEIDA CASTRO': 'IU-IC',
    'AROZ IBÁÑEZ': 'Socialista',
    'BALLETBÓ PUIG': 'Socialista',
    'BECERRIL BUSTAMANTE': 'Popular',
    'BLÁZQUEZ MARTÍNEZ': 'Socialista',
    'BRAZOVISODO': 'Socialista',
    'CAMPO CASASÚS': 'Socialista',
    'CAMPO PIÑEIRO': 'Popular',
    'CASTILLA DEL PINO': 'Socialista',
    'CONDE GUTIÉRREZ DEL ALAMO': 'Socialista',
    'CONTRERAS VILLAR': 'Socialista',
    'CRUAÑES MOLINA': 'Socialista',
    'CUENCA I VALERO': 'CiU',
    'ESTEVAN BOLEA': 'Popular',
    'FERNÁNDEZ SANZ': 'Socialista',
    'FRÍAS NAVARRETE': 'Socialista',
    'GARCÍA BLOISE': 'Socialista',
    'GARCÍA MANZANARES': 'Socialista',
    'GARCÍA-ALCAÑIZ CALVO': 'Popular',
    'GARMENDIA GALBETE': 'Mixto',
    'GERMÁN LAGUNA': 'Socialista',
    'HERMOSÍN BONO': 'Socialista',
    'IZQUIERDO ARIJA': 'Popular',
    'JUAN MILLET': 'Socialista',
    'LARRAÑAGA GALDÓS': 'Mixto',
    'MAESTRO MARTÍN': 'IU-IC',
    'MARTÍNEZ SAIZ': 'Popular',
    'MENDIZÁBAL GOROSTIAGA': 'Mixto',
    'MORAGA FERRÁNDIZ': 'Socialista',
    'MORSO PÉREZ': 'CDS',
    'NOVOA CARCACIA': 'Socialista',
    'ORTEGA PEINADO': 'Socialista',
    'PALACIO DEL VALLE-LERSUNDI': 'Popular',
    'PARDO ORTIZ': 'Socialista',
    'PELAYO DUQUE': 'Socialista',
    'PEREIRA SANTANA': 'Socialista',
    'PLA PASTOR': 'Socialista',
    'PLEGUEZUELOS AGUILAR': 'Socialista',
    'RENAU I MANÉN': 'Socialista',
    'RODRÍGUEZ CALVO': 'Socialista',
    'RODRÍGUEZ ORTEGA': 'Socialista',
    'ROMERO LÓPEZ': 'Socialista',
    'RUDI UBEDA': 'Popular',
    'SAINZ GARCÍA': 'Popular',
    'SÁNCHEZ LÓPEZ': 'Socialista',
    'SEMPER E JAÉN': 'Socialista',
    'SIMÓN CALVO': 'Socialista',
    'TOCINO BISCAROLASAGA': 'Popular',
    'VEGA RAMÓN': 'Socialista',
    'VICENTE GARCÍA': 'Socialista',
    'VILLALOBOS TALERO': 'Popular',
    'VISIEDO NIETO': 'Socialista',
}

# V LEGISLATURA (1993-1996)
DICCIONARIO_V_LEGISLATURA = {
    'ABAD PINILLOS': 'Socialista',
    'AGUILAR RIVERO': 'IU-IC',
    'ALBERDI ALONSO': 'Socialista',
    'ALEMANY I ROCA': 'CiU',
    'ÁLVAREZ GAYOL': 'Socialista',
    'AMADOR GUILLÉN': 'Popular',
    'AROZ IBÁÑEZ': 'Socialista',
    'BALLESTEROS BELINCHÓN': 'Socialista',
    'BALLETBÓ PUIG': 'Socialista',
    'BARRIOS CURBELO': 'Popular',
    'BECERRIL BUSTAMANTE': 'Popular',
    'CALDERÓN PÉREZ': 'Popular',
    'CAMPO CASASÚS': 'Socialista',
    'CAMPO PIÑEIRO': 'Popular',
    'CAVA DE LLANO Y CARRIÓ': 'Popular',
    'CONDE GUTIÉRREZ DEL ALAMO': 'Socialista',
    'CREMADES GRIÑÁN': 'Popular',
    'DÍAZ VILLANUEVA': 'Popular',
    'FERNÁNDEZ DÍAZ': 'Popular',
    'FERNÁNDEZ GONZÁLEZ': 'Popular',
    'FERNÁNDEZ SANZ': 'Socialista',
    'FERNÁNDEZ-CAPEL BAÑOS': 'Popular',
    'FIGUERAS I SIÑOL': 'Socialista',
    'FRÍAS NAVARRETE': 'Socialista',
    'GARCÍA BLOISE': 'Socialista',
    'GARCÍA MANZANARES': 'Socialista',
    'GARCÍA VILLAMAYOR': 'Socialista',
    'GARCÍA-ALCAÑIZ CALVO': 'Popular',
    'GÓMEZ GARCÍA': 'Socialista',
    'GOROSTIAGA SAIZ': 'Socialista',
    'MAESTRO MARTÍN': 'IU-IC',
    'MARTÍNEZ SAIZ': 'Popular',
    'MATO ADROVER': 'Popular',
    'MENDIZÁBAL GOROSTIAGA': 'Socialista',
    'MONTSENY MASIP': 'Popular',
    'MONZÓN SUÁREZ': 'CC',
    'MORENO GONZÁLEZ': 'Socialista',
    'NOVOA CARCACIA': 'Socialista',
    'OÑA SEVILLA': 'Popular',
    'PALACIO DEL VALLE-LERSUNDI': 'Popular',
    'PARDO ORTIZ': 'Socialista',
    'PARDO RAGA': 'Popular',
    'PELAYO DUQUE': 'Socialista',
    'PELLICER RASO': 'Socialista',
    'PÉREZ VEGA': 'Popular',
    'PLA PASTOR': 'Socialista',
    'PULGAR FRAILE': 'Popular',
    'RAHOLA I MARTÍNEZ': 'Mixto',
    'RIVADULLA GRACIA': 'IU-IC',
    'RODRÍGUEZ CALVO': 'Socialista',
    'RODRÍGUEZ ORTEGA': 'Socialista',
    'ROMACHO ROMERO': 'Socialista',
    'ROMERO LÓPEZ': 'Socialista',
    'ROMERO MARTÍNEZ': 'Socialista',
    'RUBIALES TORREJÓN': 'Socialista',
    'RUDI UBEDA': 'Popular',
    'SAINZ GARCÍA': 'Popular',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ LÓPEZ': 'Socialista',
    'SEMPER E JAÉN': 'Socialista',
    'SOLER NOMDEDEU': 'Socialista',
    'TOCINO BISCAROLASAGA': 'Popular',
    'URÁN GONZÁLEZ': 'IU-IC',
    'VICENTE GARCÍA': 'Socialista',
    'VILLALOBOS TALERO': 'Popular',
}

# VI LEGISLATURA (1996-2000)
DICCIONARIO_VI_LEGISLATURA = {
    'AGUILAR RIVERO': 'IU',
    'AGUIRRE URIBE': 'EAJ-PNV',
    'ALBERDI ALONSO': 'Socialista',
    'ALBORCH BATALLER': 'Socialista',
    'ALEDO MARTÍNEZ': 'Socialista',
    'ALMEIDA CASTRO': 'Mixto',
    'ÁLVAREZ GAYOL': 'Socialista',
    'AMADOR GUILLÉN': 'Popular',
    'AMADOR MILLÁN': 'Socialista',
    'AMORÓS I SANS': 'CiU',
    'ARAMBURU DEL RÍO': 'IU',
    'AROZ IBÁÑEZ': 'Socialista',
    'ATIENZA I GUERRERO': 'CiU',
    'BABIANO LÓPEZ': 'Socialista',
    'BALLESTEROS BELINCHÓN': 'Socialista',
    'BALLETBÓ PUIG': 'Socialista',
    'BARRIOS CURBELO': 'Popular',
    'BARTOLOMÉ NÚÑEZ': 'Socialista',
    'CALDERÓN PÉREZ': 'Popular',
    'CALLEJA DE PABLO': 'Socialista',
    'CAMACHO VÁZQUEZ': 'Socialista',
    'CAMILLERI HERNÁNDEZ': 'Popular',
    'CAMPO CASASÚS': 'Socialista',
    'CÁNOVAS MONTALBÁN': 'Socialista',
    'CASTRO MASAVEU': 'Popular',
    'CAVA DE LLANO Y CARRIÓ': 'Popular',
    'CONDE GUTIÉRREZ DEL ALAMO': 'Socialista',
    'CORTAJARENA ITURRIOZ': 'Socialista',
    'CUNILLERA I MESTRES': 'Socialista',
    'DÍEZ DE BALDEÓN GARCÍA': 'Socialista',
    'DÍEZ DE LA LASTRA BARBADILLO': 'Popular',
    'DÍEZ LÓPEZ': 'Popular',
    'FARALDO BOTANA': 'Popular',
    'FERNÁNDEZ DE LA VEGA SANZ': 'Socialista',
    'FERNÁNDEZ GONZÁLEZ': 'Popular',
    'FERNÁNDEZ RAMIRO': 'Socialista',
    'FERNÁNDEZ SANZ': 'Socialista',
    'FERNÁNDEZ-CAPEL BAÑOS': 'Popular',
    'FRÍAS NAVARRETE': 'Socialista',
    'GARCÍA LINARES': 'Socialista',
    'GARCÍA MANZANARES': 'Socialista',
    'GARCÍA-ALCAÑIZ CALVO': 'Popular',
    'GARCÍA-HIERRO CARABALLO': 'Socialista',
    'GIL I MIRÓ': 'CiU',
    'HERAS PABLO': 'Socialista',
    'HERNÁNDEZ ROZAS': 'Popular',
    'LARA CARBÓ': 'Popular',
    'LASAGABASTER OLAZÁBAL': 'Mixto',
    'LEIVA DÍEZ': 'Socialista',
    'LÓPEZ I CHAMOSA': 'Socialista',
    'MAESTRO MARTÍN': 'IU',
    'MARÓN BELTRÁN': 'Socialista',
    'MARTÍN CREVILLÉN': 'Socialista',
    'MARTÍNEZ GONZÁLEZ': 'Socialista',
    'MARTÍNEZ SAIZ': 'Popular',
    'MATADOR DE MATOS': 'Popular',
    'MATO ADROVER': 'Popular',
    'MENDIZÁBAL GOROSTIAGA': 'Socialista',
    'MONEO DÍEZ': 'Popular',
    'MONTES CONTRERAS': 'Socialista',
    'MONTSENY MASIP': 'Popular',
    'MORA DEVIS': 'Popular',
    'MULET TORRES': 'Socialista',
    'MUÑOZ SANTAMARÍA': 'Socialista',
    'NARBONA RUIZ': 'Socialista',
    'NOVOA CARCACIA': 'Socialista',
    'PALACIO DEL VALLE-LERSUNDI': 'Popular',
    'PALMA I MUÑOZ': 'Socialista',
    'PARDO RAGA': 'Popular',
    'PÉREZ VEGA': 'Popular',
    'PIN ARBOLEDAS': 'Socialista',
    'POZUELO MEÑO': 'Socialista',
    'PULGAR FRAILE': 'Popular',
    'RAHOLA I MARTÍNEZ': 'Mixto',
    'RIERA I BEN': 'CiU',
    'RIERA MADURELL': 'Socialista',
    'RIVADULLA GRACIA': 'Mixto',
    'RODRÍGUEZ-SALMONES CABEZA': 'Popular',
    'ROMERO LÓPEZ': 'Socialista',
    'RUBIALES TORREJÓN': 'Socialista',
    'RUIZ SAAVEDRA': 'Socialista',
    'SABANÉS NADAL': 'IU',
    'SAINZ GARCÍA': 'Popular',
    'SALINAS GARCÍA': 'Socialista',
    'SÁNCHEZ GARCÍA': 'Popular',
    'SÁNCHEZ LÓPEZ': 'Socialista',
    'SELLER ROCA DE TOGORES': 'Popular',
    'SILVA REGO': 'Socialista',
    'SOLSONA I PIÑOL': 'CiU',
    'TOCINO BISCAROLASAGA': 'Popular',
    'TORME PARDO': 'Popular',
    'URÁN GONZÁLEZ': 'IU',
    'URÍA ETXEBARRÍA': 'EAJ-PNV',
    'VALCARCE GARCÍA': 'Socialista',
    'VARELA VÁZQUEZ': 'Socialista',
    'VÁZQUEZ PÉREZ': 'Socialista',
    'VILLALOBOS TALERO': 'Popular',
    'VILLAR JAR': 'Popular',
}

# VII LEGISLATURA (2000-2004)
DICCIONARIO_VII_LEGISLATURA = {
    'AGUADO DEL OLMO': 'Popular',
    'ALBERDI ALONSO': 'Mixto',
    'ALBORCH BATALLER': 'Socialista',
    'AMADOR MILLÁN': 'Socialista',
    'ARÉVALO ARAYA': 'Popular',
    'ARRÚE BERGARECHE': 'Popular',
    'BÁÑEZ GARCÍA': 'Popular',
    'BARRIOS CURBELO': 'Popular',
    'BECERRIL BUSTAMANTE': 'Popular',
    'BLANCO TERÁN': 'Socialista',
    'CAMACHO VÁZQUEZ': 'Socialista',
    'CAMARERO BENÍTEZ': 'Popular',
    'CAMPO CASASÚS': 'Socialista',
    'CARABEL PEDREIRA': 'Popular',
    'CARACUEL DEL OLMO': 'Popular',
    'CASTELLANO RODRÍGUEZ': 'Popular',
    'CASTRO FONSECA': 'IU',
    'CASTRO MASAVEU': 'Popular',
    'CAVA DE LLANO Y CARRIÓ': 'Popular',
    'CHACÓN PIQUERAS': 'Socialista',
    'COMPTE LLUSA': 'Popular',
    'CONDE GUTIÉRREZ DEL ALAMO': 'Socialista',
    'CORRES VAQUERO': 'Popular',
    'CORTAJARENA ITURRIOZ': 'Socialista',
    'COSTA CAMPI': 'Socialista',
    'COSTAS MANZANARES': 'Popular',
    'CRUZ VALENTÍN': 'Socialista',
    'CUNILLERA I MESTRES': 'Socialista',
    'DANCAUSA TREVIÑO': 'Popular',
    'DÍEZ DE BALDEÓN GARCÍA': 'Socialista',
    'DÍEZ DE LA LASTRA BARBADILLO': 'Popular',
    'DURÁN SÁNCHEZ': 'Socialista',
    'ESPINOSA LÓPEZ': 'Popular',
    'ESTARÁS FERRAGUT': 'Popular',
    'EXPÓSITO MOLINA': 'CiU',
    'FARALDO BOTANA': 'Popular',
    'FEBRER SANTANDREU': 'Popular',
    'FERNÁNDEZ DE LA VEGA SANZ': 'Socialista',
    'FERNÁNDEZ GONZÁLEZ': 'Popular',
    'FERNÁNDEZ-CAPEL BAÑOS': 'Popular',
    'FERRANDO SENDRA': 'Popular',
    'GALLIZO LLAMAS': 'Socialista',
    'GARCÍA MANZANARES': 'Socialista',
    'GARCÍA PÉREZ': 'Socialista',
    'GARCÍA-ALCAÑIZ CALVO': 'Popular',
    'GARCÍA-HIERRO CARABALLO': 'Socialista',
    'GIL LÓPEZ': 'Socialista',
    'GOROSTIAGA SAIZ': 'Socialista',
    'GORRI GIL': 'Popular',
    'GRACIA JIMÉNEZ': 'Popular',
    'GUARINOS LÓPEZ': 'Popular',
    'GUERRA GALVÁN': 'Popular',
    'HERNANSANZ RUIZ': 'Popular',
    'JUANEDA ZARAGOZA': 'Socialista',
    'JULIOS REYES': 'CC',
    'LARA CARBÓ': 'Popular',
    'LASAGABASTER OLAZÁBAL': 'Mixto',
    'LÓPEZ GONZÁLEZ': 'Socialista',
    'LÓPEZ I CHAMOSA': 'Socialista',
    'MARISCAL DE GANTE MIRÓN': 'Popular',
    'MARÓN BELTRÁN': 'Socialista',
    'MARTÍN VIGIL': 'Socialista',
    'MARTÍN VIVAS': 'Popular',
    'MARTÍNEZ CERVERA': 'Popular',
    'MARTORELL PALLÁS': 'Popular',
    'MARZAL MARTÍNEZ': 'Socialista',
    'MATADOR DE MATOS': 'Popular',
    'MATO ADROVER': 'Popular',
    'MENDIZÁBAL GOROSTIAGA': 'Socialista',
    'MIQUEL SERDÁ': 'Mixto',
    'MIRALLES I GUASCH': 'Socialista',
    'MONEO DÍEZ': 'Popular',
    'MONTELONGO GONZÁLEZ': 'Popular',
    'MONTSENY MASIP': 'Popular',
    'MONZÓN CABRERA': 'CC',
    'MORENO SIRODEY': 'Socialista',
    'MUÑOZ SANTAMARÍA': 'Socialista',
    'MUÑOZ URIOL': 'Popular',
    'NAVARRO GARZÓN': 'Socialista',
    'NESTARES GARCÍA-TREVIJANO': 'Popular',
    'OLMEDO CHECA': 'Socialista',
    'OLTRA TORRES': 'Popular',
    'PAJÍN IRAOLA': 'Socialista',
    'PALMA I MUÑOZ': 'Socialista',
    'PASTOR JULIÁN': 'Popular',
    'PÉREZ BRITO': 'Popular',
    'PÉREZ DOMÍNGUEZ': 'Socialista',
    'PERIS CERVERA': 'Socialista',
    'PIGEM I PALMÉS': 'CiU',
    'PIN ARBOLEDAS': 'Socialista',
    'PISONERO RUIZ': 'Popular',
    'PLEGUEZUELOS AGUILAR': 'Socialista',
    'POL CABRER': 'Popular',
    'POZUELO MEÑO': 'Socialista',
    'QUINTANILLA BARBA': 'Popular',
    'REYES MIRANDA': 'Popular',
    'RIERA I BEN': 'CiU',
    'RIERA I REÑÉ': 'CiU',
    'RIERA MADURELL': 'Socialista',
    'RODRÍGUEZ CALLAO': 'Popular',
    'RODRÍGUEZ DÍAZ': 'Socialista',
    'RODRÍGUEZ LÓPEZ': 'Popular',
    'RODRÍGUEZ-SALMONES CABEZA': 'Popular',
    'ROGADO HERNÁNDEZ': 'Popular',
    'ROMERO LÓPEZ': 'Socialista',
    'ROMERO SÁNCHEZ': 'Popular',
    'RUBIALES TORREJÓN': 'Socialista',
    'RUDI ÚBEDA': 'Popular',
    'RUIZ RUIZ': 'Popular',
    'RUMÍ IBÁÑEZ': 'Socialista',
    'SAGARNA ALBERDI': 'Popular',
    'SAINZ GARCÍA': 'Popular',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ GARCÍA': 'Popular',
    'SELLER ROCA DE TOGORES': 'Popular',
    'SERNA MASIÁ': 'Socialista',
    'TOCINO BISCAROLASAGA': 'Popular',
    'TORME PARDO': 'Popular',
    'TORRADO REY': 'Socialista',
    'UNZURRUNZAGA CAMPOY': 'Popular',
    'URÁN GONZÁLEZ': 'IU',
    'URÍA ETXEBARRÍA': 'EAJ-PNV',
    'VALCARCE GARCÍA': 'Socialista',
    'VALENTÍN NAVARRO': 'Socialista',
    'VARELA VÁZQUEZ': 'Socialista',
    'VÁZQUEZ BLANCO': 'Popular',
    'VILLALOBOS TALERO': 'Popular',
    'VILLAR JAR': 'Popular',
    'ZARAGOZA JUNCÁ': 'Popular',
}

# VIII LEGISLATURA (2004-2008)
DICCIONARIO_VIII_LEGISLATURA = {
    'ABURTO BASELGA': 'Socialista',
    'ALBORCH BATALLER': 'Socialista',
    'ALCÁZAR ESCRIBANO': 'Socialista',
    'ÁLVAREZ ARZA': 'Socialista',
    'ÁLVAREZ OTEO': 'Socialista',
    'ARMENGOL CRIADO': 'Socialista',
    'ARNAIZ GARCÍA': 'Socialista',
    'ARRÚE BERGARECHE': 'Popular',
    'BÁÑEZ GARCÍA': 'Popular',
    'BARKOS BERRUEZO': 'Mixto',
    'BATET LAMAÑA': 'Socialista',
    'BLANCO TERÁN': 'Socialista',
    'BONÀS PAHISA': 'ERC',
    'CABRERA CALVO-SOTELO': 'Socialista',
    'CALVO POYATO': 'Socialista',
    'CAMARERO BENÍTEZ': 'Popular',
    'CAÑIGUERAL OLIVÉ': 'ERC',
    'CARACUEL DEL OLMO': 'Popular',
    'CARCEDO ROCES': 'Socialista',
    'CASAUS RODRÍGUEZ': 'Socialista',
    'CASTELLANO RODRÍGUEZ': 'Popular',
    'CASTILLEJO HERNÁNDEZ': 'Socialista',
    'CASTILLO VERA': 'Popular',
    'CASTRO MASAVEU': 'Popular',
    'CEDRÉS RODRÍGUEZ': 'Socialista',
    'CHACÓN PIQUERAS': 'Socialista',
    'COELLO FERNÁNDEZ-TRUJILLO': 'Socialista',
    'COLLDEFORNS I SOL': 'Socialista',
    'CORRAL RUIZ': 'Socialista',
    'CORTAJARENA ITURRIOZ': 'Socialista',
    'COUTO RIVAS': 'Socialista',
    'CRUZ VALENTÍN': 'Socialista',
    'CUNILLERA I MESTRES': 'Socialista',
    'DÍAZ PACHECO': 'Socialista',
    'DÍEZ DE BALDEÓN GARCÍA': 'Socialista',
    'ELÍAS CORDÓN': 'Socialista',
    'ESCUDERO SÁNCHEZ': 'Socialista',
    'ESTEVE ORTEGA': 'Socialista',
    'FARRERA GRANJA': 'Socialista',
    'FERNÁNDEZ DAVILA': 'Mixto',
    'FERNÁNDEZ DE LA VEGA SANZ': 'Socialista',
    'FERNÁNDEZ-CAPEL BAÑOS': 'Popular',
    'FERRANDO SENDRA': 'Popular',
    'FONT BONMATÍ': 'Popular',
    'FUENTES GONZÁLEZ': 'Socialista',
    'FUENTES PACHECO': 'Socialista',
    'GARCÍA SUÁREZ': 'IU-ICV',
    'GARCÍA VALLS': 'Socialista',
    'GARCÍA-ALCAÑIZ CALVO': 'Popular',
    'GARCÍA-HIERRO CARABALLO': 'Socialista',
    'GARCÍA-VALDECASAS SALGADO': 'Popular',
    'GÓMEZ SANTAMARÍA': 'Socialista',
    'GONZÁLEZ GUTIÉRREZ': 'Popular',
    'GONZÁLEZ SEGURA': 'Popular',
    'GRANDE PESQUERO': 'Socialista',
    'HERMOSÍN BONO': 'Socialista',
    'HERRERO SAINZ-ROZAS': 'Socialista',
    'HOLGADO FLORES': 'Socialista',
    'JUANEDA ZARAGOZA': 'Socialista',
    'JUANES BARCIELA': 'Socialista',
    'LARA CARBÓ': 'Popular',
    'LASAGABASTER OLAZÁBAL': 'Mixto',
    'LIZARRAGA GISBERT': 'Socialista',
    'LOPE FONTAGNE': 'Popular',
    'LÓPEZ I CHAMOSA': 'Socialista',
    'LÓPEZ RODRÍGUEZ': 'Socialista',
    'MADRAZO DÍAZ': 'Popular',
    'MALARET GARCÍA': 'Socialista',
    'MARÓN BELTRÁN': 'Socialista',
    'MARTEL GÓMEZ': 'Socialista',
    'MARTÍN MENDIZÁBAL': 'Popular',
    'MARTÍNEZ HIGUERAS': 'Socialista',
    'MATADOR DE MATOS': 'Popular',
    'MATO ADROVER': 'Popular',
    'MÉNDEZ MONASTERIO': 'Popular',
    'MENDIZÁBAL GOROSTIAGA': 'Socialista',
    'MONEO DÍEZ': 'Popular',
    'MONTESERÍN RODRÍGUEZ': 'Socialista',
    'MONTESINOS DE MIGUEL': 'Popular',
    'MONTÓN GIMÉNEZ': 'Socialista',
    'MUÑOZ DE DIEGO': 'IU-ICV',
    'MUÑOZ RESTA': 'Socialista',
    'MUÑOZ SALVÀ': 'Socialista',
    'MUÑOZ SANTAMARÍA': 'Socialista',
    'MUÑOZ URIOL': 'Popular',
    'NADAL I AYMERICH': 'Popular',
    'NAHARRO DE MORA': 'Popular',
    'NARANJO BRAVO': 'Socialista',
    'NARBONA RUIZ': 'Socialista',
    'NAVARRO CASILLAS': 'IU-ICV',
    'NAVARRO GARZÓN': 'Socialista',
    'NIÑO RICO': 'Socialista',
    'OLIVA I PEÑA': 'ERC',
    'OLIVER SAGRERAS': 'Socialista',
    'OLTRA TORRES': 'Popular',
    'ORAMAS GONZÁLEZ-MORO': 'Mixto',
    'ORTIZ RIVAS': 'Socialista',
    'PAJÍN IRAOLA': 'Socialista',
    'PALACIO VALLERSUNDI': 'Popular',
    'PALMA I MUÑOZ': 'Socialista',
    'PAN VÁZQUEZ': 'Popular',
    'PASTOR JULIÁN': 'Popular',
    'PÉREZ ANGUITA': 'Socialista',
    'PÉREZ DOMÍNGUEZ': 'Socialista',
    'PIGEM I PALMÉS': 'CiU',
    'PIN ARBOLEDAS': 'Socialista',
    'POLONIO CONTRERAS': 'Socialista',
    'PONCE AGUILERA': 'Socialista',
    'PORTEIRO GARCÍA': 'Socialista',
    'POZO FERNÁNDEZ': 'Popular',
    'POZUELO MEÑO': 'Socialista',
    'PUIG GASOL': 'Socialista',
    'QUINTANILLA BARBA': 'Popular',
    'RAMÓN-LLIN I MARTÍNEZ': 'Popular',
    'RIVERO ALCOVER': 'Socialista',
    'RODRÍGUEZ HERRER': 'Popular',
    'RODRÍGUEZ LÓPEZ': 'Popular',
    'RODRÍGUEZ RAMOS': 'Socialista',
    'RODRÍGUEZ-SALMONES CABEZA': 'Popular',
    'ROLDÓS CABALLERO': 'Popular',
    'RUDI ÚBEDA': 'Popular',
    'RUMÍ IBÁÑEZ': 'Socialista',
    'SÁENZ DE SANTAMARÍA ANTÓN': 'Popular',
    'SÁENZ ROYO': 'Socialista',
    'SAINZ GARCÍA': 'Popular',
    'SALAZAR BELLO': 'Socialista',
    'SALOM COLL': 'Popular',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ FERNÁNDEZ': 'Popular',
    'SÁNCHEZ GARCÍA': 'Popular',
    'SÁNCHEZ JÓDAR': 'Socialista',
    'SÁNCHEZ RUBIO': 'Socialista',
    'SÁNCHEZ-CAMACHO PÉREZ': 'Popular',
    'SELLER ROCA DE TOGORES': 'Popular',
    'SERNA MASIÁ': 'Socialista',
    'TOLEDO SILVESTRE': 'Socialista',
    'TORME PARDO': 'Popular',
    'TORRADO REY': 'Socialista',
    'UNZALU PÉREZ DE EULATE': 'Socialista',
    'URÍA ETXEBARRÍA': 'EAJ-PNV',
    'VALCARCE GARCÍA': 'Socialista',
    'VÁZQUEZ BLANCO': 'Popular',
    'VELASCO GARCÍA': 'Socialista',
    'VELASCO MORILLO': 'Popular',
    'VILLAGRASA PÉREZ': 'Socialista',
    'VILLALOBOS TALERO': 'Popular',
}

# IX LEGISLATURA (2008-2011)
DICCIONARIO_IX_LEGISLATURA = {
    'ABURTO BASELGA': 'Socialista',
    'ACHUTEGUI BASAGOITI': 'Socialista',
    'ALEGRÍA CONTINENTE': 'Socialista',
    'ALONSO GARCÍA': 'Popular',
    'ÁLVAREZ ARZA': 'Socialista',
    'ÁLVAREZ DE TOLEDO PERALTA RAMOS': 'Popular',
    'ÁLVAREZ OTEO': 'Socialista',
    'ÁLVAREZ-ARENAS CISNEROS': 'Popular',
    'ARIAS RODRÍGUEZ': 'Popular',
    'ARNAIZ GARCÍA': 'Socialista',
    'BÁÑEZ GARCÍA': 'Popular',
    'BAÑULS ROS': 'Popular',
    'BARKOS BERRUEZO': 'Mixto',
    'BARREIRO ÁLVAREZ': 'Popular',
    'BATET LAMAÑA': 'Socialista',
    'BECERRIL BUSTAMANTE': 'Popular',
    'BLANCO TERÁN': 'Socialista',
    'BOLARÍN SÁNCHEZ': 'Popular',
    'BONILLA DOMÍNGUEZ': 'Popular',
    'BRAVO IBÁÑEZ': 'Popular',
    'BUENAVENTURA PUIG': 'ERC-IU-ICV',
    'CABEZÓN ARBAT': 'Socialista',
    'CABEZÓN RUIZ': 'Socialista',
    'CABRERA CALVO-SOTELO': 'Socialista',
    'CABRERA NODA': 'Socialista',
    'CALVO POYATO': 'Socialista',
    'CAMARERO BENÍTEZ': 'Popular',
    'CAMPO PIÑEIRO': 'Popular',
    'CANO DÍAZ': 'Socialista',
    'CARBALLEDO BERLANGA': 'Popular',
    'CARCEDO ROCES': 'Socialista',
    'CASAUS RODRÍGUEZ': 'Socialista',
    'CASTELLANO RAMÓN': 'Socialista',
    'CASTRO DOMÍNGUEZ': 'Popular',
    'CATALÁ VERDET': 'Popular',
    'CHACÓN CARRETERO': 'Socialista',
    'CHACÓN GUTIÉRREZ': 'Popular',
    'CHACÓN PIQUERAS': 'Socialista',
    'COELLO FERNÁNDEZ-TRUJILLO': 'Socialista',
    'COLLDEFORNS I SOL': 'Socialista',
    'CORRAL RUIZ': 'Socialista',
    'CORTAJARENA ITURRIOZ': 'Socialista',
    'COSTA PALACIOS': 'Socialista',
    'CUNILLERA I MESTRES': 'Socialista',
    'DÍEZ DE BALDEÓN GARCÍA': 'Socialista',
    'DÍEZ GONZÁLEZ': 'Mixto',
    'DUEÑAS HERRANZ': 'Popular',
    'DURÁN RAMOS': 'Popular',
    'ELÍAS CORDÓN': 'Socialista',
    'ESPINOSA MANGANA': 'Socialista',
    'ESTEVE ORTEGA': 'Socialista',
    'ESTRADA IBARS': 'Socialista',
    'FABRA FERNÁNDEZ': 'Popular',
    'FELIU ÁLVAREZ DE SOTOMAYOR': 'Popular',
    'FERNÁNDEZ AGUERRI': 'Socialista',
    'FERNÁNDEZ DAVILA': 'Mixto',
    'FERNÁNDEZ DE LA VEGA SANZ': 'Socialista',
    'FERNÁNDEZ PARDO': 'Popular',
    'FERNÁNDEZ-CAPEL BAÑOS': 'Popular',
    'FERRANDO SENDRA': 'Popular',
    'FUENTES PACHECO': 'Socialista',
    'GÁMEZ GARCÍA': 'Socialista',
    'GARCÍA RUIZ': 'Socialista',
    'GARCÍA SENA': 'Popular',
    'GARCÍA VALLS': 'Socialista',
    'GASTÓN MENAL': 'Socialista',
    'GÓMEZ SANTAMARÍA': 'Socialista',
    'GONZÁLEZ SEGURA': 'Popular',
    'GRANDE PESQUERO': 'Socialista',
    'GUAITA VAÑÓ': 'Popular',
    'GUERRA GUERRA': 'Popular',
    'GUINDULÁIN GUERENDIÁIN': 'Mixto',
    'GUTIÉRREZ DEL CASTILLO': 'Socialista',
    'HERMOSÍN BONO': 'Socialista',
    'IGLESIAS FONTAL': 'Popular',
    'JIMÉNEZ GARCÍA-HERRERA': 'Socialista',
    'JUANES BARCIELA': 'Socialista',
    'LARA CARBÓ': 'Popular',
    'LIZARRAGA GISBERT': 'Socialista',
    'LÓPEZ I CHAMOSA': 'Socialista',
    'LÓPEZ RODRÍGUEZ': 'Socialista',
    'LUQUERO DE NICOLÁS': 'Socialista',
    'MADRAZO DÍAZ': 'Popular',
    'MALARET GARCÍA': 'Socialista',
    'MARAÑÓN BASARTE': 'Socialista',
    'MARÓN BELTRÁN': 'Socialista',
    'MARTÍN GONZÁLEZ': 'Socialista',
    'MARTÍNEZ LÓPEZ': 'Socialista',
    'MARTÍNEZ SAIZ': 'Popular',
    'MATO ADROVER': 'Popular',
    'MEDINA TEVA': 'Socialista',
    'MÉNDEZ MONASTERIO': 'Popular',
    'MENDIZÁBAL GOROSTIAGA': 'Socialista',
    'MERCANT NADAL': 'Popular',
    'MONEO DÍEZ': 'Popular',
    'MONTESERÍN RODRÍGUEZ': 'Socialista',
    'MONTESINOS DE MIGUEL': 'Popular',
    'MONTÓN GIMÉNEZ': 'Socialista',
    'MONTSERRAT MONTSERRAT': 'Popular',
    'MUÑOZ RESTA': 'Socialista',
    'MUÑOZ SALVÀ': 'Socialista',
    'MUÑOZ SANTAMARÍA': 'Socialista',
    'NADAL I AYMERICH': 'Popular',
    'NARBONA RUIZ': 'Socialista',
    'NAVARRO CRUZ': 'Popular',
    'ORAMAS GONZÁLEZ-MORO': 'Mixto',
    'ORTEGA RODRÍGUEZ': 'Popular',
    'PAJÍN IRAOLA': 'Socialista',
    'PALMA I MUÑOZ': 'Socialista',
    'PASTOR JULIÁN': 'Popular',
    'PEDROSA ROLDÁN': 'Popular',
    'PÉREZ DOMÍNGUEZ': 'Socialista',
    'PÉREZ HERRAIZ': 'Socialista',
    'PIGEM I PALMÉS': 'CiU',
    'PIN ARBOLEDAS': 'Socialista',
    'POZUELO MEÑO': 'Socialista',
    'PUIG GASOL': 'Socialista',
    'QUINTANILLA BARBA': 'Popular',
    'RAMALLO VÁZQUEZ': 'Popular',
    'RIERA I REÑÉ': 'CiU',
    'RIVERO ALCOVER': 'Socialista',
    'RODRÍGUEZ BARAHONA': 'Socialista',
    'RODRÍGUEZ MANIEGA': 'Popular',
    'RODRÍGUEZ RAMOS': 'Socialista',
    'RODRÍGUEZ-PIÑERO FERNÁNDEZ': 'Socialista',
    'RODRÍGUEZ-SALMONES CABEZA': 'Popular',
    'ROS MARTÍNEZ': 'Socialista',
    'RUDI ÚBEDA': 'Popular',
    'RUMÍ IBÁÑEZ': 'Socialista',
    'SÁENZ DE SANTAMARÍA ANTÓN': 'Popular',
    'SALGADO MÉNDEZ': 'Socialista',
    'SALOM COLL': 'Popular',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ GARCÍA': 'Popular',
    'SÁNCHEZ RUBIO': 'Socialista',
    'SANTA ANA FERNÁNDEZ': 'Popular',
    'SANZ CARRILLO': 'Socialista',
    'SELLER ROCA DE TOGORES': 'Popular',
    'SERNA MASIÁ': 'Socialista',
    'SOLANA BARRAS': 'Popular',
    'SURROCA I COMAS': 'CiU',
    'SUSINOS TARRERO': 'Popular',
    'TARRUELLA TOMÀS': 'CiU',
    'TORME PARDO': 'Popular',
    'TORRADO DE CASTRO': 'Popular',
    'TORRES PARADA': 'Popular',
    'TORTOSA URREA': 'Socialista',
    'TRUJILLO RINCÓN': 'Socialista',
    'UNZALU PÉREZ DE EULATE': 'Socialista',
    'VALCARCE GARCÍA': 'Socialista',
    'VALDENEBRO RODRÍGUEZ': 'Popular',
    'VALENCIANO MARTÍNEZ-OROZCO': 'Socialista',
    'VÁZQUEZ BLANCO': 'Popular',
    'VÁZQUEZ MEJUTO': 'Popular',
    'VÁZQUEZ MORILLO': 'Socialista',
    'VILLAGRASA PÉREZ': 'Socialista',
    'VILLALOBOS TALERO': 'Popular',
}

# X LEGISLATURA (2011-2016)
DICCIONARIO_X_LEGISLATURA = {
    'AGUILAR RIVERO': 'Socialista',
    'ALBERTO PÉREZ': 'Popular',
    'ALEGRÍA CONTINENTE': 'Socialista',
    'ÁLVAREZ ÁLVAREZ': 'Socialista',
    'ÁLVAREZ DE TOLEDO PERALTA RAMOS': 'Popular',
    'ÁLVAREZ-ARENAS CISNEROS': 'Popular',
    'ANGULO ROMERO': 'Popular',
    'ARES MARTÍNEZ-FORTÚN': 'Popular',
    'ARIZTEGUI LARRAÑAGA': 'Mixto',
    'ARNAIZ GARCÍA': 'Socialista',
    'ASIAN GONZÁLEZ': 'Popular',
    'BAENA AZUAGA': 'Popular',
    'BAJO PRIETO': 'Popular',
    'BÁÑEZ GARCÍA': 'Popular',
    'BARKOS BERRUEZO': 'Mixto',
    'BARREIRO ÁLVAREZ': 'Popular',
    'BATET LAMAÑA': 'Socialista',
    'BLANCO TERÁN': 'Socialista',
    'BLANQUER ALCARAZ': 'Socialista',
    'BLASCO SOTO': 'Popular',
    'BOLARÍN SÁNCHEZ': 'Popular',
    'BONILLA DOMÍNGUEZ': 'Popular',
    'BORREGO CORTÉS': 'Popular',
    'BRAVO IBÁÑEZ': 'Popular',
    'CABEZÓN RUIZ': 'Socialista',
    'CAMARERO BENÍTEZ': 'Popular',
    'CAÑIZARES CABEZAS': 'Socialista',
    'CARCEDO ROCES': 'Socialista',
    'CARREÑO FERNÁNDEZ': 'Popular',
    'CASAUS RODRÍGUEZ': 'Socialista',
    'CASTAÑO REY': 'La Izquierda Plural',
    'CASTELLANO I FERNÁNDEZ': 'CiU',
    'CASTELLANO RAMÓN': 'Socialista',
    'CHACÓN PIQUERAS': 'Socialista',
    'CID MUÑOZ': 'Popular',
    'CIURÓ I BULDÓ': 'CiU',
    'COBALEDA HERNÁNDEZ': 'Popular',
    'COBOS TRALLERO': 'Popular',
    'CONDE MARTÍNEZ': 'Popular',
    'CORRAL RUIZ': 'Socialista',
    'CORTÉS BURETA': 'Popular',
    'COSTA PALACIOS': 'Socialista',
    'CUNILLERA I MESTRES': 'Socialista',
    'DE JUAN DE MIGUEL': 'Socialista',
    'DE LAS HERAS LADERA': 'La Izquierda Plural',
    'DÍEZ GONZÁLEZ': 'UPyD',
    'DUQUE PALACIOS': 'Popular',
    'DURÁN RAMOS': 'Popular',
    'ENBEITA MAGUREGI': 'Mixto',
    'ESCUDERO BERZAL': 'Popular',
    'ESPAÑA REINA': 'Popular',
    'ESTELLER RUEDAS': 'Popular',
    'ESTEVE ORTEGA': 'Socialista',
    'FABRA FERNÁNDEZ': 'Popular',
    'FALCÓN DACAL': 'Popular',
    'FERNÁNDEZ DAVILA': 'Mixto',
    'FERNÁNDEZ GONZÁLEZ': 'Popular',
    'FERNÁNDEZ MOYA': 'Socialista',
    'FERNÁNDEZ-AHUJA GARCÍA': 'Popular',
    'FERRANDO SENDRA': 'Popular',
    'FIGUERES GÓRRIZ': 'Popular',
    'FORTEA MILLÁN': 'Popular',
    'FUMERO ROQUE': 'Popular',
    'GALLEGO ARRIOLA': 'Socialista',
    'GÁMEZ GARCÍA': 'Socialista',
    'GARCÍA ÁLVAREZ': 'La Izquierda Plural',
    'GARCÍA GÁLVEZ': 'Popular',
    'GARCÍA SENA': 'Popular',
    'GARRIDO VALENZUELA': 'Popular',
    'GOMIS DE BARBARÀ': 'CiU',
    'GONZÁLEZ GUTIÉRREZ': 'Popular',
    'GONZÁLEZ SANTÍN': 'Socialista',
    'GONZÁLEZ VÁZQUEZ': 'Popular',
    'GONZÁLEZ VERACRUZ': 'Socialista',
    'GRANDE PESQUERO': 'Socialista',
    'GUAITA VAÑÓ': 'Popular',
    'GUTIÉRREZ DEL CASTILLO': 'Socialista',
    'HEREDIA MARTÍN': 'Popular',
    'HERNÁNDEZ GUTIÉRREZ': 'Socialista',
    'HERNANZ COSTA': 'Socialista',
    'HOYO JULIÁ': 'Popular',
    'IGLESIAS FONTAL': 'Popular',
    'IGLESIAS SANTIAGO': 'Socialista',
    'JIMÉNEZ DÍAZ': 'Popular',
    'JIMÉNEZ GARCÍA-HERRERA': 'Socialista',
    'JIMÉNEZ MÍNGUEZ': 'Popular',
    'JORDÀ I ROURA': 'Mixto',
    'JUSTE PICÓN': 'Popular',
    'LAGO MARTÍNEZ': 'Popular',
    'LARA CARBÓ': 'Popular',
    'LÓPEZ GONZÁLEZ': 'Popular',
    'LÓPEZ I CHAMOSA': 'Socialista',
    'LOZANO DOMINGO': 'UPyD',
    'LUCIO CARRASCO': 'Socialista',
    'MADRAZO DÍAZ': 'Popular',
    'MARCOS DOMÍNGUEZ': 'Popular',
    'MARTÍN GONZÁLEZ': 'Socialista',
    'MARTÍN POZO': 'Popular',
    'MARTÍN REVUELTA': 'Popular',
    'MARTÍNEZ FERRO': 'Popular',
    'MARTÍNEZ SAIZ': 'Popular',
    'MATO ADROVER': 'Popular',
    'MÉNDEZ MONASTERIO': 'Popular',
    'MICHEO CARRILLO-ALBORNOZ': 'Popular',
    'MIGUÉLEZ PARIENTE': 'Popular',
    'MONEO DÍEZ': 'Popular',
    'MONTESERÍN RODRÍGUEZ': 'Socialista',
    'MONTESINOS DE MIGUEL': 'Popular',
    'MONTÓN GIMÉNEZ': 'Socialista',
    'MONTSERRAT MONTSERRAT': 'Popular',
    'MORALEJA GÓMEZ': 'Popular',
    'MORENO FELIPE': 'Popular',
    'MORO ALMARAZ': 'Popular',
    'MUÑOZ SANTAMARÍA': 'Socialista',
    'NARBONA RUIZ': 'Socialista',
    'NAVARRO CRUZ': 'Popular',
    'OÑATE MOYA': 'Socialista',
    'ORAMAS GONZÁLEZ-MORO': 'Mixto',
    'ORTIZ CASTELLVÍ': 'La Izquierda Plural',
    'PAJÍN IRAOLA': 'Socialista',
    'PASTOR JULIÁN': 'Popular',
    'PÉREZ DOMÍNGUEZ': 'Socialista',
    'PÉREZ FERNÁNDEZ': 'Mixto',
    'PÉREZ HERRAIZ': 'Socialista',
    'PÉREZ SERNA': 'Popular',
    'PIGEM I PALMÉS': 'CiU',
    'POZUELO MEÑO': 'Socialista',
    'PUYUELO DEL VAL': 'Popular',
    'QUINTANILLA BARBA': 'Popular',
    'RAMÓN UTRABO': 'Socialista',
    'REYES MIRANDA': 'Popular',
    'RIERA I REÑÉ': 'CiU',
    'RODRÍGUEZ BARAHONA': 'Socialista',
    'RODRÍGUEZ CONCEPCIÓN': 'Socialista',
    'RODRÍGUEZ FLORES': 'Popular',
    'RODRÍGUEZ GARCÍA': 'Socialista',
    'RODRÍGUEZ HERRER': 'Popular',
    'RODRÍGUEZ MANIEGA': 'Popular',
    'RODRÍGUEZ RAMÍREZ': 'Socialista',
    'RODRÍGUEZ RAMOS': 'Socialista',
    'RODRÍGUEZ SÁNCHEZ': 'Popular',
    'RODRÍGUEZ VÁZQUEZ': 'Socialista',
    'RODRÍGUEZ-PIÑERO FERNÁNDEZ': 'Socialista',
    'RODRÍGUEZ-SALMONES CABEZA': 'Popular',
    'ROMERO RODRÍGUEZ': 'Popular',
    'ROMERO SÁNCHEZ': 'Popular',
    'ROMINGUERA SALAZAR': 'Socialista',
    'ROS MARTÍNEZ': 'Socialista',
    'ROSELLÓ SAUS': 'Socialista',
    'RUMÍ IBÁÑEZ': 'Socialista',
    'SÁENZ DE SANTAMARÍA ANTÓN': 'Popular',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ GARCÍA': 'Popular',
    'SÁNCHEZ ROBLES': 'EAJ-PNV',
    'SÁNCHEZ TRANCOSO': 'Socialista',
    'SANTA ANA FERNÁNDEZ': 'Popular',
    'SAYÓS I MOTILLA': 'CiU',
    'SEARA SOBRADO': 'Socialista',
    'SERRANO ARGÜELLO': 'Popular',
    'SILVA REGO': 'Socialista',
    'SUÁREZ-BÁRCENA BLASCO': 'Popular',
    'SUMELZO JORDÁN': 'Socialista',
    'SURROCA I COMAS': 'CiU',
    'SUSINOS TARRERO': 'Popular',
    'TAPIA OTAEGI': 'EAJ-PNV',
    'TARRUELLA TOMÁS': 'CiU',
    'TORRADO DE CASTRO': 'Popular',
    'VALENCIANO MARTÍNEZ-OROZCO': 'Socialista',
    'VALERIO CORDERO': 'Socialista',
    'VARELA LEMA': 'Popular',
    'VÁZQUEZ BLANCO': 'Popular',
    'VÁZQUEZ MORILLO': 'Socialista',
    'VERAY CAMA': 'Popular',
    'VILLALOBOS TALERO': 'Popular',
}

# XI LEGISLATURA (2016)
DICCIONARIO_XI_LEGISLATURA = {
    'ALBA GOVELI': 'Podemos',
    'ALCONCHEL GONZAGA': 'Socialista',
    'ALONSO CLUSA': 'Podemos',
    'ALÓS LÓPEZ': 'Popular',
    'ÁLVAREZ SIMÓN': 'Popular',
    'ÁLVAREZ-ARENAS CISNEROS': 'Popular',
    'ANGULO ROMERO': 'Popular',
    'ARDANZA URI BARREN': 'EAJ-PNV',
    'ASIAN GONZÁLEZ': 'Popular',
    'BAJO PRIETO': 'Popular',
    'BALLESTER MUÑOZ': 'Podemos',
    'BÁÑEZ GARCÍA': 'Popular',
    'BASTIDAS BONO': 'Popular',
    'BATET LAMAÑA': 'Socialista',
    'BEITIALARRANGOITIA LIZARRALDE': 'Mixto',
    'BELARRA URTEAGA': 'Podemos',
    'BESCANSA HERNÁNDEZ': 'Podemos',
    'BLANQUER ALCARAZ': 'Socialista',
    'BONILLA DOMÍNGUEZ': 'Popular',
    'BORREGO CORTÉS': 'Popular',
    'BOSAHO GORI': 'Podemos',
    'BOTEJARA SANZ': 'Podemos',
    'BOTELLA GÓMEZ': 'Socialista',
    'BRAVO IBÁÑEZ': 'Popular',
    'CANCELA RODRÍGUEZ': 'Socialista',
    'CANTERA DE CASTRO': 'Socialista',
    'CAPELLA I FARRÉ': 'ERC',
    'CARREÑO FERNÁNDEZ': 'Popular',
    'CARREÑO VALERO': 'Podemos',
    'CASCALES MARTÍNEZ': 'Popular',
    'CHACÓN PIQUERAS': 'Socialista',
    'CIURÓ I BULDÓ': 'PDeCAT',
    'CORTÉS BURETA': 'Popular',
    'CUELLO PÉREZ': 'Socialista',
    'DE COSPEDAL GARCÍA': 'Popular',
    'DE FRUTOS MADRAZO': 'Socialista',
    'DE LA CONCHA GARCÍA-MAURIÑO': 'Podemos',
    'DEL MORAL MILLA': 'Socialista',
    'DÍAZ PÉREZ': 'Podemos',
    'DOMÍNGUEZ ÁLVAREZ': 'Podemos',
    'DUEÑAS MARTÍNEZ': 'Popular',
    'ELIZO SERRANO': 'Podemos',
    'ENBEITA MAGUREGI': 'Mixto',
    'ESCUDERO BERZAL': 'Popular',
    'ESPAÑA REINA': 'Popular',
    'FABA DE LA ENCARNACIÓN': 'Ciudadanos',
    'FERNÁNDEZ CASTAÑÓN': 'Podemos',
    'FERNÁNDEZ GÓMEZ': 'Podemos',
    'FERRER TESORO': 'Socialista',
    'FLÓREZ RODRÍGUEZ': 'Socialista',
    'FRANCO CARMONA': 'Podemos',
    'GALLEGO ARRIOLA': 'Socialista',
    'GALOVART CARRERA': 'Socialista',
    'GARCÍA PUIG': 'Podemos',
    'GARCÍA TEJERINA': 'Popular',
    'GARCÍA-PELAYO JURADO': 'Popular',
    'GARRIDO VALENZUELA': 'Popular',
    'GONZÁLEZ BAYO': 'Socialista',
    'GONZÁLEZ GUINDA': 'Popular',
    'GONZÁLEZ VÁZQUEZ': 'Popular',
    'GONZÁLEZ VERACRUZ': 'Socialista',
    'GUERRA MANSITO': 'Podemos',
    'GUINART MORENO': 'Socialista',
    'HERNANZ COSTA': 'Socialista',
    'HONORATO CHULIÁN': 'Podemos',
    'HOYO JULIÁ': 'Popular',
    'ISAC GARCÍA': 'Popular',
    'JORDÀ I ROURA': 'ERC',
    'LAFUENTE DE LA TORRE': 'Socialista',
    'LARA CARBÓ': 'Popular',
    'LASTRA FERNÁNDEZ': 'Socialista',
    'LÓPEZ ARES': 'Popular',
    'LOZANO DOMINGO': 'Socialista',
    'LUCIO CARRASCO': 'Socialista',
    'MADRAZO DÍAZ': 'Popular',
    'MARCELLO SANTOS': 'Podemos',
    'MARCOS MOYANO': 'Popular',
    'MARTÍN LLAGUNO': 'Ciudadanos',
    'MARTÍNEZ RODRÍGUEZ': 'Podemos',
    'MARTÍNEZ SAIZ': 'Popular',
    'MARTÍNEZ SEIJO': 'Socialista',
    'MEDINA SUÁREZ': 'Podemos',
    'MIGUEL MUÑOZ': 'Ciudadanos',
    'MILLÁN SALMERÓN': 'Ciudadanos',
    'MONEO DÍEZ': 'Popular',
    'MONTERO GIL': 'Podemos',
    'MONTSERRAT MONTSERRAT': 'Popular',
    'MORO ALMARAZ': 'Popular',
    'NAVARRO GARZÓN': 'Socialista',
    'NAVARRO LACOBA': 'Popular',
    'NOGUERAS I CAMERO': 'PDeCAT',
    'ORAMAS GONZÁLEZ-MORO': 'Mixto',
    'PASTOR JULIÁN': 'Popular',
    'PASTOR MUÑOZ': 'Podemos',
    'PEÑA CAMARERO': 'Socialista',
    'PEREA I CONILLAS': 'Socialista',
    'PÉREZ DOMÍNGUEZ': 'Socialista',
    'PÉREZ HERRAIZ': 'Socialista',
    'PITA CÁRDENES': 'Podemos',
    'QUINTANILLA BARBA': 'Popular',
    'RAMÓN UTRABO': 'Socialista',
    'RAYA RODRÍGUEZ': 'Socialista',
    'REYES RIVERA': 'Ciudadanos',
    'REYNÉS CALVACHE': 'Popular',
    'RIBERA I GARIJO': 'PDeCAT',
    'RIVERA ANDRÉS': 'Ciudadanos',
    'RIVERA DE LA CRUZ': 'Ciudadanos',
    'RODRÍGUEZ FERNÁNDEZ': 'Socialista',
    'RODRÍGUEZ GARCÍA': 'Socialista',
    'RODRÍGUEZ HERNÁNDEZ': 'Socialista',
    'RODRÍGUEZ MARTÍNEZ': 'Podemos',
    'RODRÍGUEZ RAMOS': 'Socialista',
    'ROJO NOGUERA': 'Popular',
    'ROMERO RODRÍGUEZ': 'Popular',
    'ROMERO SÁNCHEZ': 'Popular',
    'ROMINGUERA SALAZAR': 'Socialista',
    'ROSELL AGUILAR': 'Podemos',
    'SÁENZ DE SANTAMARÍA ANTÓN': 'Popular',
    'SÁNCHEZ MAROTO': 'Mixto',
    'SÁNCHEZ MELERO': 'Podemos',
    'SÁNCHEZ-CAMACHO PÉREZ': 'Popular',
    'SANTA ANA FERNÁNDEZ': 'Popular',
    'SERRANO BOIGAS': 'Socialista',
    'SERRANO JIMÉNEZ': 'Socialista',
    'SIBINA CAMPS': 'Podemos',
    'SIERRA ROJAS': 'Socialista',
    'SORLÍ FRESQUET': 'Mixto',
    'SUCH PALOMARES': 'Socialista',
    'SUMELZO JORDÁN': 'Socialista',
    'SURRA SPADEA': 'ERC',
    'TERRADAS VIÑALS': 'Podemos',
    'TERRÓN BERBEL': 'Podemos',
    'VALMAÑA OCHAÍTA': 'Popular',
    'VERA RUIZ-HERRERA': 'Podemos',
    'VIDAL SÁEZ': 'Podemos',
    'VILLALOBOS TALERO': 'Popular',
}

# XII LEGISLATURA (2016-2019)
DICCIONARIO_XII_LEGISLATURA = {
    'ÁBALOS MECO': 'Socialista',
    'ACEDO PENCO': 'Popular',
    'AGIRRETXEA URRESTI': 'EAJ-PNV',
    'AGUIAR RODRÍGUEZ': 'Popular',
    'AGUIRRE RODRÍGUEZ': 'Popular',
    'ALBA GOVELI': 'Podemos',
    'ALBA MULLOR': 'Popular',
    'ALBALADEJO MARTÍNEZ': 'Popular',
    'ALBERTO PÉREZ': 'Popular',
    'ALCONCHEL GONZAGA': 'Socialista',
    'ALLI MARTÍNEZ': 'Mixto',
    'ALONSO ARANEGUI': 'Popular',
    'ALONSO CANTORNÉ': 'Podemos',
    'ALONSO CLUSA': 'Podemos',
    'ALONSO DÍAZ-GUERRA': 'Popular',
    'ALONSO HERNÁNDEZ': 'Popular',
    'ALÓS LÓPEZ': 'Popular',
    'ÁLVAREZ ÁLVAREZ': 'Socialista',
    'ÁLVAREZ PALLEIRO': 'Ciudadanos',
    'ÁLVAREZ-ARENAS CISNEROS': 'Popular',
    'ANGULO ROMERO': 'Popular',
    'ANTÓN CACHO': 'Socialista',
    'ARÉVALO CARABALLO': 'Podemos',
    'ARROJO AGUDO': 'Podemos',
    'ASIAN GONZÁLEZ': 'Popular',
    'AYLLÓN MANSO': 'Popular',
    'AZPIAZU URIARTE': 'EAJ-PNV',
    'BAJO PRIETO': 'Popular',
    'BALDOVÍ RODA': 'Mixto',
    'BALLESTER MUÑOZ': 'Podemos',
    'BÁÑEZ GARCÍA': 'Popular',
    'BAÑOS RUIZ': 'Socialista',
    'BARANDIARAN BENITO': 'EAJ-PNV',
    'BARRACHINA ROS': 'Popular',
    'BARREDA DE LOS RÍOS': 'Popular',
    'BARREDA FUENTES': 'Socialista',
    'BARRIOS TEJERO': 'Popular',
    'BASTIDAS BONO': 'Popular',
    'BATALLER I RUIZ': 'Mixto',
    'BATET LAMAÑA': 'Socialista',
    'BEITIALARRANGOITIA LIZARRALDE': 'Mixto',
    'BEL ACCENSI': 'Mixto',
    'BELARRA URTEAGA': 'Podemos',
    'BELLIDO ACEVEDO': 'Socialista',
    'BERMÚDEZ DE CASTRO FERNÁNDEZ': 'Popular',
    'BERNABÉ PÉREZ': 'Popular',
    'BESCANSA HERNÁNDEZ': 'Podemos',
    'BLANCO GARRIDO': 'Popular',
    'BLANQUER ALCARAZ': 'Socialista',
    'BLASCO MARQUÉS': 'Popular',
    'BOLARÍN SÁNCHEZ': 'Popular',
    'BONILLA DOMÍNGUEZ': 'Popular',
    'BORREGO CORTÉS': 'Popular',
    'BOSAHO GORI': 'Podemos',
    'BOTEJARA SANZ': 'Podemos',
    'BOTELLA GÓMEZ': 'Socialista',
    'BRAVO BAENA': 'Popular',
    'BURGOS GALLEGO': 'Popular',
    'BUSTAMANTE MARTÍN': 'Podemos',
    'BUSTINDUY AMADOR': 'Podemos',
    'CABEZAS REGAÑO': 'Popular',
    'CABRERA CARMONA': 'Popular',
    'CALVENTE GALLEGO': 'Popular',
    'CAMACHO SÁNCHEZ': 'Socialista',
    'CÁMARA VILLAR': 'Socialista',
    'CAMPO MORENO': 'Socialista',
    'CAMPOS ARTESEROS': 'Socialista',
    'CAMPS DEVESA': 'Popular',
    'CAMPUZANO I CANADÉS': 'Mixto',
    'CANCELA RODRÍGUEZ': 'Socialista',
    'CANDELA SERNA': 'Mixto',
    'CANDÓN ADÁN': 'Popular',
    'CANO FUSTER': 'Ciudadanos',
    'CANO LEAL': 'Ciudadanos',
    'CANTERA DE CASTRO': 'Socialista',
    'CANTÓ GARCÍA DEL MORAL': 'Ciudadanos',
    'CAÑAMERO VALLE': 'Podemos',
    'CAPDEVILA I ESTEVE': 'ERC',
    'CAPELLA I FARRÉ': 'ERC',
    'CARRACEDO VERDE': 'Podemos',
    'CARRENO FERNÁNDEZ': 'Popular',
    'CARRENO VALERO': 'Podemos',
    'CASADO BLANCO': 'Popular',
    'CASCALES MARTÍNEZ': 'Popular',
    'CATALÁ POLO': 'Popular',
    'CHAIB AKHDIM': 'Socialista',
    'CHANDIRAMANI RAMESH': 'Popular',
    'CHIQUILLO BARBER': 'Popular',
    'CÍSCAR CASABÁN': 'Socialista',
    'CIURÓ I BULDÓ': 'Mixto',
    'CLAVELL LÓPEZ': 'Popular',
    'CLEMENTE GIMÉNEZ': 'Ciudadanos',
    'CORTÉS BURETA': 'Popular',
    'CORTÉS LASTRA': 'Socialista',
    'COTELO BALMASEDA': 'Popular',
    'CRUZ RODRÍGUEZ': 'Socialista',
    'CUELLO PÉREZ': 'Socialista',
    'DE ARRIBA SÁNCHEZ': 'Popular',
    'DE BARRIONUEVO GENER': 'Popular',
    'DE COSPEDAL GARCÍA': 'Popular',
    'DE FRUTOS MADRAZO': 'Socialista',
    'DE LA CONCHA GARCÍA-MAURIÑO': 'Podemos',
    'DE LA ENCINA ORTEGA': 'Socialista',
    'DE LA TORRE DÍAZ': 'Ciudadanos',
    'DEL CAMPO ESTAÚN': 'Ciudadanos',
    'DEL OLMO IBÁÑEZ': 'Podemos',
    'DEL RÍO SANZ': 'Popular',
    'DELGADO ARCE': 'Popular',
    'DELGADO RAMOS': 'Podemos',
    'DÍAZ GÓMEZ': 'Ciudadanos',
    'DÍAZ PÉREZ': 'Podemos',
    'DÍAZ TRILLO': 'Socialista',
    'DOMÈNECH SAMPERE': 'Podemos',
    'DUEÑAS MARTÍNEZ': 'Popular',
    'ECHÁNIZ SALGADO': 'Popular',
    'ELIZO SERRANO': 'Podemos',
    'ELORZA GONZÁLEZ': 'Socialista',
    'ERITJA CIURÓ': 'ERC',
    'ERREJÓN GALVÁN': 'Podemos',
    'ESCUDERO BERZAL': 'Popular',
    'ESPAÑA REINA': 'Popular',
    'ESTEBAN BRAVO': 'EAJ-PNV',
    'ESTELLER RUEDAS': 'Popular',
    'EXPÓSITO PRIETO': 'Podemos',
    'FABA DE LA ENCARNACIÓN': 'Ciudadanos',
    'FARRÉ FIDALGO': 'Podemos',
    'FERNÁNDEZ BELLO': 'Podemos',
    'FERNÁNDEZ CASTAÑÓN': 'Podemos',
    'FERNÁNDEZ DE MOYA ROMERO': 'Popular',
    'FERNÁNDEZ DÍAZ': 'Popular',
    'FERNÁNDEZ GARCÍA': 'Popular',
    'FERNÁNDEZ GÓMEZ': 'Podemos',
    'FERRER TESORO': 'Socialista',
    'FLÓREZ RODRÍGUEZ': 'Socialista',
    'FLORIANO CORRALES': 'Popular',
    'FOLE DÍAZ': 'Popular',
    'FRANCO CARMONA': 'Podemos',
    'FRANQUIS VERA': 'Socialista',
    'GALEANO GRACIA': 'Socialista',
    'GALLEGO ARRIOLA': 'Socialista',
    'GALOVART CARRERA': 'Socialista',
    'GAMAZO MICÓ': 'Popular',
    'GARAULET RODRÍGUEZ': 'Ciudadanos',
    'GARCÍA CAÑAL': 'Popular',
    'GARCÍA DÍEZ': 'Popular',
    'GARCÍA EGEA': 'Popular',
    'GARCÍA HERNÁNDEZ': 'Popular',
    'GARCÍA MIRA': 'Socialista',
    'GARCÍA PUIG': 'Podemos',
    'GARCÍA SEMPERE': 'Podemos',
    'GARCÍA TEJERINA': 'Popular',
    'GARCÍA-MARGALLO Y MARFIL': 'Popular',
    'GARCÍA-PELAYO JURADO': 'Popular',
    'GARCÍA-TIZÓN LÓPEZ': 'Popular',
    'GARRIDO VALENZUELA': 'Popular',
    'GARZÓN ESPINOSA': 'Podemos',
    'GIRAUTA VIDAL': 'Ciudadanos',
    'GÓMEZ BALSERA': 'Ciudadanos',
    'GÓMEZ GARCÍA': 'Ciudadanos',
    'GÓMEZ-REINO VARELA': 'Podemos',
    'GONZÁLEZ BAYO': 'Socialista',
    'GONZÁLEZ GARCÍA': 'Podemos',
    'GONZÁLEZ GUINDA': 'Popular',
    'GONZÁLEZ MUÑOZ': 'Popular',
    'GONZÁLEZ PELÁEZ': 'Socialista',
    'GONZÁLEZ RAMOS': 'Socialista',
    'GONZÁLEZ TEROL': 'Popular',
    'GONZÁLEZ VÁZQUEZ': 'Popular',
    'GONZÁLEZ VERACRUZ': 'Socialista',
    'GORDO PÉREZ': 'Socialista',
    'GUIJARRO GARCÍA': 'Podemos',
    'GUILLAUMES I RÀFOLS': 'Mixto',
    'GUINART MORENO': 'Socialista',
    'GUTIÉRREZ LIMONES': 'Socialista',
    'GUTIÉRREZ VIVAS': 'Ciudadanos',
    'HEREDIA DÍAZ': 'Socialista',
    'HEREDIA MARTÍN': 'Popular',
    'HERNANDO BENTO': 'Popular',
    'HERNANDO FRAILE': 'Popular',
    'HERNANDO VERA': 'Socialista',
    'HERNANZ COSTA': 'Socialista',
    'HERRERO BONO': 'Popular',
    'HOMS MOLIST': 'Mixto',
    'HONORATO CHULIÁN': 'Podemos',
    'HOYO JULIÁ': 'Popular',
    'HURTADO ZURERA': 'Socialista',
    'IGEA ARISQUETA': 'Ciudadanos',
    'IGLESIAS CAMARERO': 'Podemos',
    'IGLESIAS TURRIÓN': 'Podemos',
    'ISAC GARCÍA': 'Popular',
    'JIMÉNEZ TORTOSA': 'Socialista',
    'JORDÀ I ROURA': 'ERC',
    'JULIÀ JULIÀ': 'Ciudadanos',
    'JUNCAL RODRÍGUEZ': 'Popular',
    'LAMUÀ ESTAÑOL': 'Socialista',
    'LARA CARBÓ': 'Popular',
    'LASARTE IRIBARREN': 'Socialista',
    'LASSALLE RUIZ': 'Popular',
    'LASTRA FERNÁNDEZ': 'Socialista',
    'LEGARDA URIARTE': 'EAJ-PNV',
    'LLORENS TORRES': 'Popular',
    'LÓPEZ ÁLVAREZ': 'Socialista',
    'LÓPEZ ARES': 'Popular',
    'LÓPEZ DE URALDE GARMENDIA': 'Podemos',
    'LÓPEZ MILLA': 'Socialista',
    'LÓPEZ SOMOZA': 'Socialista',
    'LORENZO RODRÍGUEZ': 'Ciudadanos',
    'LORENZO TORRES': 'Popular',
    'LUCIO CARRASCO': 'Socialista',
    'LUENA LÓPEZ': 'Socialista',
    'LUIS BAIL': 'Podemos',
    'LUIS RODRÍGUEZ': 'Popular',
    'MADINA MUÑOZ': 'Socialista',
    'MARCELLO SANTOS': 'Podemos',
    'MARCOS DOMÍNGUEZ': 'Popular',
    'MARCOS MOYANO': 'Popular',
    'MARGALL SASTRE': 'ERC',
    'MARÍ BOSÓ': 'Popular',
    'MARISCAL ANAYA': 'Popular',
    'MAROTO ARANZÁBAL': 'Popular',
    'MARTÍN LLAGUNO': 'Ciudadanos',
    'MARTÍNEZ FERRO': 'Popular',
    'MARTÍNEZ GONZÁLEZ': 'Ciudadanos',
    'MARTÍNEZ OBLANCA': 'Mixto',
    'MARTÍNEZ RODRÍGUEZ': 'Podemos',
    'MARTÍNEZ SAIZ': 'Popular',
    'MARTÍNEZ SEIJO': 'Socialista',
    'MARTÍNEZ VÁZQUEZ': 'Popular',
    'MARTÍNEZ-MAILLO TORIBIO': 'Popular',
    'MARTÍN-TOLEDANO SUÁREZ': 'Popular',
    'MATARÍ SÁEZ': 'Popular',
    'MATEU ISTÚRIZ': 'Popular',
    'MATOS MASCAREÑO': 'Popular',
    'MATUTE GARCÍA DE JALÓN': 'Mixto',
    'MAURA BARANDIARÁN': 'Ciudadanos',
    'MAURA ZORITA': 'Podemos',
    'MAYORAL PERALES': 'Podemos',
    'MEIJÓN COUSELO': 'Socialista',
    'MENA ARCA': 'Podemos',
    'MÉNDEZ DE VIGO MONTOJO': 'Popular',
    'MERCHÁN MESÓN': 'Socialista',
    'MERINO LÓPEZ': 'Popular',
    'MILLÁN SALMERÓN': 'Ciudadanos',
    'MIQUEL I VALENTÍ': 'Mixto',
    'MOLINERO HOYOS': 'Popular',
    'MONEO DÍEZ': 'Popular',
    'MONEREO PÉREZ': 'Podemos',
    'MONTERO GIL': 'Podemos',
    'MONTERO SOLER': 'Podemos',
    'MONTORO ROMERO': 'Popular',
    'MONTSERRAT MONTSERRAT': 'Popular',
    'MORAGAS SÁNCHEZ': 'Popular',
    'MORALEJA GÓMEZ': 'Popular',
    'MORENO BUSTOS': 'Popular',
    'MORENO PALANQUES': 'Popular',
    'MORO ALMARAZ': 'Popular',
    'MOVELLÁN LOMBILLA': 'Popular',
    'MOYA MATAS': 'Podemos',
    'MUÑOZ GONZÁLEZ': 'Socialista',
    'NADAL BELDA': 'Popular',
    'NAVARRO CRUZ': 'Popular',
    'NAVARRO FERNÁNDEZ-RODRÍGUEZ': 'Ciudadanos',
    'NAVARRO GARZÓN': 'Socialista',
    'NAVARRO LACOBA': 'Popular',
    'NIETO BALLESTEROS': 'Popular',
    'NOGUERAS I CAMERO': 'Mixto',
    'NÚÑEZ JIMÉNEZ': 'Popular',
    'OLANO VELA': 'Popular',
    'OLÒRIZ SERRA': 'ERC',
    'ORAMAS GONZÁLEZ-MORO': 'Mixto',
    'PALACÍN GUARNÉ': 'Socialista',
    'PALMER TOUS': 'Popular',
    'PANIAGUA NÚÑEZ': 'Popular',
    'PASCUAL PEÑA': 'Podemos',
    'PASTOR JULIÁN': 'Popular',
    'PASTOR MUÑOZ': 'Podemos',
    'PEÑA CAMARERO': 'Socialista',
    'PEREA I CONILLAS': 'Socialista',
    'PÉREZ ARAS': 'Popular',
    'PÉREZ DOMÍNGUEZ': 'Socialista',
    'PÉREZ HERRAIZ': 'Socialista',
    'PÉREZ LÓPEZ': 'Popular',
    'PÉREZ-HICKMAN SILVÁN': 'Popular',
    'PIQUER SANCHO': 'Socialista',
    'PÍRIZ MAYA': 'Popular',
    'PITA CÁRDENES': 'Podemos',
    'PONS SAMPIETRO': 'Socialista',
    'POSADA MORENO': 'Popular',
    'POSTIGO QUINTANA': 'Popular',
    'POSTIUS TERRADO': 'Mixto',
    'PRADAS TORRES': 'Socialista',
    'PRENDES PRENDES': 'Ciudadanos',
    'QUEVEDO ITURBE': 'Mixto',
    'QUINTANA MARTÍNEZ': 'Socialista',
    'QUINTANILLA BARBA': 'Popular',
    'RAJOY BREY': 'Popular',
    'RALLO LOMBARTE': 'Socialista',
    'RAMÍREZ DEL MOLINO MORÁN': 'Popular',
    'RAMÍREZ FREIRE': 'Ciudadanos',
    'RAMÓN UTRABO': 'Socialista',
    'RAMOS ESTEBAN': 'Socialista',
    'RAMOS JORDÁN': 'Podemos',
    'RAYA RODRÍGUEZ': 'Socialista',
    'REYES RIVERA': 'Ciudadanos',
    'REYNÉS CALVACHE': 'Popular',
    'RIVERA ANDRÉS': 'Ciudadanos',
    'RIVERA DE LA CRUZ': 'Ciudadanos',
    'RIVERA DÍAZ': 'Ciudadanos',
    'ROBLES FERNÁNDEZ': 'Socialista',
    'ROCA MAS': 'Popular',
    'RODRÍGUEZ GARCÍA': 'Socialista',
    'RODRÍGUEZ HERNÁNDEZ': 'Socialista',
    'RODRÍGUEZ MARTÍNEZ': 'Podemos',
    'RODRÍGUEZ RAMOS': 'Socialista',
    'RODRÍGUEZ RODRÍGUEZ': 'Podemos',
    'ROJAS GARCÍA': 'Popular',
    'ROJO NOGUERA': 'Popular',
    'ROLDÁN MONÉS': 'Ciudadanos',
    'ROMANÍ CANTERA': 'Popular',
    'ROMERO HERNÁNDEZ': 'Popular',
    'ROMERO RODRÍGUEZ': 'Popular',
    'ROMERO SÁNCHEZ': 'Popular',
    'ROMINGUERA SALAZAR': 'Socialista',
    'RUANO GARCÍA': 'Popular',
    'RUFIÁN ROMERO': 'ERC',
    'RUIZ I CARBONELL': 'Socialista',
    'SÁENZ DE SANTAMARÍA ANTÓN': 'Popular',
    'SAGASTIZABAL UNZETABARRENETXEA': 'EAJ-PNV',
    'SAHUQUILLO GARCÍA': 'Socialista',
    'SALUD ARESTE': 'Podemos',
    'SALVADOR ARMENDÁRIZ': 'Mixto',
    'SALVADOR GARCÍA': 'Ciudadanos',
    'SALVADOR I DUCH': 'ERC',
    'SÁNCHEZ AMOR': 'Socialista',
    'SÁNCHEZ MAROTO': 'Podemos',
    'SÁNCHEZ MELERO': 'Podemos',
    'SÁNCHEZ PÉREZ-CASTEJÓN': 'Socialista',
    'SÁNCHEZ SERNA': 'Podemos',
    'SÁNCHEZ-CAMACHO PÉREZ': 'Popular',
    'SANTA ANA FERNÁNDEZ': 'Popular',
    'SANTOS ITOIZ': 'Podemos',
    'SAURA GARCÍA': 'Socialista',
    'SERRADA PARIENTE': 'Socialista',
    'SERRANO JIMÉNEZ': 'Socialista',
    'SERRANO MARTÍNEZ': 'Socialista',
    'SIBINA CAMPS': 'Podemos',
    'SICILIA ALFÉREZ': 'Socialista',
    'SIERRA ROJAS': 'Socialista',
    'SIMANCAS SIMANCAS': 'Socialista',
    'SIXTO IGLESIAS': 'Podemos',
    'SORLÍ FRESQUET': 'Mixto',
    'SUÁREZ LAMATA': 'Popular',
    'SUCH PALOMARES': 'Socialista',
    'SUMELZO JORDÁN': 'Socialista',
    'SURRA SPADEA': 'ERC',
    'TARDÀ I COMA': 'ERC',
    'TARNO BLANCO': 'Popular',
    'TELECHEA I LOZANO': 'ERC',
    'TEN OLIVER': 'Ciudadanos',
    'TERRÓN BERBEL': 'Podemos',
    'TORRES FERNÁNDEZ': 'Popular',
    'TORRES HERRERA': 'Popular',
    'TORRES MORA': 'Socialista',
    'TORRES TEJADA': 'Popular',
    'TREMIÑO GÓMEZ': 'Popular',
    'TREVÍN LOMBÁN': 'Socialista',
    'TUNDIDOR MORENO': 'Socialista',
    'URQUIZU SANCHO': 'Socialista',
    'VALIDO PÉREZ': 'Podemos',
    'VALMAÑA OCHAÍTA': 'Popular',
    'VAÑÓ FERRE': 'Popular',
    'VÁZQUEZ BLANCO': 'Popular',
    'VÁZQUEZ ROJAS': 'Popular',
    'VELASCO BAIDES': 'Socialista',
    'VENDRELL GARDEÑES': 'Podemos',
    'VERA PRÓ': 'Popular',
    'VERA RUIZ-HERRERA': 'Podemos',
    'VIDAL SÁEZ': 'Podemos',
    'VIEJO VIÑAS': 'Podemos',
    'VILA GÓMEZ': 'Podemos',
    'VILLALOBOS TALERO': 'Popular',
    'VILLEGAS PÉREZ': 'Ciudadanos',
    'VISO DIÉGUEZ': 'Popular',
    'XUCLÀ I COSTA': 'Mixto',
    'YLLANES SUÁREZ': 'Podemos',
    'ZARAGOZA ALONSO': 'Socialista',
    'ZOIDO ÁLVAREZ': 'Popular',
    'ZURITA EXPÓSITO': 'Popular',
}

# XIII LEGISLATURA (2019)
DICCIONARIO_XIII_LEGISLATURA = {
    'ABADES MARTÍNEZ': 'Popular',
    'ÁBALOS MECO': 'Mixto',
    'ABASCAL CONDE': 'VOX',
    'ACEDO REYES': 'Popular',
    'ACEVES GALINDO': 'Socialista',
    'ADRIO TARACIDO': 'Socialista',
    'AGIRRETXEA URRESTI': 'EAJ-PNV',
    'AGÜERA GAGO': 'Popular',
    'AGUIRRE GIL DE BIEDMA': 'VOX',
    'AIZCORBE TORRA': 'VOX',
    'AIZPURUA ARZALLUS': 'EH Bildu',
    'ALCARAZ MARTOS': 'VOX',
    'ALFONSO CENDÓN': 'Socialista',
    'ALFONSO SILVESTRE': 'Popular',
    'ALÍA AGUADO': 'Popular',
    'ALMIRÓN RUIZ': 'Socialista',
    'ALMODÓVAR SÁNCHEZ': 'Socialista',
    'ALONSO CANTORNÉ': 'Plurinacional SUMAR',
    'ALÓS LÓPEZ': 'Popular',
    'ÁLVAREZ DE TOLEDO PERALTA-RAMOS': 'Popular',
    'ÁLVAREZ FANJUL': 'Popular',
    'ÁLVAREZ GONZÁLEZ': 'Socialista',
    'ÁLVARO VIDAL': 'ERC',
    'ANDALA UBBI': 'Plurinacional SUMAR',
    'ANDRÉS AÑÓN': 'Socialista',
    'ANTUÑANO COLINA': 'Socialista',
    'ARAGONÉS MENDIGUCHÍA': 'Popular',
    'ARANDA VARGAS': 'Socialista',
    'ARGOTA CASTRO': 'Socialista',
    'ARGÜELLES GARCÍA': 'Popular',
    'ARMARIO GONZÁLEZ': 'VOX',
    'ARMENGOL SOCIAS': 'Socialista',
    'ARRABAS MAROTO': 'Socialista',
    'ASARTA CUEVAS': 'VOX',
    'AZORÍN SALAR': 'Socialista',
    'BADIA CASAS': 'Plurinacional SUMAR',
    'BARRIO BAROJA': 'Popular',
    'BAYÓN ROLO': 'Popular',
    'BEAMONTE MESA': 'Popular',
    'BELARRA URTEAGA': 'Mixto',
    'BELDA PÉREZ-PEDRERO': 'Popular',
    'BELMONTE GÓMEZ': 'Popular',
    'BENDODO BENASAYAG': 'Popular',
    'BERMÚDEZ DE CASTRO FERNÁNDEZ': 'Popular',
    'BLANCO ARRÚE': 'Socialista',
    'BLANQUER ALCARAZ': 'Socialista',
    'BOADA DANÉS': 'Plurinacional SUMAR',
    'BOLAÑOS GARCÍA': 'Socialista',
    'BORREGO CORTÉS': 'Popular',
    'BRAVO BAENA': 'Popular',
    'CABEZÓN CASAS': 'Popular',
    'CACHO ISLA': 'Socialista',
    'CALVO GÓMEZ': 'Junts per Catalunya',
    'CAMINO MIÑANA': 'Socialista',
    'CAMPOS ASENSI': 'VOX',
    'CANELO MATITO': 'Socialista',
    'CANTALAPIEDRA ÁLVAREZ': 'Popular',
    'CARAZO HERMOSO': 'Popular',
    'CARBALLEDO BERLANGA': 'Popular',
    'CASTILLA ÁLVAREZ': 'Socialista',
    'CATALÁN HIGUERAS': 'Mixto',
    'CAVACASILLAS RODRÍGUEZ': 'Popular',
    'CELAYA BREY': 'Popular',
    'CERCAS MENA': 'Socialista',
    'CERVERA PINART': 'Junts per Catalunya',
    'CHAMORRO DELMO': 'VOX',
    'CLAVELL LÓPEZ': 'Popular',
    'CLEMENTE MUÑOZ': 'Popular',
    'COBO CARMONA': 'Socialista',
    'COBO PÉREZ': 'Socialista',
    'COBO VEGA': 'Popular',
    'COFIÑO FERNÁNDEZ': 'Plurinacional SUMAR',
    'CONDE BAJÉN': 'Popular',
    'CONDE LÓPEZ': 'Popular',
    'CONESA COMA': 'Socialista',
    'CORTÉS CARBALLO': 'Popular',
    'CORUJO BERRIEL': 'Socialista',
    'CRESPÍN RUBIO': 'Socialista',
    'CRUSET DOMÈNECH': 'Junts per Catalunya',
    'CRUZ SANTANA': 'Socialista',
    'CRUZ-GUZMÁN GARCÍA': 'Popular',
    'CUESTA RODRÍGUEZ': 'Popular',
    'CUEVAS LARROSA': 'Popular',
    'DE LA ROSA BAENA': 'Socialista',
    'DE LAS CUEVAS CORTÉS': 'Popular',
    'DE LOS SANTOS GONZÁLEZ': 'Popular',
    'DE LUNA TOBARRA': 'Popular',
    'DE MEER MÉNDEZ': 'VOX',
    'DE ROSA TORNER': 'Popular',
    'DEL VALLE RODRÍGUEZ': 'VOX',
    'DELGADO ARCE': 'Popular',
    'DELGADO-TARAMONA HERNÁNDEZ': 'Popular',
    'DÍAZ MARÍN': 'Socialista',
    'DÍAZ PÉREZ': 'Plurinacional SUMAR',
    'DIOUF DIOH': 'Socialista',
    'ESTREMS FAYOS': 'ERC',
    'FABRA PART': 'Popular',
    'FAGÚNDEZ CAMPO': 'Socialista',
    'FANECA LÓPEZ': 'Socialista',
    'FERNÁNDEZ BENÉITEZ': 'Socialista',
    'FERNÁNDEZ GONZÁLEZ': 'Popular',
    'FERNÁNDEZ HERNÁNDEZ': 'VOX',
    'FERNÁNDEZ RÍOS': 'VOX',
    'FIGAREDO ÁLVAREZ-SALA': 'VOX',
    'FLORES JUBERÍAS': 'VOX',
    'FLORIANO CORRALES': 'Popular',
    'FOLCH BLANC': 'Popular',
    'FRANCO GONZÁLEZ': 'Popular',
    'FULLAONDO LA CRUZ': 'EH Bildu',
    'FÚNEZ DE GREGORIO': 'Popular',
    'GALLARDO BARRENA': 'Popular',
    'GAMARRA RUIZ-CLAVIJO': 'Popular',
    'GARCÍA ADANERO': 'Popular',
    'GARCÍA CHAVARRÍA': 'Socialista',
    'GARCÍA FÉLIX': 'Popular',
    'GARCÍA GOMIS': 'VOX',
    'GARCÍA GURRUTXAGA': 'Socialista',
    'GARCÍA LÓPEZ': 'Socialista',
    'GARCÍA MORÍS': 'Socialista',
    'GARRE MURCIA': 'Popular',
    'GARRIDO JIMÉNEZ': 'Socialista',
    'GARRIDO VALENZUELA': 'Popular',
    'GAVIN I VALLS': 'Junts per Catalunya',
    'GIL DE REBOLEÑO LASTORTRE': 'Plurinacional SUMAR',
    'GIL LÁZARO': 'VOX',
    'GIL SANTIAGO': 'Popular',
    'GÓMEZ PIÑA': 'Socialista',
    'GONZÁLEZ GRACIA': 'Socialista',
    'GONZÁLEZ LÓPEZ': 'Plurinacional SUMAR',
    'GONZÁLEZ VÁZQUEZ': 'Popular',
    'GONZÁLEZ-ROBATTO PEROTE': 'VOX',
    'GRACIA BLANCO': 'Socialista',
    'GRANOLLERS CUNILLERA': 'ERC',
    'GUARDIOLA SALMERÓN': 'Popular',
    'GUIJARRO GARCÍA': 'Plurinacional SUMAR',
    'GUINART MORENO': 'Socialista',
    'GUTIÉRREZ PRIETO': 'Socialista',
    'GUTIÉRREZ SANTIAGO': 'Socialista',
    'HERNÁNDEZ QUERO': 'VOX',
    'HERNANDO FRAILE': 'Popular',
    'HERRERA GARCÍA': 'Socialista',
    'HERRERO BONO': 'Popular',
    'HISPÁN IGLESIAS DE USSEL': 'Popular',
    'HITA TÉLLEZ': 'Socialista',
    'HOCES ÍÑIGUEZ': 'VOX',
    'HOYO JULIÁ': 'Popular',
    'IBÁÑEZ HERNANDO': 'Popular',
    'IBÁÑEZ MEZQUITA': 'Plurinacional SUMAR',
    'INIESTA EGIDO': 'Socialista',
    'IÑARRITU GARCÍA': 'EH Bildu',
    'JEREZ ANTEQUERA': 'Socialista',
    'JIMÉNEZ LINUESA': 'Popular',
    'JÓDAR PÉREZ': 'Socialista',
    'JORDÀ I ROURA': 'ERC',
    'LAGO PEÑAS': 'Plurinacional SUMAR',
    'LAMUÀ ESTAÑOL': 'Socialista',
    'LEAL FERNÁNDEZ': 'Socialista',
    'LEGARDA URIARTE': 'EAJ-PNV',
    'LIMA GARCÍA': 'Popular',
    'LLAMAZARES DOMINGO': 'Popular',
    'LÓPEZ ÁLVAREZ': 'Socialista',
    'LÓPEZ CANO': 'Socialista',
    'LÓPEZ MARAVER': 'VOX',
    'LÓPEZ TAGLIAFICO': 'Plurinacional SUMAR',
    'LÓPEZ ZAMORA': 'Socialista',
    'LORENTE ANAYA': 'Popular',
    'LORENZO CAZORLA': 'Socialista',
    'LOSADA FERNÁNDEZ': 'Socialista',
    'MACÍAS GATA': 'Popular',
    'MADRENAS I MIR': 'Junts per Catalunya',
    'MADRID OLMO': 'Popular',
    'MALDONADO LÓPEZ': 'Socialista',
    'MARCOS ORTEGA': 'Popular',
    'MARÍ BOSÓ': 'Popular',
    'MARISCAL ANAYA': 'Popular',
    'MARISCAL ZABALA': 'VOX',
    'MARQUÉS ATÉS': 'Socialista',
    'MARTÍN BLANCO': 'Popular',
    'MARTÍN GARCÍA': 'Popular',
    'MARTÍN MARTÍNEZ': 'Socialista',
    'MARTÍN RODRÍGUEZ': 'Socialista',
    'MARTÍN URRIZA': 'Plurinacional SUMAR',
    'MARTÍNEZ BARBERO': 'Plurinacional SUMAR',
    'MARTÍNEZ GÓMEZ': 'Popular',
    'MARTÍNEZ HIERRO': 'Plurinacional SUMAR',
    'MARTÍNEZ LABELLA': 'Popular',
    'MARTÍNEZ RAMÍREZ': 'Socialista',
    'MARTÍNEZ SALMERÓN': 'Socialista',
    'MARTÍNEZ SEIJO': 'Socialista',
    'MATUTE GARCÍA DE JALÓN': 'EH Bildu',
    'MAYORAL DE LAMO': 'Socialista',
    'MAYORAL PÉREZ': 'Socialista',
    'MEJÍAS SÁNCHEZ': 'VOX',
    'MELGAREJO MORENO': 'Popular',
    'MELLADO SIERRA': 'Socialista',
    'MÉNDEZ MONASTERIO': 'VOX',
    'MERCADAL BAQUERO': 'Socialista',
    'MERINO MARTÍNEZ': 'Popular',
    'MESQUIDA MAYANS': 'Popular',
    'MICÓ MICÓ': 'Mixto',
    'MÍNGUEZ GARCÍA': 'Socialista',
    'MOLINA LEÓN': 'Popular',
    'MONEO DÍEZ': 'Popular',
    'MONTÁVEZ AGUILLAUME': 'Socialista',
    'MONTERO CUADRADO': 'Socialista',
    'MONTESINOS DE MIGUEL': 'Popular',
    'MORALEJA GÓMEZ': 'Popular',
    'MORALES ÁLVAREZ': 'Socialista',
    'MORENO BORRÁS': 'Popular',
    'MORENO FERNÁNDEZ': 'Socialista',
    'MORO ALMARAZ': 'Popular',
    'MUÑOZ ABRINES': 'Popular',
    'MUÑOZ DE LA IGLESIA': 'Popular',
    'NACARINO-BRABO JIMÉNEZ': 'Popular',
    'NARBONA RUIZ': 'Socialista',
    'NASARRE OLIVA': 'Socialista',
    'NAVARRO LACOBA': 'Popular',
    'NAVARRO LÓPEZ': 'Popular',
    'NOGUERAS I CAMERO': 'Junts per Catalunya',
    'NORIEGA GÓMEZ': 'Popular',
    'NÚÑEZ FEIJÓO': 'Popular',
    'NÚÑEZ GUIJARRO': 'Popular',
    'OGOU I CORBI': 'Plurinacional SUMAR',
    'OLANO VELA': 'Popular',
    'ORTEGA SMITH-MOLINA': 'VOX',
    'OTERO GABIRONDO': 'EH Bildu',
    'OTERO GARCÍA': 'Socialista',
    'OTERO RODRÍGUEZ': 'Socialista',
    'PAGÈS I MASSÓ': 'Junts per Catalunya',
    'PALENCIA RUBIO': 'Popular',
    'PANIAQUA NÚÑEZ': 'Popular',
    'PARÉ AREGALL': 'Socialista',
    'PARRA APARICIO': 'Popular',
    'PARRA GALLEGO': 'Popular',
    'PASCUAL ROCAMORA': 'Popular',
    'PEDREÑO MOLINA': 'Popular',
    'PEÑA CAMARERO': 'Socialista',
    'PEREA I CONILLAS': 'Socialista',
    'PÉREZ CORONADO': 'Popular',
    'PÉREZ LÓPEZ': 'Popular',
    'PÉREZ ORTIZ': 'Socialista',
    'PÉREZ RECUERDA': 'Popular',
    'PISARELLO PRADOS': 'Plurinacional SUMAR',
    'PLAZA GARCÍA': 'Socialista',
    'POBLADOR PACHECO': 'Socialista',
    'POSE MESURA': 'Socialista',
    'POZUETA FERNÁNDEZ': 'EH Bildu',
    'PRIETO SERRANO': 'Popular',
    'PUENTE SANTIAGO': 'Socialista',
    'PUEO SANZ': 'Plurinacional SUMAR',
    'PUYO FRAGA': 'Popular',
    'QUINTANA CARBALLO': 'Popular',
    'QUINTANILLA NAVARRO': 'Popular',
    'QUINTERO HERNÁNDEZ': 'Socialista',
    'RALLO LOMBARTE': 'Socialista',
    'RAMAJO PRADA': 'Popular',
    'RAMIJO CARNER': 'Socialista',
    'RAMÍREZ DEL RÍO': 'VOX',
    'RAMÍREZ MARTÍN': 'Popular',
    'RAMÍREZ MORENO': 'Socialista',
    'RAMOS ESTEBAN': 'Socialista',
    'RECAS MARTÍN': 'Plurinacional SUMAR',
    'REDONDO CÁRDENAS': 'Socialista',
    'REGO CANDAMIL': 'Mixto',
    'RENTERIA LASANTA': 'EAJ-PNV',
    'REQUENA RUIZ': 'Popular',
    'REY DE LAS HERAS': 'Socialista',
    'REYNAL REILLO': 'Popular',
    'RIVERA ARIAS': 'Plurinacional SUMAR',
    'RIVES ARCAYNA': 'Socialista',
    'ROBLES LÓPEZ': 'VOX',
    'RODRÍGUEZ ALMEIDA': 'VOX',
    'RODRÍGUEZ CALLEJA': 'Popular',
    'RODRÍGUEZ DE MILLÁN PARRO': 'VOX',
    'RODRÍGUEZ GÓMEZ DE CELIS': 'Socialista',
    'RODRÍGUEZ HERRER': 'Popular',
    'RODRÍGUEZ PALACIOS': 'Socialista',
    'RODRÍGUEZ SALAS': 'Socialista',
    'RODRÍGUEZ SERRA': 'Popular',
    'RODRÍGUEZ SUÁREZ': 'Socialista',
    'ROJAS GARCÍA': 'Popular',
    'ROJAS MANRIQUE': 'Popular',
    'ROJO BLAS': 'Socialista',
    'ROMÁN JASANADA': 'Popular',
    'ROMANÍ CANTERA': 'Popular',
    'ROMERO POZO': 'Socialista',
    'ROMERO VILCHES': 'VOX',
    'ROS MARTÍNEZ': 'Socialista',
    'RUEDA PERELLÓ': 'VOX',
    'RUFIÁN ROMERO': 'ERC',
    'RUIZ BOIX': 'Socialista',
    'RUIZ DE DIEGO': 'Socialista',
    'RUIZ SOLÁS': 'VOX',
    'SÁEZ ALONSO-MUÑUMER': 'VOX',
    'SÁEZ CRUZ': 'Socialista',
    'SAGASTIZABAL UNZETABARRENETXEA': 'EAJ-PNV',
    'SAHUQUILLO GARCÍA': 'Socialista',
    'SAINZ MARTÍN': 'Socialista',
    'SALVADOR I DUCH': 'ERC',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ GARCÍA': 'VOX',
    'SÁNCHEZ OJEDA': 'Popular',
    'SÁNCHEZ PÉREZ': 'Popular',
    'SÁNCHEZ PÉREZ-CASTEJÓN': 'Socialista',
    'SÁNCHEZ SERNA': 'Mixto',
    'SÁNCHEZ SIERRA': 'Popular',
    'SÁNCHEZ TORREGROSA': 'Popular',
    'SANCHO ÍÑIGUEZ': 'Socialista',
    'SANTANA AGUILERA': 'Socialista',
    'SANTANA PERERA': 'Mixto',
    'SANTIAGO ROMERO': 'Plurinacional SUMAR',
    'SANTOS MARAVER': 'Plurinacional SUMAR',
    'SANZ MARTÍNEZ': 'Socialista',
    'SARRIÀ MORELL': 'Socialista',
    'SASTRE UYÁ': 'Popular',
    'SAYAS LÓPEZ': 'Popular',
    'SÉMPER PASCUAL': 'Popular',
    'SENDEROS ORAÁ': 'Socialista',
    'SERRADA PARIENTE': 'Socialista',
    'SERRANO MARTÍNEZ': 'Socialista',
    'SIERRA CABALLERO': 'Plurinacional SUMAR',
    'SIMANCAS SIMANCAS': 'Socialista',
    'SIMARRO VICENS': 'Popular',
    'SOLDEVILLA NOVIALS': 'Socialista',
    'SOLER MUR': 'Socialista',
    'TABOADELA ÁLVAREZ': 'Socialista',
    'TARNO BLANCO': 'Popular',
    'TELLADO FILGUEIRA': 'Popular',
    'TENIENTE SÁNCHEZ': 'Popular',
    'TOMÁS OLIVARES': 'Popular',
    'TORRES TEJADA': 'Popular',
    'TRENZANO RUBIO': 'Socialista',
    'URIARTE BENGOECHEA': 'Popular',
    'VALERO MORALES': 'Plurinacional SUMAR',
    'VALIDO GARCÍA': 'Mixto',
    'VALLUGERA BALAÑÀ': 'ERC',
    'VAQUERO MONTERO': 'EAJ-PNV',
    'VARELA PAZOS': 'Popular',
    'VÁZQUEZ BLANCO': 'Popular',
    'VÁZQUEZ JIMÉNEZ': 'Popular',
    'VEDRINA CONESA': 'Popular',
    'VELARDE GÓMEZ': 'Mixto',
    'VELASCO MORILLO': 'Popular',
    'VELASCO RETAMOSA': 'Popular',
    'VERANO DOMÍNGUEZ': 'Popular',
    'VERDEJO VICENTE': 'Socialista',
    'VIDAL MATAS': 'Plurinacional SUMAR',
    'VIDAL SÁEZ': 'Plurinacional SUMAR',
    'ZARAGOZA ALONSO': 'Socialista',
}

# XIV LEGISLATURA (2019-2023)
DICCIONARIO_XIV_LEGISLATURA = {
    'ÁBALOS MECO': 'Socialista',
    'ABASCAL CONDE': 'VOX',
    'ACEVES GALINDO': 'Socialista',
    'AGIRRETXEA URRESTI': 'EAJ-PNV',
    'AIZCORBE TORRA': 'VOX',
    'AIZPURUA ARZALLUS': 'EH Bildu',
    'ALCARAZ MARTOS': 'VOX',
    'ALFONSO CENDÓN': 'Socialista',
    'ALMODÓBAR BARCELÓ': 'Popular',
    'ALONSO PÉREZ': 'Popular',
    'ALONSO SUÁREZ': 'Socialista',
    'ALONSO-CUEVILLAS I SAYROL': 'Plural',
    'ÁLVAREZ DE TOLEDO PERALTA-RAMOS': 'Popular',
    'ÁLVAREZ FANJUL': 'Popular',
    'ÁLVAREZ I GARCÍA': 'ERC',
    'ANDRÉS AÑÓN': 'Socialista',
    'ANDRÉS BAREA': 'Socialista',
    'ANGUITA PÉREZ': 'Socialista',
    'ANGULO ROMERO': 'Popular',
    'ANTÓN CACHO': 'Socialista',
    'ARAGONÉS MENDIGUCHÍA': 'Popular',
    'ARANDA VARGAS': 'Socialista',
    'ARANGÜENA FERNÁNDEZ': 'Socialista',
    'ARAUJO MORALES': 'Socialista',
    'ARRABAS MAROTO': 'Socialista',
    'ARRIMADAS GARCÍA': 'Ciudadanos',
    'ASARTA CUEVAS': 'VOX',
    'ASENS LLODRÀ': 'Podemos',
    'AZORÍN SALAR': 'Socialista',
    'BAL FRANCÉS': 'Ciudadanos',
    'BALDOVÍ RODA': 'Plural',
    'BAÑOS RUIZ': 'Socialista',
    'BARANDIARAN BENITO': 'EAJ-PNV',
    'BAS CORUGEIRA': 'Popular',
    'BASSA COLL': 'ERC',
    'BATET LAMAÑA': 'Socialista',
    'BEL ACCENSI': 'Plural',
    'BELARRA URTEAGA': 'Podemos',
    'BELTRÁN VILLALBA': 'Popular',
    'BERJA VEGA': 'Socialista',
    'BERMÚDEZ DE CASTRO FERNÁNDEZ': 'Popular',
    'BETORET COLL': 'Popular',
    'BLANQUER ALCARAZ': 'Socialista',
    'BOADELLA ESTEVE': 'Plural',
    'BORRÀS CASTANYER': 'Plural',
    'BORRÁS PABÓN': 'VOX',
    'BORREGO CORTÉS': 'Popular',
    'BOTELLA GÓMEZ': 'Socialista',
    'BOTRAN PAHISSA': 'Mixto',
    'BRAVO BARCO': 'Socialista',
    'BUENO CAMPANARIO': 'Socialista',
    'BUENO PINTO': 'Podemos',
    'BUSTAMANTE MARTÍN': 'Podemos',
    'CABALLERO GUTIÉRREZ': 'Socialista',
    'CABEZÓN CASAS': 'Popular',
    'CALLEJAS CANO': 'Popular',
    'CALVO GÓMEZ': 'Plural',
    'CALVO LISTE': 'VOX',
    'CALVO POYATO': 'Socialista',
    'CAMBRONERO PIQUERAS': 'Mixto',
    'CAMPO MORENO': 'Socialista',
    'CANALES DUQUE': 'Socialista',
    'CANCELA RODRÍGUEZ': 'Socialista',
    'CANTERA DE CASTRO': 'Socialista',
    'CAÑADELL SALVIA': 'Plural',
    'CAÑIZARES PACHECO': 'VOX',
    'CAPDEVILA I ESTEVE': 'ERC',
    'CARAZO HERMOSO': 'Popular',
    'CARCEDO ROCES': 'Socialista',
    'CARRILLO DE LOS REYES': 'Socialista',
    'CARVALHO DANTAS': 'ERC',
    'CASADO BLANCO': 'Popular',
    'CASARES HONTAÑÓN': 'Socialista',
    'CASERO ÁVILA': 'Popular',
    'CASTELLÓN RUBIO': 'Popular',
    'CASTILLO LÓPEZ': 'Popular',
    'CELAA DIÉGUEZ': 'Socialista',
    'CERDÁN LEÓN': 'Socialista',
    'CERQUEIRO GONZÁLEZ': 'Socialista',
    'CHAMORRO DELMO': 'VOX',
    'CLAVELL LÓPEZ': 'Popular',
    'CONSTENLA CARBÓN': 'Popular',
    'CONTRERAS PELÁEZ': 'VOX',
    'CORREDOR SIERRA': 'Socialista',
    'CORTÉS CARBALLO': 'Popular',
    'CORTÉS GÓMEZ': 'Podemos',
    'CRESPÍN RUBIO': 'Socialista',
    'CRUZ-GUZMÁN GARCÍA': 'Popular',
    'CUATRECASAS ASUA': 'Socialista',
    'DE LAS HERAS FERNÁNDEZ': 'VOX',
    'DE LUNA TOBARRA': 'Popular',
    'DE MEER MÉNDEZ': 'VOX',
    'DE QUINTO ROMERO': 'Ciudadanos',
    'DEL VALLE DE ISCAR': 'Socialista',
    'DEL VALLE RODRÍGUEZ': 'VOX',
    'DELGADO ARCE': 'Popular',
    'DELGADO GARCÍA': 'Socialista',
    'DELGADO RAMOS': 'Podemos',
    'DÍAZ GÓMEZ': 'Ciudadanos',
    'DÍAZ PÉREZ': 'Podemos',
    'DÍAZ RODRÍGUEZ': 'Socialista',
    'DIOUF DIOH': 'Socialista',
    'DUQUE DUQUE': 'Socialista',
    'DUQUE MORÁN': 'Socialista',
    'DURÁN PERALTA': 'Socialista',
    'ECHÁNIZ SALGADO': 'Popular',
    'ECHENIQUE ROBBA': 'Podemos',
    'ELIZO SERRANO': 'Podemos',
    'ELORRIAGA PISARIK': 'Popular',
    'ELORZA GONZÁLEZ': 'Socialista',
    'ERITJA CIURÓ': 'ERC',
    'ERREJÓN GALVÁN': 'Plural',
    'ESPAÑA REINA': 'Popular',
    'ESPEJO-SAAVEDRA CONESA': 'Ciudadanos',
    'ESPINOSA DE LOS MONTEROS DE SIMÓN': 'VOX',
    'ESTEBAN BRAVO': 'EAJ-PNV',
    'ESTEBAN CALONJE': 'VOX',
    'FAGÚNDEZ CAMPO': 'Socialista',
    'FANECA LÓPEZ': 'Socialista',
    'FERNÁNDEZ BENÉITEZ': 'Socialista',
    'FERNÁNDEZ CASERO': 'Socialista',
    'FERNÁNDEZ CASTAÑÓN': 'Podemos',
    'FERNÁNDEZ HERNÁNDEZ': 'VOX',
    'FERNÁNDEZ PÉREZ': 'Mixto',
    'FERNÁNDEZ RÍOS': 'VOX',
    'FERNÁNDEZ-LOMANA GUTIÉRREZ': 'VOX',
    'FERNÁNDEZ-ROCA SUÁREZ': 'VOX',
    'FERRER TESORO': 'Socialista',
    'FIGAREDO ÁLVAREZ-SALA': 'VOX',
    'FRANCO CARMONA': 'Podemos',
    'FRANCO PARDO': 'Socialista',
    'FUENTES CURBELO': 'Socialista',
    'GAGO BUGARÍN': 'Popular',
    'GAMARRA RUIZ-CLAVIJO': 'Popular',
    'GAMAZO MICÓ': 'Popular',
    'GARCÉS SANAGUSTÍN': 'Popular',
    'GARCÍA ADANERO': 'Mixto',
    'GARCÍA CHAVARRÍA': 'Socialista',
    'GARCÍA DÍEZ': 'Popular',
    'GARCÍA EGEA': 'Popular',
    'GARCÍA GÓMEZ': 'Socialista',
    'GARCÍA GURRUTXAGA': 'Socialista',
    'GARCÍA LÓPEZ': 'Socialista',
    'GARCÍA MORÍS': 'Socialista',
    'GARCÍA NIETO': 'Podemos',
    'GARCÍA PUIG': 'Podemos',
    'GARCÍA RODRÍGUEZ': 'Popular',
    'GARCÍA TEJERINA': 'Popular',
    'GARCÍA-PELAYO JURADO': 'Popular',
    'GARRIDO GUTIÉRREZ': 'Podemos',
    'GARRIGA VAZ DE CONCICAO': 'VOX',
    'GARZÓN ESPINOSA': 'Podemos',
    'GÁZQUEZ COLLADO': 'Popular',
    'GESTOSO DE MIGUEL': 'VOX',
    'GIL LÁZARO': 'VOX',
    'GIMÉNEZ GIMÉNEZ': 'Ciudadanos',
    'GÓMEZ HERNÁNDEZ': 'Socialista',
    'GÓMEZ-REINO VARELA': 'Podemos',
    'GONZÁLEZ CABALLERO': 'Socialista',
    'GONZÁLEZ COELLO DE PORTUGAL': 'VOX',
    'GONZÁLEZ GUINDA': 'Popular',
    'GONZÁLEZ LASO': 'Socialista',
    'GONZÁLEZ MUÑOZ': 'Popular',
    'GONZÁLEZ PÉREZ': 'Socialista',
    'GONZÁLEZ RAMOS': 'Socialista',
    'GONZÁLEZ TEROL': 'Popular',
    'GONZÁLEZ VÁZQUEZ': 'Popular',
    'GOROSPE ELEZCANO': 'EAJ-PNV',
    'GRANDE-MARLASKA GÓMEZ': 'Socialista',
    'GRANOLLERS CUNILLERA': 'ERC',
    'GUAITA ESTERUELAS': 'Socialista',
    'GUERRA LÓPEZ': 'Socialista',
    'GUIJARRO CEBALLOS': 'Socialista',
    'GUIJARRO GARCÍA': 'Podemos',
    'GUINART MORENO': 'Socialista',
    'GUIRAO CABRERA': 'Socialista',
    'GUITARTE GIMENO': 'Mixto',
    'GUTIÉRREZ DÍAZ DE OTAZU': 'Popular',
    'GUTIÉRREZ PRIETO': 'Socialista',
    'GUTIÉRREZ SALINAS': 'Socialista',
    'GUTIÉRREZ VIVAS': 'Ciudadanos',
    'HERNANZ COSTA': 'Socialista',
    'HERRERO BONO': 'Popular',
    'HISPÁN IGLESIAS DE USSEL': 'Popular',
    'HONRUBIA HURTADO': 'Podemos',
    'HOYO JULIÁ': 'Popular',
    'HURTADO ZURERA': 'Socialista',
    'IGLESIAS TURRIÓN': 'Podemos',
    'ILLAMOLA DAUSÀ': 'Plural',
    'ILLUECA BALLESTER': 'Podemos',
    'IÑARRITU GARCÍA': 'EH Bildu',
    'IZQUIERDO RONCERO': 'Socialista',
    'JARA MORENO': 'VOX',
    'JEREZ JUAN': 'Popular',
    'JIMÉNEZ LINUESA': 'Popular',
    'JIMÉNEZ REVUELTA': 'VOX',
    'JIMÉNEZ-BECERRIL BARRIO': 'Popular',
    'JOVER DÍAZ': 'Podemos',
    'LAMUÀ ESTAÑOL': 'Socialista',
    'LASTRA FERNÁNDEZ': 'Socialista',
    'LEAL FERNÁNDEZ': 'Socialista',
    'LEDESMA MARTÍN': 'Popular',
    'LEGARDA URIARTE': 'EAJ-PNV',
    'LIMA CID': 'Socialista',
    'LÓPEZ ÁLVAREZ': 'Socialista',
    'LÓPEZ CANO': 'Socialista',
    'LÓPEZ DE URALDE GARMENDIA': 'Podemos',
    'LÓPEZ DOMÍNGUEZ': 'Podemos',
    'LÓPEZ MARAVER': 'VOX',
    'LÓPEZ MOYA': 'Popular',
    'LÓPEZ SOMOZA': 'Socialista',
    'LÓPEZ ZAMORA': 'Socialista',
    'LÓPEZ-BAS VALERO': 'Ciudadanos',
    'LORITE LORITE': 'Popular',
    'LOSADA FERNÁNDEZ': 'Socialista',
    'MAESTRO MOLINER': 'Podemos',
    'MANGLANO ALBACAR': 'Popular',
    'MANSO OLIVAR': 'VOX',
    'MÁÑEZ RODRÍGUEZ': 'Socialista',
    'MARCOS DOMÍNGUEZ': 'Popular',
    'MARCOS MOYANO': 'Popular',
    'MARCOS ORTEGA': 'Popular',
    'MARGALL SASTRE': 'ERC',
    'MARÍ KLOSE': 'Socialista',
    'MARISCAL ANAYA': 'Popular',
    'MARISCAL ZABALA': 'VOX',
    'MAROTO ILLERA': 'Socialista',
    'MÁRQUEZ GUERRERO': 'Podemos',
    'MARRA DOMÍNGUEZ': 'Socialista',
    'MARRODÁN FUNES': 'Socialista',
    'MARTÍN LLAGUNO': 'Ciudadanos',
    'MARTÍNEZ DE TEJADA PÉREZ': 'Popular',
    'MARTÍNEZ FERRO': 'Popular',
    'MARTÍNEZ GRANADOS': 'Ciudadanos',
    'MARTÍNEZ OBLANCA': 'Mixto',
    'MARTÍNEZ SALMERÓN': 'Socialista',
    'MARTÍNEZ SEIJO': 'Socialista',
    'MATARÍ SÁEZ': 'Popular',
    'MATEU ISTÚRIZ': 'Popular',
    'MATUTE GARCÍA DE JALÓN': 'Euskal Herria Bildu',
    'MAYORAL PERALES': 'Podemos',
    'MAZÓN RAMOS': 'Mixto',
    'MEDEL PÉREZ': 'Podemos',
    'MEIJÓN COUSELO': 'Socialista',
    'MENA ARCA': 'Podemos',
    'MÉNDEZ MONASTERIO': 'VOX',
    'MERINO MARTÍNEZ': 'Popular',
    'MESTRE BAREA': 'VOX',
    'MÍNGUEZ GARCÍA': 'Socialista',
    'MIQUEL I VALENTÍ': 'Plural',
    'MIRALLES MARTÍN': 'VOX',
    'MONEO DÍEZ': 'Popular',
    'MONTERO CUADRADO': 'Socialista',
    'MONTERO GIL': 'Podemos',
    'MONTESINOS AGUAYO': 'Popular',
    'MONTESINOS DE MIGUEL': 'Popular',
    'MONTILLA MARTOS': 'Socialista',
    'MORALEJA GÓMEZ': 'Popular',
    'MORO ALMARAZ': 'Popular',
    'MOVELLÁN LOMBILLA': 'Popular',
    'MUÑOZ DALDA': 'Podemos',
    'MUÑOZ VIDAL': 'Ciudadanos',
    'NARVÁEZ BANDERA': 'Socialista',
    'NASARRE OLIVA': 'Socialista',
    'NAVALPOTRO GÓMEZ': 'Socialista',
    'NAVARRO LACOBA': 'Popular',
    'NAVARRO LÓPEZ': 'Popular',
    'NEVADO DEL CAMPO': 'VOX',
    'NOGUERAS I CAMERO': 'Plural',
    'NUET PUJALS': 'ERC',
    'OLANO VELA': 'Popular',
    'OLONA CHOCLÁN': 'VOX',
    'ORAMAS GONZÁLEZ-MORO': 'Mixto',
    'ORIA LÓPEZ': 'Socialista',
    'ORTEGA DOMÍNGUEZ': 'Socialista',
    'ORTEGA OTERO': 'Socialista',
    'ORTEGA SMITH-MOLINA': 'VOX',
    'ORTIZ GALVÁN': 'Popular',
    'PADILLA RUIZ': 'Socialista',
    'PAGÈS I MASSÓ': 'Plural',
    'PANIAQUA NÚÑEZ': 'Popular',
    'PASTOR JULIÁN': 'Popular',
    'PEDRAJA SÁINZ': 'Socialista',
    'PEDREÑO MOLINA': 'Popular',
    'PEÑA CAMARERO': 'Socialista',
    'PEREA I CONILLAS': 'Socialista',
    'PÉREZ ABELLÁS': 'Socialista',
    'PÉREZ DÍAZ': 'Popular',
    'PÉREZ MERINO': 'Podemos',
    'PÉREZ RECUERDA': 'Popular',
    'PICÓ GARCÉS': 'Plural',
    'PÍRIZ MAYA': 'Popular',
    'PISARELLO PRADOS': 'Podemos',
    'PITA CÁRDENES': 'Mixto',
    'PLANAS PUCHADES': 'Socialista',
    'POLO LLAVATA': 'Socialista',
    'PONS SAMPIETRO': 'Socialista',
    'POSTIGO QUINTANA': 'Popular',
    'POZUETA FERNÁNDEZ': 'Euskal Herria Bildu',
    'PRIETO NIETO': 'Socialista',
    'PROHENS RIGO': 'Popular',
    'PUJOL I FARRÉ': 'ERC',
    'QUEVEDO ITURBE': 'Mixto',
    'QUINTANILLA NAVARRO': 'Popular',
    'RAMALLO VÁZQUEZ': 'Popular',
    'RAMÍREZ CARNER': 'Socialista',
    'RAMÍREZ DEL RÍO': 'VOX',
    'RAMÓN UTRABO': 'Socialista',
    'RAMOS ESTEBAN': 'Socialista',
    'RAMOS RODRÍGUEZ': 'Socialista',
    'RAYA RODRÍGUEZ': 'Socialista',
    'REDONDO CALVILLO': 'Popular',
    'REGO CANDAMIL': 'Plural',
    'RENAU MARTÍNEZ': 'Socialista',
    'REQUEJO NOVOA': 'VOX',
    'REQUENA RUIZ': 'Popular',
    'RIBERA RODRÍGUEZ': 'Socialista',
    'RIOLOBOS REGADERA': 'Popular',
    'ROBLES FERNÁNDEZ': 'Socialista',
    'ROBLES LÓPEZ': 'VOX',
    'RODRÍGUEZ ALMEIDA': 'VOX',
    'RODRÍGUEZ GÓMEZ': 'Socialista',
    'RODRÍGUEZ GÓMEZ DE CELIS': 'Socialista',
    'RODRÍGUEZ HERRER': 'Popular',
    'RODRÍGUEZ RODRÍGUEZ': 'Podemos',
    'RODRÍGUEZ SALAS': 'Socialista',
    'ROJAS GARCÍA': 'Popular',
    'ROMANÍ CANTERA': 'Popular',
    'ROMERO HERNÁNDEZ': 'Popular',
    'ROMERO SÁNCHEZ': 'Popular',
    'ROMERO VILCHES': 'VOX',
    'ROS MARTÍNEZ': 'Socialista',
    'ROSELL AGUILAR': 'Podemos',
    'ROSELLÓ BOUSO': 'Podemos',
    'ROSETY FERNÁNDEZ DE CASTRO': 'VOX',
    'ROSIQUE I SALTOR': 'ERC',
    'RUEDA PERELLÓ': 'VOX',
    'RUFIÁN ROMERO': 'ERC',
    'RUIZ DE PINEDO UNDIANO': 'Euskal Herria Bildu',
    'RUIZ I CARBONELL': 'Socialista',
    'RUIZ LÓPEZ': 'Socialista',
    'RUIZ NAVARRO': 'VOX',
    'RUIZ SOLÁS': 'VOX',
    'SAAVEDRA MUÑOZ': 'Podemos',
    'SABANÉS NADAL': 'Plural',
    'SÁEZ ALONSO-MUÑUMER': 'VOX',
    'SAGASTIZABAL UNZETABARRENETXEA': 'EAJ-PNV',
    'SAHUQUILLO GARCÍA': 'Socialista',
    'SALAZAR RODRÍGUEZ': 'Socialista',
    'SALVÁ VERD': 'VOX',
    'SALVADOR I DUCH': 'ERC',
    'SÁNCHEZ DEL REAL': 'VOX',
    'SÁNCHEZ ESCOBAR': 'Socialista',
    'SÁNCHEZ GARCÍA': 'VOX',
    'SÁNCHEZ JÓDAR': 'Socialista',
    'SÁNCHEZ PÉREZ': 'Popular',
    'SÁNCHEZ PÉREZ-CASTEJÓN': 'Socialista',
    'SÁNCHEZ SERNA': 'Podemos',
    'SANCHO GUARDIA': 'Socialista',
    'SANCHO ÍÑIGUEZ': 'Socialista',
    'SANTAMARÍA RUIZ': 'Popular',
    'SANTIAGO ROMERO': 'Podemos',
    'SARRIÀ MORELL': 'Socialista',
    'SAURA GARCÍA': 'Socialista',
    'SAYAS LÓPEZ': 'Mixto',
    'SEGURA JUST': 'VOX',
    'SENDEROS ORAÁ': 'Socialista',
    'SERRADA PARIENTE': 'Socialista',
    'SERRANO MARTÍNEZ': 'Socialista',
    'SEVA RUIZ': 'Socialista',
    'SICILIA ALFÉREZ': 'Socialista',
    'SIMANCAS SIMANCAS': 'Socialista',
    'SOLER MUR': 'Socialista',
    'SOTO BURILLO': 'Socialista',
    'STEEGMANN OLMEDILLAS': 'VOX',
    'SUÁREZ ILLANA': 'Popular',
    'SUÁREZ LAMATA': 'Popular',
    'SUMELZO JORDÁN': 'Socialista',
    'TAIBO MONELOS': 'Socialista',
    'TARNO BLANCO': 'Popular',
    'TELECHEA I LOZANO': 'ERC',
    'TIRADO OCHOA': 'Popular',
    'TIZÓN VÁZQUEZ': 'Socialista',
    'TOSCANO DE BALBÍN': 'VOX',
    'TRÍAS GIL': 'VOX',
    'URIARTE BENGOCHEA': 'Popular',
    'URIARTE TORREALDAY': 'Podemos',
    'UTRILLA CANO': 'VOX',
    'VALERIO CORDERO': 'Socialista',
    'VALLUGERA BALAÑÀ': 'ERC',
    'VÁZQUEZ BLANCO': 'Popular',
    'VEGA ARIAS': 'VOX',
    'VEHÍ CANTENYS': 'Mixto',
    'VELARDE GÓMEZ': 'Podemos',
    'VELASCO MORILLO': 'Popular',
    'VERA RUIZ-HERRERA': 'Podemos',
    'VICENTE VIONDI': 'Socialista',
    'VIDAL SÁEZ': 'Podemos',
    'VILCHES RUIZ': 'Socialista',
    'VILLAGRASA QUERO': 'Socialista',
    'ZAMARRÓN MORENO': 'Socialista',
    'ZAMBRANO GARCÍA-RAEZ': 'VOX',
    'ZAPATA SIMÓN': 'Socialista',
    'ZARAGOZA ALONSO': 'Socialista',
    'ZURITA EXPÓSITO': 'Popular',
}

# XV LEGISLATURA (2023-actualidad)
DICCIONARIO_XV_LEGISLATURA = {
    'ABADES MARTÍNEZ': 'Popular',
    'ÁBALOS MECO': 'Mixto',
    'ABASCAL CONDE': 'VOX',
    'ACEDO REYES': 'Popular',
    'ACEVES GALINDO': 'Socialista',
    'ADRIO TARACIDO': 'Socialista',
    'AGIRRETXEA URRESTI': 'Vasco (EAJ-PNV)',
    'AGÜERA GAGO': 'Popular',
    'AGUIRRE GIL DE BIEDMA': 'VOX',
    'AIZCORBE TORRA': 'VOX',
    'AIZPURUA ARZALLUS': 'Euskal Herria Bildu',
    'ALCARAZ MARTOS': 'VOX',
    'ALFONSO CENDÓN': 'Socialista',
    'ALFONSO SILVESTRE': 'Popular',
    'ALÍA AGUADO': 'Popular',
    'ALMIRÓN RUIZ': 'Socialista',
    'ALMODÓVAR SÁNCHEZ': 'Socialista',
    'ALONSO CANTORNÉ': 'SUMAR',
    'ALÓS LÓPEZ': 'Popular',
    'ÁLVAREZ DE TOLEDO PERALTA-RAMOS': 'Popular',
    'ÁLVAREZ FANJUL': 'Popular',
    'ÁLVAREZ GONZÁLEZ': 'Socialista',
    'ÁLVARO VIDAL': 'Republicano',
    'ANDALA UBBI': 'SUMAR',
    'ANDRÉS AÑÓN': 'Socialista',
    'ANTUÑANO COLINA': 'Socialista',
    'ARAGONÉS MENDIGUCHÍA': 'Popular',
    'ARANDA VARGAS': 'Socialista',
    'ARGOTA CASTRO': 'Socialista',
    'ARGÜELLES GARCÍA': 'Popular',
    'ARMARIO GONZÁLEZ': 'VOX',
    'ARMENGOL SOCIAS': 'Socialista',
    'ARRIBAS MAROTO': 'Socialista',
    'ASARTA CUEVAS': 'VOX',
    'AZORÍN SALAR': 'Socialista',
    'BADIA CASAS': 'SUMAR',
    'BARRIO BAROJA': 'Popular',
    'BAYÓN ROLO': 'Popular',
    'BEAMONTE MESA': 'Popular',
    'BELARRA URTEAGA': 'Mixto',
    'BELDA PÉREZ-PEDRERO': 'Popular',
    'BELMONTE GÓMEZ': 'Popular',
    'BENDODO BENASAYAG': 'Popular',
    'BERMÚDEZ DE CASTRO FERNÁNDEZ': 'Popular',
    'BLANCO ARRÚE': 'Socialista',
    'BLANQUER ALCARAZ': 'Socialista',
    'BOADA DANÉS': 'SUMAR',
    'BOLAÑOS GARCÍA': 'Socialista',
    'BORREGO CORTÉS': 'Popular',
    'BRAVO BAENA': 'Popular',
    'CABEZÓN CASAS': 'Popular',
    'CACHO ISLA': 'Socialista',
    'CALVO GÓMEZ': 'Junts per Catalunya',
    'CAMINO MIÑANA': 'Socialista',
    'CAMPOS ASENSI': 'VOX',
    'CANELO MATITO': 'Socialista',
    'CANTALAPIEDRA ÁLVAREZ': 'Popular',
    'CARAZO HERMOSO': 'Popular',
    'CARBALLEDO BERLANGA': 'Popular',
    'CASTILLA ÁLVAREZ': 'Socialista',
    'CATALÁN HIGUERAS': 'Mixto',
    'CAVACASILLAS RODRÍGUEZ': 'Popular',
    'CELAYA BREY': 'Popular',
    'CERCAS MENA': 'Socialista',
    'CERVERA PINART': 'Junts per Catalunya',
    'CHAMORRO DELMO': 'VOX',
    'CLAVELL LÓPEZ': 'Popular',
    'CLEMENTE MUÑOZ': 'Popular',
    'COBO CARMONA': 'Socialista',
    'COBO PÉREZ': 'Socialista',
    'COFIÑO FERNÁNDEZ': 'SUMAR',
    'CONDE BAJÉN': 'Popular',
    'CONDE LÓPEZ': 'Popular',
    'CONESA COMA': 'Socialista',
    'CORTÉS CARBALLO': 'Popular',
    'CORUJO BERRIEL': 'Socialista',
    'CRESPÍN RUBIO': 'Socialista',
    'CRUSET DOMÈNECH': 'Junts per Catalunya',
    'CRUZ SANTANA': 'Socialista',
    'CRUZ-GUZMÁN GARCÍA': 'Popular',
    'CUESTA RODRÍGUEZ': 'Popular',
    'CUEVAS LARROSA': 'Popular',
    'DE LA ROSA BAENA': 'Socialista',
    'DE LAS CUEVAS CORTÉS': 'Popular',
    'DE LOS SANTOS GONZÁLEZ': 'Popular',
    'DE LUNA TOBARRA': 'Popular',
    'DE MEER MÉNDEZ': 'VOX',
    'DE ROSA TORNER': 'Popular',
    'DEL VALLE RODRÍGUEZ': 'VOX',
    'DELGADO ARCE': 'Popular',
    'DELGADO-TARAMONA HERNÁNDEZ': 'Popular',
    'DÍAZ MARÍN': 'Socialista',
    'DÍAZ PÉREZ': 'SUMAR',
    'DIOUF DIOH': 'Socialista',
    'ESTREMS FAYOS': 'Republicano',
    'FABRA PART': 'Popular',
    'FAGÚNDEZ CAMPO': 'Socialista',
    'FANECA LÓPEZ': 'Socialista',
    'FERNÁNDEZ BENÉITEZ': 'Socialista',
    'FERNÁNDEZ GONZÁLEZ': 'Popular',
    'FERNÁNDEZ HERNÁNDEZ': 'VOX',
    'FERNÁNDEZ RÍOS': 'VOX',
    'FIGAREDO ÁLVAREZ-SALA': 'VOX',
    'FLORES JUBERÍAS': 'VOX',
    'FLORIANO CORRALES': 'Popular',
    'FOLCH BLANC': 'Popular',
    'FRANCO GONZÁLEZ': 'Popular',
    'FULLAONDO LA CRUZ': 'Euskal Herria Bildu',
    'FÚNEZ DE GREGORIO': 'Popular',
    'GALLARDO BARRENA': 'Popular',
    'GAMARRA RUIZ-CLAVIJO': 'Popular',
    'GARCÍA ADANERO': 'Popular',
    'GARCÍA CHAVARRÍA': 'Socialista',
    'GARCÍA FÉLIX': 'Popular',
    'GARCÍA GOMIS': 'VOX',
    'GARCÍA GURRUTXAGA': 'Socialista',
    'GARCÍA LÓPEZ': 'Socialista',
    'GARCÍA MORÍS': 'Socialista',
    'GARRE MURCIA': 'Popular',
    'GARRIDO JIMÉNEZ': 'Socialista',
    'GARRIDO VALENZUELA': 'Popular',
    'GAVIN I VALLS': 'Junts per Catalunya',
    'GIL DE REBOLEÑO LASTORTRIES': 'SUMAR',
    'GIL LÁZARO': 'VOX',
    'GIL SANTIAGO': 'Popular',
    'GÓMEZ PIÑA': 'Socialista',
    'GONZÁLEZ GRACIA': 'Socialista',
    'GONZÁLEZ LÓPEZ': 'SUMAR',
    'GONZÁLEZ VÁZQUEZ': 'Popular',
    'GONZÁLEZ-ROBATTO PEROTE': 'VOX',
    'GRACIA BLANCO': 'Socialista',
    'GRANOLLERS CUNILLERA': 'Republicano',
    'GUARDIOLA SALMERÓN': 'Popular',
    'GUIJARRO GARCÍA': 'SUMAR',
    'GUINART MORENO': 'Socialista',
    'GUTIÉRREZ PRIETO': 'Socialista',
    'GUTIÉRREZ SANTIAGO': 'Socialista',
    'HERNÁNDEZ QUERO': 'VOX',
    'HERNANDO FRAILE': 'Popular',
    'HERRERA GARCÍA': 'Socialista',
    'HERRERO BONO': 'Popular',
    'HISPÁN IGLESIAS DE USSEL': 'Popular',
    'HITA TÉLLEZ': 'Socialista',
    'HOCES ÍÑIGUEZ': 'VOX',
    'HOYO JULIÁ': 'Popular',
    'IBÁÑEZ HERNANDO': 'Popular',
    'IBÁÑEZ MEZQUITA': 'SUMAR',
    'INIESTA EGIDO': 'Socialista',
    'IÑARRITU GARCÍA': 'Euskal Herria Bildu',
    'JEREZ ANTEQUERA': 'Socialista',
    'JIMÉNEZ LINUESA': 'Popular',
    'JÓDAR PÉREZ': 'Socialista',
    'JORDÀ I ROURA': 'Republicano',
    'LAGO PEÑAS': 'SUMAR',
    'LAMUÀ ESTAÑOL': 'Socialista',
    'LEAL FERNÁNDEZ': 'Socialista',
    'LEGARDA URIARTE': 'Vasco (EAJ-PNV)',
    'LIMA GARCÍA': 'Popular',
    'LLAMAZARES DOMINGO': 'Popular',
    'LÓPEZ ÁLVAREZ': 'Socialista',
    'LÓPEZ CANO': 'Socialista',
    'LÓPEZ MARAVER': 'VOX',
    'LÓPEZ TAGLIAFICO': 'SUMAR',
    'LÓPEZ ZAMORA': 'Socialista',
    'LORENTE ANAYA': 'Popular',
    'LORENZO CAZORLA': 'Socialista',
    'LOSADA FERNÁNDEZ': 'Socialista',
    'MACÍAS GATA': 'Popular',
    'MADRENAS I MIR': 'Junts per Catalunya',
    'MADRID OLMO': 'Popular',
    'MALDONADO LÓPEZ': 'Socialista',
    'MARCOS ORTEGA': 'Popular',
    'MARÍ BOSÓ': 'Popular',
    'MARISCAL ANAYA': 'Popular',
    'MARISCAL ZABALA': 'VOX',
    'MARQUÉS ATÉS': 'Socialista',
    'MARTÍN BLANCO': 'Popular',
    'MARTÍN GARCÍA': 'Popular',
    'MARTÍN MARTÍNEZ': 'Socialista',
    'MARTÍN RODRÍGUEZ': 'Socialista',
    'MARTÍN URRIZA': 'SUMAR',
    'MARTÍNEZ BARBERO': 'SUMAR',
    'MARTÍNEZ GÓMEZ': 'Popular',
    'MARTÍNEZ HIERRO': 'SUMAR',
    'MARTÍNEZ LABELLA': 'Popular',
    'MARTÍNEZ RAMÍREZ': 'Socialista',
    'MARTÍNEZ SALMERÓN': 'Socialista',
    'MARTÍNEZ SEIJO': 'Socialista',
    'MATUTE GARCÍA DE JALÓN': 'Euskal Herria Bildu',
    'MAYORAL DE LAMO': 'Socialista',
    'MAYORAL PÉREZ': 'Socialista',
    'MEJÍAS SÁNCHEZ': 'VOX',
    'MELGAREJO MORENO': 'Popular',
    'MELLADO SIERRA': 'Socialista',
    'MÉNDEZ MONASTERIO': 'VOX',
    'MERCADAL BAQUERO': 'Socialista',
    'MERINO MARTÍNEZ': 'Popular',
    'MESQUIDA MAYANS': 'Popular',
    'MICÓ MICÓ': 'Mixto',
    'MÍNGUEZ GARCÍA': 'Socialista',
    'MOLINA LEÓN': 'Popular',
    'MONEO DÍEZ': 'Popular',
    'MONTÁVEZ AGUILLAUME': 'Socialista',
    'MONTERO CUADRADO': 'Socialista',
    'MONTESINOS DE MIGUEL': 'Popular',
    'MORALEJA GÓMEZ': 'Popular',
    'MORALES ÁLVAREZ': 'Socialista',
    'MORENO BORRÁS': 'Popular',
    'MORENO FERNÁNDEZ': 'Socialista',
    'MORO ALMARAZ': 'Popular',
    'MUÑOZ ABRINES': 'Popular',
    'MUÑOZ DE LA IGLESIA': 'Popular',
    'NACARINO-BRABO JIMÉNEZ': 'Popular',
    'NARBONA RUIZ': 'Socialista',
    'NASARRE OLIVA': 'Socialista',
    'NAVARRO LACOBA': 'Popular',
    'NAVARRO LÓPEZ': 'Popular',
    'NOGUERAS I CAMERO': 'Junts per Catalunya',
    'NORIEGA GÓMEZ': 'Popular',
    'NÚÑEZ FEIJÓO': 'Popular',
    'NÚÑEZ GUIJARRO': 'Popular',
    'OGOU I CORBI': 'SUMAR',
    'OLANO VELA': 'Popular',
    'ORTEGA SMITH-MOLINA': 'VOX',
    'OTERO GABIRONDO': 'Euskal Herria Bildu',
    'OTERO GARCÍA': 'Socialista',
    'OTERO RODRÍGUEZ': 'Socialista',
    'PAGÈS I MASSÓ': 'Junts per Catalunya',
    'PALENCIA RUBIO': 'Popular',
    'PANIAQUA NÚÑEZ': 'Popular',
    'PARÉ AREGALL': 'Socialista',
    'PARRA APARICIO': 'Popular',
    'PARRA GALLEGO': 'Popular',
    'PASCUAL ROCAMORA': 'Popular',
    'PEDREÑO MOLINA': 'Popular',
    'PEÑA CAMARERO': 'Socialista',
    'PEREA I CONILLAS': 'Socialista',
    'PÉREZ CORONADO': 'Popular',
    'PÉREZ LÓPEZ': 'Popular',
    'PÉREZ ORTIZ': 'Socialista',
    'PÉREZ OSMA': 'Popular',
    'PÉREZ RECUERDA': 'Popular',
    'PISARELLO PRADOS': 'SUMAR',
    'PLAZA GARCÍA': 'Socialista',
    'POBLADOR PACHECO': 'Socialista',
    'POSE MESURA': 'Socialista',
    'POZUETA FERNÁNDEZ': 'Euskal Herria Bildu',
    'PRIETO SERRANO': 'Popular',
    'PUENTE SANTIAGO': 'Socialista',
    'PUEYO SANZ': 'SUMAR',
    'PUY FRAGA': 'Popular',
    'QUINTANA CARBALLO': 'Popular',
    'QUINTANILLA NAVARRO': 'Popular',
    'QUINTERO HERNÁNDEZ': 'Socialista',
    'RALLO LOMBARTE': 'Socialista',
    'RAMAJO PRADA': 'Popular',
    'RAMÍREZ CARNER': 'Socialista',
    'RAMÍREZ DEL RÍO': 'VOX',
    'RAMÍREZ MARTÍN': 'Popular',
    'RAMÍREZ MORENO': 'Socialista',
    'RAMOS ESTEBAN': 'Socialista',
    'RECAS MARTÍN': 'SUMAR',
    'REDONDO CÁRDENAS': 'Socialista',
    'REGO CANDAMIL': 'Mixto',
    'RENTERIA LASANTA': 'Vasco (EAJ-PNV)',
    'REQUENA RUIZ': 'Popular',
    'REY DE LAS HERAS': 'Socialista',
    'REYNAL REILLO': 'Popular',
    'RIVERA ARIAS': 'SUMAR',
    'RIVES ARCAYNA': 'Socialista',
    'ROBLES LÓPEZ': 'VOX',
    'RODRÍGUEZ ALMEIDA': 'VOX',
    'RODRÍGUEZ CALLEJA': 'Popular',
    'RODRÍGUEZ DE MILLÁN PARRO': 'VOX',
    'RODRÍGUEZ GÓMEZ DE CELIS': 'Socialista',
    'RODRÍGUEZ HERRER': 'Popular',
    'RODRÍGUEZ PALACIOS': 'Socialista',
    'RODRÍGUEZ SALAS': 'Socialista',
    'RODRÍGUEZ SERRA': 'Popular',
    'RODRÍGUEZ SUÁREZ': 'Socialista',
    'ROJAS GARCÍA': 'Popular',
    'ROJAS MANRIQUE': 'Popular',
    'ROJO BLAS': 'Socialista',
    'ROMÁN JASANADA': 'Popular',
    'ROMANÍ CANTERA': 'Popular',
    'ROMERO POZO': 'Socialista',
    'ROMERO VILCHES': 'VOX',
    'ROS MARTÍNEZ': 'Socialista',
    'RUEDA PERELLÓ': 'VOX',
    'RUFIÁN ROMERO': 'Republicano',
    'RUIZ BOIX': 'Socialista',
    'RUIZ DE DIEGO': 'Socialista',
    'RUIZ SOLÁS': 'VOX',
    'SÁEZ ALONSO-MUÑUMER': 'VOX',
    'SÁEZ CRUZ': 'Socialista',
    'SAGASTIZABAL UNZETABARRENETXEA': 'Vasco (EAJ-PNV)',
    'SAHUQUILLO GARCÍA': 'Socialista',
    'SAINZ MARTÍN': 'Socialista',
    'SALVADOR I DUCH': 'Republicano',
    'SÁNCHEZ DÍAZ': 'Socialista',
    'SÁNCHEZ GARCÍA': 'VOX',
    'SÁNCHEZ OJEDA': 'Popular',
    'SÁNCHEZ PÉREZ': 'Popular',
    'SÁNCHEZ PÉREZ-CASTEJÓN': 'Socialista',
    'SÁNCHEZ SERNA': 'Mixto',
    'SÁNCHEZ SIERRA': 'Popular',
    'SÁNCHEZ TORREGRASA': 'Popular',
    'SANCHO ÍÑIGUEZ': 'Socialista',
    'SANTANA AGUILERA': 'Socialista',
    'SANTANA PERERA': 'Mixto',
    'SANTIAGO ROMERO': 'SUMAR',
    'SANTOS MARAVER': 'SUMAR',
    'SANZ MARTÍNEZ': 'Socialista',
    'SARRIÀ MORELL': 'Socialista',
    'SASTRE UYÁ': 'Popular',
    'SAYAS LÓPEZ': 'Popular',
    'SÉMPER PASCUAL': 'Popular',
    'SENDEROS ORAÁ': 'Socialista',
    'SERRADA PARIENTE': 'Socialista',
    'SERRANO MARTÍNEZ': 'Socialista',
    'SIERRA CABALLERO': 'SUMAR',
    'SIMANCAS SIMANCAS': 'Socialista',
    'SIMARRO VICENS': 'Popular',
    'SOLDEVILLA NOVIALS': 'Socialista',
    'SOLER MUR': 'Socialista',
    'TABOADELA ÁLVAREZ': 'Socialista',
    'TARNO BLANCO': 'Popular',
    'TELLADO FILGUEIRA': 'Popular',
    'TENIENTE SÁNCHEZ': 'Popular',
    'TOMÁS OLIVARES': 'Popular',
    'TORRES TEJADA': 'Popular',
    'TRENZANO RUBIO': 'Socialista',
    'URIARTE BENGOCHEA': 'Popular',
    'VALERO MORALES': 'SUMAR',
    'VALIDO GARCÍA': 'Mixto',
    'VALLUGERA BALAÑÀ': 'Republicano',
    'VAQUERO MONTERO': 'Vasco (EAJ-PNV)',
    'VARELA PAZOS': 'Popular',
    'VÁZQUEZ BLANCO': 'Popular',
    'VÁZQUEZ JIMÉNEZ': 'Popular',
    'VEDRINA CONESA': 'Popular',
    'VELARDE GÓMEZ': 'Mixto',
    'VELASCO MORILLO': 'Popular',
    'VELASCO RETAMOSA': 'Popular',
    'VERANO DOMÍNGUEZ': 'Popular',
    'VERDEJO VICENTE': 'Socialista',
    'VIDAL MATAS': 'SUMAR',
    'VIDAL SÁEZ': 'SUMAR',
    'ZARAGOZA ALONSO': 'Socialista',
}

# Diccionario por defecto
DICCIONARIO_POR_DEFECTO = DICCIONARIO_XV_LEGISLATURA

# Rangos de fechas por legislatura (sin huecos)
DICCIONARIOS_POR_FECHA = [
    ((1977, 7, 13), (1979, 3, 22),  DICCIONARIO_CONSTITUYENTE),
    ((1979, 3, 23), (1982, 11, 17), DICCIONARIO_I_LEGISLATURA),
    ((1982, 11, 18),(1986, 7, 14),  DICCIONARIO_II_LEGISLATURA),
    ((1986, 7, 15), (1989, 11, 20), DICCIONARIO_III_LEGISLATURA),
    ((1989, 11, 21),(1993, 6, 27),  DICCIONARIO_IV_LEGISLATURA),
    ((1993, 6, 28), (1996, 3, 26),  DICCIONARIO_V_LEGISLATURA),
    ((1996, 3, 27), (2000, 4, 4),   DICCIONARIO_VI_LEGISLATURA),
    ((2000, 4, 5),  (2004, 4, 1),   DICCIONARIO_VII_LEGISLATURA),
    ((2004, 4, 2),  (2008, 3, 31),  DICCIONARIO_VIII_LEGISLATURA),
    ((2008, 4, 1),  (2011, 12, 12), DICCIONARIO_IX_LEGISLATURA),
    ((2011, 12, 13),(2016, 1, 12),  DICCIONARIO_X_LEGISLATURA),
    ((2016, 1, 13), (2016, 7, 18),  DICCIONARIO_XI_LEGISLATURA),
    ((2016, 7, 19), (2019, 5, 20),  DICCIONARIO_XII_LEGISLATURA),
    ((2019, 5, 21), (2019, 12, 2),  DICCIONARIO_XIII_LEGISLATURA),
    ((2019, 12, 3), (2023, 8, 16),  DICCIONARIO_XIV_LEGISLATURA),
    ((2023, 8, 17), None,           DICCIONARIO_XV_LEGISLATURA),
]

print("✅ PARTE 4 CARGADA: Diccionarios de oradores (1977-2026)")


# ==============================================================
# PARTE 5 — FUNCIONES DE IDENTIFICACIÓN DE ORADORES
# ==============================================================

def normalizar_nombre_orador_completo(nombre):
    """Normalización completa de nombres."""
    nombre = re.sub(r'\s+', ' ', nombre.strip())
    nombre = re.sub(r':$', '', nombre)
    normalizaciones = {
        'MONTSERRAT MONTSERRAT': 'MONTSERRAT MONTSERRAT',
        'MIQUEL I VALENTÍ': 'MIQUEL I VALENTÍ',
        'FERNÁNDEZ CASTAÑÓN': 'FERNÁNDEZ CASTAÑÓN',
        'RODRÍGUEZ MARTÍNEZ': 'RODRÍGUEZ MARTÍNEZ',
        'CANCELA RODRÍ GUEZ': 'CANCELA RODRÍGUEZ',
        'ALLI MRTÍNEZ': 'ALLI MARTÍNEZ',
        'ROSSELL AGUILAR': 'ROSSELL AGUILAR',
        'ORDÓÑEZ CARBAJAL': 'ORDÓÑEZ CARBAJAL',
        'VELARDE GÓMEZ': 'VELARDE GÓMEZ',
        'BELARRA URTEAGA': 'BELARRA URTEAGA',
        'VERSTRYNGE REVUELTA': 'VERSTRYNGE REVUELTA',
        'LOIS GONZÁLEZ': 'LOIS GONZÁLEZ',
        'CALVO POYATO': 'CALVO POYATO',
        'MONTERO GIL': 'MONTERO GIL',
        'DÍAZ PÉREZ': 'DÍAZ PÉREZ',
        'SÁNCHEZ-CAMACHO PÉREZ': 'SÁNCHEZ-CAMACHO PÉREZ',
        'GUITARTE GIMENO': 'GUITARTE GIMENO',
        'SORLÍ FRESQUET': 'SORLÍ FRESQUET',
        'AGUIRRE GIL DE BIEDMA': 'AGUIRRE GIL DE BIEDMA',
        'ACEDO REYES': 'ACEDO REYES',
        'RIVERA ARIAS': 'RIVERA ARIAS',
        'MARTÍN RODRÍGUEZ': 'MARTÍN RODRÍGUEZ',
        'SAINZ MARTÍN': 'SAINZ MARTÍN',
        'NACARINO-BRABO JIMÉNEZ': 'NACARINO-BRABO JIMÉNEZ',
        'GUINART MORENO': 'GUINART MORENO',
        'ÁBALOS MECO': 'ÁBALOS MECO',
        'ÁLVAREZ PALLEIRO': 'ÁLVAREZ PALLEIRO',
    }
    return normalizaciones.get(nombre.upper(), nombre.upper())


def extraer_grupo_del_contexto_corregido(contexto):
    """Extrae grupo parlamentario del contexto de presidenta."""
    contexto_lower = contexto.lower()
    patrones_grupos = {
        'Socialista': [r'presentada por el grupo parlamentario socialista', r'grupo parlamentario socialista'],
        'Popular': [r'presentada por el grupo parlamentario popular', r'grupo parlamentario popular'],
        'Podemos': [r'presentada por el grupo parlamentario confederal de unidos podemos-en comú podem-en marea',
                    r'presentada por el grupo parlamentario confederal', r'unidos podemos', 'unidas podemos',
                    'en comú podem', 'en marea', 'grupo podemos', r'grupo confederal'],
        'VOX': [r'presentada por el grupo parlamentario vox', r'grupo parlamentario vox'],
        'Ciudadanos': [r'presentada por el grupo parlamentario ciudadanos', r'grupo parlamentario ciudadanos'],
        'EAJ-PNV': [r'presentada por el grupo parlamentario vasco', 'eaj-pnv', r'grupo parlamentario vasco'],
        'ERC': [r'presentada por el grupo parlamentario republicano', 'esquerra republicana', r'grupo parlamentario republicano'],
        'EH Bildu': [r'presentada por el grupo parlamentario eh bildu', r'grupo parlamentario eh bildu'],
        'Plurinacional SUMAR': [r'grupo plurinacional sumar', r'grupo sumar'],
        'Plural': [r'presentada por el grupo parlamentario plural', r'grupo parlamentario plural'],
        'Junts per Catalunya': [r'presentada por junts per catalunya', 'junts per catalunya'],
        'Mixto': [r'presentada por el grupo mixto', 'grupo mixto'],
    }
    for grupo, patrones in patrones_grupos.items():
        if any(re.search(patron, contexto_lower) for patron in patrones):
            return grupo
    patrones_genericos = {
        'Socialista': ['grupo socialista'],
        'Popular': ['grupo popular'],
        'Podemos': ['unidos podemos', 'unidas podemos', 'en comú podem', 'en marea', 'podemos', 'grupo confederal'],
        'VOX': ['grupo vox'],
        'Ciudadanos': ['grupo ciudadanos'],
        'EAJ-PNV': ['grupo vasco'],
        'ERC': ['grupo republicano'],
        'EH Bildu': ['grupo eh bildu'],
        'Plurinacional SUMAR': ['sumar'],
        'Plural': ['grupo plural'],
        'Mixto': ['grupo mixto'],
    }
    for grupo, patrones in patrones_genericos.items():
        if any(re.search(r'\b' + patron + r'\b', contexto_lower) for patron in patrones):
            return grupo
    return "Desconocido"

print("✅ PARTE 5 CARGADA: Funciones de identificación")


# ==============================================================
# PARTE 6 — FUNCIONES DE EXTRACCIÓN DE INTERVENCIONES
# ==============================================================

def detectar_fin_debate_inteligente(linea_actual, intervenciones_extraidas, lineas_siguientes=None, es_inicio_debate=False):
    """Detección inteligente del fin de debate."""
    linea_limpia = linea_actual.strip()
    contexto_completo = linea_limpia
    if lineas_siguientes:
        contexto_completo += " " + " ".join([ls.strip() for ls in lineas_siguientes[:2]])
    contexto_upper = contexto_completo.upper()

    if len(intervenciones_extraidas) < 5:
        return False
    if es_inicio_debate and len(intervenciones_extraidas) == 0:
        return False

    patrones_fin_absolutos = [
        r'\(NÚMERO DE EXPEDIENTE\s+\d+/\d+\)',
        r'—.*\(NÚMERO DE EXPEDIENTE\s+\d+/\d+\)',
        r'^[A-Z][A-ZÁÉÍÓÚÑ\s]+\(CONTINUACIÓN\):?',
        r'^—\s*[A-Z][A-ZÁÉÍÓÚÑ\s]+[\.\)]\s*\(CONTINUACIÓN\)',
        r'^ENMIENDAS DEL SENADO',
        r'^— PROYECTO DE LEY',
        r'^— PROPOSICIÓN DE LEY',
    ]
    for patron in patrones_fin_absolutos:
        if re.search(patron, contexto_upper):
            print(f"        🔚 FIN DEBATE (Absoluto): {linea_limpia[:80]}...")
            return True

    patrones_nuevos_fin = [
        r'^MOCIONES CONSECUENCIA DE',
        r'^DEBATES DE TOTALIDAD DE',
        r'^PROPOSICIONES NO DE LEY',
        r'^ENMIENDAS DEL SENADO',
        r'^SOLICITUDES DE',
        r'^— DEL GRUPO PARLAMENTARIO.*\(NÚMERO DE EXPEDIENTE',
    ]
    for patron in patrones_nuevos_fin:
        if re.search(patron, contexto_upper):
            print(f"        🔚 CAMBIO DE SECCIÓN: {linea_limpia[:80]}...")
            return True

    if linea_limpia.startswith('—') and len(intervenciones_extraidas) > 2:
        contenido = linea_limpia[1:].strip()
        if (contenido[:20].isupper() or
            any(palabra in contenido.upper() for palabra in ['PROYECTO', 'PROPOSICIÓN', 'ENMIENDA', 'MOCIÓN'])) and \
           '(Número de expediente' in contenido:
            print(f"        🔚 NUEVO TÍTULO CON EXPEDIENTE: {linea_limpia[:80]}...")
            return True

    if len(intervenciones_extraidas) > 3:
        patrones_contextuales = [
            r'^[A-Z][A-ZÁÉÍÓÚÑ\s]+\:?$',
            r'^—\s*[A-Z][A-ZÁÉÍÓÚÑ\s]+[\.\)]',
        ]
        for patron in patrones_contextuales:
            if re.match(patron, linea_limpia):
                if not es_inicio_intervencion(linea_limpia, contexto_completo):
                    print(f"        🔚 NUEVO TÍTULO: {linea_limpia[:80]}...")
                    return True

    return False


def es_inicio_intervencion(linea, contexto_previo=""):
    """Determina si una línea es el inicio de una intervención real."""
    patrones_intervencion = [
        r'^(?:El señor|La señora)\s+[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]*\s*\([^)]+\)\s*:',
        r'^(?:El señor|La señora|SECRETARI[OA]|MINISTR[OA]|DIRECTOR[RA]|PRESIDENTE?|VICEPRESIDENTE?)\s+[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+\s*:',
        r'^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s*:',
    ]
    for patron in patrones_intervencion:
        if re.search(patron, linea):
            return True

    patrones_falsos_positivos = [
        r'\(El señor[^)]+ocupa la Presidencia\)',
        r'\(La señora[^)]+abandona el hemiciclo\)',
        r'\([^)]*toma asiento[^)]*\)',
        r'\.\s*\([^)]+\)',
    ]
    for patron in patrones_falsos_positivos:
        if re.search(patron, contexto_previo + " " + linea):
            return False

    return False


def detectar_fin_intervencion_mejorado(linea_actual, texto_intervencion, lineas_siguientes=None):
    """Detección precisa del fin de una intervención."""
    linea = linea_actual.strip()

    if len(texto_intervencion) < 3:
        return False

    patrones_fin_absolutos = [
        r'^\(.*aplausos.*\)$',
        r'^\(.*ovación.*\)$',
        r'^\(.*risas.*\)$',
        r'^\(.*rumores.*\)$',
        r'^\.{10,}$',
        r'^Tiene la palabra el señor',
        r'^Tiene la palabra la señora',
        r'^Cedo la palabra al señor',
        r'^Cedo la palabra a la señora',
    ]
    for patron in patrones_fin_absolutos:
        if re.search(patron, linea, re.IGNORECASE):
            print(f"        🔚 FIN ABSOLUTO: '{linea}'")
            return True

    if re.match(r'^\d+-\w+-\d+-\w+$', linea) or re.match(r'^:\w+$', linea):
        return False

    if (len(texto_intervencion) > 10 and len(linea) < 30 and
        linea.isupper() and not any(palabra in linea.lower() for palabra in ['president', 'secretari', 'ministr'])):
        print(f"        🔚 LÍNEA CORTA MAYÚSCULAS: '{linea}'")
        return True

    patrones_falsos_orador = [
        r'^—\s*como decía', r'^—\s*a ver si', r'^—\s*señorías', r'^—\s*compañeras',
        r'^—\s*quería', r'^—\s*en definitiva', r'^—\s*sin embargo', r'^—\s*por ejemplo',
        r'^—\s*además', r'^—\s*pero', r'^—\s*y',
    ]
    for patron in patrones_falsos_orador:
        if re.match(patron, linea.lower()):
            return False

    if lineas_siguientes:
        siguientes_str = " ".join(lineas_siguientes[:2])
        patrones_continuacion = ['como decía', 'a ver si', 'es decir', 'por ejemplo',
                                  'además', 'sin embargo', 'no obstante', 'por otro lado']
        if any(patron in siguientes_str.lower() for patron in patrones_continuacion):
            return False

    patrones_fin_falsos = [r'finalizo ya', r'termino ya', r'para concluir', r'en conclusión', r'para terminar']
    for patron in patrones_fin_falsos:
        if re.search(patron, linea.lower()):
            if lineas_siguientes and len(lineas_siguientes) > 0:
                if len(lineas_siguientes[0].strip()) > 50:
                    return False
            return True

    return False


def extraer_texto_intervencion_inteligente_mejorado(lineas, indice_inicio):
    """Extrae texto de intervención con detección inteligente de fin."""
    texto = []
    i = indice_inicio

    if i < len(lineas):
        texto.append(lineas[i])
        i += 1

    while i < len(lineas):
        linea_actual = lineas[i].strip()

        if re.match(r'^\d+-\w+-\d+-\w+$', linea_actual) or re.match(r'^:\w+$', linea_actual):
            i += 1
            continue

        lineas_siguientes = lineas[i+1:i+3] if i+1 < len(lineas) else []
        contexto = "\n".join(texto[-3:]) if len(texto) >= 3 else ""

        if es_inicio_intervencion(linea_actual, contexto):
            if not es_falso_positivo_intervencion(linea_actual, texto):
                break

        if detectar_fin_intervencion_mejorado(linea_actual, "\n".join(texto), lineas_siguientes):
            break

        texto.append(lineas[i])
        i += 1

    resultado = '\n'.join(texto)
    print(f"        📝 Intervención: {len(resultado)} chars, {i - indice_inicio} líneas")
    return resultado, i


def es_falso_positivo_intervencion(linea, texto_previo):
    """Detecta falsos positivos de inicio de intervención."""
    patrones_falsos = [
        r'^—\s*[Ee]l señor',
        r'^—\s*[Ll]a señora',
        r'^\([^)]*dice[^)]*\)',
    ]
    for patron in patrones_falsos:
        if re.search(patron, linea):
            return True
    return False


def normalizar_cargo_complejo(orador_completo):
    """Normaliza nombres de cargos complejos."""
    orador_upper = orador_completo.upper()
    cargos = {
        'SECRETARIA DE ESTADO': 'SECRETARIA DE ESTADO',
        'SECRETARIO DE ESTADO': 'SECRETARIO DE ESTADO',
        'MINISTRO': 'MINISTRO',
        'MINISTRA': 'MINISTRA',
        'DIRECTOR': 'DIRECTOR',
        'DIRECTORA': 'DIRECTORA',
        'VICEPRESIDENTA': 'VICEPRESIDENTA',
        'VICEPRESIDENTE': 'VICEPRESIDENTE',
        'PRESIDENTA': 'PRESIDENTA',
        'PRESIDENTE': 'PRESIDENTE',
    }
    for cargo_key, cargo_display in cargos.items():
        if cargo_key in orador_upper:
            match_nombre = re.search(r'\((.*?)\)', orador_upper)
            if match_nombre:
                return f"{cargo_display} ({match_nombre.group(1).strip()})"
            return cargo_display
    return orador_completo


def detectar_anuncio_mesa_mejorado(linea, texto_completo, estado_actual):
    """Detecta anuncios de la Mesa analizando texto completo."""
    cargos = ['PRESIDENTA', 'PRESIDENTE', 'VICEPRESIDENTE', 'VICEPRESIDENTA', 'SECRETARIO', 'SECRETARIA']
    if not any(cargo in linea.upper() for cargo in cargos):
        return False

    texto_lower = texto_completo.lower()
    if any(p in texto_lower for p in ['turno de fijación', 'fijación de posiciones', 'fijar posición']):
        estado_actual['tipo_turno'] = 'Fijación de posición'
        return True
    elif any(p in texto_lower for p in ['enmiendas', 'debate de las enmiendas', 'enmienda aprobada']):
        estado_actual['tipo_turno'] = 'Enmiendas'
        return True
    elif any(p in texto_lower for p in ['preguntas', 'turno de preguntas', 'interrogaciones']):
        estado_actual['tipo_turno'] = 'Preguntas'
        return True
    elif any(p in texto_lower for p in ['comparecencia', 'comparecientes']):
        estado_actual['tipo_turno'] = 'Comparecencias'
        return True
    elif any(p in texto_lower for p in ['interpelación']):
        estado_actual['tipo_turno'] = 'Interpelaciones'
        return True
    elif any(p in texto_lower for p in ['proposición de ley', 'proposición no de ley']):
        estado_actual['tipo_turno'] = 'Defensa Proposición de Ley'
        return True
    elif any(p in texto_lower for p in ['moción']):
        estado_actual['tipo_turno'] = 'Defensa Moción'
        return True
    return False

print("✅ PARTE 6 CARGADA: Funciones de extracción de intervenciones")


# ==============================================================
# PARTE 7 — FUNCIÓN PRINCIPAL DE EXTRACCIÓN
# ==============================================================

def extraer_intervenciones_mejorada_completa(texto_debate, titulo_corto, archivo, metadata, analizador, numero_expediente_titulo=""):
    """Extrae intervenciones de un debate."""
    intervenciones = []
    lineas = texto_debate.split('\n')
    i = 0

    # Usar número del título si está disponible; si no, buscarlo en el texto
    if numero_expediente_titulo:
        id_discurso = numero_expediente_titulo
        print(f"      ✅ ID del título: {id_discurso}")
    else:
        # Fallback: buscar en el texto completo usando buscar_numero_expediente
        id_discurso = buscar_numero_expediente(texto_debate, titulo_corto)
        print(f"      🔍 ID del texto: {id_discurso}")

    lineas_limpias = limpiar_cabeceras_documento_mejorado(lineas, archivo)
    estado_actual = {'tipo_turno': 'Intervención en debate'}
    orden_intervencion = 1
    lineas_procesadas = 0
    max_lineas_por_debate = 500
    debate_recien_iniciado = True

    while i < len(lineas_limpias) and lineas_procesadas < max_lineas_por_debate:
        linea = lineas_limpias[i].strip()
        lineas_procesadas += 1

        lineas_siguientes = lineas_limpias[i+1:i+3] if i+1 < len(lineas_limpias) else []
        if detectar_fin_debate_inteligente(linea, intervenciones, lineas_siguientes, debate_recien_iniciado):
            break

        if len(intervenciones) > 0:
            debate_recien_iniciado = False

        patron_orador_flexible = r'^(?:El señor|La señora|SECRETARI[OA]|MINISTR[OA]|DIRECTOR[RA]|PRESIDENTE?|VICEPRESIDENTE?)\s+(.+?)\s*:\s*(.*)'
        match_orador = re.match(patron_orador_flexible, linea, re.IGNORECASE)

        if match_orador:
            orador_completo = match_orador.group(1).strip()
            orador_normalizado = normalizar_cargo_complejo(orador_completo)
            cargos_mesa_o_gobierno = ['PRESIDENTA', 'PRESIDENTE', 'VICEPRESIDENTE', 'VICEPRESIDENTA',
                                       'SECRETARIO DE ESTADO', 'SECRETARIA DE ESTADO', 'MINISTRO', 'MINISTRA',
                                       'DIRECTOR', 'DIRECTORA', 'SECRETARIO', 'SECRETARIA']
            es_cargo = any(cargo in orador_normalizado.upper() for cargo in cargos_mesa_o_gobierno)

            if es_cargo:
                print(f"        🏛️  CARGO: {orador_normalizado}")
                texto_intervencion, nuevo_indice = extraer_texto_intervencion_inteligente_mejorado(lineas_limpias, i)
                texto_limpio = limpiar_texto_intervencion_final(texto_intervencion)
                if detectar_anuncio_mesa_mejorado(linea, texto_limpio, estado_actual):
                    print(f"        🔄 MESA → estado: {estado_actual['tipo_turno']}")

                grupo = 'MESA/GOBIERNO'
                match_nombre_cargo = re.search(r'\((.*?)\)', orador_normalizado)
                if match_nombre_cargo:
                    nombre_persona_cargo = match_nombre_cargo.group(1).strip()
                    grupo_persona = obtener_grupo_preciso(nombre_persona_cargo, "", metadata.get('fecha', '1/1/2020'))
                    if grupo_persona != "Desconocido":
                        grupo = grupo_persona

                if len(texto_limpio) > 20:
                    intervenciones.append({
                        'Fecha': metadata.get('fecha', 'Fecha Desconocida'),
                        'Fase temporal': '1',
                        'Tipo': metadata.get('comision_especifica', 'Tipo Desconocido'),
                        'ID Documento': archivo,
                        'ID Discurso': id_discurso,
                        'Título': titulo_corto,
                        'Orador': orador_normalizado,
                        'Grupo Parlamentario': grupo,
                        'Tipo Intervención': 'Gestión de debate',
                        'Intervención': texto_limpio,
                        'Corpus': '',
                        'Orden Intervención': orden_intervencion,
                    })
                    orden_intervencion += 1
                i = nuevo_indice

            else:
                print(f"        🎙️ ORADOR: {orador_normalizado}")
                orador = orador_normalizado
                grupo = obtener_grupo_preciso(orador, "", metadata.get('fecha', '1/1/2020'))
                texto_intervencion, nuevo_indice = extraer_texto_intervencion_inteligente_mejorado(lineas_limpias, i)
                texto_limpio = limpiar_texto_intervencion_final(texto_intervencion)
                tipo = estado_actual['tipo_turno']

                if len(texto_limpio) > 50:
                    intervenciones.append({
                        'Fecha': metadata.get('fecha', 'Fecha Desconocida'),
                        'Fase temporal': '1',
                        'Tipo': metadata.get('comision_especifica', 'Tipo Desconocido'),
                        'ID Documento': archivo,
                        'ID Discurso': id_discurso,
                        'Título': titulo_corto,
                        'Orador': orador,
                        'Grupo Parlamentario': grupo,
                        'Tipo Intervención': tipo,
                        'Intervención': texto_limpio,
                        'Corpus': '',
                        'Orden Intervención': orden_intervencion,
                    })
                    print(f"        💾 {orador} - {grupo} - {tipo}")
                    orden_intervencion += 1
                i = nuevo_indice
        else:
            i += 1

    if lineas_procesadas >= max_lineas_por_debate:
        print(f"        ⚠️  Límite de {max_lineas_por_debate} líneas alcanzado")

    print(f"        📊 Total intervenciones: {len(intervenciones)}")
    return intervenciones


def eliminar_duplicados_inteligente(intervenciones):
    """Elimina solo duplicados exactos."""
    if not intervenciones:
        return []
    intervenciones_unicas = []
    claves_vistas = set()
    for intervencion in intervenciones:
        texto_hash = hash(intervencion['Intervención'])
        clave = f"{intervencion['Orador']}_{texto_hash}_{intervencion['Orden Intervención']}"
        if clave not in claves_vistas:
            claves_vistas.add(clave)
            intervenciones_unicas.append(intervencion)
        else:
            print(f"        🧹 DUPLICADO EXACTO: {intervencion['Orador']}")
    return intervenciones_unicas

print("✅ PARTE 7 CARGADA: Función principal de extracción")


# ==============================================================
# PARTE 8 — CLASE ANALIZADORCOMPLETO
# ==============================================================

class AnalizadorCompleto:
    def __init__(self, palabras_clave):
        self.palabras_clave = palabras_clave

    def extraer_numero_expediente_del_titulo(self, titulo):
        """Extrae el número de expediente del título."""
        if not titulo:
            return ""
        print(f"      🔍 Buscando expediente en título...")

        patron = r'\(Número de expediente\s+(\d{3}/\d{6})\)\s*\.+'
        matches = list(re.finditer(patron, titulo))
        if matches:
            numero = matches[-1].group(1)
            print(f"      ✅ ID EXTRAÍDO: {numero}")
            return numero

        patron_sin_puntos = r'\(Número de expediente\s+(\d{3}/\d{6})\)'
        matches = list(re.finditer(patron_sin_puntos, titulo))
        if matches:
            numero = matches[-1].group(1)
            print(f"      ✅ ID EXTRAÍDO (sin puntos): {numero}")
            return numero

        print(f"      ❌ ID no encontrado en título")
        return ""

    def extraer_metadata_completa(self, texto):
        """Extrae metadatos completos del documento."""
        metadata = {'fecha': None, 'tipo_sesion': None, 'comision_especifica': None}

        patron_fecha = r'celebrada el (?:lunes|martes|miércoles|jueves|viernes|sábado|domingo)\s+(\d{1,2}\s+de\s+[a-z]+\s+de\s+\d{4})'
        match_fecha = re.search(patron_fecha, texto, re.IGNORECASE)
        if match_fecha:
            metadata['fecha'] = match_fecha.group(1)

        if 'PLENO Y DIPUTACIÓN PERMANENTE' in texto:
            metadata['tipo_sesion'] = 'PLENO'
            metadata['comision_especifica'] = 'PLENO'
        elif 'COMISIONES' in texto:
            metadata['tipo_sesion'] = 'COMISIONES'
            metadata['comision_especifica'] = 'COMISIÓN'
        else:
            metadata['tipo_sesion'] = 'SESIÓN PARLAMENTARIA'
            metadata['comision_especifica'] = 'NO IDENTIFICADA'

        return metadata

    def extraer_indice_completo(self, texto):
        """Extrae todos los títulos del índice."""
        titulos = []
        if not texto or 'ORDEN DEL DÍA:' not in texto:
            return titulos

        try:
            inicio = texto.find('ORDEN DEL DÍA:')
            fin = texto.find('SUMARIO', inicio)
            if fin == -1:
                fin = len(texto)

            seccion_orden = texto[inicio:fin]
            lineas = seccion_orden.split('\n')
            i = 0

            while i < len(lineas):
                linea = lineas[i].strip()
                if linea.startswith('—'):
                    titulo_completo = linea[1:].strip()
                    j = i + 1

                    while j < len(lineas):
                        linea_siguiente = lineas[j].strip()
                        if '....' in linea_siguiente or '...' in linea_siguiente:
                            partes = re.split(r'\.{3,}', linea_siguiente)
                            if len(partes) > 1:
                                ultimo = partes[-1].strip()
                                if ultimo.isdigit():
                                    titulo_completo += ' ' + linea_siguiente
                                    pagina = int(ultimo)
                                    titulo_limpio = re.sub(r'\s+', ' ', titulo_completo).strip()
                                    if len(titulo_limpio) > 10:
                                        titulos.append({'titulo': titulo_limpio, 'pagina': pagina})
                                    break
                        elif linea_siguiente.startswith('—'):
                            break
                        if (linea_siguiente and len(linea_siguiente) > 2 and
                            not linea_siguiente.replace('.', '').strip().isdigit() and
                            not re.match(r'^\.+$', linea_siguiente)):
                            titulo_completo += ' ' + linea_siguiente
                        j += 1
                    else:
                        j = len(lineas)
                    i = j
                else:
                    i += 1

        except Exception as e:
            print(f"      ❌ Error extrayendo índice: {e}")

        return titulos

    def es_titulo_relevante(self, titulo):
        """Verifica si el título contiene palabras clave."""
        if not titulo:
            return False
        titulo_lower = titulo.lower()
        for palabra in self.palabras_clave:
            if palabra.lower() in titulo_lower:
                print(f"      ✅ RELEVANTE: '{palabra}' en: {titulo[:80]}...")
                return True
        print(f"      ❌ NO RELEVANTE: {titulo}")
        return False

    def procesar_pdf(self, pdf_path, nombre_archivo):
        """Procesa un PDF completo."""
        print(f"\n📄 Procesando: {nombre_archivo}")
        try:
            with pdfplumber.open(pdf_path) as pdf:
                texto_indice = ""
                for i in range(min(5, len(pdf.pages))):
                    texto = pdf.pages[i].extract_text()
                    if texto:
                        texto_indice += texto + "\n"

                metadata = self.extraer_metadata_completa(texto_indice)
                print(f"📅 Fecha: {metadata['fecha']} | 🏛️  Tipo: {metadata['tipo_sesion']}")

                titulos = self.extraer_indice_completo(texto_indice) or []
                titulos_relevantes = []
                for titulo in titulos:
                    titulo_texto = titulo.get('titulo', '')
                    if self.es_titulo_relevante(titulo_texto):
                        numero_expediente = self.extraer_numero_expediente_del_titulo(titulo_texto)
                        titulos_relevantes.append({
                            'titulo': titulo_texto,
                            'pagina': titulo.get('pagina', 1),
                            'numero_expediente': numero_expediente,
                        })

                print(f"📋 Títulos: {len(titulos)} | Relevantes: {len(titulos_relevantes)}")
                return titulos_relevantes, metadata

        except Exception as e:
            print(f"❌ Error procesando PDF: {e}")
            return [], {}

print("✅ PARTE 8 CARGADA: Clase AnalizadorCompleto")


# ==============================================================
# PARTE 9 — FUNCIONES DE APOYO
# ==============================================================

import locale

def obtener_diccionario_por_fecha(fecha_str):
    """Obtiene el diccionario de oradores correspondiente a una fecha dada."""
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES')
        except locale.Error:
            pass

    fecha_obj = None
    try:
        patron_fecha_es = r'(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})'
        match_es = re.search(patron_fecha_es, fecha_str.lower())
        if match_es:
            dia_str, mes_str, año_str = match_es.groups()
            meses = {
                'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
                'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
                'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
            }
            mes_num = meses.get(mes_str)
            if mes_num:
                fecha_obj = datetime(int(año_str), mes_num, int(dia_str))

        if fecha_obj is None:
            for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%d %B %Y', '%B %d, %Y', '%d-%m-%Y', '%Y/%m/%d']:
                try:
                    fecha_obj = datetime.strptime(fecha_str, fmt)
                    break
                except ValueError:
                    pass

        if fecha_obj is None:
            print(f"      ⚠️  No se pudo parsear: '{fecha_str}'. Usando diccionario por defecto.")
            return DICCIONARIO_POR_DEFECTO

        for rango_inicio, rango_fin, diccionario in DICCIONARIOS_POR_FECHA:
            inicio_obj = datetime(*rango_inicio).date()
            fin_obj = datetime(*rango_fin).date() if rango_fin else None
            if fin_obj is None:
                if fecha_obj.date() >= inicio_obj:
                    return diccionario
            elif inicio_obj <= fecha_obj.date() <= fin_obj:
                return diccionario

        print(f"      ⚠️  Fecha {fecha_str} fuera de rango. Usando diccionario por defecto.")
        return DICCIONARIO_POR_DEFECTO

    except Exception as e:
        print(f"      ❌ Error con fecha '{fecha_str}': {e}. Usando diccionario por defecto.")
        return DICCIONARIO_POR_DEFECTO


def obtener_grupo_preciso(nombre_orador, texto_contexto, fecha_str):
    """Busca el grupo parlamentario de un orador por fecha."""
    nombre_normalizado = normalizar_nombre_orador_completo(nombre_orador)
    diccionario_legislatura = obtener_diccionario_por_fecha(fecha_str)

    if nombre_normalizado in diccionario_legislatura:
        return diccionario_legislatura[nombre_normalizado]

    for _, _, diccionario_alt in DICCIONARIOS_POR_FECHA:
        if diccionario_alt is not diccionario_legislatura:
            if nombre_normalizado in diccionario_alt:
                print(f"        ℹ️  Grupo encontrado en otra legislatura: {diccionario_alt[nombre_normalizado]}")
                return diccionario_alt[nombre_normalizado]

    grupo_contexto = extraer_grupo_del_contexto_corregido(texto_contexto)
    if grupo_contexto != "Desconocido":
        print(f"        ⚠️  Grupo del contexto: {grupo_contexto}")
        return grupo_contexto

    print(f"        ❌ Grupo no encontrado: {nombre_orador}")
    return "Desconocido"


def identificar_tipo_intervencion_corregido(texto_intervencion, nombre_orador):
    """Identifica el tipo de intervención basado en el texto."""
    texto_lower = texto_intervencion.lower()
    if any(p in texto_lower for p in ['turno a favor', 'intervengo a favor', 'para apoyar la propuesta']):
        return 'Intervención a favor'
    elif any(p in texto_lower for p in ['turno en contra', 'intervengo en contra', 'para oponerme a la propuesta']):
        return 'Intervención en contra'
    elif any(p in texto_lower for p in ['rectificación', 'para rectificar']):
        return 'Rectificación'
    elif any(p in texto_lower for p in ['réplica', 'para replicar']):
        return 'Réplica'
    elif any(p in texto_lower for p in ['dúplica', 'para duplicar']):
        return 'Dúplica'
    elif any(p in texto_lower for p in ['por alusiones', 'mencionado', 'aludido']):
        return 'Por alusiones'
    elif any(p in texto_lower for p in ['cuestión de orden', 'para una cuestión de orden']):
        return 'Cuestión de orden'
    elif any(p in texto_lower for p in ['pregunta', 'formular una pregunta', 'interrogar']):
        return 'Pregunta'
    elif any(p in texto_lower for p in ['respuesta', 'para responder', 'en respuesta a la pregunta']):
        return 'Respuesta'
    elif 'enmienda' in texto_lower:
        return 'Enmiendas'
    elif any(p in texto_lower for p in ['comparecencia', 'comparecer']):
        return 'Comparecencias'
    elif 'interpelación' in texto_lower:
        return 'Interpelaciones'
    elif any(p in texto_lower for p in ['proposición de ley', 'proposición no de ley']):
        return 'Defensa Proposición de Ley'
    elif 'moción' in texto_lower:
        return 'Defensa Moción'
    if any(cargo in nombre_orador.upper() for cargo in
           ['PRESIDENTA', 'PRESIDENTE', 'VICEPRESIDENTE', 'VICEPRESIDENTA', 'SECRETARIO', 'SECRETARIA']):
        return 'Gestión de debate'
    return 'Intervención en debate'

print("✅ PARTE 9 CARGADA: Funciones de apoyo")


# ==============================================================
# PARTE 10 — FUNCIÓN PRINCIPAL DE PROCESAMIENTO
# ==============================================================

def mostrar_configuracion():
    print("⚙️ CONFIGURACIÓN ACTUAL:")
    print(f"   • Ruta Drive: {ENLACE_DRIVE}")
    print(f"   • Palabras clave: {len(PALABRAS_CLAVE)}")
    for i, palabra in enumerate(PALABRAS_CLAVE[:6]):
        print(f"     {i+1}. {palabra}")
    if len(PALABRAS_CLAVE) > 6:
        print(f"     ... y {len(PALABRAS_CLAVE) - 6} más")
    print(f"   • Archivos a procesar: {'Todos' if NUMERO_ARCHIVOS_A_PROCESAR is None else NUMERO_ARCHIVOS_A_PROCESAR}")
    print("-" * 50)


def generar_metricas_calidad(df):
    """Genera métricas de calidad del procesamiento."""
    if len(df) == 0:
        print("📊 No hay datos para generar métricas")
        return
    total = len(df)
    sin_grupo = len(df[df['Grupo Parlamentario'] == 'Desconocido'])
    sin_tipo = len(df[df['Tipo Intervención'] == 'Intervención en debate'])
    textos_cortos = len(df[df['Intervención'].str.len() < 50])
    print(f"\n📊 MÉTRICAS DE CALIDAD:")
    print(f"   • Total intervenciones: {total}")
    print(f"   • Precisión grupos: {((total-sin_grupo)/total*100):.1f}% ({sin_grupo} sin grupo)")
    print(f"   • Precisión tipos: {((total-sin_tipo)/total*100):.1f}% ({sin_tipo} sin tipo)")
    print(f"   • Textos válidos: {((total-textos_cortos)/total*100):.1f}% ({textos_cortos} cortos)")
    print(f"   • Oradores únicos: {df['Orador'].nunique()}")
    print(f"   • Grupos detectados: {df['Grupo Parlamentario'].nunique()}")


def procesar_pdfs_simple_mejorada(lista_archivos_a_procesar=None):
    """Procesa PDFs y extrae corpus parlamentario."""
    print("PROCESANDO ARCHIVOS...")
    print("=" * 50)

    if not os.path.exists(ENLACE_DRIVE):
        print(f"Error: Ruta no existe: {ENLACE_DRIVE}")
        return

    analizador = AnalizadorCompleto(PALABRAS_CLAVE)
    archivos_pdf = [f for f in os.listdir(ENLACE_DRIVE) if f.lower().endswith('.pdf')]

    if lista_archivos_a_procesar is not None:
        archivos_a_procesar = lista_archivos_a_procesar
    elif NUMERO_ARCHIVOS_A_PROCESAR is None:
        archivos_a_procesar = archivos_pdf
    else:
        archivos_a_procesar = archivos_pdf[:NUMERO_ARCHIVOS_A_PROCESAR]

    todas_intervenciones = []
    debates_procesados_global = set()

    print(f"Archivos a procesar: {len(archivos_a_procesar)}")
    print("-" * 50)

    for archivo in archivos_a_procesar:
        print(f"Procesando: {archivo}")
        pdf_path = os.path.join(ENLACE_DRIVE, archivo)
        titulos_relevantes, metadata = analizador.procesar_pdf(pdf_path, archivo)

        if not titulos_relevantes:
            print(f"  ⚠️  Sin títulos relevantes en {archivo}")
            continue

        for titulo in titulos_relevantes:
            if not titulo or not isinstance(titulo, dict):
                continue

            titulo_texto = titulo.get('titulo', '')
            if not titulo_texto:
                continue

            titulo_hash = hash(titulo_texto.strip().lower()[:100])
            titulo_key = f"{archivo}_{titulo.get('pagina', 0)}_{titulo_hash}"

            if titulo_key in debates_procesados_global:
                print(f"  ⚠️  SALTO: debate ya procesado - {titulo_texto[:60]}...")
                continue
            debates_procesados_global.add(titulo_key)

            print(f"  Debate: {titulo_texto[:80]}...")
            print(f"  Fecha: {metadata.get('fecha', 'Desconocida')} | Sesión: {metadata.get('tipo_sesion', 'Desconocida')}")

            numero_expediente = titulo.get('numero_expediente', '')
            if numero_expediente:
                print(f"  📎 Expediente: {numero_expediente}")

            texto_debate = ""
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    start_page = max(0, titulo.get('pagina', 1) - 1)
                    end_page = min(len(pdf.pages), start_page + 35)
                    print(f"  📄 Páginas {start_page+1}–{end_page}")
                    for page_num in range(start_page, end_page):
                        texto = pdf.pages[page_num].extract_text()
                        if texto:
                            texto_debate += texto + "\n"
            except Exception as e:
                print(f"  Error leyendo PDF: {e}")
                continue

            if texto_debate and len(texto_debate) > 100:
                try:
                    intervenciones = extraer_intervenciones_mejorada_completa(
                        texto_debate,
                        titulo_texto,
                        archivo,
                        metadata,
                        analizador,
                        numero_expediente,
                    )
                    if intervenciones:
                        intervenciones_unicas = eliminar_duplicados_inteligente(intervenciones)
                        duplicados = len(intervenciones) - len(intervenciones_unicas)
                        if duplicados > 0:
                            print(f"  🧹 {duplicados} duplicados eliminados")
                        todas_intervenciones.extend(intervenciones_unicas)
                        grupos = set(iv['Grupo Parlamentario'] for iv in intervenciones_unicas if iv['Grupo Parlamentario'] != 'MESA')
                        print(f"  Intervenciones: {len(intervenciones_unicas)} | Grupos: {', '.join(sorted(grupos))}")
                    else:
                        print("  Sin intervenciones")
                except Exception as e:
                    print(f"  ❌ Error en extracción: {e}")
                    continue
            else:
                print("  Texto insuficiente")

    if todas_intervenciones:
        print(f"📊 Total antes de deduplicación final: {len(todas_intervenciones)}")
        todas_intervenciones = eliminar_duplicados_inteligente(todas_intervenciones)
        print(f"📊 Total final: {len(todas_intervenciones)}")

        df = pd.DataFrame(todas_intervenciones)
        generar_metricas_calidad(df)

        df.to_excel(OUTPUT_PATH, index=False)
        print("\n" + "=" * 50)
        print("✅ PROCESAMIENTO COMPLETADO")
        print("=" * 50)
        print(f"Total intervenciones: {len(todas_intervenciones)}")
        print(f"Archivos procesados:  {len(archivos_a_procesar)}")
        print(f"Oradores únicos:      {df['Orador'].nunique()}")
        print(f"Grupos detectados:    {df['Grupo Parlamentario'].nunique()}")
        print(f"Con ID expediente:    {(df['ID Discurso'] != '').sum()}")
        print(f"Guardado en:          {OUTPUT_PATH}")

        grupos_data = df[df['Grupo Parlamentario'] != 'MESA']
        if len(grupos_data) > 0:
            print(f"\nDISTRIBUCIÓN:")
            for grupo, count in grupos_data['Grupo Parlamentario'].value_counts().items():
                print(f"  {grupo}: {count}")

        return df
    else:
        print("\nNo se encontraron intervenciones.")
        return None


# ==============================================================
# EJECUCIÓN
# ==============================================================

mostrar_configuracion()
resultado = procesar_pdfs_simple_mejorada(archivos_a_procesar_aleatorios)
