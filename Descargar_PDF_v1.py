# ============================================================
# Descargar_PDF — From PDF to Corpus
# Sara Sampayo Sande · Centro Crímina, UMH Elche · MIT License
# https://github.com/[URL]
# ============================================================
#
# Descarga automatizada de Diarios de Sesiones del Congreso
# de los Diputados filtrados por palabra clave, legislatura
# y cámara. Diseñada para Google Colaboratory.
#
# NOTA: La búsqueda por texto libre solo está disponible
# a partir de la VI Legislatura.
# ============================================================


# ==============================================================
# CELDA 1 — INSTALACIÓN (ejecutar solo una vez por sesión)
# ==============================================================

# !pip install requests beautifulsoup4 tqdm tabulate -q

import requests
import os
import time
import json
from tqdm import tqdm
from urllib.parse import quote

print("✅ Librerías cargadas")

# Obtener cookies frescas del portal del Congreso
print("\n🍪 Obteniendo cookies del Congreso...")

try:
    session = requests.Session()
    response = session.get(
        'https://www.congreso.es/gl/busqueda-de-publicaciones',
        timeout=10
    )
    cookies = session.cookies.get_dict()
    JSESSIONID = cookies.get('JSESSIONID', '')

    if JSESSIONID:
        print(f"✅ Cookies obtenidas correctamente")
    else:
        print("⚠️ No se obtuvo JSESSIONID, usando cookie por defecto")
        JSESSIONID = 'esRur9kVmDnbOYAqryVEbekB9nENIQPNV48PDFDn.cgdpjbnode2pro'

except Exception as e:
    print(f"⚠️ Error obteniendo cookies: {e}. Usando cookie por defecto.")
    JSESSIONID = 'esRur9kVmDnbOYAqryVEbekB9nENIQPNV48PDFDn.cgdpjbnode2pro'

COOKIES = {'JSESSIONID': JSESSIONID}
print("✅ Listo para continuar\n")


# ==============================================================
# CELDA 2 — CONFIGURACIÓN DE BÚSQUEDA (rellenar y ejecutar)
# ==============================================================

# 🎯 PALABRAS CLAVE — separar con punto y coma (;)
PALABRAS_CLAVE_INPUT = "Libertad Sexual; Violencia de género"  # @param {type:"string"}

# 📅 LEGISLATURAS — números separados por comas
LEGISLATURAS_INPUT = "15,14,13,12"  # @param {type:"string"}

# 📄 TIPO DE PUBLICACIÓN
TIPO_PUBLICACION = "Diarios de Sesiones"  # @param ["Diarios de Sesiones", "Boletines Oficiales", "Todos"]

# 🏛️ CÁMARA
CAMARA = "Congreso de los Diputados"  # @param ["Congreso de los Diputados", "Senado", "Cortes Xerais", "Todas"]

# 📁 CARPETA DE SALIDA
CARPETA_BASE = "descargas_congreso"  # @param {type:"string"}

# 🔧 OPCIONES AVANZADAS
PAUSA_ENTRE_DESCARGAS = 0.5  # @param {type:"slider", min:0.1, max:2.0, step:0.1}

# ------------------------------------------------------------------
# Procesamiento interno (no modificar)
# ------------------------------------------------------------------
TIPO_CODIGO = {
    "Diarios de Sesiones": "D",
    "Boletines Oficiales": "B",
    "Todos": ""
}
CAMARA_CODIGO = {
    "Congreso de los Diputados": "CONGRESO",
    "Senado": "SENADO",
    "Cortes Xerais": "CORTES",
    "Todas": ""
}

PALABRAS_CLAVE = [p.strip() for p in PALABRAS_CLAVE_INPUT.split(';') if p.strip()]
LEGISLATURAS = [int(x.strip()) for x in LEGISLATURAS_INPUT.split(',') if x.strip().isdigit()]
TIPO_SELECCIONADO = TIPO_CODIGO[TIPO_PUBLICACION]
CAMARA_SELECCIONADA = CAMARA_CODIGO[CAMARA]

print("=" * 60)
print("✅ CONFIGURACIÓN CARGADA")
print("=" * 60)
print(f"📌 Palabras clave: {PALABRAS_CLAVE}")
print(f"📅 Legislaturas:   {LEGISLATURAS}")
print(f"📄 Tipo:           {TIPO_PUBLICACION}")
print(f"🏛️  Cámara:         {CAMARA}")
print(f"📁 Carpeta salida: {CARPETA_BASE}")
print("=" * 60)


# ==============================================================
# CELDA 3 — FUNCIONES AUXILIARES (no modificar)
# ==============================================================

HEADERS = {
    'accept': 'application/json, text/javascript, */*; q=0.01',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'x-requested-with': 'XMLHttpRequest',
}


def buscar_en_legislatura(legislatura, palabra_clave, pagina=1):
    """Busca documentos en una legislatura usando la API interna del Congreso."""
    url = "https://www.congreso.es/gl/busqueda-de-publicaciones"
    params = {
        'p_p_id': 'publicaciones',
        'p_p_lifecycle': '2',
        'p_p_resource_id': 'filtrarListado',
    }
    data = {
        '_publicaciones_legislatura': str(legislatura),
        '_publicaciones_texto': f'"{palabra_clave}"',
        '_publicaciones_tipoBusqueda': '0',
        '_publicaciones_publicacion': TIPO_SELECCIONADO,
        '_publicaciones_seccion': CAMARA_SELECCIONADA,
        '_publicaciones_paginaActual': str(pagina),
    }
    data = {k: v for k, v in data.items() if v != ''}

    try:
        respuesta = requests.post(
            url, params=params, data=data,
            headers=HEADERS, cookies=COOKIES, timeout=15
        )
        respuesta.raise_for_status()
        return respuesta.json()
    except Exception as e:
        print(f"   ❌ Error en búsqueda (página {pagina}): {e}")
        return None


def extraer_documentos(data):
    """Extrae la lista de documentos de la respuesta JSON de la API."""
    docs = []
    if isinstance(data, dict):
        for key, value in data.items():
            if key.startswith('documento') and isinstance(value, dict):
                if 'diario' in value and 'cve' in value:
                    docs.append({
                        'cve': value['cve'],
                        'url': f"https://www.congreso.es{value['diario']}",
                        'orga': value.get('orga', ''),
                        'fecha': value.get('fecha', '')
                    })
    return docs


def obtener_total_paginas(data):
    """Calcula el número total de páginas a partir del campo publicaciones_encontradas."""
    if isinstance(data, dict) and 'publicaciones_encontradas' in data:
        try:
            total = int(data['publicaciones_encontradas'])
            if total > 0:
                return (total + 19) // 20
        except (ValueError, TypeError):
            pass
    return 1


def descargar_pdf(url, ruta, verbose=False):
    """
    Descarga un PDF y lo guarda en la ruta indicada.
    Retorna True si la descarga fue exitosa, False en caso contrario.
    El parámetro verbose controla si se imprimen mensajes de progreso.
    """
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        resp = requests.get(
            url,
            headers={'User-Agent': 'Mozilla/5.0'},
            stream=True,
            timeout=30
        )
        resp.raise_for_status()

        with open(ruta, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        if os.path.exists(ruta) and os.path.getsize(ruta) > 1000:
            if verbose:
                kb = os.path.getsize(ruta) / 1024
                print(f"      ✅ {os.path.basename(ruta)} ({kb:.1f} KB)")
            return True
        else:
            if os.path.exists(ruta):
                os.remove(ruta)
            return False

    except Exception as e:
        if verbose:
            print(f"      ❌ Error: {e}")
        return False


def generar_enlace_supervision(legislatura, palabra):
    """Genera un enlace para verificar la búsqueda manualmente en el navegador."""
    palabra_codificada = quote(palabra)
    return (
        f"https://www.congreso.es/gl/busqueda-de-publicaciones"
        f"?p_p_id=publicaciones&p_p_lifecycle=0"
        f"&_publicaciones_legislatura={legislatura}"
        f"&_publicaciones_texto=%22{palabra_codificada}%22"
        f"&_publicaciones_publicacion={TIPO_SELECCIONADO}"
        f"&_publicaciones_seccion={CAMARA_SELECCIONADA}"
    )


print("✅ Funciones cargadas correctamente")


# ==============================================================
# CELDA 4 — BÚSQUEDA Y DESCARGA
# ==============================================================

os.makedirs(CARPETA_BASE, exist_ok=True)
total_global = 0
resultados = {}

print("=" * 60)
print("🚀 INICIANDO BÚSQUEDA Y DESCARGA")
print("=" * 60)

for palabra in PALABRAS_CLAVE:
    print(f"\n📌 BUSCANDO: '{palabra}'")

    for legislatura in LEGISLATURAS:
        print(f"\n   📄 Legislatura {legislatura}")
        enlace = generar_enlace_supervision(legislatura, palabra)
        print(f"      🔗 {enlace}")

        carpeta = os.path.join(
            CARPETA_BASE,
            palabra.replace(' ', '_'),
            f"leg_{legislatura}"
        )
        os.makedirs(carpeta, exist_ok=True)

        # Primera página para conocer el total
        data = buscar_en_legislatura(legislatura, palabra, pagina=1)
        if not data:
            print("      ❌ Sin respuesta del servidor")
            continue

        total_docs = int(data.get('publicaciones_encontradas', 0))
        print(f"      📊 Documentos encontrados: {total_docs}")
        if total_docs == 0:
            continue

        total_paginas = obtener_total_paginas(data)

        # Recopilar todos los documentos paginando
        todos_docs = extraer_documentos(data)
        for pagina in range(2, total_paginas + 1):
            data_pag = buscar_en_legislatura(legislatura, palabra, pagina=pagina)
            if data_pag:
                todos_docs.extend(extraer_documentos(data_pag))
            time.sleep(0.5)

        print(f"      📋 Total en lista: {len(todos_docs)} documentos")

        # Descarga con barra de progreso
        descargados = 0
        with tqdm(total=len(todos_docs), desc="      Descargando", unit="PDF", leave=True) as pbar:
            for idx, doc in enumerate(todos_docs, 1):
                nombre = f"{idx:04d}_{doc['cve']}.PDF"
                ruta = os.path.join(carpeta, nombre)
                if descargar_pdf(doc['url'], ruta):
                    descargados += 1
                pbar.update(1)
                time.sleep(PAUSA_ENTRE_DESCARGAS)

        print(f"      ✅ {descargados}/{len(todos_docs)} descargados correctamente")

        # Guardar lista de CVEs
        with open(os.path.join(carpeta, 'lista_completa.txt'), 'w', encoding='utf-8') as f:
            for idx, doc in enumerate(todos_docs, 1):
                f.write(f"{idx:04d}\t{doc['cve']}\t{doc.get('orga', '')}\n")

        resultados[(palabra, legislatura)] = len(todos_docs)
        total_global += len(todos_docs)

print("\n" + "=" * 60)
print("📊 RESUMEN FINAL")
print("=" * 60)
for (palabra, leg), count in resultados.items():
    print(f"  '{palabra}' — Legislatura {leg}: {count} documentos")
print(f"\n✅ TOTAL: {total_global} documentos | Carpeta: {CARPETA_BASE}")
print("=" * 60)


# ==============================================================
# CELDA 5 — RESUMEN DE DESCARGAS (opcional)
# ==============================================================

import pandas as pd
from tabulate import tabulate

print("=" * 60)
print("📊 RESUMEN DE DESCARGAS EN DISCO")
print("=" * 60)

if not os.path.exists(CARPETA_BASE):
    print(f"❌ La carpeta '{CARPETA_BASE}' no existe.")
else:
    resumen = []
    for palabra in os.listdir(CARPETA_BASE):
        ruta_palabra = os.path.join(CARPETA_BASE, palabra)
        if not os.path.isdir(ruta_palabra):
            continue
        for leg_folder in os.listdir(ruta_palabra):
            ruta_leg = os.path.join(ruta_palabra, leg_folder)
            if not os.path.isdir(ruta_leg):
                continue
            pdfs = [f for f in os.listdir(ruta_leg) if f.upper().endswith('.PDF')]
            if pdfs:
                resumen.append({
                    'Palabra clave': palabra.replace('_', ' '),
                    'Legislatura': leg_folder.replace('leg_', ''),
                    'PDFs descargados': len(pdfs)
                })

    if resumen:
        df = pd.DataFrame(resumen).sort_values(['Palabra clave', 'Legislatura'])
        print(tabulate(df, headers='keys', tablefmt='grid', showindex=False))
        print(f"\n✅ TOTAL: {df['PDFs descargados'].sum()} PDFs")

        csv_path = "/content/resumen_descargas.csv"
        df.to_csv(csv_path, index=False)
        print(f"💾 Resumen guardado en: {csv_path}")
    else:
        print("❌ No se encontraron archivos descargados.")


# ==============================================================
# CELDA 6 — DETECCIÓN Y ELIMINACIÓN DE DUPLICADOS (opcional)
# ==============================================================

from collections import defaultdict

print("=" * 60)
print("🔍 DETECCIÓN DE DUPLICADOS")
print("=" * 60)

# Opciones: "Solo listar" o "Eliminar duplicados"
MODO = "Solo listar"  # @param ["Solo listar", "Eliminar duplicados"]

archivos = []
for root, dirs, files in os.walk(CARPETA_BASE):
    for file in files:
        if file.upper().endswith('.PDF'):
            legislatura = next(
                (p.replace('leg_', '') for p in root.split(os.sep) if p.startswith('leg_')),
                "desconocida"
            )
            cve = file.split('_', 1)[1].replace('.PDF', '') if '_' in file else file.replace('.PDF', '')
            indice = file.split('_', 1)[0] if '_' in file else "0000"
            archivos.append({
                'leg': legislatura,
                'cve': cve,
                'indice': indice,
                'ruta': os.path.join(root, file),
                'tamaño_kb': os.path.getsize(os.path.join(root, file)) / 1024
            })

print(f"📊 Total archivos analizados: {len(archivos)}")

total_eliminados = 0
espacio_liberado = 0

for leg in sorted(set(a['leg'] for a in archivos)):
    archivos_leg = [a for a in archivos if a['leg'] == leg]
    cve_dict = defaultdict(list)
    for a in archivos_leg:
        cve_dict[a['cve']].append(a)

    for cve, lista in cve_dict.items():
        if len(lista) > 1:
            lista_ord = sorted(lista, key=lambda x: x['indice'])
            for a in lista_ord[1:]:
                if MODO == "Eliminar duplicados":
                    try:
                        os.remove(a['ruta'])
                        total_eliminados += 1
                        espacio_liberado += a['tamaño_kb'] / 1024
                    except Exception:
                        pass

if MODO == "Solo listar":
    print("ℹ️  Modo solo listar — no se eliminó nada")
else:
    print(f"✅ Eliminados: {total_eliminados} archivos ({espacio_liberado:.2f} MB liberados)")


# ==============================================================
# CELDA 7 — DESCARGAR CORPUS COMPLETO COMO ZIP
# ==============================================================

import zipfile
from datetime import datetime
from google.colab import files  # Solo funciona en Google Colab

print("=" * 60)
print("📦 COMPRIMIR Y DESCARGAR CORPUS")
print("=" * 60)

INCLUIR_FECHA = True   # @param {type:"boolean"}
ELIMINAR_DESPUES = False  # @param {type:"boolean"}

fecha = datetime.now().strftime("%Y%m%d_%H%M%S") if INCLUIR_FECHA else ""
nombre_zip = f"corpus_congreso_{fecha}.zip" if fecha else "corpus_congreso.zip"
ruta_zip = f"/content/{nombre_zip}"

if not os.path.exists(CARPETA_BASE):
    print(f"❌ La carpeta '{CARPETA_BASE}' no existe.")
else:
    # Estadísticas
    total_archivos = sum(
        1 for _, _, fs in os.walk(CARPETA_BASE)
        for f in fs if f.upper().endswith('.PDF')
    )
    print(f"📊 Archivos a comprimir: {total_archivos}")
    print("⏳ Comprimiendo...")

    comprimidos = 0
    with zipfile.ZipFile(ruta_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, fs in os.walk(CARPETA_BASE):
            for file in fs:
                if file.upper().endswith('.PDF'):
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, start=os.path.dirname(CARPETA_BASE))
                    zipf.write(file_path, arcname)
                    comprimidos += 1

    tamaño_zip = os.path.getsize(ruta_zip) / (1024 * 1024)
    print(f"✅ ZIP creado: {nombre_zip} ({tamaño_zip:.2f} MB)")

    try:
        files.download(ruta_zip)
        print("✅ Descarga iniciada")
    except Exception as e:
        print(f"⚠️ Descarga automática fallida: {e}")
        print(f"   → Descarga manual: panel izquierdo de Colab > busca '{nombre_zip}'")

    if ELIMINAR_DESPUES:
        os.remove(ruta_zip)
        print("🧹 ZIP eliminado del servidor")
