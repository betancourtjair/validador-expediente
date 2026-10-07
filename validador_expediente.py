#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validador de Expediente de Reclutamiento — Fitness Para Todos
================================================================

Lee un conjunto de PDFs (uno por documento), identifica de qué documento se
trata, extrae texto (con OCR automático si el PDF es una imagen escaneada),
corre las reglas de validación que pidió el equipo de reclutamiento, y
genera un Excel (.xlsx) con el resultado para anexar al expediente.

REGLAS IMPLEMENTADAS
---------------------
1. Legibilidad: se extrae texto nativo del PDF; si una página no trae texto
   (documento escaneado) se corre OCR (Tesseract, español). Si ni el texto
   nativo ni el OCR logran leer nada útil, se marca "ILEGIBLE".
2. Pertenencia: se compara el nombre encontrado en cada documento contra el
   nombre registrado del candidato (comparación por tokens, tolerante a
   orden Nombre/Apellido y a acentos). El Comprobante de domicilio se
   excluye de esta regla (puede estar a nombre de otra persona).
3. Completitud: se coteja contra la lista de documentos obligatorios y se
   reporta qué falta. Cada documento debe llegar en un solo PDF (esto se
   valida por diseño: se procesa un archivo = un documento).
4. INE: debe traer 2 páginas en el mismo PDF (frente y reverso) y la
   vigencia impresa no debe haber pasado ya.
5. CSF: debe traer 2 páginas en el mismo PDF y la fecha de emisión no debe
   tener más de 3 meses. Se lee también el "Estatus en el padrón" que la
   propia Constancia imprime (Activo / Suspendido / Cancelado) como
   verificación de autenticidad de primer nivel.
6. Comprobante de domicilio: no debe tener más de 3 meses de antigüedad.
7. Cuenta bancaria: se identifica el banco a partir de la CLABE (los
   primeros 3 dígitos son el código de institución bancaria, es más
   confiable que buscar el nombre del banco en el texto). Se valida que
   existan cuenta, CLABE y nombre del colaborador, y se rechaza si el banco
   es Nu, Spin by OXXO o Mercado Pago.
8. Datos extraídos para el resumen: además de las reglas de arriba, el
   resumen muestra el RFC (leído de la CSF, junto a la etiqueta "RFC"), el
   CURP (leído del documento CURP, formato de 18 caracteres), la fecha de
   nacimiento (leída del acta, junto a la etiqueta "FECHA DE NACIMIENTO") y
   la dirección (leída por separado del comprobante de domicilio y del
   INE). De la CSF también se leen, de la propia constancia: el CURP, el
   Código Postal (bajo "Datos del domicilio registrado") y el Régimen
   (bajo "Regímenes", en la hoja 2). De los avisos de retención -ahora
   documentos separados- se lee el Número de crédito: en Infonavit, bajo
   el apartado "Información del crédito del trabajador"; en Fonacot, junto
   a la etiqueta "número de crédito". Igual que el resto de los datos
   leídos por OCR, son una propuesta a confirmar contra el documento
   original, no un dato ya verificado.

LIMITACIONES IMPORTANTES (léelas antes de confiar 100% en el resultado)
------------------------------------------------------------------------
- La verificación de que la carátula bancaria "incluya el logo del banco"
  es una revisión VISUAL; el OCR no puede confirmar que hay un logotipo.
  El script sí valida el banco (vía CLABE), cuenta, CLABE y nombre, pero
  marca el logo como "revisar visualmente".
- La autenticidad plena de la CSF ante el SAT requiere consultar el
  portal oficial (siat.sat.gob.mx); este script no hace esa consulta en
  vivo (no es un servicio público disponible para automatizar). Usa el
  estatus impreso en el documento como primera señal y genera, si el QR
  es legible, el enlace para que alguien lo confirme con un clic.
- El OCR de documentos escaneados no es perfecto. Toda fecha o nombre
  extraído por OCR debe leerse como "propuesta a confirmar", no como
  verdad absoluta — por eso el reporte siempre incluye el fragmento de
  texto de donde salió el dato.

USO
---
    python3 validador_expediente.py --nombre "Alam Naresh Poot Cauich" \
        --rfc POCA0201084WA --curp POCA020108HYNTCLA6 \
        --salida expediente_alam.xlsx \
        archivo1.pdf archivo2.pdf ...

Requiere: pdfplumber, pytesseract, pdf2image, openpyxl, y tener instalados
en el sistema los binarios `tesseract` (con el paquete de idioma `spa`) y
`poppler-utils` (pdftoppm). En Debian/Ubuntu:
    apt-get install -y tesseract-ocr tesseract-ocr-spa poppler-utils
    pip install pdfplumber pytesseract pdf2image openpyxl
"""

import argparse
import datetime
import io
import os
import re
import sys
import time
import unicodedata
import zipfile

import pdfplumber
from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageStat, ImageFilter
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from zoneinfo import ZoneInfo
    _ZONA_CDMX = ZoneInfo("America/Mexico_City")
except Exception:
    # Si por algún motivo el contenedor no trae la base de datos de zonas
    # horarias (IANA tzdata), no queremos que la app truene: seguimos con la
    # fecha del sistema (normalmente UTC) en vez de la de CDMX.
    _ZONA_CDMX = None

# "Hoy", pero en la zona horaria de Ciudad de México (no la del contenedor,
# que en Cloud Run corre en UTC) — de aquí salen tanto la fecha que se
# imprime en "Generado: ..." como todos los cálculos de vigencia/antigüedad
# (días vencidos, edad, etc.) que usan HOY en este archivo.
HOY = (
    datetime.datetime.now(_ZONA_CDMX).date()
    if _ZONA_CDMX is not None
    else datetime.date.today()
)

# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------

def normaliza(txt):
    """Mayúsculas, sin acentos, espacios colapsados."""
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", txt).upper().strip()


def _sin_acentos_mismo_largo(txt):
    """Como normaliza(), pero SIN colapsar espacios ni recortar: cada
    carácter acentuado (á, é, ñ, ü, ...) se reduce a su base de 1 solo
    carácter, así que el resultado queda con exactamente el mismo largo
    (y los mismos índices) que 'txt'. Se usa para poder buscar una
    etiqueta ignorando acentos/mayúsculas y luego recortar el fragmento
    que sigue directamente del texto ORIGINAL (con acentos y mayúsculas
    tal cual los trae el documento), en vez de quedarse con la versión
    en mayúsculas sin acentos -mucho menos legible para quien revisa el
    reporte, por ejemplo al extraer una dirección."""
    if not txt:
        return ""
    return unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii").upper()


MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}

MESES_ABREV = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}

# Alternativa explícita de los nombres de mes (en vez de "cualquier letra")
# para el patrón de fecha "<dia> DE <mes> DE <anio>" — ver el porqué en
# PATRON_FECHA_LARGA más abajo. Se ordenan de más largo a más corto nomás
# por higiene (ninguno es prefijo de otro, así que en la práctica no cambia
# el resultado, pero evita sorpresas si algún día se agrega un mes con
# nombre parecido a otro).
_MESES_ALTERNATIVAS = "|".join(sorted(MESES.keys(), key=len, reverse=True))

# Patrón de fecha "<dia> DE <mes> DE <anio>" (p.ej. "21 DE AGOSTO DE 2026").
# Importante: el grupo del mes usa la alternativa explícita de nombres de
# mes (_MESES_ALTERNATIVAS), NO "cualquier letra" ([A-ZÑ]+) como se hacía
# antes. La razón: el texto extraído de algunos PDFs (la CSF en particular)
# a veces pierde los espacios entre palabras ("21DEAGOSTODE2026"). Con
# "cualquier letra", el grupo del mes se comía de forma "greedy" el "DE"
# que le sigue (dejando "AGOSTODE", que no es un mes válido) y la fecha se
# perdía por completo -eso hacía que, en la CSF, la fecha real de emisión
# no se detectara y el script cayera en otra fecha del documento (como la
# fecha de inicio de operaciones del contribuyente, mucho más vieja),
# marcando por error la constancia como fuera de vigencia. Al limitar el
# grupo del mes a los nombres de mes reales, el regex ya no puede "comerse"
# el "DE" siguiente y la fecha se reconoce igual, con o sin espacios.
PATRON_FECHA_LARGA = re.compile(
    r"(\d{1,2})\s*(?:DE|DEL)?\s*(" + _MESES_ALTERNATIVAS + r")\s*(?:DE|DEL)?\s*(\d{4})",
    re.IGNORECASE,
)


def busca_fechas(texto):
    """Devuelve una lista de (date, fragmento_texto) encontradas en el texto,
    soportando '20/05/2026' y '20 de mayo de 2026' / '20 DE MAYO DE 2026'."""
    fechas = []
    t = texto or ""

    for m in re.finditer(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", t):
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            fechas.append((datetime.date(y, mo, d), m.group(0)))
        except ValueError:
            pass

    for m in PATRON_FECHA_LARGA.finditer(normaliza(t)):
        dia, mes_txt, anio = m.group(1), m.group(2), m.group(3)
        mes = MESES.get(mes_txt.upper())
        if mes:
            try:
                fechas.append((datetime.date(int(anio), mes, int(dia)), m.group(0)))
            except ValueError:
                pass

    # formato abreviado tipo recibo (ej. CFE): "03 MAY 26", "03/MAY/26" y, como
    # lo imprimen algunos recibos (Naturgy/CFE) y lo lee el OCR, sin
    # separadores ("25Ago26") o con un cero de más en el año ("25Ago026").
    patron_abrev = re.compile(r"(\d{1,2})\s*[/\s\-\.]?\s*([A-Z]{3})\s*[/\s\-\.]?\s*(\d{4}|0?\d{2})\b")
    for m in patron_abrev.finditer(normaliza(t)):
        dia, mes_txt, anio_txt = m.group(1), m.group(2), m.group(3)
        mes = MESES_ABREV.get(mes_txt.upper())
        if mes:
            anio = int(anio_txt) if len(anio_txt) == 4 else 2000 + int(anio_txt[-2:])
            try:
                fechas.append((datetime.date(anio, mes, int(dia)), m.group(0)))
            except ValueError:
                pass

    # ISO "2026-08-12" y la fecha de timbrado de una factura electrónica
    # ("Fecha Timbrado 20260812 23:53:05" / "Fecha Timbrado: 2026-08-12").
    for m in re.finditer(r"\b(20\d{2})-(\d{2})-(\d{2})\b", t):
        try:
            fechas.append((datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))), m.group(0)))
        except ValueError:
            pass
    for m in re.finditer(r"FECHA\s*(?:DE\s*)?(?:TIMBRADO|EMISION|EXPEDICION)\D{0,6}(20\d{2})(\d{2})(\d{2})\b", normaliza(t)):
        try:
            fechas.append((datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))), m.group(0)))
        except ValueError:
            pass

    return fechas


def fecha_mas_reciente_razonable(fechas, no_futuras=True):
    """De una lista de (date, frag), regresa la más reciente que no sea del
    futuro lejano (para evitar folios/números mal interpretados como fecha)."""
    candidatas = [f for f in fechas if f[0].year >= 2015 and f[0].year <= HOY.year + 1]
    if no_futuras:
        candidatas = [f for f in candidatas if f[0] <= HOY]
    if not candidatas:
        return None
    return max(candidatas, key=lambda f: f[0])


# ---------------------------------------------------------------------------
# Extracción de texto (nativo + OCR de respaldo)
# ---------------------------------------------------------------------------

# Límite de tiempo (segundos) por cada llamada a Tesseract. Sin esto, una
# imagen "difícil" (foto muy ruidosa, documento con un layout raro) puede
# hacer que el análisis de la página se tarde muchísimo y — como ya pasó una
# vez con Render — deje la generación del Excel como pasmada. Con el límite,
# si una sola llamada se pasa de tiempo simplemente se descarta ese intento
# (se trata como si no hubiera podido leer nada) y se sigue con el siguiente,
# en vez de quedarse trabada ahí.
OCR_TIMEOUT_SEGUNDOS = 12


def _ocr_con_confianza(imagen, lang="spa", config="", timeout=None):
    """Corre OCR y regresa (texto, confianza_promedio_0_a_100).

    A diferencia de adivinar qué tan "limpio" se ve un texto contando tipos
    de caracteres (eso falla: un OCR que lee mal puede seguir escupiendo
    letras y espacios "normales", solo que las palabras equivocadas — se
    ve limpio pero está mal), aquí se usa la confianza que el propio
    Tesseract calcula por palabra (0-100, viene de image_to_data) y se
    promedia. Es la señal más confiable para decidir si vale la pena
    escalar a un pase más pesado, o para elegir cuál de varias
    configuraciones de Tesseract leyó mejor una misma imagen."""
    try:
        datos = pytesseract.image_to_data(
            imagen, lang=lang, config=config, timeout=timeout or OCR_TIMEOUT_SEGUNDOS,
            output_type=pytesseract.Output.DICT,
        )
    except Exception:
        return "", -1.0

    lineas = {}
    confianzas = []
    n = len(datos.get("text", []))
    for i in range(n):
        texto = (datos["text"][i] or "").strip()
        if not texto:
            continue
        clave_linea = (datos["block_num"][i], datos["par_num"][i], datos["line_num"][i])
        lineas.setdefault(clave_linea, []).append(texto)
        try:
            conf = float(datos["conf"][i])
        except (TypeError, ValueError):
            conf = -1.0
        if conf >= 0:
            confianzas.append(conf)

    texto_completo = "\n".join(" ".join(palabras) for palabras in lineas.values())
    confianza_prom = sum(confianzas) / len(confianzas) if confianzas else -1.0
    return texto_completo, confianza_prom


def _preprocesa_para_ocr(imagen, nivel="ligero"):
    """Prepara una imagen antes de mandarla al OCR.
    - 'ligero': escala de grises + autocontraste. No es OCR, es solo
      procesamiento de imagen (barato en CPU), así que se aplica siempre
      desde el pase rápido — ayuda bastante con fotos de celular con poco
      contraste o fondos de color, sin costo extra relevante.
    - 'fuerte': lo anterior + binarización (blanco/negro, usando el brillo
      promedio de la imagen para fijar el umbral) + nitidez. Es lo que más
      ayuda con documentos que a simple vista se ven bien pero el OCR lee
      mal: fondos de color en credenciales, sombras, brillo del flash, papel
      térmico desteñido. Es más pesado, por eso solo se usa en el pase HD, y
      solo para las páginas que ya fallaron el pase rápido."""
    gris = ImageOps.grayscale(imagen)
    # OJO: cutoff>0 aquí puede ser CONTRAPRODUCENTE. Se probó con cutoff=2
    # (recorta el 2% más oscuro/claro del histograma antes de estirar el
    # contraste) y en un documento de fondo casi blanco con poco texto
    # oscuro, ese recorte metía un patrón de moteado en las letras que hacía
    # que Tesseract no reconociera NADA en un documento perfectamente
    # legible a simple vista. cutoff=0 (usar el mínimo/máximo real de la
    # imagen para estirar el contraste, sin recortar nada) no tiene ese
    # problema y sigue ayudando con fotos de bajo contraste.
    gris = ImageOps.autocontrast(gris, cutoff=0)
    if nivel == "ligero":
        return gris
    brillo_medio = ImageStat.Stat(gris).mean[0]
    umbral = max(100, min(200, brillo_medio * 0.85))
    binaria = gris.point(lambda x: 255 if x > umbral else 0)
    return binaria.filter(ImageFilter.SHARPEN)


def _corrige_rotacion(imagen):
    """Detecta si la imagen viene rotada 90/180/270° —común cuando el
    candidato fotografía su credencial o comprobante con el celular en la
    orientación equivocada, lo cual hace que Tesseract no reconozca casi nada
    aunque el documento sea perfectamente legible— y la corrige antes del
    OCR. Si Tesseract no logra determinar la orientación (imagen muy
    ruidosa) simplemente se deja igual."""
    try:
        osd = pytesseract.image_to_osd(
            imagen, output_type=pytesseract.Output.DICT, timeout=OCR_TIMEOUT_SEGUNDOS
        )
        angulo = int(osd.get("rotate", 0) or 0)
        # PIL's Image.rotate(x) gira en sentido antihorario x grados; el
        # campo "rotate" que regresa Tesseract ya viene en la convención que
        # hace falta pasarle directo (probado empíricamente: usar -angulo
        # deja la imagen al revés). Por eso aquí es rotate(angulo), sin signo
        # invertido.
        if angulo:
            return imagen.rotate(angulo, expand=True)
    except Exception:
        pass
    return imagen


def _extraer_imagenes_incrustadas(ruta, indice_pagina):
    """Extrae, tal cual, las imágenes que ya vienen incrustadas en el PDF
    para esa página (JPEG/PNG originales) -sin volver a renderizar la
    página completa con poppler.

    Se descubrió con un caso real (INE fotografiada/escaneada e insertada
    como imagen en el PDF) que renderizar la página vía convert_from_path
    (aunque sea a 300 dpi) puede salir con MENOS calidad para el OCR que
    la imagen original: el remuestreo/recompresión de poppler, sumado al
    autocontraste, puede volver el fondo decorativo de una credencial
    (mapa de México, patrón de fondo) casi indistinguible del texto,
    mientras que Tesseract lee bien la imagen JPEG original tal cual.
    Por eso esto se usa como una candidata MÁS a competir por confianza
    junto a los pases de poppler, nunca como reemplazo directo (hay PDFs
    sin imágenes incrustadas -texto nativo-, o con imágenes en formatos
    que no se pueden decodificar aquí).

    Se descartan imágenes muy chicas (sellos/logos) porque no traen el
    contenido principal del documento. Regresa una lista de imágenes PIL
    en RGB, ordenadas de arriba hacia abajo según su posición en la
    página (para páginas con más de una imagen, p.ej. anverso y reverso
    de una credencial escaneados juntos)."""
    imagenes = []
    try:
        with pdfplumber.open(ruta) as pdf:
            if indice_pagina >= len(pdf.pages):
                return []
            pagina = pdf.pages[indice_pagina]
            candidatas = sorted(pagina.images, key=lambda im: im.get("top", 0))
            for im in candidatas:
                ancho, alto = im.get("srcsize", (0, 0))
                if ancho < 200 or alto < 200:
                    continue
                try:
                    datos = im["stream"].get_data()
                    imagen = Image.open(io.BytesIO(datos)).convert("RGB")
                    imagenes.append(imagen)
                except Exception:
                    continue
    except Exception:
        return []
    return imagenes


def _ocr_imagenes_incrustadas(ruta, indice_pagina, lang="spa", config="--psm 3"):
    """OCR directo (sin preprocesar) sobre las imágenes incrustadas de la
    página -ver _extraer_imagenes_incrustadas-. Si hay varias (anverso y
    reverso en la misma página), se corre OCR por separado en cada una y
    se concatena el texto; la confianza reportada es el promedio
    ponderado por longitud de texto de cada imagen, para que se pueda
    comparar contra las demás candidatas del pase HD con el mismo
    criterio (mayor confianza gana). Regresa (texto, confianza); texto
    vacío y confianza -1.0 si no hay imágenes o ninguna se pudo leer."""
    imagenes = _extraer_imagenes_incrustadas(ruta, indice_pagina)
    if not imagenes:
        return "", -1.0
    textos = []
    for imagen in imagenes:
        try:
            texto, confianza = _ocr_con_confianza(imagen, lang=lang, config=config)
        except Exception:
            texto, confianza = "", -1.0
        if texto:
            textos.append((texto, confianza))
    if not textos:
        return "", -1.0
    texto_completo = "\n".join(t for t, _ in textos)
    peso_total = sum(len(t) for t, _ in textos)
    if peso_total <= 0:
        confianza_prom = max(c for _, c in textos)
    else:
        confianza_prom = sum(len(t) * c for t, c in textos) / peso_total
    return texto_completo, confianza_prom


# ---------------------------------------------------------------------------
# OCR reforzado: SOLO se usa cuando la primera lectura no deja legibles los
# datos clave (ver extraer_texto_pdf y procesar_documento). Nace de casos
# reales de escáneres de celular (Adobe Scan y similares): el PDF trae una
# imagen de página completa + una capa de texto "invisible" que el propio
# escáner generó, casi siempre de muy mala calidad (la credencial es chica
# dentro de una hoja blanca; el recibo viene con barras de color y QR). La
# app confiaba en esa capa por traer "suficiente texto" y nunca corría su
# OCR; y cuando sí corría, lo hacía sobre la página completa reducida, donde
# el texto de una credencial queda demasiado chico para Tesseract.
# ---------------------------------------------------------------------------

# Cada llamada a Tesseract en el pase reforzado puede ser más lenta que en el
# rápido (imágenes más grandes), así que tiene su propio límite; y hay un
# tope total por página para que un documento imposible no trabe el Excel.
OCR_TIMEOUT_REFORZADO_SEGUNDOS = 40
OCR_PRESUPUESTO_REFORZADO_SEGUNDOS = 120
# Por debajo de esta calificación (ver _calidad_texto) el texto nativo de un
# escaneo se considera basura y se vuelve a leer con el OCR reforzado.
CALIDAD_MINIMA_TEXTO_NATIVO = 0.25

_VOCALES = set("AEIOUÁÉÍÓÚÜaeiouáéíóúü")


def _calidad_texto(texto):
    """Calificación 0-1 de qué tan "texto de verdad" es una cadena: proporción
    de palabras reconocibles (letras, con vocales, sin rachas de consonantes
    raras) y de números/códigos, penalizando renglones de 1-2 caracteres y
    símbolos sueltos (~ • \\ < > ^ ...), que es lo que escupe un OCR malo.
    No sabe de idioma ni de diccionarios: solo distingue basura obvia de
    texto aprovechable (un documento digital normal sale >0.5; la capa de
    texto de un escáner de celular fallido sale ~0)."""
    toks = re.findall(r"\S+", texto or "")
    if not toks:
        return 0.0
    puntos = 0.0
    for t in toks:
        s = t.strip(".,:;()\"'“”-/¿?¡!*•")
        if (len(s) >= 3 and re.fullmatch(r"[^\W\d_]+", s) and any(c in _VOCALES for c in s)
                and not re.search(r"[^\WAEIOUÁÉÍÓÚÜaeiouáéíóúü\d_]{5,}", s)):
            puntos += 1.0
        elif len(s) >= 2 and re.fullmatch(r"[\dA-Za-z/.,:\-$%]+", s) and re.search(r"\d", s):
            puntos += 0.6
        elif re.fullmatch(r"[^\W\d_]{2}", s):
            puntos += 0.4  # conectores cortos: DE, LA, EL, EN, AL, UN
    base = puntos / len(toks)
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    cortas = sum(1 for l in lineas if len(l) <= 2) / max(1, len(lineas))
    simbolos = sum(texto.count(c) for c in "~•\\|<>^_{}[]¬°") / max(1, len(re.sub(r"\s", "", texto)))
    return max(0.0, base - 0.5 * cortas - 3 * simbolos)


def _pagina_parece_escaneo(pagina):
    """True si alguna imagen incrustada cubre la mayor parte de la página
    (foto/escaneo), o sea, si el texto que trae el PDF no es "nativo" sino
    el que le pegó el escáner encima."""
    area_pagina = float(pagina.width * pagina.height) or 1.0
    for im in pagina.images:
        try:
            if (im["x1"] - im["x0"]) * (im["bottom"] - im["top"]) >= 0.6 * area_pagina:
                return True
        except Exception:
            continue
    return False


def _tramos(ocupado, hueco_min):
    """Rangos (ini, fin) de posiciones consecutivas ocupadas, uniendo los que
    estén separados por menos de hueco_min posiciones vacías."""
    tramos, ini, ult = [], None, None
    for i, v in enumerate(ocupado):
        if not v:
            continue
        if ini is None:
            ini = i
        elif i - ult - 1 >= hueco_min:
            tramos.append((ini, ult))
            ini = i
        ult = i
    if ini is not None:
        tramos.append((ini, ult))
    return tramos


def _bloques_contenido(imagen, min_fraccion=0.03):
    """Separa una hoja escaneada con contenido disperso -p.ej. anverso y
    reverso de una credencial sobre una hoja casi blanca- en recortes, uno
    por bloque, para poder leer cada uno AMPLIADO. Si no se distingue más de
    un bloque (documento que ocupa toda la hoja, foto con fondo no blanco...)
    regresa la imagen completa. Solo usa PIL (sin numpy/OpenCV)."""
    w, h = imagen.size
    gris = ImageOps.grayscale(imagen)
    f = max(1, max(w, h) // 500)
    peq = gris.resize((max(1, w // f), max(1, h // f)), Image.BOX)
    pw, ph = peq.size
    hist = peq.histogram()
    total, acum, papel = pw * ph, 0, 255
    for v in range(255, -1, -1):
        acum += hist[v]
        if acum >= total * 0.20:
            papel = v
            break
    umbral = papel - 35
    mascara = peq.point(lambda x: 255 if x < umbral else 0)
    mascara = mascara.filter(ImageFilter.MinFilter(3)).filter(ImageFilter.MaxFilter(5))

    filas = list(mascara.resize((1, ph), Image.BOX).getdata())
    bandas = _tramos([v > 255 * 0.015 for v in filas], max(2, int(ph * 0.03)))
    bloques = []
    for y0, y1 in bandas:
        if (y1 - y0 + 1) < ph * 0.06:
            continue
        franja = mascara.crop((0, y0, pw, y1 + 1))
        cols = list(franja.resize((pw, 1), Image.BOX).getdata())
        ocupadas = [i for i, v in enumerate(cols) if v > 255 * 0.015]
        if not ocupadas:
            continue
        x0, x1 = ocupadas[0], ocupadas[-1]
        if (x1 - x0 + 1) * (y1 - y0 + 1) < pw * ph * min_fraccion:
            continue
        mx, my = int(pw * 0.015), int(ph * 0.015)
        caja = (max(0, x0 - mx) * f, max(0, y0 - my) * f, min(pw, x1 + 1 + mx) * f, min(ph, y1 + 1 + my) * f)
        bloques.append(imagen.crop((caja[0], caja[1], min(w, caja[2]), min(h, caja[3]))))
    if not bloques:
        return [imagen]
    if len(bloques) == 1 and bloques[0].width * bloques[0].height >= 0.8 * w * h:
        return [imagen]
    return bloques


def _escala_a_ancho(imagen, ancho):
    if imagen.width == ancho:
        return imagen
    return imagen.resize((ancho, max(1, round(imagen.height * ancho / imagen.width))), Image.LANCZOS)


def _binariza_local(gris, factor):
    """Binariza con el umbral calculado con el brillo de ESA imagen (no de la
    página completa): sobre fondos de color/guilloché separa mejor el texto."""
    gris = ImageOps.autocontrast(gris, cutoff=0)
    brillo = ImageStat.Stat(gris).mean[0]
    return gris.point(lambda x, u=brillo * factor: 255 if x > u else 0).filter(ImageFilter.SHARPEN)


def _pasos_ocr_reforzado(bloque, limitado=False):
    """Lista de (imagen_preparada, psm) a probar, de la más prometedora a la
    menos. Para un bloque con forma de credencial (relación ~1.6) se lee la
    columna de texto -sin la foto ni la marca de agua- a ~3x, que es donde
    mejor sale el domicilio y la vigencia, y la tarjeta completa a ~1800 px.
    Para un documento de hoja completa se lee por franjas a resolución casi
    nativa (reducir toda la hoja a la vez deja el texto chico y, además,
    Tesseract tarda tanto que se pasa del límite de tiempo)."""
    ancho, alto = bloque.size
    pasos = []
    if 1.3 <= ancho / max(1, alto) <= 1.85 and ancho <= 2600:
        col = _escala_a_ancho(bloque.crop((int(ancho * .33), int(alto * .16), int(ancho * .80), int(alto * .97))), 1200)
        completa = _escala_a_ancho(bloque, 1800)
        g_col, g_comp = ImageOps.grayscale(col), ImageOps.grayscale(completa)
        a_col = ImageOps.autocontrast(g_col, cutoff=0)
        a_comp = ImageOps.autocontrast(g_comp, cutoff=0)
        pasos = [(a_col, "6"), (_binariza_local(g_col, 0.7), "6"), (a_comp, "6"),
                 (_binariza_local(g_col, 0.8), "4"), (a_col, "3"), (a_comp, "3")]
        return pasos[:3] if limitado else pasos
    g = ImageOps.autocontrast(ImageOps.grayscale(_escala_a_ancho(bloque, min(2000, max(1600, ancho)))), cutoff=0)
    w, h = g.size
    franjas = [(0, h)] if h <= w * 0.8 else [(int(h * a), int(h * b)) for a, b in ((0, .36), (.30, .66), (.62, 1.0))]
    for psm in (("3",) if limitado else ("3", "6")):
        for y0, y1 in franjas:
            pasos.append((g.crop((0, y0, w, y1)), psm))
    return pasos


def _ocr_reforzado(ruta, indice_pagina, campos_ok=None):
    """Igual que _ocr_reforzado_candidatos pero con todo el texto unido en una
    sola cadena (lo que necesita el pase automático de extraer_texto_pdf)."""
    return "\n".join(_ocr_reforzado_candidatos(ruta, indice_pagina, campos_ok))


def _ocr_reforzado_candidatos(ruta, indice_pagina, campos_ok=None):
    """Lectura reforzada de UNA página: recorta/amplía la imagen incrustada y
    prueba varias variantes de OCR, deteniéndose en cuanto 'campos_ok'
    (callable texto -> (cuantos_ok, total)) dice que ya salieron todos los
    campos clave; si no hay campos_ok (no se sabe qué buscar) solo corre las
    primeras variantes. Regresa la LISTA de textos de las variantes que se
    corrieron (la de mejor calidad primero; [] si no se pudo leer nada): así
    quien la usa puede quedarse con la mejor lectura de CADA campo en vez de
    mezclarlas todas."""
    inicio = time.monotonic()
    imagenes = _extraer_imagenes_incrustadas(ruta, indice_pagina)
    if not imagenes:
        try:
            pag = convert_from_path(ruta, dpi=300, first_page=indice_pagina + 1, last_page=indice_pagina + 1)
            imagenes = [pag[0].convert("RGB")] if pag else []
        except Exception:
            imagenes = []
    candidatos = []

    def completo():
        if campos_ok is None:
            return False
        ok, total = campos_ok("\n".join(candidatos))
        return ok >= total

    for imagen in imagenes:
        for bloque in _bloques_contenido(imagen):
            for preparada, psm in _pasos_ocr_reforzado(bloque, limitado=campos_ok is None):
                if time.monotonic() - inicio > OCR_PRESUPUESTO_REFORZADO_SEGUNDOS:
                    break
                try:
                    texto, _ = _ocr_con_confianza(preparada, lang="spa", config=f"--psm {psm}",
                                                  timeout=OCR_TIMEOUT_REFORZADO_SEGUNDOS)
                except Exception:
                    texto = ""
                if texto and len(texto.strip()) >= 10:
                    candidatos.append(texto)
                if completo():
                    break
            if completo():
                break
        if completo():
            break
    candidatos.sort(key=_calidad_texto, reverse=True)
    return candidatos


def _tokens_largos(texto):
    return len(re.findall(r"[A-Za-zÁÉÍÓÚÑáéíóúñ:]{22,}", texto or ""))


def _extrae_texto_pagina(pagina):
    """Texto nativo de la página. Algunos PDF (la CSF del SAT, por ejemplo)
    traen los caracteres tan juntos que con la tolerancia de espacio normal
    de pdfplumber (3) las palabras salen pegadas ("NombredelaColonia:
    PENSILSUR"). Si el texto trae palabras absurdamente largas, se vuelve a
    extraer con una tolerancia menor (2) y se usa esa versión solo si deja
    menos palabras pegadas; así los PDF que ya salían bien no cambian."""
    texto = (pagina.extract_text() or "").strip()
    if _tokens_largos(texto) >= 1:
        try:
            alt = (pagina.extract_text(x_tolerance=2) or "").strip()
        except Exception:
            alt = ""
        if alt and _tokens_largos(alt) < _tokens_largos(texto):
            return alt
    return texto


def extraer_texto_pdf(ruta):
    """Regresa (lista_de_texto_por_pagina, num_paginas, uso_ocr_por_pagina).

    Estrategia de OCR en 2 pasos:
      1) Pase rápido (200 dpi + gris/autocontraste ligero) — resuelve la
         gran mayoría de documentos escaneados o fotografiados con luz
         decente, y sigue siendo barato en CPU (funciona incluso en un
         hosting con muy pocos recursos).
      2) Solo para las páginas que el pase rápido no leyó bien —ya sea
         porque salió muy poco texto, o porque Tesseract mismo reporta poca
         confianza en lo que leyó (típico de una foto con glare, sombra o
         fondo de color, aunque el documento sea legible a simple vista)—
         un pase HD (300 dpi) que además:
           - corrige automáticamente la rotación (fotos tomadas de lado),
           - prueba la imagen con y sin binarizar, y
           - prueba dos configuraciones de segmentación de Tesseract
             (--psm 3 y 6),
         quedándose con la combinación de mayor confianza reportada por el
         propio Tesseract (no la primera que salga ni la más larga). Cada
         intento individual tiene un límite de tiempo (OCR_TIMEOUT_SEGUNDOS)
         para que una imagen realmente mala no trabe la generación completa
         del Excel.
    """
    paginas_texto = []
    ocr_usado = []
    es_escaneo = []
    with pdfplumber.open(ruta) as pdf:
        num_paginas = len(pdf.pages)
        for pagina in pdf.pages:
            texto = _extrae_texto_pagina(pagina)
            paginas_texto.append(texto)
            ocr_usado.append(False)
            es_escaneo.append(_pagina_parece_escaneo(pagina))

    # Escáneres de celular (Adobe Scan y similares) pegan sobre la imagen una
    # capa de texto propia que, si la foto no salió perfecta, es casi pura
    # basura ("BECERR.\\.", "CUALHTEMOC, COMX"). Traer "suficientes
    # caracteres" no significa que sea legible: si la página es un escaneo y
    # ese texto califica muy mal, se lee de nuevo con el OCR reforzado y se
    # reemplaza SOLO si lo nuevo califica claramente mejor. Un PDF digital
    # normal (sin imagen de página completa) nunca entra aquí.
    for i, t in enumerate(paginas_texto):
        if len(t) >= 20 and es_escaneo[i] and _calidad_texto(t) < CALIDAD_MINIMA_TEXTO_NATIVO:
            alt = _ocr_reforzado(ruta, i)
            if alt and _calidad_texto(alt) > _calidad_texto(t) + 0.10:
                paginas_texto[i] = alt
                ocr_usado[i] = True

    # Umbral de confianza (0-100, escala propia de Tesseract) bajo el cual se
    # considera que un pase de OCR "no se puede confiar" y conviene escalar
    # al siguiente pase, aunque ya haya salido algo de texto.
    CONFIANZA_MINIMA = 60

    # IMPORTANTE (memoria): se rasteriza UNA página a la vez (first_page/
    # last_page) y directo en escala de grises (grayscale=True en
    # convert_from_path, en vez de rasterizar en color y convertir después)
    # — no el PDF completo ni todas las variantes juntas. En un hosting con
    # RAM limitada (Cloud Run con 1 GiB, Render free con 512 MiB) rasterizar
    # TODO el documento en color a 300 dpi puede fácilmente pasar de 1 GiB
    # entre las distintas copias en memoria (rápida + HD + rotada + ligera +
    # fuerte) y tronar el contenedor — que es justo lo que le pasó a esta app
    # en producción. Procesando de a una página, y liberando cada imagen con
    # "del" apenas se deja de necesitar, el uso pico de memoria se queda muy
    # por debajo de eso sin sacrificar la calidad del OCR.
    necesita_ocr = [i for i, t in enumerate(paginas_texto) if len(t) < 20]
    if necesita_ocr:
        reintentar = []
        for i in necesita_ocr:
            try:
                imagenes_pagina = convert_from_path(
                    ruta, dpi=200, first_page=i + 1, last_page=i + 1, grayscale=True
                )
            except Exception as e:
                imagenes_pagina = []
                print(f"  [aviso] no se pudo rasterizar la página {i+1} para OCR ({e})", file=sys.stderr)
            if not imagenes_pagina:
                continue
            imagen_prep = _preprocesa_para_ocr(imagenes_pagina[0], nivel="ligero")
            del imagenes_pagina
            texto_ocr, confianza = _ocr_con_confianza(imagen_prep, lang="spa")
            del imagen_prep
            if len(texto_ocr) > len(paginas_texto[i]):
                paginas_texto[i] = texto_ocr
                ocr_usado[i] = True
            # Se reintenta con el pase HD no solo si vino muy poco texto,
            # sino también si Tesseract mismo reporta poca confianza en lo
            # que leyó —eso es justo lo que pasa con documentos legibles a
            # simple vista pero que el pase rápido lee mal (poca luz, glare,
            # fondo de color)—.
            if len(texto_ocr) < 20 or confianza < CONFIANZA_MINIMA:
                reintentar.append(i)

        for i in reintentar:
            mejor_texto, mejor_confianza = paginas_texto[i], -1.0
            mejoro = False
            terminar = False

            # Candidata adicional (y la más barata: no rasteriza la página):
            # OCR directo sobre la(s) imagen(es) tal como vienen incrustadas
            # en el PDF (ver _extraer_imagenes_incrustadas). En documentos
            # escaneados de una sola imagen -típico en fotos de INE o
            # comprobante de domicilio- esto en la práctica gana casi
            # siempre contra los pases de poppler de abajo, porque no
            # arrastra el remuestreo/recompresión de volver a rasterizar la
            # página completa. Se prueba primero: si ya sale confiable, se
            # evita de plano rasterizar en HD.
            for psm in ("3", "6"):
                try:
                    candidato, confianza = _ocr_imagenes_incrustadas(ruta, i, config=f"--psm {psm}")
                except Exception as e:
                    candidato, confianza = "", -1.0
                    print(f"  [aviso] OCR (imagen incrustada, psm {psm}) falló en página {i+1} ({e})", file=sys.stderr)
                if candidato and len(candidato) >= 20 and confianza > mejor_confianza:
                    mejor_texto, mejor_confianza = candidato, confianza
                    mejoro = True
                if mejor_confianza >= 75 and len(mejor_texto) >= 40:
                    terminar = True
                    break

            # Se prueban dos variantes de preprocesamiento (con y sin
            # binarizar — binarizar ayuda mucho con fondos de color pero en
            # documentos de fondo claro a veces borra más de lo que ayuda),
            # UNA A LA VEZ (no las dos en memoria simultáneamente), cada una
            # con dos configuraciones de segmentación de Tesseract, y se
            # elige la de mayor confianza reportada por el propio Tesseract
            # (no la primera que salga ni la más larga). Solo se llega aquí
            # -y solo se rasteriza en HD- si la imagen incrustada no bastó.
            if not terminar:
                try:
                    imagenes_pagina_hd = convert_from_path(
                        ruta, dpi=300, first_page=i + 1, last_page=i + 1, grayscale=True
                    )
                except Exception as e:
                    imagenes_pagina_hd = []
                    print(f"  [aviso] no se pudo rasterizar en HD la página {i+1} para OCR ({e})", file=sys.stderr)
                imagen_hd = _corrige_rotacion(imagenes_pagina_hd[0]) if imagenes_pagina_hd else None
                del imagenes_pagina_hd

                if imagen_hd is not None:
                    for nivel in ("ligero", "fuerte"):
                        imagen_candidata = _preprocesa_para_ocr(imagen_hd, nivel=nivel)
                        for psm in ("3", "6"):
                            try:
                                candidato, confianza = _ocr_con_confianza(
                                    imagen_candidata, lang="spa", config=f"--psm {psm}"
                                )
                            except Exception as e:
                                candidato, confianza = "", -1.0
                                print(f"  [aviso] OCR (HD, psm {psm}) falló en página {i+1} ({e})", file=sys.stderr)
                            if candidato and len(candidato) >= 20 and confianza > mejor_confianza:
                                mejor_texto, mejor_confianza = candidato, confianza
                                mejoro = True
                            # Ya se ve confiable y con longitud razonable: no
                            # vale la pena seguir probando más configuraciones.
                            if mejor_confianza >= 75 and len(mejor_texto) >= 40:
                                terminar = True
                                break
                        del imagen_candidata
                        if terminar:
                            break
                    del imagen_hd

            if mejoro:
                paginas_texto[i] = mejor_texto
                ocr_usado[i] = True

    return paginas_texto, num_paginas, ocr_usado


def _ocr_hd_forzado(ruta, indice_pagina=0, todas_las_variantes=False):
    """Repite el pase HD de OCR (300 dpi, corrección de rotación, y las 4
    combinaciones de preprocesado/segmentación que usa extraer_texto_pdf)
    para UNA página específica, sin importar la confianza que haya
    reportado el pase rápido.

    Se usa como último recurso para campos puntuales que en la práctica
    resultan más sensibles a la calidad del OCR que el resto del
    documento — el caso encontrado fue el "Número de crédito" del aviso
    de Fonacot: viene en una tabla angosta (dos columnas muy juntas) que
    el pase rápido a veces lee mal justo ahí (la confianza PROMEDIO de la
    página sale alta porque el resto del documento sí se lee bien, así
    que extraer_texto_pdf nunca escala al pase HD).

    Por default regresa el texto de mayor confianza entre las 4
    combinaciones (o "" si no se pudo rasterizar la página). Si
    todas_las_variantes=True regresa, en cambio, la lista completa de los 4
    textos (puede traer "" para alguna combinación que haya fallado) — útil
    cuando la palabra clave que se busca resultó más estable en los DÍGITOS
    que en la ETIQUETA según la combinación (ver PATRON_NUMERO_CREDITO_FONACOT_ALT),
    así que conviene probar el patrón contra las 4 en vez de solo la de
    mayor confianza general."""
    try:
        imagenes_hd = convert_from_path(
            ruta, dpi=300, first_page=indice_pagina + 1, last_page=indice_pagina + 1, grayscale=True
        )
    except Exception as e:
        print(f"  [aviso] no se pudo rasterizar en HD (forzado) la página {indice_pagina+1} ({e})", file=sys.stderr)
        return [] if todas_las_variantes else ""
    if not imagenes_hd:
        return [] if todas_las_variantes else ""
    imagen_hd = _corrige_rotacion(imagenes_hd[0])
    del imagenes_hd

    variantes = []
    mejor_texto, mejor_confianza = "", -1.0
    for nivel in ("ligero", "fuerte"):
        imagen_candidata = _preprocesa_para_ocr(imagen_hd, nivel=nivel)
        for psm in ("3", "6"):
            try:
                candidato, confianza = _ocr_con_confianza(imagen_candidata, lang="spa", config=f"--psm {psm}")
            except Exception:
                candidato, confianza = "", -1.0
            variantes.append(candidato)
            if candidato and confianza > mejor_confianza:
                mejor_texto, mejor_confianza = candidato, confianza
        del imagen_candidata
    del imagen_hd
    if todas_las_variantes:
        return variantes
    return mejor_texto


# ---------------------------------------------------------------------------
# Clasificación de documento
# ---------------------------------------------------------------------------

# clave -> (nombre legible, lista de palabras/frases clave, es obligatorio)
CATEGORIAS = {
    "CV": ("CV", ["EXPERIENCIA LABORAL", "REFERENCIAS LABORALES", "PERFIL PROFESIONAL", "CURRICULUM", "HABILIDADES", "FORMACION ACADEMICA", "ACERCA DE MI"], True),
    "ACTA_NACIMIENTO": ("Acta de nacimiento", ["ACTA DE NACIMIENTO", "REGISTRO CIVIL", "OFICIALIA", "NACIMIENTOS"], True),
    "INE": ("INE", ["INSTITUTO NACIONAL ELECTORAL", "CREDENCIAL PARA VOTAR", "CLAVE DE ELECTOR"], True),
    "COMPROBANTE_DOMICILIO": ("Comprobante de domicilio", ["COMISION FEDERAL DE ELECTRICIDAD", "CFE", "TELMEX", "RECIBO", "TOTAL A PAGAR", "PERIODO FACTURADO", "IZZI", "TELEFONOS DE MEXICO", "AGUA"], True),
    "COMPROBANTE_ESTUDIOS": ("Certificado de estudios", ["CERTIFICADO DE ESTUDIOS", "CURSO Y ACREDITO", "UNIVERSIDAD", "LICENCIATURA", "CEDULA PROFESIONAL", "SECRETARIA DE EDUCACION", "BACHILLERATO", "PROMEDIO GENERAL"], True),
    "CURP": ("CURP", ["CLAVE UNICA DE REGISTRO DE POBLACION", "CURP CERTIFICADA", "CLAVE:"], True),
    "CSF": ("CSF", ["CONSTANCIA DE SITUACION FISCAL", "CEDULA DE IDENTIFICACION FISCAL", "REGISTRO FEDERAL DE CONTRIBUYENTES", "ACUSE UNICO DE INSCRIPCION", "INSCRIPCION AL REGISTRO FEDERAL", "IDCIF"], True),
    "NSS": ("NSS", ["NUMERO DE SEGURIDAD SOCIAL", "INSTITUTO MEXICANO DEL SEGURO SOCIAL", "IMSS"], True),
    "CUENTA_BANCARIA": ("Cuenta bancaria", ["CLABE", "ESTADO DE CUENTA", "NO. DE CUENTA", "CARATULA"], True),
    "INFONAVIT": ("Aviso de retención Infonavit", ["INFONAVIT", "INSTITUTO DEL FONDO NACIONAL DE LA VIVIENDA", "AVISO PARA RETENCION DE DESCUENTOS"], False),
    "FONACOT": ("Aviso de retención Fonacot", ["FONACOT", "INSTITUTO FONACOT"], False),
    # condicionales: aplican según el puesto (entrenador, barbero, estilista) o la situación del candidato
    "CERTIFICADO_MEDICO": ("Certificado médico", ["CERTIFICADO MEDICO", "RECONOCIMIENTO MEDICO", "MEDICO CIRUJANO"], False),
    "CERTIFICADO_INSTRUCTOR": ("Certificado de entrenador / barbero / estilista", ["CERTIFICADO", "FITNESS COACH", "ENTRENADOR", "ESTILISTA", "BARBERO", "BARBER", "COSMETOLOGIA", "COSMETOLOGO", "DIPLOMADO"], False),
    "CONSTANCIA_LABORAL": ("Constancia(s) laboral(es) / cartas de referencia", ["CONSTANCIA LABORAL", "CARTA LABORAL", "HACE CONSTAR QUE", "CARTA DE RECOMENDACION", "RECURSOS HUMANOS"], False),
}

# orden e integrantes de la checklist obligatoria que pidió el negocio
CHECKLIST_OBLIGATORIO = [
    "CV", "ACTA_NACIMIENTO", "INE", "COMPROBANTE_DOMICILIO",
    "COMPROBANTE_ESTUDIOS", "CURP", "CSF", "NSS", "CUENTA_BANCARIA",
    "INFONAVIT", "FONACOT", "CERTIFICADO_MEDICO", "CERTIFICADO_INSTRUCTOR",
    "CONSTANCIA_LABORAL",
]

# de estos, cuáles son condicionales (no siempre aplican) en vez de siempre-obligatorios
CONDICIONALES = {"INFONAVIT", "FONACOT", "CERTIFICADO_MEDICO", "CERTIFICADO_INSTRUCTOR", "CONSTANCIA_LABORAL"}


def _palabra_clave_presente(palabra, texto_norm):
    """Busca una palabra clave en el texto normalizado. Las claves cortas
    (CFE, AGUA, IMSS...) exigen límites de palabra para no activarse dentro
    de otras palabras (p. ej. AGUA dentro de "AGUASCALIENTES"); las frases
    largas se buscan como subcadena, tolerando variantes pegadas."""
    clave = normaliza(palabra)
    if len(clave) <= 5 and clave.replace(" ", "").isalnum():
        return re.search(r"(?<![A-Z0-9])" + re.escape(clave) + r"(?![A-Z0-9])", texto_norm) is not None
    return clave in texto_norm


def sugerir_categoria_por_contenido(texto_completo, categoria_cargada):
    """Compara la categoría en la que el documento se CARGÓ (campo del
    formulario o nombre del archivo) contra lo que dice su contenido. Regresa
    (clave_sugerida, nombre_legible) solo cuando hay evidencia clara de que es
    otro documento; si no, None. Se ignora el nombre del archivo a propósito
    (es justo lo que puede estar equivocado) y se exige evidencia fuerte para
    no dar falsas alarmas con escaneos de texto pobre: que el contenido no
    respalde en absoluto la categoría cargada y que lo sugerido coincida con
    2+ palabras clave o con una frase larga (p. ej. "CERTIFICADO MEDICO"), o
    que lo sugerido supere por 2+ coincidencias a lo cargado."""
    t = normaliza(texto_completo)
    puntajes = {}
    for clave, (_, palabras, _) in CATEGORIAS.items():
        aciertos = [normaliza(pal) for pal in palabras if _palabra_clave_presente(pal, t)]
        puntajes[clave] = (len(aciertos), sum(len(a) for a in aciertos))
    mejor = max(puntajes, key=lambda c: puntajes[c])
    n_mejor, chars_mejor = puntajes[mejor]
    n_cargada = puntajes.get(categoria_cargada, (0, 0))[0]
    if mejor == categoria_cargada or n_mejor == 0:
        return None
    if puntajes.get(categoria_cargada, (0, 0)) >= puntajes[mejor]:
        return None
    fuerte = n_mejor >= 2 or chars_mejor >= 12
    if fuerte and (n_cargada == 0 or n_mejor >= n_cargada + 2):
        return mejor, CATEGORIAS[mejor][0]
    return None


def clasificar(texto_completo, nombre_archivo):
    t = normaliza(texto_completo)
    nombre_arch_norm = normaliza(nombre_archivo)
    mejor_clave, mejor_score = "DESCONOCIDO", 0
    for clave, (_, palabras, _) in CATEGORIAS.items():
        score = sum(1 for p in palabras if _palabra_clave_presente(p, t))
        # pequeño empujón si el nombre del archivo ya lo sugiere
        pistas_nombre = {
            "CV": ["CV", "CURRICULUM"], "ACTA_NACIMIENTO": ["ACTA"], "INE": ["INE", "IFE"],
            "COMPROBANTE_DOMICILIO": ["DOMICILIO", "RECIBO", "CFE", "LOCALIZACION"],
            "COMPROBANTE_ESTUDIOS": ["GRADO", "ESTUDIOS", "TITULO", "CEDULA"],
            "CURP": ["CURP"], "CSF": ["CSF", "CONSTANCIA", "FISCAL"], "NSS": ["NSS", "SEGURIDAD SOCIAL", "LOCALIZACION"],
            "CUENTA_BANCARIA": ["CUENTA", "BANCO", "ESTADO DE CUENTA", "CARATULA"],
            "INFONAVIT": ["INFONAVIT"],
            "FONACOT": ["FONACOT"],
            "CERTIFICADO_INSTRUCTOR": ["CERTIFICADO", "BARBER", "ESTILISTA", "COACH"],
            "CERTIFICADO_MEDICO": ["MEDICO"],
            "CONSTANCIA_LABORAL": ["LABORAL", "CONSTANCIA"],
        }
        for pista in pistas_nombre.get(clave, []):
            if pista in nombre_arch_norm:
                score += 1
        if score > mejor_score:
            mejor_clave, mejor_score = clave, score
    return mejor_clave if mejor_score > 0 else "DESCONOCIDO"


# ---------------------------------------------------------------------------
# Carga masiva por ZIP: convención de nombres de archivo
# ---------------------------------------------------------------------------
# Para la carga masiva (un ZIP por candidato) el sistema identifica cada
# documento por el NOMBRE del archivo dentro del ZIP, en vez de tener que
# adivinarlo por el contenido — así el resultado es predecible y no depende
# de qué tan bien se leyó el OCR. Cada clave de CATEGORIAS tiene una lista de
# "alias" (palabras que se buscan como token completo, o como subcadena si el
# alias ya trae guion bajo) dentro del nombre del archivo ya normalizado
# (minúsculas, sin acentos, cualquier separador -espacio, guion, punto-
# convertido a "_"). Ejemplos de nombres que SÍ matchean:
#   cv.pdf, CV_Juan.pdf, 01-cv.pdf, curriculum.pdf, comprobante_domicilio.pdf,
#   Comprobante-Domicilio (2).pdf, domicilio.pdf, ine.pdf, INE_frente_reverso.pdf
# Si el nombre del archivo no matchea ningún alias, se cae al reconocimiento
# por CONTENIDO (la misma lógica que ya usa la carga de un candidato a la
# vez) y, si tampoco así se identifica, el documento se reporta como
# "Desconocido / no identificado" (equivalente a la carpeta "otros").
ALIAS_ARCHIVO = {
    "CV": ["cv", "curriculum", "resume"],
    "ACTA_NACIMIENTO": ["acta_nacimiento", "acta"],
    "INE": ["ine", "ife", "credencial"],
    "COMPROBANTE_DOMICILIO": ["comprobante_domicilio", "domicilio", "recibo_luz", "recibo_agua", "recibo_cfe", "cfe", "comprobante"],
    "COMPROBANTE_ESTUDIOS": ["comprobante_estudios", "estudios", "titulo", "cedula_profesional", "certificado_estudios"],
    "CURP": ["curp"],
    "CSF": ["csf", "constancia_situacion_fiscal", "situacion_fiscal", "fiscal"],
    "NSS": ["nss", "seguridad_social", "imss"],
    "CUENTA_BANCARIA": ["cuenta_bancaria", "cuenta", "caratula_bancaria", "caratula", "clabe", "estado_cuenta"],
    "INFONAVIT": ["infonavit"],
    "FONACOT": ["fonacot"],
    "CERTIFICADO_MEDICO": ["certificado_medico", "medico"],
    "CERTIFICADO_INSTRUCTOR": ["certificado_instructor", "certificado_entrenador", "entrenador", "barbero", "estilista", "coach"],
    "CONSTANCIA_LABORAL": ["constancia_laboral", "carta_laboral", "laboral", "referencia_laboral"],
}

# orden en el que se revisan los alias: los más largos/específicos primero,
# para que "comprobante_domicilio" gane sobre el genérico "comprobante" si
# ambos aparecen en el nombre del archivo.
_ALIAS_ORDENADOS = sorted(
    ((clave, alias) for clave, alias_list in ALIAS_ARCHIVO.items() for alias in alias_list),
    key=lambda par: -len(par[1]),
)


def _normaliza_nombre_archivo(nombre_archivo):
    """minúsculas, sin acentos, sin extensión, separadores -> '_' """
    base = os.path.splitext(os.path.basename(nombre_archivo))[0]
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode("ascii")
    base = base.lower()
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    return base


def detectar_categoria_por_nombre_archivo(nombre_archivo):
    """Regresa la clave de CATEGORIAS sugerida por el NOMBRE del archivo
    (convención de la carga masiva por ZIP), o None si el nombre no da
    ninguna pista clara (en ese caso el caller debe caer al reconocimiento
    por contenido, ver clasificar())."""
    base = "_" + _normaliza_nombre_archivo(nombre_archivo) + "_"
    for clave, alias in _ALIAS_ORDENADOS:
        if ("_" + alias + "_") in base:
            return clave
    return None


# ---------------------------------------------------------------------------
# Nombre: extracción y comparación
# ---------------------------------------------------------------------------

def tokens_nombre(nombre):
    return [t for t in normaliza(nombre).split(" ") if len(t) > 1]


def nombre_coincide(texto_doc, nombre_candidato):
    """Compara por tokens: cuenta cuántas palabras del nombre del candidato
    aparecen en el texto del documento. Tolerante a orden y a acentos."""
    tks = tokens_nombre(nombre_candidato)
    if not tks:
        return None, 0
    texto_norm = normaliza(texto_doc)
    encontrados = sum(1 for tk in tks if re.search(r"\b" + re.escape(tk) + r"\b", texto_norm))
    proporcion = encontrados / len(tks)
    if proporcion >= 0.75:
        return True, proporcion
    elif proporcion >= 0.4:
        return None, proporcion  # dudoso -> revisar a mano
    else:
        return False, proporcion


# ---------------------------------------------------------------------------
# Extracción de identificadores y datos personales (RFC, CURP, fecha de
# nacimiento, dirección) — se agregan al resumen igual que los datos
# bancarios de la carátula: son "propuesta a confirmar" (ver limitaciones
# del OCR en el docstring del módulo), pensadas para que el equipo de
# reclutamiento no tenga que volver a abrir cada PDF para copiar el dato,
# pero sin reemplazar una revisión rápida contra el documento original.
# ---------------------------------------------------------------------------

# RFC de persona física: 4 letras (o 3 si es persona moral) + 6 dígitos
# (fecha AAMMDD) + 3 caracteres alfanuméricos (homoclave). Ejemplo:
# XAXX010101000 (el RFC genérico que usa el SAT para "público en general").
PATRON_RFC = re.compile(r"\bRFC\b[:\.\-]?\s*([A-Z&]{3,4}\d{6}[A-Z0-9]{3})\b")

# CURP: 18 caracteres siempre en este orden — 4 letras, 6 dígitos (fecha de
# nacimiento AAMMDD), 1 letra H/M (sexo), 2 letras (código de entidad), 3
# consonantes, 1 carácter alfanumérico (homoclave) y 1 dígito verificador.
# Ejemplo: XEXX010101HNEXXXA4.
PATRON_CURP = re.compile(r"\b([A-Z]{4}\d{6}[HM][A-Z]{5}[A-Z0-9]\d)\b")


def extraer_rfc(texto_completo):
    """Extrae el RFC del texto de la CSF, anclado a la etiqueta "RFC" (el
    campo que la propia Constancia imprime con ese nombre). Regresa el RFC
    detectado o None si no se encontró nada con el formato esperado."""
    m = PATRON_RFC.search(normaliza(texto_completo))
    return m.group(1) if m else None


def extraer_curp(texto_completo):
    """Extrae el CURP del texto del documento CURP (o de la CSF, que también
    lo trae impreso), buscando la cadena de 18 caracteres con el formato
    oficial. Regresa el CURP detectado o None si no se encontró nada con ese
    formato."""
    t = normaliza(texto_completo)
    m = PATRON_CURP.search(t)
    if m:
        return m.group(1)
    return _curp_reparado_por_ocr(t)


_CURP_A_DIGITO = {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "B": "8", "S": "5", "Z": "2"}
_CURP_A_LETRA = {"0": "O", "1": "I", "8": "B", "5": "S", "2": "Z", "6": "G"}


def _curp_valido_estricto(c):
    if not re.fullmatch(r"[A-Z]{4}\d{6}[HM][A-Z]{2}[B-DF-HJ-NP-TV-Z]{3}[A-Z0-9]\d", c):
        return False
    mes, dia = int(c[6:8]), int(c[8:10])
    return 1 <= mes <= 12 and 1 <= dia <= 31 and c[11:13] in CURP_ESTADOS


def _curp_por_posicion(s):
    """Corrige confusiones típicas del OCR (O/0, I/1, B/8...) según lo que
    cada posición del CURP puede ser (letra o dígito)."""
    out = list(s)
    for i, ch in enumerate(out):
        if 4 <= i <= 9 or i == 17:
            out[i] = _CURP_A_DIGITO.get(ch, ch)
        elif i in (0, 1, 2, 3, 10, 11, 12, 13, 14, 15):
            out[i] = _CURP_A_LETRA.get(ch, ch)
    return "".join(out)


def _curp_reparado_por_ocr(t_norm):
    """Respaldo cuando el OCR leyó mal un carácter del CURP (p.ej.
    "LOCR0O80416MMCPDNAO", con un 0 de más y una O por 0 al final). Busca
    cadenas de 18 o 19 caracteres, corrige por posición y, si son 19, prueba
    quitar cada carácter; solo acepta un resultado que cumple TODA la
    estructura (fecha real, sexo, clave de entidad válida, consonantes)."""
    for m in re.finditer(r"(?<![A-Z0-9])[A-Z0-9]{18,19}(?![A-Z0-9])", t_norm):
        token = m.group(0)
        variantes = [token] if len(token) == 18 else [token[:i] + token[i + 1:] for i in range(len(token))]
        mejor = None
        for v in variantes:
            c = _curp_por_posicion(v)
            if _curp_valido_estricto(c):
                costo = sum(1 for a, b in zip(v, c) if a != b)
                if mejor is None or costo < mejor[0]:
                    mejor = (costo, c)
        if mejor:
            return mejor[1]
    return None


# Número de Seguridad Social (NSS): 11 dígitos, impreso en la constancia de
# "Asignación de Número de Seguridad Social" del IMSS, junto a la etiqueta
# "Número de Seguridad Social".
PATRON_NSS = re.compile(r"NUMERO\s*DE\s*SEGURIDAD\s*SOCIAL\s*:?\s*(\d{11})\b")


def extraer_nss(texto_completo):
    """Extrae el Número de Seguridad Social del documento de "Asignación de
    NSS" del IMSS, anclado a la etiqueta "Número de Seguridad Social".
    Regresa el NSS detectado (11 dígitos) o None si no se encontró."""
    t_norm = normaliza(texto_completo)
    m = PATRON_NSS.search(t_norm)
    if m:
        return m.group(1)
    # Formatos tipo "NSS: 0421069193 -9" (aviso/solicitud del IMSS que imprime
    # el dígito verificador separado por un guion).
    m = re.search(r"\bNSS\s*:?\s*(\d{10})\s*[-–]?\s*(\d)\b", t_norm)
    return m.group(1) + m.group(2) if m else None


# Códigos de entidad federativa que usa el CURP (posiciones 12-13), definidos
# por RENAPO -no son los mismos que las abreviaturas postales comunes-. "NE"
# es el código reservado para quien nació en el extranjero.
CURP_ESTADOS = {
    "AS": "Aguascalientes", "BC": "Baja California", "BS": "Baja California Sur",
    "CC": "Campeche", "CL": "Coahuila", "CS": "Chiapas", "CH": "Chihuahua",
    "DF": "Ciudad de México", "CM": "Ciudad de México", "DG": "Durango",
    "GT": "Guanajuato", "GR": "Guerrero", "HG": "Hidalgo", "JC": "Jalisco",
    "MC": "México", "MN": "Michoacán", "MS": "Morelos", "NT": "Nayarit",
    "NL": "Nuevo León", "OC": "Oaxaca", "PL": "Puebla", "QO": "Querétaro",
    "QR": "Quintana Roo", "SP": "San Luis Potosí", "SL": "Sinaloa",
    "SR": "Sonora", "TC": "Tabasco", "TS": "Tamaulipas", "TL": "Tlaxcala",
    "VZ": "Veracruz", "YN": "Yucatán", "ZS": "Zacatecas", "NE": "Extranjero",
}


def datos_derivados_de_curp(curp):
    """A partir de un CURP con formato válido (18 caracteres), deriva
    género, estado de nacimiento y si corresponde a alguien nacido en
    México -sin depender de que el documento CURP se haya leído bien en
    esos renglones-, usando las posiciones fijas que define RENAPO (11:
    sexo H/M: 12-13: entidad). Se usa solo como PROPUESTA para precargar la
    hoja "Datos completos" del Excel; no reemplaza los datos que el propio
    candidato capture. Regresa {} si el CURP no trae un formato reconocible."""
    if not curp or len(curp) != 18:
        return {}
    sexo = curp[10]
    entidad = curp[11:13]
    genero = "Hombre" if sexo == "H" else "Mujer" if sexo == "M" else None
    estado_nacimiento = CURP_ESTADOS.get(entidad)
    nacionalidad = None
    if estado_nacimiento:
        nacionalidad = "Extranjera" if entidad == "NE" else "Mexicana"

    # Fecha de nacimiento: posiciones 5-10 del CURP (después de las 4 letras
    # iniciales) traen AAMMDD. El propio CURP no dice el siglo, pero la
    # convención de RENAPO para el diferenciador (posición 18, el
    # penúltimo... en realidad posición 17, un carácter antes del dígito
    # verificador) es que sea una LETRA para quien nació en o después del
    # 2000, y un DÍGITO para quien nació antes -así se resuelve el siglo sin
    # ambigüedad-.
    aa, mm, dd = curp[4:6], curp[6:8], curp[8:10]
    diferenciador = curp[16]
    siglo = 2000 if diferenciador.isalpha() else 1900
    fecha_nacimiento = None
    try:
        fecha_nacimiento = datetime.date(siglo + int(aa), int(mm), int(dd))
    except ValueError:
        fecha_nacimiento = None

    edad = None
    if fecha_nacimiento:
        edad = HOY.year - fecha_nacimiento.year - ((HOY.month, HOY.day) < (fecha_nacimiento.month, fecha_nacimiento.day))

    return {
        "genero": genero, "estado_nacimiento": estado_nacimiento, "nacionalidad": nacionalidad,
        "fecha_nacimiento": fecha_nacimiento, "edad": edad,
    }


# Código Postal de la CSF: bajo "Datos del domicilio registrado", etiqueta
# "Código Postal:" seguida directamente de 5 dígitos (a veces sin espacio,
# p.ej. "Código Postal:97203").
PATRON_CP_CSF = re.compile(r"CODIGO\s*POSTAL\s*:?\s*(\d{5})\b")


def extraer_codigo_postal_csf(texto_completo):
    """Extrae el Código Postal impreso en la CSF, anclado a la etiqueta
    "Código Postal" de la sección "Datos del domicilio registrado". Regresa
    el CP detectado o None si no se encontró."""
    m = PATRON_CP_CSF.search(normaliza(texto_completo))
    return m.group(1) if m else None


def _campo_csf_por_regex(t_norm, patron_etiqueta, patrones_fin, ventana=250):
    """Como una versión de _campo_entre_etiquetas que tolera que la
    extracción del PDF haya pegado las palabras de la etiqueta sin espacio
    entre ellas -el mismo fenómeno ya documentado en otros campos de la CSF,
    ver extraer_nombres_csf_bien_espaciado más abajo, y que en esta sección
    de "Datos del domicilio registrado" resultó afectar a la mayoría de las
    etiquetas-: 'patron_etiqueta' y cada elemento de 'patrones_fin' son
    expresiones regulares (normalmente con \\s* entre palabras, que también
    matchea CERO espacios) en vez de texto literal.

    Regresa el texto que sigue a la etiqueta, cortado en la primera
    coincidencia de cualquiera de 'patrones_fin' (se prueban todas y se usa
    la que aparezca más cerca); si ninguna aparece, se recorta a 120
    caracteres como respaldo. None si la etiqueta de inicio no se
    encontró."""
    m = re.search(patron_etiqueta, t_norm)
    if not m:
        return None
    resto = t_norm[m.end():m.end() + ventana]
    fin = 120
    for patron_fin in patrones_fin:
        m_fin = re.search(patron_fin, resto)
        if m_fin and m_fin.start() < fin:
            fin = m_fin.start()
    return resto[:fin].strip(" :.-\t")


_CSF_ETQ_TIPO_VIALIDAD = r"TIPO\s*DE\s*VIALIDAD"
_CSF_ETQ_NOMBRE_VIALIDAD = r"NOMBRE\s*DE\s*VIALIDAD"
_CSF_ETQ_NUMERO_EXTERIOR = r"NUMERO\s*EXTERIOR"
_CSF_ETQ_NUMERO_INTERIOR = r"NUMERO\s*INTERIOR"
_CSF_ETQ_COLONIA = r"NOMBRE\s*DE\s*LA\s*COLONIA"
_CSF_ETQ_LOCALIDAD = r"NOMBRE\s*DE\s*LA\s*LOCALIDAD"
_CSF_ETQ_MUNICIPIO = r"NOMBRE\s*DEL\s*MUNICIPIO\s*O\s*DEMARCACION\s*TERRITORIAL"
# la palabra FEDERATIVA a veces sale deformada del OCR ("ENTIDAD EEN:")
_CSF_ETQ_ENTIDAD = r"NOMBRE\s*DE\s*LA\s*ENTIDAD(?:\s*FEDERATIVA|\s*[A-Z]{1,12}(?=\s*:))?"
_CSF_ETQ_ENTRE_CALLE = r"ENTRE\s*CALLE"
_CSF_ETQ_Y_CALLE = r"Y\s*CALLE"
_CSF_ETQ_FIN_SECCION = r"ACTIVIDADES\s*ECONOMICAS|REGIMENES"


def extraer_domicilio_csf(texto_completo):
    """Extrae los campos de la sección "Datos del domicilio registrado" de
    la CSF (Tipo/Nombre de Vialidad, Número Exterior/Interior, Colonia,
    Localidad, Municipio, Entidad Federativa) y arma con ellos, junto con
    el Código Postal (ver extraer_codigo_postal_csf), una "Dirección
    fiscal" legible en una sola línea para la hoja "Datos completos".

    Cuando "Nombre de la Colonia" trae un nombre largo (de fraccionamiento)
    que ocupa 2 renglones en el PDF, la extracción de texto a veces
    intercala ahí en medio la etiqueta "Número Interior" -con su valor,
    casi siempre vacío- porque en el PDF esa etiqueta está en la columna de
    al lado, a la altura del primer renglón de la Colonia. Por eso la
    Colonia se captura sin cortar en esa etiqueta (para no perder el
    segundo renglón) y luego se le quita, si quedó pegada adentro; y el
    "Número Interior" en sí solo se acepta si el texto capturado de veras
    parece un número interior (empieza con dígito, o es un texto corto tipo
    "PB"/"A") -si no, se descarta como el residuo de la Colonia que
    probablemente es, y se deja en blanco (a revisar a mano) en vez de
    asignarle por error un pedazo del nombre de la Colonia.

    Regresa un dict con cada campo por separado (código_postal, tipo/
    nombre de vialidad, número exterior/interior, colonia, localidad,
    municipio, entidad_federativa) y la llave 'direccion_fiscal' con el
    texto ya armado; si no se encontró la sección, todo queda en None."""
    t = normaliza(texto_completo)

    tipo_vialidad = _campo_csf_por_regex(t, _CSF_ETQ_TIPO_VIALIDAD, [_CSF_ETQ_NOMBRE_VIALIDAD])
    nombre_vialidad = _campo_csf_por_regex(t, _CSF_ETQ_NOMBRE_VIALIDAD, [_CSF_ETQ_NUMERO_EXTERIOR])
    numero_exterior = _campo_csf_por_regex(
        t, _CSF_ETQ_NUMERO_EXTERIOR, [_CSF_ETQ_NUMERO_INTERIOR, _CSF_ETQ_COLONIA]
    )

    colonia_bruta = _campo_csf_por_regex(t, _CSF_ETQ_COLONIA, [_CSF_ETQ_LOCALIDAD])
    colonia = None
    if colonia_bruta is not None:
        colonia = re.sub(_CSF_ETQ_NUMERO_INTERIOR + r"\s*:?\s*", " ", colonia_bruta)
        colonia = re.sub(r"\s+", " ", colonia).strip(" ,") or None

    numero_interior_bruto = _campo_csf_por_regex(
        t, _CSF_ETQ_NUMERO_INTERIOR, [_CSF_ETQ_COLONIA, _CSF_ETQ_LOCALIDAD]
    )
    numero_interior = None
    if numero_interior_bruto:
        # la CSF suele imprimir el valor con prefijo ("INT 14", "INT14")
        numero_interior_bruto = re.sub(r"^(?:INTERIOR|INT)\.?\s*", "", numero_interior_bruto).strip(" .:-") or ""
    if numero_interior_bruto and re.match(r"^(?:[0-9][0-9A-Z\-\.]{0,9}|[A-Z]{1,2})$", numero_interior_bruto):
        numero_interior = numero_interior_bruto

    localidad = _campo_csf_por_regex(t, _CSF_ETQ_LOCALIDAD, [_CSF_ETQ_MUNICIPIO])
    municipio = _campo_csf_por_regex(t, _CSF_ETQ_MUNICIPIO, [_CSF_ETQ_ENTIDAD])
    entidad = _campo_csf_por_regex(
        t, _CSF_ETQ_ENTIDAD, [_CSF_ETQ_ENTRE_CALLE, _CSF_ETQ_Y_CALLE, _CSF_ETQ_FIN_SECCION]
    )
    codigo_postal = extraer_codigo_postal_csf(texto_completo)

    resultado = {
        "codigo_postal": codigo_postal,
        "tipo_vialidad": tipo_vialidad or None,
        "nombre_vialidad": nombre_vialidad or None,
        "numero_exterior": numero_exterior or None,
        "numero_interior": numero_interior or None,
        "colonia": colonia or None,
        "localidad": localidad or None,
        "municipio": municipio or None,
        "entidad_federativa": entidad or None,
        "direccion_fiscal": None,
    }

    if not any([tipo_vialidad, nombre_vialidad, numero_exterior, colonia, municipio, entidad, codigo_postal]):
        return resultado

    calle = " ".join(p for p in (tipo_vialidad, nombre_vialidad) if p)
    numero = numero_exterior or ""
    if numero_interior:
        numero = f"{numero} Int. {numero_interior}".strip()
    fragmento_calle = " ".join(p for p in (calle, numero) if p)

    piezas = [
        fragmento_calle or None,
        f"Col. {colonia}" if colonia else None,
        localidad if localidad and localidad != municipio else None,
        municipio,
        entidad,
        f"C.P. {codigo_postal}" if codigo_postal else None,
    ]
    resultado["direccion_fiscal"] = ", ".join(p for p in piezas if p).strip(", ") or None
    return resultado


# Régimen de la CSF: en la hoja 2, bajo la sección "Regímenes", una tabla con
# encabezados "Régimen | Fecha Inicio | Fecha Fin" y debajo el nombre del
# régimen (texto libre, p.ej. "Régimen de Sueldos y Salarios e Ingresos
# Asimilados a Salarios") seguido de su fecha de inicio. Se ancla a los
# encabezados de la tabla y se toma el texto hasta la primera fecha
# (DD/MM/AAAA) como el nombre del régimen, en vez de una lista fija de
# nombres posibles (el SAT tiene varios regímenes y el contribuyente puede
# estar en cualquiera).
PATRON_REGIMEN_CSF = re.compile(
    r"REGIMEN(?:ES)?\s*:?\s*REGIMEN\s*FECHA\s*INICIO\s*FECHA\s*FIN\s*(.+?)\s*\d{2}/\d{2}/\d{4}"
)


def extraer_regimen_csf(texto_completo):
    """Extrae el nombre del Régimen impreso en la hoja 2 de la CSF, bajo la
    sección "Regímenes". Regresa el texto del régimen detectado o None si no
    se encontró la tabla esperada."""
    m = PATRON_REGIMEN_CSF.search(normaliza(texto_completo))
    return m.group(1).strip() if m else None


# Nombre(s), Primer y Segundo Apellido de la CSF: bajo "Datos de
# Identificación del Contribuyente", en ese orden y cada uno con su propia
# etiqueta ("Nombre (s):", "Primer Apellido:", "Segundo Apellido:"). Se
# ancla cada campo a la etiqueta que le sigue, en vez de a un largo fijo,
# porque el nombre y los apellidos pueden traer más de una palabra.
_SEP_CSF = r"[\s:|!\[\]\.\-_]*"
PATRON_NOMBRES_CSF = re.compile(r"NOMBRE\s*\(S\)" + _SEP_CSF + r"(.{1,80}?)\s*PRIMER\s*APELLIDO")
PATRON_APELLIDO_PATERNO_CSF = re.compile(r"PRIMER\s*APELLIDO" + _SEP_CSF + r"(.{1,60}?)\s*SEGUNDO\s*APELLIDO")
PATRON_APELLIDO_MATERNO_CSF = re.compile(
    r"SEGUNDO\s*APELLIDO" + _SEP_CSF + r"(.{0,60}?)\s*(?:FECHA\s*INICIO\s*DE\s*OPERACIONES|NOMBRE\s*COMERCIAL|ESTATUS|FECHA\s*DE\s*ULTIMO)"
)


def _limpia_nombre_csf(valor):
    """Deja solo letras y espacios de un nombre/apellido capturado: el texto
    de la CSF o del acuse del RFC puede traer ahí separadores de tabla leídos
    como '|', '1' o ':' (los nombres no llevan dígitos ni símbolos). Si lo que
    queda es demasiado largo para ser un nombre, regresa None."""
    if not valor:
        return None
    limpio = re.sub(r"[^A-Z ]+", " ", valor.upper())
    limpio = re.sub(r"\s+", " ", limpio).strip()
    # una letra suelta al final (o una I/L al inicio) es un separador mal leído
    limpio = re.sub(r"(?:\s+[A-Z])+$", "", limpio)
    limpio = re.sub(r"^(?:[IL]\s+)+", "", limpio)
    if not limpio or len(limpio) > 50:
        return None
    return limpio


def extraer_nombres_csf(texto_completo):
    """Extrae el/los Nombre(s) impresos en la CSF, anclado entre la etiqueta
    "Nombre (s):" y la etiqueta "Primer Apellido:" que le sigue. Regresa el
    texto detectado o None si no se encontró."""
    m = PATRON_NOMBRES_CSF.search(normaliza(texto_completo))
    return _limpia_nombre_csf(m.group(1)) if m else None


def extraer_apellido_paterno_csf(texto_completo):
    """Extrae el Primer Apellido impreso en la CSF, anclado entre las
    etiquetas "Primer Apellido:" y "Segundo Apellido:". Regresa el texto
    detectado o None si no se encontró."""
    m = PATRON_APELLIDO_PATERNO_CSF.search(normaliza(texto_completo))
    return _limpia_nombre_csf(m.group(1)) if m else None


def extraer_apellido_materno_csf(texto_completo):
    """Extrae el Segundo Apellido impreso en la CSF, anclado entre la
    etiqueta "Segundo Apellido:" y la etiqueta que le sigue en la plantilla
    del SAT ("Fecha inicio de operaciones:" en la CSF; "Nombre Comercial:" en
    el acuse de inscripción). Regresa el texto detectado o None si no se
    encontró (por ejemplo, si el contribuyente solo tiene un apellido, ese
    renglón viene vacío)."""
    m = PATRON_APELLIDO_MATERNO_CSF.search(normaliza(texto_completo))
    return _limpia_nombre_csf(m.group(1)) if m else None


# El renglón "Nombre(s): ..." de arriba a veces pierde el espacio ENTRE
# palabras del propio valor (p.ej. "Nombre(s): ALAMNARESH" en vez de "ALAM
# NARESH") -es la extracción de texto del PDF la que junta esas dos
# palabras si el hueco entre ellas es muy angosto en ese renglón en
# particular, no un error de la expresión regular-. La CSF trae ese mismo
# nombre bien espaciado en el encabezado, justo debajo de "Registro Federal
# de Contribuyentes" (es el campo "Nombre, denominación o razón social" que
# imprime la Cédula de Identificación Fiscal), así que se usa esa copia
# como fuente preferida para el nombre completo, y de ahí se le quitan los
# apellidos (ya extraídos aparte, sin ese problema porque son una sola
# palabra) para quedarse solo con los Nombres, bien espaciados.
PATRON_NOMBRE_COMPLETO_CSF_ENCABEZADO = re.compile(
    r"REGISTRO\s*FEDERAL\s*DE\s*CONTRIBUYENTES\s*(.+?)\s*NOMBRE\s*,\s*DENOMINACION\s*O\s*RAZON"
)


def extraer_nombre_completo_csf_encabezado(texto_completo):
    """Extrae el nombre completo tal como lo imprime el encabezado de la
    Cédula de Identificación Fiscal (bien espaciado, a diferencia del
    renglón "Nombre(s):" de más abajo). Regresa el texto detectado o None
    si no se encontró ese encabezado."""
    m = PATRON_NOMBRE_COMPLETO_CSF_ENCABEZADO.search(normaliza(texto_completo))
    if not m:
        return None
    valor = m.group(1).strip()
    # un nombre real son pocas palabras sin símbolos; si el patrón abarcó
    # texto de otras secciones (OCR/capa de texto basura), se descarta
    if not valor or len(valor) > 60 or re.search(r"[^A-Z ]", valor):
        return None
    return valor


def extraer_nombres_csf_bien_espaciado(texto_completo, apellido_paterno, apellido_materno):
    """Nombres(s) de la CSF, evitando el problema de espacio perdido descrito
    arriba: parte del nombre completo bien espaciado del encabezado y le
    quita los apellidos (segundo y luego primero) del final. Si ese
    encabezado no se encontró, cae de vuelta a extraer_nombres_csf (el
    renglón "Nombre(s):", que puede venir con las palabras pegadas si el
    nombre trae más de una)."""
    nombre_completo = extraer_nombre_completo_csf_encabezado(texto_completo)
    if nombre_completo:
        for apellido in (apellido_materno, apellido_paterno):
            if apellido and nombre_completo.upper().endswith(" " + apellido.upper()):
                nombre_completo = nombre_completo[: -len(apellido)].rstrip()
        if nombre_completo:
            return nombre_completo
    return extraer_nombres_csf(texto_completo)


# Número de crédito (Infonavit y Fonacot): etiqueta "Número de crédito"
# seguida de una secuencia de dígitos. En Infonavit aparece limpio bajo el
# apartado "Información del crédito del trabajador"; en Fonacot el
# formulario es más variado (a veces la tabla superior no se lee bien por
# OCR), pero el propio documento repite la etiqueta de forma clara más
# adelante ("...corresponde al número de crédito 17507"), así que se busca
# en todo el texto y se toma la primera coincidencia clara.
PATRON_NUMERO_CREDITO = re.compile(r"NUMERO\s*DE\s*CREDITO\s*:?\s*(\d{4,12})\b")

# Respaldo específico para Fonacot: en fotos/escaneos de calidad pareja el
# OCR a veces lee mal justo la palabra "CREDITO" de esa frase (p.ej.
# "CRODIO", "NURNERO DE CREDITO") aunque el resto de la frase y, sobre
# todo, los DÍGITOS del número salgan bien en las 4 combinaciones de
# preprocesado probadas. Por eso este patrón se ancla en "CORRESPONDE AL"
# (estable en las pruebas) y tolera hasta 5 palabras cualesquiera antes de
# los dígitos, en vez de exigir que "NUMERO DE CREDITO" se haya leído
# perfecto.
PATRON_NUMERO_CREDITO_FONACOT_ALT = re.compile(r"CORRESPONDE AL(?:\s+\S+){0,5}?\s+(\d{4,12})\b")


def extraer_numero_credito(texto_completo, etiqueta_seccion=None, admite_respaldo_fonacot=False):
    """Extrae el "Número de crédito" del texto. Si se da etiqueta_seccion
    (p.ej. "INFORMACION DEL CREDITO DEL TRABAJADOR" para Infonavit), primero
    recorta el texto a partir de esa etiqueta para evitar falsos positivos;
    si no se da (Fonacot), busca en todo el documento. Si
    admite_respaldo_fonacot=True y no se encontró nada con el patrón
    principal, se intenta también PATRON_NUMERO_CREDITO_FONACOT_ALT (ver su
    comentario). Regresa el número detectado (como texto, para no perder
    ceros a la izquierda) o None."""
    t = normaliza(texto_completo)
    if etiqueta_seccion:
        idx = t.find(normaliza(etiqueta_seccion))
        if idx != -1:
            t = t[idx:]
    m = PATRON_NUMERO_CREDITO.search(t)
    if m:
        return m.group(1)
    if admite_respaldo_fonacot:
        m_alt = PATRON_NUMERO_CREDITO_FONACOT_ALT.search(t)
        if m_alt:
            return m_alt.group(1)
    return None


def busca_fecha_nacimiento(texto):
    """Busca la fecha de nacimiento en el acta, anclada a la etiqueta
    'FECHA DE NACIMIENTO'. A diferencia de fecha_mas_reciente_razonable
    (pensada para vigencias de 3 meses), aquí SÍ se aceptan fechas de hace
    varias décadas -es justo lo esperado en una fecha de nacimiento-; solo
    se descartan fechas futuras o absurdamente viejas (antes de 1900).
    Regresa un date o None si no se encontró la etiqueta o no había una
    fecha legible cerca."""
    t_norm = normaliza(texto)
    for etiqueta in ("FECHA DE NACIMIENTO", "FECHADENACIMIENTO"):
        idx = t_norm.find(etiqueta)
        if idx == -1:
            continue
        ventana = t_norm[idx: idx + len(etiqueta) + 40]
        candidatas = [f for f in busca_fechas(ventana) if f[0] <= HOY and f[0].year >= 1900]
        if candidatas:
            return candidatas[0][0]
    return None


# Palabras que, si aparecen en un renglón después de la etiqueta de
# domicilio, casi seguro ya indican que se salió del bloque de dirección y
# se entró a otro campo del documento (periodo facturado, clave de
# elector, etc.) -así no se arrastra ese texto pegado a la dirección-.
_DIRECCION_ETIQUETAS_CORTE = [
    "PERIODO", "TOTAL A PAGAR", "LIMITE DE PAGO", "PAGAR ANTES", "CORTE A PARTIR",
    "CLAVE DE ELECTOR", "VIGENCIA", "SECCION", "CURP", "FOLIO", "ESTADO DE CUENTA",
    "SUBTOTAL", "MEDIDOR", "CONSUMO", "REFERENCIA", "RFC", "FECHA DE NACIMIENTO",
    "SEXO", "CLAVE DE LA ELECTORA", "INSTITUTO NACIONAL ELECTORAL", "CREDENCIAL PARA VOTAR",
    "AVISO DE PRIVACIDAD", "ANO DE REGISTRO", "EMISION", "MUNICIPIO EMISOR",
    "ELECTOR",  # "CLAVE DE ELECTOR" mal leída por el OCR ("CUVEDE ELECTOR")
]


def extraer_direccion(texto_completo, etiquetas, ventana=220, max_lineas=4):
    """Extracción best-effort de una dirección: busca la primera de las
    'etiquetas' dadas (ignorando acentos/mayúsculas) y arma la dirección
    tomando, renglón por renglón, el texto ORIGINAL que sigue -con sus
    acentos y mayúsculas/minúsculas tal cual los trae el documento, para
    que sea legible-. Se detiene en el primer renglón vacío o en el primer
    renglón que ya se ve como otro campo del documento (ver
    _DIRECCION_ETIQUETAS_CORTE), para no arrastrar texto de otras secciones
    pegado a la dirección; como respaldo también se limita a 'max_lineas'
    renglones y a 'ventana' caracteres.

    No existe un formato único de dirección en México que se pueda validar
    con una expresión regular (a diferencia del RFC o el CURP), así que
    esto SIEMPRE debe tratarse como una propuesta a confirmar contra el
    documento original, nunca como un dato ya verificado."""
    texto_original = texto_completo or ""
    texto_plano = _sin_acentos_mismo_largo(texto_original)
    for etiqueta in etiquetas:
        idx = texto_plano.find(etiqueta)
        if idx == -1:
            continue
        fin = idx + len(etiqueta) + ventana
        lineas_orig = texto_original[idx + len(etiqueta): fin].split("\n")
        lineas_plano = texto_plano[idx + len(etiqueta): fin].split("\n")

        recolectadas = []
        for linea_orig, linea_plano in zip(lineas_orig, lineas_plano):
            # Símbolos sueltos que mete el OCR (| \\ < > ~ _ * ^ ! ¡) no son parte de
            # ninguna dirección; se quitan para que no ensucien el resultado ni
            # hagan pasar por "texto" un renglón que solo trae una raya.
            linea_orig_limpia = re.sub(r"\s+", " ", re.sub(r"[|\\<>~_*^!¡]+", " ", linea_orig)).strip(" :.-\t")
            linea_plano_limpia = re.sub(r"[|\\<>~_*^!¡]+", " ", linea_plano).strip()
            if not linea_plano_limpia:
                if recolectadas:
                    break
                continue  # renglón vacío antes de empezar: se ignora, no corta
            if any(corte in linea_plano_limpia for corte in _DIRECCION_ETIQUETAS_CORTE):
                break
            if linea_orig_limpia:
                recolectadas.append(linea_orig_limpia)
            if len(recolectadas) >= max_lineas:
                break

        fragmento = ", ".join(recolectadas).strip(" ,")
        if len(fragmento) >= 8:
            return fragmento
    return None


def extraer_direccion_recibo_cfe(texto):
    """Recibo de CFE: el domicilio NO lleva etiqueta, va en el encabezado
    debajo del nombre del titular y antes de "NO. DE SERVICIO":

        IDALIA PEREZ DE LOPE                      TOTAL A PAGAR:
        DOROTEA MZ 17 LT 12                       $410
        CABALLO CALVILENO Y GUILLERMO DE MAMBRINO
        LA MANCHA II. C.P. 53717                  (CUATROCIENTOS DIEZ PESOS M.N)
        NAUCALPAN DE JUAREZ, MEX.
        NO. DE SERVICIO: 140890300274

    Regresa "calle, colonia C.P. nnnnn, municipio, estado" (se omite la línea
    de "entre calles" de en medio) o None si el texto no es un recibo de CFE
    o no se encuentra el bloque."""
    if not texto:
        return None
    plano = _sin_acentos_mismo_largo(texto).upper()
    if "ELECTRICIDAD" not in plano and "CFE" not in plano:
        return None
    lineas = texto.split("\n")
    lineas_plano = plano.split("\n")
    idx_serv = next((i for i, l in enumerate(lineas_plano) if re.search(r"N[O0]\.?\s*DE\s*SERVICIO", l)), None)
    if idx_serv is None:
        return None
    desde = max(0, idx_serv - 9)
    idx_total = idx_rfc = None
    for i in range(idx_serv - 1, desde - 1, -1):
        if idx_total is None and "TOTAL A PAGAR" in lineas_plano[i]:
            idx_total = i
        if idx_rfc is None and re.search(r"RFC\s*:?\s*CFE", lineas_plano[i]):
            idx_rfc = i
    if idx_total is not None:
        inicio = idx_total + 1
    elif idx_rfc is not None:
        inicio = idx_rfc + 2
    else:
        return None

    limpias = []
    for linea in lineas[inicio:idx_serv]:
        l = re.sub(r"TOTAL\s*A\s*PAGAR\s*:?", " ", linea, flags=re.IGNORECASE)
        l = re.sub(r"\([^)]*PESOS[^)]*\)?", " ", l, flags=re.IGNORECASE)
        l = re.sub(r"\$\s*[\d.,\s]*", " ", l)
        l = re.sub(r"[|\\<>~_*^!¡]+", " ", l)
        l = re.sub(r"\s+", " ", l).strip(" :-\t")
        l = re.sub(r"\s+\d{1,3}[A-Z]{1,2}$", "", l)  # "25H": resto del monto
        l = re.sub(r"\bl{2}\b", "II", l)  # "LA MANCHA ll." -> "LA MANCHA II."
        l = l.rstrip(" .")
        if l:
            limpias.append(l)
    if len(limpias) < 2:
        return None
    ultimo = limpias[-1]
    previas = limpias[:-1]
    if len(previas) >= 3:
        previas = [previas[0], previas[-1]]  # se omite "entre calles"
    direccion = ", ".join(previas + [ultimo])
    if not re.search(r"C\.?\s*P\.?\s*:?\s*\d{5}", direccion, flags=re.IGNORECASE):
        return None
    return direccion


def _direccion_comprobante(texto):
    return (extraer_domicilio_etiquetado(texto)
            or extraer_direccion_recibo_cfe(texto)
            or extraer_direccion(texto, _ETIQUETAS_DOMICILIO_COMPROBANTE))


def extraer_domicilio_etiquetado(texto):
    """Recibos (Naturgy, CFE, agua...) que imprimen el domicilio como campos
    con etiqueta en vez de un bloque "DOMICILIO":

        Calle: RAMON FABIE  Núm: 0014
        Colonia: VISTA ALEGRE  C.P.: 06860
        Mpo/Edo: CUAUHTEMOC, CD. DE MEX.

    Regresa la dirección en una línea ("calle núm, colonia, municipio, estado
    C.P. 06860", el mismo formato que extraer_direccion, para que lo
    descomponga descomponer_direccion) o None si no se encuentran, al menos,
    la calle y la colonia o el municipio. Tolera que el OCR pegue palabras
    ("RAMONFABIE") o suelte puntos/dos puntos."""
    t = re.sub(r"[ \t]+", " ", texto or "")

    def _uno(patron):
        m = re.search(patron, t, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", m.group(1)).strip(" :.-,") if m else None

    calle = _uno(r"\bCalle\s*[:.]\s*(.+?)(?=\s+N[uú]m(?:ero)?\b|\n|$)")
    numero = _uno(r"\bN[uú]m(?:ero)?\.?\s*[:.]\s*([A-Za-z0-9\-]+)")
    colonia = _uno(r"\bColonia\s*[:.]\s*(.+?)(?=\s+C\.?\s*P\b|\n|$)")
    cp = _uno(r"\bC\.?\s*P\.?\s*[:.]*\s*(\d{5})\b")
    municipio = _uno(r"\b(?:Mpo|Municipio|Delegaci[oó]n|Alcald[ií]a)\s*(?:/\s*Edo\.?)?\s*[:.]\s*(.+?)(?=\n|$)")
    if not calle or not (colonia or municipio):
        return None
    piezas = [" ".join(p for p in (calle, numero) if p), colonia, municipio]
    direccion = ", ".join(p for p in piezas if p)
    if cp:
        direccion += f" C.P. {cp}"
    return direccion


def _normaliza_etiqueta_domicilio(texto):
    """El OCR suele leer mal la etiqueta del domicilio del INE ("DOMICIIO",
    "DOMICUIO", "DOMCILIO"...). Se corrige a "DOMICILIO" para que
    extraer_direccion la encuentre."""
    return re.sub(r"(?<![A-Za-z])DOM[ILU1]?C[ILU1]{1,3}O(?![A-Za-z])", "DOMICILIO", texto or "", flags=re.IGNORECASE)


def _direccion_ine(texto):
    return extraer_direccion(_normaliza_etiqueta_domicilio(texto), ["DOMICILIO"])


def _puntaje_direccion(direccion):
    """0-1: qué tan "dirección de verdad" se ve una dirección extraída (sin
    símbolos de OCR, con estado reconocible, número de casa, municipio corto
    y limpio, código postal). Sirve para elegir entre varias lecturas de OCR
    de la misma credencial la que dejó el domicilio mejor leído."""
    if not direccion:
        return 0.0
    dom = descomponer_direccion(direccion)
    puntos = 0.0
    if not re.search(r"[><|!~*{}\[\]¡¿^]", direccion):
        puntos += 0.3
    if dom.get("estado") and _es_estado_mexico(normaliza(dom["estado"])):
        puntos += 0.25
    muni = dom.get("municipio") or ""
    if muni and len(muni.split()) <= 4 and re.fullmatch(r"[A-Za-zÁÉÍÓÚÑáéíóúñ .\-]+", muni):
        puntos += 0.2
    if dom.get("numero_exterior") or re.search(r"\d", dom.get("calle") or ""):
        puntos += 0.15
    if dom.get("codigo_postal"):
        puntos += 0.1
    return puntos


def _campos_clave_ine(texto):
    """(campos_ok, total) de lo que se necesita leer de la cara frontal del
    INE para el reporte: domicilio (calle + municipio o estado) y vigencia
    (dos años junto a VIGENCIA). Lo usa el pase de OCR reforzado para saber
    si ya terminó."""
    domicilio_ok = _puntaje_direccion(_direccion_ine(texto)) >= 0.7
    vigencia_ok = _busca_vigencia_ine(normaliza(texto)) is not None
    return int(domicilio_ok) + int(vigencia_ok), 2


def _campos_clave_comprobante(texto):
    """(campos_ok, total) para el comprobante de domicilio: una fecha que
    permita calcular la vigencia y un domicilio."""
    _, _, fecha = analiza_comprobante_domicilio(texto)
    direccion = _direccion_comprobante(texto)
    return int(fecha is not None) + int(bool(direccion)), 2


def _campos_clave_cuenta(texto):
    """(campos_ok, total) para el documento bancario: basta una CLABE
    válida o un número de cuenta explícito (el banco se deduce de la CLABE)."""
    r = analiza_cuenta_bancaria(texto, "")
    ok = bool(r["clabe_detectada"] and _clabe_valida(r["clabe_detectada"])) or bool(r["numero_cuenta"])
    return int(ok), 1


def _campos_clave_curp(texto):
    """(campos_ok, total) para el documento CURP: la clave de 18 caracteres."""
    return int(extraer_curp(texto) is not None), 1


def _campos_clave_csf(texto):
    """(campos_ok, total) para la hoja 1 de la CSF / acuse del RFC: el
    domicilio fiscal armado sin símbolos de ruido (con calle, municipio y
    entidad) y el código postal. En una CSF bajada del SAT siempre salen
    completos, así que el OCR reforzado casi nunca se activa."""
    d = extraer_domicilio_csf(texto)
    dir_ok = bool(
        d["direccion_fiscal"] and d["nombre_vialidad"] and d["municipio"] and d["entidad_federativa"]
        and not re.search(r"[~_|\\<>\[\]{}]", d["direccion_fiscal"])
        and not re.search(r"[:]", d["entidad_federativa"])
        and len(d["entidad_federativa"]) <= 30
    )
    return int(dir_ok) + int(bool(d["codigo_postal"])), 2


_ETIQUETAS_DOMICILIO_COMPROBANTE = [
    "DOMICILIO DEL SERVICIO", "DOMICILIO DEL USUARIO", "NOMBRE Y DOMICILIO DEL USUARIO",
    "DIRECCION DEL SERVICIO", "DOMICILIO DE INSTALACION", "DOMICILIO",
]

_CAMPOS_CLAVE_POR_CATEGORIA = {
    "INE": _campos_clave_ine,
    "COMPROBANTE_DOMICILIO": _campos_clave_comprobante,
    "CUENTA_BANCARIA": _campos_clave_cuenta,
    "CSF": _campos_clave_csf,
    "CURP": _campos_clave_curp,
}


def _mejores_lecturas_ine(candidatos):
    """De varias lecturas de OCR de la misma credencial, se queda con la que
    dejó mejor el DOMICILIO y con la primera que trae la VIGENCIA (pueden ser
    distintas: cada variante de OCR acierta en renglones diferentes), y las
    une con la del domicilio primero. Las demás lecturas -casi siempre ruido
    del reverso o de la marca de agua- se descartan."""
    if not candidatos:
        return ""
    mejor_dom = max(candidatos, key=lambda c: (_puntaje_direccion(_direccion_ine(c)), _calidad_texto(c)))
    elegidos = [mejor_dom]
    if _busca_vigencia_ine(normaliza(mejor_dom)) is None:
        for c in candidatos:
            if c is not mejor_dom and _busca_vigencia_ine(normaliza(c)) is not None:
                elegidos.append(c)
                break
    return "\n".join(elegidos)


def _refuerza_ocr_si_hace_falta(ruta, clave, paginas_texto, ocr_usado):
    """Si la categoría tiene campos clave (ver _CAMPOS_CLAVE_POR_CATEGORIA) y
    el texto que ya se tiene NO los deja a todos legibles, corre el OCR
    reforzado (ver _ocr_reforzado) y antepone su texto al de la página. Si la
    primera lectura ya salió completa NO hace nada (cero costo extra).
    Regresa (paginas_texto, ocr_usado, aplicado)."""
    chequeo = _CAMPOS_CLAVE_POR_CATEGORIA.get(clave)
    if chequeo is None:
        return paginas_texto, ocr_usado, False
    ok, total = chequeo("\n".join(paginas_texto))
    if ok >= total:
        return paginas_texto, ocr_usado, False
    paginas, ocr, aplicado = list(paginas_texto), list(ocr_usado), False
    # INE: solo la cara frontal (página 1); comprobante: hasta 2 páginas.
    limite = 1 if clave in ("INE", "CSF", "CURP") else min(len(paginas), 2)
    for i in range(limite):
        original = paginas[i]
        candidatos = _ocr_reforzado_candidatos(ruta, i, campos_ok=lambda t, o=original: chequeo(t + "\n" + o))
        if clave == "INE":
            alt = _mejores_lecturas_ine(candidatos)
        elif clave == "CSF" and candidatos:
            mejor = max(candidatos, key=lambda c: (chequeo(c + "\n" + original)[0], _calidad_texto(c)))
            # la mejor lectura va primero (manda en el domicilio); las demás
            # franjas aportan el resto de los campos (nombre, apellidos)
            alt = "\n".join([mejor] + [c for c in candidatos if c is not mejor])
        else:
            alt = "\n".join(candidatos)
        if alt:
            paginas[i] = alt + "\n" + original
            ocr[i] = True
            aplicado = True
        if chequeo("\n".join(paginas))[0] >= total:
            break
    return paginas, ocr, aplicado


# Nombres/abreviaturas de los 32 estados de México, tal como suelen
# aparecer impresos al final de una dirección (INE, comprobante de
# domicilio). Se usan SOLO para reconocer cuál pedazo de la dirección es
# el Estado (para poder separar el Municipio, que va justo antes); el
# texto que se guarda en la columna "Estado" siempre es el original tal
# como lo trae el documento (abreviado o completo), nunca esta lista.
ESTADOS_MEXICO_ALIAS = {
    "AGUASCALIENTES", "AGS",
    "BAJA CALIFORNIA", "BC", "BCN",
    "BAJA CALIFORNIA SUR", "BCS",
    "CAMPECHE", "CAMP",
    "CIUDAD DE MEXICO", "CDMX", "COMX", "CD DE MEXICO", "CD DE MEX", "CIUDAD DE MEX", "DISTRITO FEDERAL", "DF",
    "COAHUILA", "COAH", "COAHUILA DE ZARAGOZA",
    "COLIMA", "COL",
    "CHIAPAS", "CHIS",
    "CHIHUAHUA", "CHIH",
    "DURANGO", "DGO",
    "GUANAJUATO", "GTO",
    "GUERRERO", "GRO",
    "HIDALGO", "HGO",
    "JALISCO", "JAL",
    "ESTADO DE MEXICO", "EDOMEX", "EDO MEX", "EDO DE MEXICO", "MEX",
    "MICHOACAN", "MICH", "MICHOACAN DE OCAMPO",
    "MORELOS", "MOR",
    "NAYARIT", "NAY",
    "NUEVO LEON", "NL",
    "OAXACA", "OAX",
    "PUEBLA", "PUE",
    "QUERETARO", "QRO",
    "QUINTANA ROO", "QROO", "Q ROO",
    "SAN LUIS POTOSI", "SLP",
    "SINALOA", "SIN",
    "SONORA", "SON",
    "TABASCO", "TAB",
    "TAMAULIPAS", "TAMPS",
    "TLAXCALA", "TLAX",
    "VERACRUZ", "VER", "VERACRUZ DE IGNACIO DE LA LLAVE",
    "YUCATAN", "YUC",
    "ZACATECAS", "ZAC",
}


def _es_estado_mexico(texto_normalizado):
    """True si 'texto_normalizado' (mayúsculas, sin acentos) es -tal cual o
    quitándole un punto final- uno de los nombres/abreviaturas de estado en
    ESTADOS_MEXICO_ALIAS."""
    return re.sub(r"\s+", " ", texto_normalizado.replace(".", " ")).strip() in ESTADOS_MEXICO_ALIAS


def descomponer_direccion(direccion):
    """Best-effort: separa una dirección de una sola línea -como la que
    regresa extraer_direccion, con cada renglón del documento unido por
    ", "- en sus componentes (Calle, Número exterior, Número interior,
    Colonia, Municipio, Estado, Código Postal) para las columnas de la
    hoja "Datos completos".

    No existe un formato único de domicilio en México (a diferencia del
    RFC o el CURP), así que esto SIEMPRE es una propuesta a confirmar
    contra el documento original -por eso, igual que el resto de "Datos
    completos", se resalta en amarillo-, nunca un dato ya verificado.
    Cualquier campo que no se pudo reconocer queda en None; el Código
    Postal se reconoce junto a la etiqueta "C.P."/"CP" o, en su defecto,
    como número suelto de 5 dígitos al final del renglón de la Colonia
    (formato común en el INE, que casi nunca imprime la etiqueta "C.P."
    pero sí el número) -nunca se adivina de un domicilio de un solo
    renglón, para no confundirlo con un Número exterior largo.

    Regresa un dict con las llaves calle, numero_exterior,
    numero_interior, colonia, municipio, estado y codigo_postal."""
    resultado = {
        "calle": None, "numero_exterior": None, "numero_interior": None,
        "colonia": None, "municipio": None, "estado": None, "codigo_postal": None,
    }
    if not direccion:
        return resultado

    texto = direccion

    m_cp = re.search(r"C\.?\s*P\.?\s*:?\s*(\d{5})\b", texto, flags=re.IGNORECASE)
    if m_cp:
        resultado["codigo_postal"] = m_cp.group(1)
        texto = (texto[:m_cp.start()] + texto[m_cp.end():]).strip(" ,")

    # "México" como país no se necesita aislar aquí (va en su propia
    # columna, con un valor fijo); se quita si aparece suelto al final
    # para que no se confunda con el Estado.
    texto = re.sub(r",\s*M[EÉ]XICO\s*$", "", texto, flags=re.IGNORECASE).strip(" ,")

    partes = [p.strip() for p in texto.split(",") if p.strip()]
    if not partes:
        return resultado

    # Basura del OCR pegada después del estado ("CDMX Oe", "MEX e"): se quitan
    # las palabras de 1-2 letras del final solo si lo que queda termina en un
    # estado reconocible.
    palabras_fin = partes[-1].split()
    for quitar in (1, 2):
        if len(palabras_fin) > quitar and all(len(w) <= 2 and w.isalpha() for w in palabras_fin[-quitar:]):
            resto_fin = palabras_fin[:-quitar]
            if any(_es_estado_mexico(normaliza(" ".join(resto_fin[-n:]))) for n in (1, 2, 3) if len(resto_fin) >= n):
                partes[-1] = " ".join(resto_fin)
                break

    idx_estado = None
    estado_texto = None
    municipio_de_ultimo_segmento = None

    for i in (len(partes) - 1, len(partes) - 2):
        if i < 0:
            continue
        segmento = partes[i]
        if _es_estado_mexico(normaliza(segmento)):
            idx_estado, estado_texto = i, segmento
            break
        palabras = segmento.split()
        encontrado = False
        for n in (3, 2, 1):
            if len(palabras) > n:  # > n: que siempre quede algo de municipio
                cola = " ".join(palabras[-n:])
                if _es_estado_mexico(normaliza(cola)):
                    idx_estado, estado_texto = i, cola
                    municipio_de_ultimo_segmento = " ".join(palabras[:-n])
                    encontrado = True
                    break
        if encontrado:
            break

    resultado["estado"] = estado_texto

    if idx_estado is not None:
        if municipio_de_ultimo_segmento:
            resultado["municipio"] = municipio_de_ultimo_segmento or None
            resto = partes[:idx_estado]
        else:
            resto = partes[:idx_estado]
            if resto:
                resultado["municipio"] = resto[-1]
                resto = resto[:-1]
    else:
        # No se reconoció ningún estado: se asume, como respaldo, que el
        # último pedazo ya es el Municipio (p.ej. una alcaldía de CDMX que
        # no trae el estado explícito en el domicilio).
        resto = partes[:-1]
        if partes:
            resultado["municipio"] = partes[-1]

    # Muchas credenciales de INE sí traen el Código Postal, pero SIN la
    # etiqueta "C.P.": va pegado como número suelto de 5 dígitos al final
    # del renglón de la Colonia (p.ej. "FRACC SIAN KAAN V 97314"). Esto
    # solo se acepta cuando la Colonia viene en su propio renglón (resto
    # con 2+ pedazos): así no se confunde con un Número exterior de 5
    # dígitos en domicilios de un solo renglón (calle+número), donde un
    # número exterior tan largo sería rarísimo pero no imposible.
    if resultado["codigo_postal"] is None and len(resto) >= 2:
        m_cp_suelto = re.search(r"(?<!\d)(\d{5})(?:\s+[A-Za-z]{1,2})?\s*$", resto[-1])
        if m_cp_suelto:
            resultado["codigo_postal"] = m_cp_suelto.group(1)
            recorte = resto[-1][:m_cp_suelto.start()].strip(" ,.-")
            if recorte:
                resto[-1] = recorte
            else:
                resto = resto[:-1]

    primero = None
    if resto:
        primero = resto[0]
        if len(resto) > 1:
            resultado["colonia"] = ", ".join(resto[1:])

    if primero:
        # El número interior (si viene pegado en la misma línea, p.ej. "AV
        # REFORMA 100 INT 4B") se separa ANTES que el exterior: si no, el
        # exterior se lo llevaría por error (es el que queda más a la
        # derecha del renglón).
        m_int = re.search(r"\bINT(?:ERIOR)?\.?\s*([A-Za-z0-9\-]+(?:\s+[A-Za-z])?)$", primero, flags=re.IGNORECASE)
        if m_int:
            resultado["numero_interior"] = m_int.group(1)
            primero = primero[:m_int.start()].strip(" ,.-")
        # "DOROTEA MZ 17 LT 12": manzana y lote juntos son el número exterior
        m_mzlt = re.search(r"\s+((?:MZA?|MANZANA)\.?\s*\d+\s*(?:LT|LOTE)\.?\s*\d+)$", primero, flags=re.IGNORECASE)
        if m_mzlt:
            resultado["numero_exterior"] = m_mzlt.group(1)
            primero = primero[:m_mzlt.start()].strip(" ,.-")
        palabras = primero.split()
        if resultado["numero_exterior"] is None and len(palabras) >= 2 and re.match(r"^\d+[A-Za-z]?$", palabras[-1]):
            resultado["numero_exterior"] = palabras[-1]
            primero = " ".join(palabras[:-1])
        resultado["calle"] = primero or None

    return resultado


CODIGOS_CLABE_BANCOS = {
    "002": "Banamex/Citibanamex", "006": "Bancomext", "009": "Banobras",
    "012": "BBVA México", "014": "Santander", "019": "Banjercito",
    "021": "HSBC", "030": "Bajío", "036": "Inbursa", "037": "Interbanco",
    "042": "Mifel", "044": "Scotiabank", "058": "Banregio", "059": "Invex",
    "060": "Bansi", "062": "Afirme", "072": "Banorte", "102": "Multiva",
    "103": "American Express", "106": "Bank of America", "108": "MUFG",
    "110": "JP Morgan", "112": "BMONEX", "113": "VE POR MAS", "124": "Deutsche Bank",
    "127": "Banco Azteca", "128": "Autofin", "129": "Barclays", "130": "Compartamos",
    "131": "Banco Famsa", "132": "BMULTIVA", "133": "Actinver", "134": "Wal-Mart (Bancoppel línea)",
    "135": "Nafin", "136": "Intercam Banco", "137": "Bankaool", "138": "ABC Capital",
    "140": "Consubanco", "141": "Volkswagen Bank", "143": "CIBanco", "145": "BBASE",
    "147": "Bankaool", "148": "PagaTodo", "150": "Inmobiliario", "151": "Donde",
    "152": "Bancrea", "154": "Banco Covalto", "155": "ICBC", "156": "Sabadell",
    "157": "Shinhan", "158": "Mizuho Bank", "160": "Banco S3", "166": "Bansefi/Bienestar",
    "168": "Hipotecaria Federal",
    # Fintechs / billeteras que el negocio no acepta como cuenta de nómina
    "638": "NVIO Pagos México (Nu)", "722": "Mercado Pago W Digital", "728": "Spin by OXXO",
    "659": "Openpay/Klar (según convenio)", "646": "STP (posible operador de una fintech)",
}

BANCOS_EXCLUIDOS = {"NVIO Pagos México (Nu)", "Mercado Pago W Digital", "Spin by OXXO"}
NOMBRES_EXCLUIDOS_TEXTO = ["NU MEXICO", "NU BANK", "SPIN BY OXXO", "MERCADO PAGO", "MERCADOPAGO"]


def _clabe_valida(clabe):
    """Valida los 18 dígitos y el dígito verificador de una CLABE."""
    if not clabe or len(clabe) != 18 or not clabe.isdigit():
        return False
    pesos = [3, 7, 1] * 6
    suma = sum((int(d) * pesos[i]) % 10 for i, d in enumerate(clabe[:17]))
    return (10 - suma % 10) % 10 == int(clabe[17])


def analiza_cuenta_bancaria(texto_completo, nombre_candidato):
    obs = []
    t_norm = normaliza(texto_completo)

    # la etiqueta puede venir mal leída por el OCR (CLASE, CLABF, CLAVE...)
    m_clabe = re.search(r"CLA[BV8S][EF3]\b[^0-9]{0,15}(\d[\d\s]{16,22}\d)", t_norm)
    clabe = re.sub(r"\s", "", m_clabe.group(1)) if m_clabe else None
    if clabe and len(clabe) >= 18:
        clabe = clabe[:18]
    if not clabe or not _clabe_valida(clabe):
        # respaldo: cualquier número de 18 dígitos con dígito verificador
        # válido y código de banco conocido (no depende de la etiqueta)
        for m in re.finditer(r"(?<!\d)(\d{18})(?!\d)", t_norm):
            if _clabe_valida(m.group(1)) and m.group(1)[:3] in CODIGOS_CLABE_BANCOS:
                clabe = m.group(1)
                break

    banco = None
    if clabe:
        banco = CODIGOS_CLABE_BANCOS.get(clabe[:3], f"Código de banco no identificado ({clabe[:3]})")

    m_cuenta = re.search(r"(?:NO\.?\s*DE\s*CUENTA|NUMERO\s*DE\s*CUENTA|CUENTA)\s*:?\s*(\d{6,20})", t_norm)
    numero_cuenta = m_cuenta.group(1) if m_cuenta else None
    cuenta_derivada = False
    if numero_cuenta is None and clabe and _clabe_valida(clabe):
        # en una CLABE, los dígitos 7 a 17 son el número de cuenta (11 dígitos)
        numero_cuenta = clabe[6:17]
        cuenta_derivada = True
    tiene_cuenta = numero_cuenta is not None
    nombre_ok, proporcion = nombre_coincide(texto_completo, nombre_candidato)

    banco_rechazado = False
    if banco in BANCOS_EXCLUIDOS:
        banco_rechazado = True
    elif not clabe:
        # sin CLABE detectable por regex, buscamos mención directa de banco excluido en el encabezado
        encabezado = t_norm[:600]
        if any(k in encabezado for k in NOMBRES_EXCLUIDOS_TEXTO):
            banco_rechazado = True
            banco = "Nu / Spin / Mercado Pago (detectado por texto, sin CLABE confirmada)"

    if banco_rechazado:
        obs.append(f"Banco NO permitido para depósito de nómina: {banco}.")
    elif banco is None:
        obs.append("No se detectó CLABE ni banco; revisar manualmente que sea un banco permitido.")
    else:
        obs.append(f"Banco detectado: {banco} (permitido).")

    if not clabe:
        obs.append("No se encontró una CLABE de 18 dígitos legible.")
    if cuenta_derivada:
        obs.append(f"Número de cuenta tomado de la CLABE (dígitos 7 a 17): {numero_cuenta}.")
    elif not tiene_cuenta:
        obs.append("No se encontró un número de cuenta explícito.")
    obs.append("El logo del banco no puede confirmarse por OCR: revisar visualmente el PDF.")

    return {
        "banco": banco,
        "banco_rechazado": banco_rechazado,
        "clabe_detectada": clabe,
        "numero_cuenta": numero_cuenta,
        "tiene_cuenta": tiene_cuenta,
        "nombre_coincide": nombre_ok,
        "observaciones": " ".join(obs),
    }


def _fecha_emision_csf_por_etiqueta(t_norm):
    """Busca el campo "Lugar y Fecha de Emisión" de la CSF (arriba a la
    derecha del documento, junto al código de barras) y regresa la fecha que
    lo acompaña, con formato "<lugar> A <dia> DE <mes> DE <anio>" (por
    ejemplo "TLALPAN, CIUDAD DE MEXICO A 21 DE AGOSTO DE 2026").

    Esta es la fecha correcta para calcular la vigencia de 3 meses: es la
    fecha en que se generó/descargó la constancia desde el portal del SAT,
    NO hay que confundirla con otras fechas que trae el documento (como la
    fecha de inicio de operaciones del contribuyente, que puede tener años
    de antigüedad).

    Se busca dentro de una ventana amplia de texto después de la etiqueta
    (en vez de exigir que la fecha venga inmediatamente después) porque la
    extracción de texto de un PDF con columnas -como el encabezado de la
    CSF, que trae la columna del QR/RFC a la izquierda y esta etiqueta a la
    derecha- a veces intercala renglones de una columna con la otra; la
    fecha en sí (p.ej. "21 DE AGOSTO DE 2026") sigue apareciendo completa y
    en orden, solo que puede no estar pegada a la etiqueta.

    Se prueba también la etiqueta sin espacios ("LUGARYFECHADEEMISION") por
    si la pérdida de espacios de la extracción llega a afectar hasta al
    propio texto de la etiqueta, no solo a la fecha."""
    idx = t_norm.find("LUGAR Y FECHA DE EMISION")
    if idx == -1:
        idx = t_norm.find("LUGARYFECHADEEMISION")
    if idx == -1:
        return None
    ventana = t_norm[idx:idx + 500]
    m = PATRON_FECHA_LARGA.search(ventana)
    if not m:
        return None
    mes = MESES.get(m.group(2))
    if not mes:
        return None
    try:
        return datetime.date(int(m.group(3)), mes, int(m.group(1)))
    except ValueError:
        return None


def analiza_csf(paginas_texto):
    texto_completo = "\n".join(paginas_texto)
    t_norm = normaliza(texto_completo)
    obs = []
    # El "Acuse único de inscripción al RFC" es una sola hoja, sin régimen
    # ni estatus: no es la CSF completa pero se acepta como comprobante del RFC.
    es_acuse = re.search(r"ACUSE\s*UNICO\W*DE\s*INSCRIPCION", t_norm) is not None
    if es_acuse:
        obs.append(
            "El documento es el Acuse único de inscripción al RFC (una sola hoja; no trae régimen ni estatus en el "
            "padrón), no la Constancia de Situación Fiscal completa."
        )

    rfc = extraer_rfc(texto_completo)
    if rfc:
        obs.append(f"RFC detectado: {rfc}.")
    else:
        obs.append("No se detectó un RFC con el formato esperado junto a la etiqueta 'RFC'; revisar manualmente.")

    curp = extraer_curp(texto_completo)
    if curp:
        obs.append(f"CURP detectado: {curp}.")
    else:
        obs.append("No se detectó un CURP con el formato esperado en la CSF; revisar manualmente.")

    codigo_postal = extraer_codigo_postal_csf(texto_completo)
    if codigo_postal:
        obs.append(f"Código Postal detectado: {codigo_postal}.")
    else:
        obs.append("No se detectó el Código Postal junto a la etiqueta 'Código Postal'; revisar manualmente.")

    domicilio = extraer_domicilio_csf(texto_completo)
    if domicilio["direccion_fiscal"]:
        obs.append(f"Dirección fiscal armada con los datos del domicilio registrado: {domicilio['direccion_fiscal']}.")
    else:
        obs.append(
            "No se detectó la sección 'Datos del domicilio registrado' con el formato esperado; "
            "revisar manualmente la Dirección fiscal."
        )

    regimen = extraer_regimen_csf(texto_completo)
    if regimen:
        obs.append(f"Régimen detectado (hoja 2): {regimen}.")
    elif not es_acuse:
        obs.append("No se detectó la tabla de 'Regímenes' (hoja 2) con el formato esperado; revisar manualmente.")

    apellido_paterno = extraer_apellido_paterno_csf(texto_completo)
    apellido_materno = extraer_apellido_materno_csf(texto_completo)
    nombres = extraer_nombres_csf_bien_espaciado(texto_completo, apellido_paterno, apellido_materno)
    if nombres or apellido_paterno or apellido_materno:
        obs.append(
            "Nombre(s)/Apellidos detectados en la CSF: "
            f"{nombres or '—'} / {apellido_paterno or '—'} / {apellido_materno or '—'}."
        )
    else:
        obs.append(
            "No se detectaron Nombre(s)/Primer Apellido/Segundo Apellido junto a esas etiquetas en la CSF; "
            "revisar manualmente."
        )

    # el PDF de la CSF a veces pierde los espacios entre palabras al extraer texto
    # ("Estatusenelpadrón:ACTIVO"), así que probamos con y sin espacios.
    m_estatus = re.search(r"ESTATUS\s*EN\s*EL\s*PADRON\s*:?\s*([A-Z]+)", t_norm)
    estatus = m_estatus.group(1) if m_estatus else None
    if estatus == "ACTIVO":
        obs.append("La CSF indica estatus ACTIVO en el padrón del SAT.")
    elif estatus:
        obs.append(f"La CSF indica estatus '{estatus}' (revisar, no es ACTIVO).")
    elif not es_acuse:
        obs.append("No se pudo leer el estatus del contribuyente en el texto; revisar manualmente.")

    # Se prioriza la fecha junto a "Lugar y Fecha de Emisión" (ver docstring
    # de _fecha_emision_csf_por_etiqueta). Antes se usaba "la fecha más
    # reciente de todo el documento", lo cual a veces tomaba por error una
    # fecha vieja de otra sección de la CSF y marcaba la constancia como
    # fuera de vigencia sin estarlo en realidad. Si por algún motivo no se
    # encuentra la etiqueta (formato distinto, texto muy degradado), se cae
    # de vuelta al método anterior como respaldo, dejando claro en las
    # observaciones que debe revisarse a mano.
    emision_fecha = _fecha_emision_csf_por_etiqueta(t_norm)
    if emision_fecha:
        obs.append(f"Fecha de emisión (\"Lugar y fecha de emisión\"): {emision_fecha.isoformat()}.")
    else:
        fechas = busca_fechas(texto_completo)
        candidata = fecha_mas_reciente_razonable(fechas)
        emision_fecha = candidata[0] if candidata else None
        if emision_fecha:
            obs.append(
                f"No se encontró la etiqueta \"Lugar y fecha de emisión\"; se usó como respaldo la "
                f"fecha más reciente detectada en el documento: {emision_fecha.isoformat()} "
                "(revisar manualmente que sea correcta)."
            )

    dentro_3_meses = None
    if emision_fecha:
        dias = (HOY - emision_fecha).days
        dentro_3_meses = dias <= 92
        obs.append(f"{dias} día(s) de antigüedad respecto a hoy ({HOY.isoformat()}).")
    else:
        obs.append("No se detectó una fecha de emisión clara; revisar manualmente.")

    obs.append("Autenticidad ante el SAT: no se consulta en vivo el portal del SAT desde este script; "
                "si el documento trae QR legible, ábranlo en el validador oficial para confirmar.")

    return {
        "estatus": estatus,
        "activo": estatus == "ACTIVO",
        "rfc": rfc,
        "curp": curp,
        "codigo_postal": codigo_postal,
        "direccion_fiscal": domicilio["direccion_fiscal"],
        "regimen": regimen,
        "nombres": nombres,
        "apellido_paterno": apellido_paterno,
        "apellido_materno": apellido_materno,
        "fecha_emision": emision_fecha,
        "dentro_3_meses": dentro_3_meses,
        "num_paginas_ok": len(paginas_texto) >= 2 or es_acuse,
        "observaciones": " ".join(obs),
    }


def _ocr_vigencia_ine(ruta):
    """Reintento localizado, solo para el renglón de Vigencia de la cara
    frontal del INE: se usa cuando el texto que ya se extrajo (ver
    extraer_texto_pdf) no trajo dos años legibles junto a "VIGENCIA".

    Ese renglón ("FECHA DE NACIMIENTO / SECCIÓN / VIGENCIA") suele ir
    impreso sobre un fondo de color (mapa de México, franjas) que un
    umbral de binarización calculado con el brillo de TODA la credencial
    -como hace _preprocesa_para_ocr- no necesariamente separa bien: puede
    quedar bien para el resto del documento y mal justo ahí. Aquí se
    recorta solo el tercio inferior de la imagen incrustada (donde
    siempre va esa fila) y se calcula el umbral con el brillo de ESE
    recorte, probando varios factores -no uno solo- y quedándose con la
    primera combinación que efectivamente encuentra "VIGENCIA" seguida de
    2 años, en vez de la de mayor confianza general (aquí solo importa
    acertar ese campo puntual).

    Regresa el texto reconocido (para que analiza_ine lo procese con la
    misma lógica de siempre) o None si ninguna combinación funcionó."""
    try:
        imagenes = _extraer_imagenes_incrustadas(ruta, 0)
    except Exception:
        imagenes = []
    if not imagenes:
        return None
    frente = imagenes[0]
    ancho, alto = frente.size
    recorte = frente.crop((0, int(alto * 0.65), ancho, alto))
    gris = ImageOps.autocontrast(ImageOps.grayscale(recorte), cutoff=0)
    brillo = ImageStat.Stat(gris).mean[0]
    for factor in (0.7, 0.85, 1.0):
        umbral = brillo * factor
        binaria = gris.point(lambda x, u=umbral: 255 if x > u else 0)
        candidata = binaria.filter(ImageFilter.SHARPEN)
        for psm in ("6", "3"):
            try:
                texto, _ = _ocr_con_confianza(candidata, lang="spa", config=f"--psm {psm}")
            except Exception:
                texto = ""
            t_norm = normaliza(texto)
            idx = t_norm.find("VIGENCIA")
            if idx != -1 and len(re.findall(r"\b(19\d{2}|20\d{2})\b", t_norm[idx: idx + 80])) >= 2:
                return texto
    return None


def _busca_vigencia_ine(texto_norm):
    """(anio_inicio, anio_fin) de la primera aparición de "VIGENCIA" que
    traiga dos años (19xx/20xx) en los 80 caracteres siguientes; None si no
    hay ninguna. Se prueban TODAS las apariciones (no solo la primera): con
    varias lecturas de OCR concatenadas, una puede traer la palabra sin los
    años y otra con ellos."""
    for m in re.finditer("VIGENCIA", texto_norm):
        anios = re.findall(r"\b(19\d{2}|20\d{2})\b", texto_norm[m.start(): m.start() + 80])
        if len(anios) >= 2:
            return int(anios[-2]), int(anios[-1])
    return None


def analiza_ine(paginas_texto, ruta=None):
    obs = []

    dos_paginas = len(paginas_texto) >= 2
    if dos_paginas:
        obs.append("PDF trae 2 páginas: se asume frente y reverso en un solo archivo.")
    else:
        obs.append("El PDF trae solo 1 página: falta el reverso (o viene en archivo separado).")

    # La vigencia SIEMPRE se busca en la cara frontal (página 1) — el reverso
    # es la franja MRZ y no trae el campo "VIGENCIA", así que buscarla ahí
    # solo metería ruido. Si el PDF llegó con una sola página, se analiza esa
    # misma como frente.
    texto_frente = paginas_texto[0] if paginas_texto else ""
    t_norm = normaliza(texto_frente)

    # El OCR de la credencial suele meter ruido entre "VIGENCIA" y los años
    # (números de sección, caracteres mal leídos), así que buscamos los dos
    # años (19xx/20xx) más cercanos después de la palabra VIGENCIA en vez de
    # exigir que estén pegados a ella.
    def _busca_vigencia(texto_norm):
        return _busca_vigencia_ine(texto_norm)

    encontrada = _busca_vigencia(t_norm)
    recuperada_con_reintento = False
    # Ese renglón suele ir sobre un fondo de color que un umbral de
    # binarización global (calculado con el brillo de TODA la credencial)
    # no siempre separa bien; si no se encontró aquí, vale la pena un
    # reintento localizado sobre esa franja antes de rendirse (ver
    # _ocr_vigencia_ine). Solo se hace cuando de verdad hace falta -no en
    # cada INE- porque implica volver a rasterizar/OCR-ear.
    if encontrada is None and ruta:
        texto_retry = _ocr_vigencia_ine(ruta)
        if texto_retry:
            encontrada = _busca_vigencia(normaliza(texto_retry))
            recuperada_con_reintento = encontrada is not None

    vigente = None
    anio_fin = None
    if encontrada:
        anio_inicio, anio_fin = encontrada
        vigente = anio_fin >= HOY.year
        dias_para_vencer = (datetime.date(anio_fin, 12, 31) - HOY).days
        sufijo = " (recuperada con un reintento localizado sobre ese renglón)" if recuperada_con_reintento else ""
        if vigente:
            obs.append(f"Vigencia impresa en la cara frontal: {anio_inicio}-{anio_fin} (vigente hoy {HOY.isoformat()}){sufijo}.")
        else:
            obs.append(f"Vigencia impresa en la cara frontal: {anio_inicio}-{anio_fin} — VENCIDA (venció hace {abs(dias_para_vencer)} días respecto a hoy {HOY.isoformat()}){sufijo}.")
    elif t_norm.find("VIGENCIA") != -1:
        obs.append("Se encontró la palabra VIGENCIA en la cara frontal pero no dos años legibles junto a ella; revisar manualmente.")
    else:
        obs.append("No se detectó el campo de vigencia en la cara frontal; revisar manualmente.")

    return {
        "vigente": vigente,
        "vigencia_anio_fin": anio_fin,
        "dos_paginas": dos_paginas,
        "observaciones": " ".join(obs),
    }


def analiza_fecha_limite(texto_completo, dias_limite, etiqueta):
    """Busca el campo de vigencia/fecha en el texto y lo compara contra el día
    de hoy. Regresa (dentro_del_limite, texto_de_observacion, fecha_encontrada)."""
    fechas = busca_fechas(texto_completo)
    fecha = fecha_mas_reciente_razonable(fechas)
    if not fecha:
        return None, f"No se detectó el campo de vigencia/fecha en {etiqueta}; revisar manualmente.", None
    dias = (HOY - fecha[0]).days
    dentro = dias <= dias_limite
    if dentro:
        obs = (f"Vigencia de {etiqueta}: fecha detectada {fecha[0].isoformat()}, hoy es {HOY.isoformat()} "
               f"({dias} días de antigüedad, dentro del límite de {dias_limite} días).")
    else:
        obs = (f"Vigencia de {etiqueta}: fecha detectada {fecha[0].isoformat()}, hoy es {HOY.isoformat()} "
               f"— FUERA DE VIGENCIA ({dias} días de antigüedad, supera el límite de {dias_limite} días).")
    return dentro, obs, fecha[0]


# Etiquetas que suelen imprimir los recibos (CFE, agua, etc.) cuando NO traen
# una fecha de emisión clara, pero sí una fecha de vencimiento del pago o del
# periodo. A propósito no se exige que la fecha sea "DE" o "DEL" exacto: al
# buscar por subcadena ("LIMITE DE PAGO" dentro de "FECHA LIMITE DE PAGO",
# o "PAGAR ANTES DE" dentro de "PAGAR ANTES DEL DÍA") ya cubre variantes
# comunes sin necesitar una lista larguísima.
ETIQUETAS_FECHA_LIMITE_DOMICILIO = [
    "LIMITE DE PAGO",
    "PAGAR ANTES DE",
    "CORTE A PARTIR DE",
]


def busca_fecha_por_etiquetas(texto, etiquetas, ventana=50):
    """Busca una fecha que aparezca justo después de alguna de las
    etiquetas dadas (p.ej. 'LÍMITE DE PAGO: 15 JUL 26'). A diferencia de
    fecha_mas_reciente_razonable, aquí SÍ se aceptan fechas futuras: estas
    etiquetas casi siempre marcan una fecha de vencimiento que cae después
    de la fecha de emisión del recibo, incluso después de hoy si el recibo
    es reciente.
    Regresa (fecha, etiqueta_encontrada) o None si no se encontró nada."""
    t_norm = normaliza(texto)
    for etiqueta in etiquetas:
        idx = t_norm.find(normaliza(etiqueta))
        if idx == -1:
            continue
        ventana_texto = t_norm[idx: idx + len(etiqueta) + ventana]
        fechas = busca_fechas(ventana_texto)
        candidatas = [f for f in fechas if 2015 <= f[0].year <= HOY.year + 1]
        if candidatas:
            return candidatas[0][0], etiqueta
    return None


def analiza_comprobante_domicilio(texto_completo):
    """Regla de vigencia del comprobante de domicilio.

    Primero se usa el criterio general (cualquier fecha del texto, la más
    reciente que no sea futura). Muchos recibos (CFE, agua, etc.) no traen
    una fecha de emisión explícita, solo una fecha límite de pago o de
    corte —y esa fecha casi siempre cae DESPUÉS de hoy si el recibo es
    reciente—, así que si el criterio general no encuentra nada se busca
    puntualmente junto a las etiquetas 'LÍMITE DE PAGO', 'PAGAR ANTES DE' o
    'CORTE A PARTIR DE', y en ese caso SÍ se acepta aunque sea una fecha
    futura: que la fecha de pago todavía no llegue no invalida el recibo,
    al contrario, es señal de que es reciente. En ambos casos la regla de
    vigencia es la misma: si la fecha encontrada tiene más de 3 meses
    (92 días) de antigüedad respecto a hoy, se marca fuera de vigencia."""
    dentro, obs, fecha = analiza_fecha_limite(texto_completo, 92, "el comprobante de domicilio")
    if fecha is not None:
        return dentro, obs, fecha

    encontrada = busca_fecha_por_etiquetas(texto_completo, ETIQUETAS_FECHA_LIMITE_DOMICILIO)
    if not encontrada:
        return dentro, obs, fecha

    fecha_etiqueta, etiqueta = encontrada
    dias = (HOY - fecha_etiqueta).days
    dentro2 = dias <= 92
    if dias >= 0:
        estado = "dentro del límite de 92 días." if dentro2 else "— FUERA DE VIGENCIA, supera el límite de 92 días."
        obs2 = (f"No se detectó una fecha de emisión explícita, pero se encontró la fecha de "
                f"'{etiqueta.title()}': {fecha_etiqueta.isoformat()} (hoy es {HOY.isoformat()}, "
                f"{dias} días de antigüedad, {estado}")
    else:
        obs2 = (f"No se detectó una fecha de emisión explícita, pero se encontró la fecha de "
                f"'{etiqueta.title()}': {fecha_etiqueta.isoformat()}, que todavía no llega (hoy es "
                f"{HOY.isoformat()}) — el recibo es reciente, se considera dentro de vigencia.")
    return dentro2, obs2, fecha_etiqueta


# ---------------------------------------------------------------------------
# Procesamiento principal
# ---------------------------------------------------------------------------

def procesar_documento(ruta, nombre_candidato, categoria_forzada=None):
    """Si se da categoria_forzada (clave de CATEGORIAS), se usa esa
    categoría directamente en vez de intentar reconocerla por contenido —
    esto es lo que usa la carga masiva por ZIP, donde la categoría ya viene
    determinada por el NOMBRE del archivo (ver
    detectar_categoria_por_nombre_archivo). La carga de un candidato a la
    vez sigue funcionando igual que siempre (sin categoria_forzada)."""
    nombre_archivo = os.path.basename(ruta)
    paginas_texto, num_paginas, ocr_usado = extraer_texto_pdf(ruta)
    texto_completo = "\n".join(paginas_texto)
    legible = any(len(p.strip()) >= 20 for p in paginas_texto)
    clave = categoria_forzada or clasificar(texto_completo, nombre_archivo)
    nombre_legible, _, _ = CATEGORIAS.get(clave, ("Desconocido / no identificado", [], False))

    # OCR reforzado SOLO si esta categoría tiene campos clave (INE: domicilio y
    # vigencia; comprobante: fecha y domicilio) y la primera lectura no los
    # dejó legibles. Si ya salieron bien, no se gasta ni un segundo extra.
    paginas_texto, ocr_usado, ocr_reforzado_aplicado = _refuerza_ocr_si_hace_falta(
        ruta, clave, paginas_texto, ocr_usado)
    texto_completo = "\n".join(paginas_texto)

    fila = {
        "archivo": nombre_archivo,
        "categoria_clave": clave,
        "categoria": nombre_legible,
        "num_paginas": num_paginas,
        "uso_ocr": any(ocr_usado),
        "legible": legible,
        "texto_muestra": texto_completo[:400].replace("\n", " ").strip(),
        "nombre_coincide": None,
        "detalle": "",
    }

    if not legible:
        fila["detalle"] = "El documento no pudo leerse (ni texto nativo ni OCR); solicitar de nuevo, escaneo/foto más nítida."
        return fila

    excluye_nombre = clave == "COMPROBANTE_DOMICILIO"
    if not excluye_nombre:
        coincide, proporcion = nombre_coincide(texto_completo, nombre_candidato)
        fila["nombre_coincide"] = coincide
        fila["detalle"] += f"Coincidencia de nombre: {proporcion*100:.0f}% de las palabras del nombre capturado se encontraron en el documento. "

    detalles_extra = []
    if categoria_forzada in CATEGORIAS:
        sugerida = sugerir_categoria_por_contenido(texto_completo, categoria_forzada)
        if sugerida:
            fila["categoria_sugerida"] = sugerida[0]
            fila["categoria_sugerida_nombre"] = sugerida[1]
            detalles_extra.append(
                f"⚠ POSIBLE DOCUMENTO EN EL CAMPO EQUIVOCADO: se cargó como «{nombre_legible}», pero el "
                f"contenido parece «{sugerida[1]}». Revisar el PDF y pedir que se cargue en el campo correcto."
            )
    if ocr_reforzado_aplicado:
        detalles_extra.append("La primera lectura no dejó legibles los datos clave; se aplicó una lectura reforzada (OCR) a la imagen.")

    if clave == "CSF":
        r = analiza_csf(paginas_texto)
        fila["num_paginas_ok"] = r["num_paginas_ok"]
        fila["vigencia_ok"] = r["dentro_3_meses"]
        fila["vigencia_fecha_texto"] = r["fecha_emision"].isoformat() if r["fecha_emision"] else None
        fila["estatus_sat"] = r["estatus"]
        fila["rfc"] = r["rfc"]
        fila["curp"] = r["curp"]
        fila["codigo_postal"] = r["codigo_postal"]
        fila["direccion_fiscal"] = r["direccion_fiscal"]
        fila["regimen"] = r["regimen"]
        fila["nombres"] = r["nombres"]
        fila["apellido_paterno"] = r["apellido_paterno"]
        fila["apellido_materno"] = r["apellido_materno"]
        detalles_extra.append(r["observaciones"])
        if not r["num_paginas_ok"]:
            detalles_extra.append("Falta la segunda hoja de la CSF en este PDF (debe traer ambas en 1 solo archivo).")

    elif clave == "CURP":
        curp = extraer_curp(texto_completo)
        fila["curp"] = curp
        if curp:
            detalles_extra.append(f"CURP detectado: {curp}.")
        else:
            detalles_extra.append("No se detectó un CURP con el formato esperado (18 caracteres); revisar manualmente.")

    elif clave == "NSS":
        nss = extraer_nss(texto_completo)
        fila["numero_seguridad_social"] = nss
        if nss:
            detalles_extra.append(f"Número de Seguridad Social detectado: {nss}.")
        else:
            detalles_extra.append(
                "No se detectó el Número de Seguridad Social (11 dígitos) junto a esa etiqueta; revisar manualmente."
            )

    elif clave == "FONACOT":
        numero_credito = extraer_numero_credito(texto_completo, admite_respaldo_fonacot=True)
        if not numero_credito:
            # El pase rápido de OCR a veces lee mal justo la tabla angosta
            # donde viene el número de crédito, aunque el resto del
            # documento salga con buena confianza (ver _ocr_hd_forzado). Se
            # prueban las 4 combinaciones del pase HD (no solo la de mayor
            # confianza general): en la práctica los DÍGITOS del número
            # salen estables en las 4, aunque la palabra "crédito" de la
            # etiqueta salga mal leída justo en la de mayor confianza.
            for texto_variante in _ocr_hd_forzado(ruta, indice_pagina=0, todas_las_variantes=True):
                if not texto_variante:
                    continue
                numero_credito = extraer_numero_credito(texto_variante, admite_respaldo_fonacot=True)
                if numero_credito:
                    break
        fila["numero_credito"] = numero_credito
        if numero_credito:
            detalles_extra.append(f"Número de crédito Fonacot detectado: {numero_credito}.")
        else:
            detalles_extra.append("No se detectó el número de crédito Fonacot; revisar manualmente.")

    elif clave == "INFONAVIT":
        numero_credito = extraer_numero_credito(texto_completo, etiqueta_seccion="INFORMACION DEL CREDITO DEL TRABAJADOR")
        if not numero_credito:
            texto_hd = _ocr_hd_forzado(ruta, indice_pagina=0)
            if texto_hd:
                numero_credito = extraer_numero_credito(texto_hd, etiqueta_seccion="INFORMACION DEL CREDITO DEL TRABAJADOR")
        fila["numero_credito"] = numero_credito
        if numero_credito:
            detalles_extra.append(f"Número de crédito Infonavit detectado: {numero_credito}.")
        else:
            detalles_extra.append(
                "No se detectó el número de crédito bajo 'Información del crédito del trabajador'; revisar manualmente."
            )

    elif clave == "INE":
        # La vigencia SIEMPRE se revisa en la cara frontal (página 1) contra
        # la fecha de hoy; que el PDF traiga ambas caras solo se checa por
        # separado (num_paginas_ok) para no mezclar los dos criterios.
        r = analiza_ine(paginas_texto, ruta=ruta)
        fila["vigencia_ok"] = r["vigente"]
        fila["vigencia_fecha_texto"] = f"Vigente hasta {r['vigencia_anio_fin']}" if r["vigencia_anio_fin"] else None
        fila["num_paginas_ok"] = r["dos_paginas"]
        detalles_extra.append(r["observaciones"])
        direccion_ine = _direccion_ine(texto_completo)
        fila["direccion"] = direccion_ine
        fila["domicilio"] = descomponer_direccion(direccion_ine)
        if direccion_ine:
            detalles_extra.append(f"Dirección detectada en el INE (revisar contra el documento): {direccion_ine}")
        else:
            detalles_extra.append("No se detectó la etiqueta 'DOMICILIO' en el INE; revisar manualmente.")

    elif clave == "COMPROBANTE_DOMICILIO":
        dentro, obs, fecha = analiza_comprobante_domicilio(texto_completo)
        fila["vigencia_ok"] = dentro
        fila["vigencia_fecha_texto"] = fecha.isoformat() if fecha else None
        detalles_extra.append(obs)
        direccion_domicilio = _direccion_comprobante(texto_completo)
        fila["direccion"] = direccion_domicilio
        fila["domicilio"] = descomponer_direccion(direccion_domicilio)
        if direccion_domicilio:
            detalles_extra.append(f"Dirección detectada en el comprobante de domicilio (revisar contra el documento): {direccion_domicilio}")
        else:
            detalles_extra.append("No se detectó una etiqueta de domicilio reconocible en el comprobante; revisar manualmente.")

    elif clave == "CUENTA_BANCARIA":
        r = analiza_cuenta_bancaria(texto_completo, nombre_candidato)
        fila["banco"] = r["banco"]
        fila["banco_rechazado"] = r["banco_rechazado"]
        fila["clabe"] = r["clabe_detectada"]
        fila["numero_cuenta"] = r["numero_cuenta"]
        detalles_extra.append(r["observaciones"])
        if num_paginas > 2:
            detalles_extra.append(
                f"El archivo trae {num_paginas} páginas (estado de cuenta completo). "
                "Se recomienda pedir solo la carátula (1 página) para no exponer el detalle de movimientos."
            )

    elif clave == "ACTA_NACIMIENTO":
        dentro, obs, fecha = analiza_fecha_limite(texto_completo, 365 * 5, "el acta de nacimiento")
        fila["vigencia_ok"] = dentro
        fila["vigencia_fecha_texto"] = fecha.isoformat() if fecha else None
        detalles_extra.append(obs + " (regla de 5 años detectada como práctica común; confirmar si aplica formalmente).")

        fecha_nac = busca_fecha_nacimiento(texto_completo)
        fila["fecha_nacimiento"] = fecha_nac
        fila["fecha_nacimiento_texto"] = fecha_nac.strftime("%d/%m/%Y") if fecha_nac else None
        if fecha_nac:
            detalles_extra.append(f"Fecha de nacimiento detectada: {fecha_nac.strftime('%d/%m/%Y')}.")
        else:
            detalles_extra.append(
                "No se detectó la fecha de nacimiento (etiqueta 'FECHA DE NACIMIENTO' no encontrada o sin "
                "fecha legible cerca); revisar manualmente."
            )

    fila["detalle"] += " ".join(detalles_extra)
    return fila


# ---------------------------------------------------------------------------
# Carga masiva por ZIP: extracción y procesamiento de un candidato
# ---------------------------------------------------------------------------

def _es_pdf_valido_en_zip(nombre_interno):
    """Filtra basura común de los ZIPs: carpetas, archivos ocultos
    (._archivo.pdf que crea macOS al comprimir) y la carpeta __MACOSX."""
    if nombre_interno.endswith("/"):
        return False
    base = os.path.basename(nombre_interno)
    if not base or base.startswith("."):
        return False
    if "__MACOSX" in nombre_interno:
        return False
    return base.lower().endswith(".pdf")


def listar_pdfs_en_zip(ruta_zip):
    """Regresa la lista de nombres internos (rutas dentro del ZIP) que son
    PDFs a procesar. Si el ZIP viene corrupto o no es un ZIP real, deja
    que zipfile.BadZipFile se propague — el caller lo captura para reportar
    ese candidato en particular como error, sin tronar el resto del lote."""
    with zipfile.ZipFile(ruta_zip) as zf:
        return [n for n in zf.namelist() if _es_pdf_valido_en_zip(n)]


def procesar_zip_candidato_progresivo(ruta_zip, nombre_candidato, dir_extraccion, nombres_pdf=None):
    """Generador: extrae y procesa cada PDF de un ZIP de un candidato.

    Cede (yield) el nombre del archivo justo antes de procesarlo —así quien
    lo consuma con un bucle 'for'/'next()' puede reportar avance en vivo
    (ver /validar_lote en app.py, que lo usa con 'yield from' dentro de su
    propio generador de eventos SSE)— y al terminar REGRESA (return) la
    lista de filas ya procesadas (recuperable como el .value de la
    StopIteration, o con 'resultado = yield from procesar_zip_candidato_progresivo(...)').

    La categoría de cada documento se determina primero por el NOMBRE del
    archivo (convención cv.pdf / ine.pdf / comprobante_domicilio.pdf / etc,
    ver detectar_categoria_por_nombre_archivo) y, si el nombre no da ninguna
    pista, se cae al reconocimiento por CONTENIDO de siempre (procesar_documento
    sin categoria_forzada)."""
    if nombres_pdf is None:
        nombres_pdf = listar_pdfs_en_zip(ruta_zip)
    filas = []
    with zipfile.ZipFile(ruta_zip) as zf:
        for nombre_interno in nombres_pdf:
            nombre_base = os.path.basename(nombre_interno)
            yield nombre_base
            try:
                ruta_extraida = zf.extract(nombre_interno, dir_extraccion)
                categoria = detectar_categoria_por_nombre_archivo(nombre_base)
                fila = procesar_documento(ruta_extraida, nombre_candidato, categoria_forzada=categoria)
            except Exception as e:
                fila = {
                    "archivo": nombre_base, "categoria_clave": "ERROR",
                    "categoria": "Error al procesar", "num_paginas": 0, "uso_ocr": False,
                    "legible": False, "texto_muestra": "", "nombre_coincide": None,
                    "detalle": f"Error: {e}",
                }
            filas.append(fila)
    return filas


def _concilia_fecha_nacimiento(filas):
    """Valida la fecha de nacimiento leída del acta contra la que lleva
    incrustada el CURP (AAMMDD). El OCR del acta confunde dígitos con
    facilidad (2008 -> 2003) y esa fecha alimenta la edad y "Datos
    completos", mientras que el CURP (documento o CSF) la trae de forma
    estructurada. Si difieren, se usa la del CURP y se deja un aviso en las
    observaciones del acta. Modifica 'filas' en sitio."""
    acta = next((f for f in filas if f["categoria_clave"] == "ACTA_NACIMIENTO"), None)
    if not acta or not acta.get("fecha_nacimiento"):
        return
    curp_fila = next((f for f in filas if f["categoria_clave"] == "CURP" and f.get("curp")), None)
    csf_fila = next((f for f in filas if f["categoria_clave"] == "CSF" and f.get("curp")), None)
    curp = (curp_fila or {}).get("curp") or (csf_fila or {}).get("curp")
    if not curp:
        return
    derivada = datos_derivados_de_curp(curp).get("fecha_nacimiento")
    if not derivada or derivada == acta["fecha_nacimiento"]:
        return
    leida = acta.get("fecha_nacimiento_texto")
    acta["fecha_nacimiento"] = derivada
    acta["fecha_nacimiento_texto"] = derivada.strftime("%d/%m/%Y")
    acta["detalle"] = (acta.get("detalle") or "") + (
        f" AVISO: la fecha de nacimiento leída del acta ({leida}) no coincide con la que trae el CURP "
        f"({acta['fecha_nacimiento_texto']}); se usó la del CURP (el OCR del acta pudo confundir un dígito). "
        "Revisar contra el PDF."
    )


def construir_reporte(candidato, filas):
    _concilia_fecha_nacimiento(filas)
    encontrados_por_categoria = {}
    for fila in filas:
        encontrados_por_categoria.setdefault(fila["categoria_clave"], []).append(fila)

    checklist = []
    for clave in CHECKLIST_OBLIGATORIO:
        nombre_legible, _, obligatorio = CATEGORIAS[clave]
        docs = encontrados_por_categoria.get(clave, [])
        checklist.append({
            "clave": clave,
            "categoria": nombre_legible,
            "obligatorio": obligatorio,
            "recibido": len(docs) > 0,
            "archivos": [d["archivo"] for d in docs],
            # archivos cargados aquí cuyo contenido parece de otra categoría
            "revisar": [d["archivo"] for d in docs if d.get("categoria_sugerida")],
        })

    extra = []
    for clave, docs in encontrados_por_categoria.items():
        if clave not in CHECKLIST_OBLIGATORIO:
            nombre_legible, _, _ = CATEGORIAS.get(clave, ("Desconocido / no identificado", [], False))
            extra.append({"categoria": nombre_legible, "archivos": [d["archivo"] for d in docs]})

    return checklist, extra


# ---------------------------------------------------------------------------
# Salida a Excel
# ---------------------------------------------------------------------------

VERDE = PatternFill("solid", fgColor="C6EFCE")
ROJO = PatternFill("solid", fgColor="FFC7CE")
AMARILLO = PatternFill("solid", fgColor="FFEB9C")
GRIS_ENCABEZADO = PatternFill("solid", fgColor="1F6B4C")
FUENTE_ENCABEZADO = Font(color="FFFFFF", bold=True)


def _set_encabezados(ws, encabezados, fila=1):
    for col, texto in enumerate(encabezados, start=1):
        c = ws.cell(row=fila, column=col, value=texto)
        c.font = FUENTE_ENCABEZADO
        c.fill = GRIS_ENCABEZADO
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[fila].height = 30


def _autoancho(ws, anchos):
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho


def _llenar_hoja_resumen(ws, candidato, filas, checklist, extra, fila_inicio=1):
    """Escribe el bloque de resumen (datos del candidato, completitud,
    checklist, faltantes, vigencias, datos bancarios) empezando en
    fila_inicio. Regresa la siguiente fila libre después de todo el bloque
    -así generar_excel_lote puede seguir escribiendo el detalle por
    documento justo abajo, en la MISMA hoja, para tener un solo tab por
    candidato-. Usado tanto por generar_excel (un candidato, hoja propia)
    como por generar_excel_lote (varios candidatos, una hoja por cada uno)."""
    r0 = fila_inicio
    ws.cell(row=r0, column=1, value="Expediente de reclutamiento — validación automática").font = Font(bold=True, size=14)
    ws.cell(row=r0 + 1, column=1, value=f"Candidato: {candidato['nombre']}")
    ws.cell(row=r0 + 2, column=1, value=(
        f"E-mail: {candidato.get('email') or '—'}    Nacionalidad: {candidato.get('nacionalidad') or '—'}    "
        f"Estado civil: {candidato.get('estado_civil') or '—'}    Teléfono: {candidato.get('telefono') or '—'}"
    ))
    ws.cell(row=r0 + 3, column=1, value=f"Generado: {HOY.isoformat()}")

    faltantes = [c for c in checklist if c["obligatorio"] and not c["recibido"]]
    completo = len(faltantes) == 0
    ws.cell(row=r0 + 5, column=1, value="Estado de completitud:").font = Font(bold=True)
    c_estado_general = ws.cell(row=r0 + 5, column=2, value="COMPLETO" if completo else f"INCOMPLETO — faltan {len(faltantes)} documento(s)")
    c_estado_general.fill = VERDE if completo else ROJO
    c_estado_general.font = Font(bold=True)

    _set_encabezados(ws, ["Documento", "Obligatorio", "¿Recibido?", "Archivo(s)"], fila=r0 + 7)
    r = r0 + 8
    for item in checklist:
        ws.cell(row=r, column=1, value=item["categoria"])
        ws.cell(row=r, column=2, value="Sí" if item["obligatorio"] else "Condicional (solo si aplica)")
        if item.get("revisar"):
            celda_recibido = ws.cell(row=r, column=3, value="Sí — REVISAR (parece otro documento)")
            celda_recibido.fill = AMARILLO
        else:
            celda_recibido = ws.cell(row=r, column=3, value="Sí" if item["recibido"] else "No")
            if item["obligatorio"]:
                celda_recibido.fill = VERDE if item["recibido"] else ROJO
            else:
                celda_recibido.fill = VERDE if item["recibido"] else AMARILLO
        ws.cell(row=r, column=4, value=", ".join(item["archivos"]) or "—")
        r += 1

    r += 2
    ws.cell(row=r, column=1, value="Documentos faltantes (obligatorios):").font = Font(bold=True)
    r += 1
    if faltantes:
        for item in faltantes:
            c = ws.cell(row=r, column=1, value="• " + item["categoria"])
            c.fill = ROJO
            r += 1
    else:
        ws.cell(row=r, column=1, value="Ninguno — la documentación obligatoria está completa.")
        r += 1

    condicionales_faltantes = [c for c in checklist if (not c["obligatorio"]) and not c["recibido"]]
    r += 1
    ws.cell(row=r, column=1, value="Documentos condicionales NO recibidos (confirmar con el candidato si le aplican):").font = Font(bold=True)
    r += 1
    if condicionales_faltantes:
        for item in condicionales_faltantes:
            c = ws.cell(row=r, column=1, value="• " + item["categoria"])
            c.fill = AMARILLO
            r += 1
    else:
        ws.cell(row=r, column=1, value="Ninguno pendiente.")
        r += 1

    mal_cargados = [f for f in filas if f.get("categoria_sugerida")]
    if mal_cargados:
        r += 1
        ws.cell(row=r, column=1, value="⚠ Posibles documentos cargados en el campo equivocado (revisar el PDF):").font = Font(bold=True)
        r += 1
        for f in mal_cargados:
            c = ws.cell(row=r, column=1, value=f"• {f['archivo']}: cargado como «{f['categoria']}», parece «{f['categoria_sugerida_nombre']}»")
            c.fill = AMARILLO
            r += 1

    if extra:
        r += 1
        ws.cell(row=r, column=1, value="Documentos adicionales recibidos (no están en la lista de arriba):").font = Font(bold=True, italic=True)
        r += 1
        for item in extra:
            ws.cell(row=r, column=1, value=item["categoria"])
            ws.cell(row=r, column=4, value=", ".join(item["archivos"]))
            r += 1

    # --- Vigencias verificadas (INE cara frontal, comprobante de domicilio,
    # CSF, acta) contra la fecha de hoy. Se listan aunque no se haya podido
    # leer la fecha, para que quede visible qué falta revisar a mano. -------
    DOCS_CON_REGLA_VIGENCIA = {"INE", "COMPROBANTE_DOMICILIO", "CSF", "ACTA_NACIMIENTO"}
    con_vigencia = [f for f in filas if f["categoria_clave"] in DOCS_CON_REGLA_VIGENCIA]
    r += 2
    ws.cell(row=r, column=1, value=f"Vigencias verificadas contra hoy ({HOY.isoformat()}):").font = Font(bold=True)
    r += 1
    if con_vigencia:
        _set_encabezados(ws, ["Documento", "Fecha / vigencia detectada", "Estado", "Archivo"], fila=r)
        r += 1
        for f in con_vigencia:
            ws.cell(row=r, column=1, value=f["categoria"])
            ws.cell(row=r, column=2, value=f.get("vigencia_fecha_texto") or "No se detectó")
            vig = f.get("vigencia_ok")
            c_estado = ws.cell(row=r, column=3, value="OK" if vig else "FUERA DE VIGENCIA — revisar" if vig is False else "Sin fecha detectada — revisar a mano")
            c_estado.fill = VERDE if vig else ROJO if vig is False else AMARILLO
            ws.cell(row=r, column=4, value=f["archivo"])
            r += 1
    else:
        ws.cell(row=r, column=1, value="No se recibieron documentos con regla de vigencia.")
        r += 1

    # --- Datos extraídos de los documentos (RFC de la CSF, CURP del
    # documento CURP, fecha de nacimiento del acta, dirección del
    # comprobante de domicilio y del INE por separado). Igual que los datos
    # bancarios de abajo, son "propuesta a confirmar": se muestran para que
    # el equipo de reclutamiento no tenga que abrir cada PDF, pero siguen
    # marcados como pendientes de revisar contra el documento original. ---
    csf_fila = next((f for f in filas if f["categoria_clave"] == "CSF"), None)
    curp_fila = next((f for f in filas if f["categoria_clave"] == "CURP"), None)
    acta_fila = next((f for f in filas if f["categoria_clave"] == "ACTA_NACIMIENTO"), None)
    domicilio_fila = next((f for f in filas if f["categoria_clave"] == "COMPROBANTE_DOMICILIO"), None)
    ine_fila = next((f for f in filas if f["categoria_clave"] == "INE"), None)
    infonavit_fila = next((f for f in filas if f["categoria_clave"] == "INFONAVIT"), None)
    fonacot_fila = next((f for f in filas if f["categoria_clave"] == "FONACOT"), None)
    nss_fila = next((f for f in filas if f["categoria_clave"] == "NSS"), None)

    r += 2
    ws.cell(row=r, column=1, value="Datos extraídos de los documentos (propuesta a confirmar contra el PDF):").font = Font(bold=True)
    r += 1
    _set_encabezados(ws, ["Dato", "Valor detectado", "Fuente"], fila=r)
    r += 1
    def _dato(fila_doc, campo):
        """(valor, sin_documento) — distingue "no se recibió el documento
        fuente" (condicionales como Infonavit/Fonacot que pueden no aplicar)
        de "se recibió pero no se pudo extraer el dato"."""
        if fila_doc is None:
            return None, True
        return fila_doc.get(campo), False

    datos_extraidos = [
        ("RFC", *_dato(csf_fila, "rfc"), "CSF"),
        ("CURP", *_dato(curp_fila, "curp"), "Documento CURP"),
        ("CURP (CSF)", *_dato(csf_fila, "curp"), "CSF"),
        ("Código Postal (CSF)", *_dato(csf_fila, "codigo_postal"), "CSF"),
        ("Dirección fiscal (CSF)", *_dato(csf_fila, "direccion_fiscal"), "CSF"),
        ("NSS (Número de Seguridad Social)", *_dato(nss_fila, "numero_seguridad_social"), "Asignación de NSS (IMSS)"),
        ("Régimen (CSF, hoja 2)", *_dato(csf_fila, "regimen"), "CSF"),
        ("Fecha de nacimiento", *_dato(acta_fila, "fecha_nacimiento_texto"), "Acta de nacimiento"),
        ("Dirección", *_dato(domicilio_fila, "direccion"), "Comprobante de domicilio"),
        ("Dirección (INE)", *_dato(ine_fila, "direccion"), "INE"),
        ("Número de crédito Infonavit", *_dato(infonavit_fila, "numero_credito"), "Infonavit"),
        ("Número de crédito Fonacot", *_dato(fonacot_fila, "numero_credito"), "Fonacot"),
    ]
    for etiqueta, valor, sin_documento, fuente in datos_extraidos:
        ws.cell(row=r, column=1, value=etiqueta)
        if sin_documento:
            ws.cell(row=r, column=2, value="No se recibió el documento")
        else:
            c_valor = ws.cell(row=r, column=2, value=valor or "No detectado")
            c_valor.number_format = "@"  # forzar texto: RFC/CURP no deben interpretarse como número
            if not valor:
                c_valor.fill = AMARILLO
        ws.cell(row=r, column=3, value=fuente)
        r += 1

    # --- Datos bancarios de la carátula, exportados como texto para que no
    # se corrompan en Excel (CLABE/cuenta con ceros a la izquierda, etc.) --
    bancarios = [f for f in filas if f["categoria_clave"] == "CUENTA_BANCARIA"]
    r += 2
    ws.cell(row=r, column=1, value="Datos bancarios (carátula) — CLABE y cuenta exportados como texto:").font = Font(bold=True)
    r += 1
    if bancarios:
        _set_encabezados(ws, ["Archivo", "Banco", "Número de cuenta", "CLABE"], fila=r)
        r += 1
        for f in bancarios:
            ws.cell(row=r, column=1, value=f["archivo"])
            ws.cell(row=r, column=2, value=f.get("banco") or "No identificado")
            c_cuenta = ws.cell(row=r, column=3, value=f.get("numero_cuenta") or "No detectado")
            c_cuenta.number_format = "@"  # forzar texto: evita notación científica o pérdida de ceros
            c_clabe = ws.cell(row=r, column=4, value=f.get("clabe") or "No detectada")
            c_clabe.number_format = "@"
            r += 1
    else:
        ws.cell(row=r, column=1, value="No se recibió carátula bancaria.")
        r += 1

    _autoancho(ws, [42, 26, 20, 45])
    return r + 1


def _llenar_hoja_detalle(ws, filas, fila_inicio=1):
    """Escribe la tabla de detalle por documento (una fila por archivo)
    empezando en fila_inicio. Regresa la siguiente fila libre. Igual que
    _llenar_hoja_resumen, la usan tanto generar_excel (hoja propia) como
    generar_excel_lote (debajo del resumen, en la misma hoja del candidato)."""
    encabezados2 = [
        "Archivo", "Documento identificado", "Páginas", "¿Usó OCR?", "Legible",
        "Nombre coincide", "Vigencia OK", "Fecha/vigencia detectada", "Banco (si aplica)",
        "Número de cuenta", "CLABE", "Observaciones",
    ]
    _set_encabezados(ws, encabezados2, fila=fila_inicio)
    r = fila_inicio + 1
    fila_primera_dato = r
    for fila in filas:
        ws.cell(row=r, column=1, value=fila["archivo"])
        ws.cell(row=r, column=2, value=fila["categoria"])
        ws.cell(row=r, column=3, value=fila["num_paginas"])
        ws.cell(row=r, column=4, value="Sí" if fila["uso_ocr"] else "No")

        c_leg = ws.cell(row=r, column=5, value="Sí" if fila["legible"] else "NO — revisar")
        c_leg.fill = VERDE if fila["legible"] else ROJO

        nc = fila.get("nombre_coincide")
        if fila["categoria_clave"] == "COMPROBANTE_DOMICILIO":
            c_nom = ws.cell(row=r, column=6, value="N/A (puede ser otra persona)")
            c_nom.fill = AMARILLO
        else:
            texto_nc = "Sí" if nc is True else "NO — revisar" if nc is False else "Dudoso — revisar"
            c_nom = ws.cell(row=r, column=6, value=texto_nc)
            c_nom.fill = VERDE if nc is True else ROJO if nc is False else AMARILLO

        vig = fila.get("vigencia_ok")
        if vig is None:
            c_vig = ws.cell(row=r, column=7, value="—")
        else:
            c_vig = ws.cell(row=r, column=7, value="OK" if vig else "FUERA DE RANGO — revisar")
            c_vig.fill = VERDE if vig else ROJO

        ws.cell(row=r, column=8, value=fila.get("vigencia_fecha_texto") or "—")

        banco = fila.get("banco")
        c_banco = ws.cell(row=r, column=9, value=banco or "—")
        if fila.get("banco_rechazado"):
            c_banco.fill = ROJO

        # CLABE y número de cuenta se exportan como TEXTO (number_format "@")
        # para que Excel no los convierta a notación científica ni les
        # quite ceros a la izquierda.
        c_cta = ws.cell(row=r, column=10, value=fila.get("numero_cuenta") or "—")
        c_cta.number_format = "@"
        c_clabe = ws.cell(row=r, column=11, value=fila.get("clabe") or "—")
        c_clabe.number_format = "@"

        ws.cell(row=r, column=12, value=fila["detalle"] + (" | Muestra: " + fila["texto_muestra"] if fila["legible"] else ""))
        ws.cell(row=r, column=12).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    for fila_idx in range(fila_primera_dato, r):
        ws.row_dimensions[fila_idx].height = 45

    return r


# ---------------------------------------------------------------------------
# Tercera hoja: "Datos completos" — misma plantilla de alta en nómina/RH que
# usa Fitness Para Todos (columnas y notas de origen tomadas tal cual de su
# archivo de plantilla), precargada con lo que ya se puede derivar de los
# documentos y del formulario. NO sustituye el alta real: los campos
# operativos (puesto, club, salario, fechas de contrato, etc.) los sigue
# llenando RH a mano, porque esos datos no vienen en ningún documento del
# candidato ni en este formulario.
# ---------------------------------------------------------------------------

# (encabezado, nota de origen tal como la trae la plantilla de RH, o None si
# esa columna no traía nota). Se excluyen de la plantilla original una
# columna en blanco y unos valores de referencia sueltos al final (UMA, IMSS)
# que no son campos por candidato.
CAMPOS_DATOS_COMPLETOS = [
    ("Estatus", None),
    ("Número de empleado", None),
    ("Nombres", None),
    ("Apellido paterno", None),
    ("Apellido materno", None),
    ("Nombre completo en Worky", None),
    ("CURP", "Se obtiene del CURP"),
    ("RFC", "Se obtiene de CSF"),
    ("Número seguridad social", "Se obtiene de Asignación de IMSS"),
    ("Nombre fiscal", "Se obtiene de la constancia de situación fiscal"),
    ("RETENCIÓN INFONAVIT", "se obtiene del aviso de retencion INFONAVIT"),
    ("RETENCIÓN FONACOT", "se obtiene del aviso de retencion FONACOT"),
    ("Régimen fiscal", "Se obtiene de la constancia de situación fiscal"),
    ("C.P. fiscal", "Se obtiene de la constancia de situación fiscal"),
    ("Género", "Se obtiene de CURP"),
    ("Estado de nacimiento", "Se obtiene de CURP"),
    ("Nacionalidad", "Se obtiene de CURP"),
    ("Fecha de nacimiento", None),
    ("Edad", None),
    ("Estado civil", None),
    ("Correo personal", None),
    ("Correo corporativo", None),
    ("Teléfono", None),
    ("Calle", "Se obtiene del comprobante de domicilio"),
    ("Número exterior", None),
    ("Número interior", None),
    ("CP", None),
    ("Estado", None),
    ("Municipio", None),
    ("Colonia", None),
    ("País", None),
    ("DIRECCION FISCAL", "Se obtiene de CSF"),
    ("Método de pago", None),
    ("Banco", "Se obtiene de estado de cuenta"),
    ("Número de cuenta", None),
    ("Número de cuenta Clabe", None),
    ("Nombre de la empresa", None),
    ("REGISTRO PATRONAL", None),
    ("CI departamento", None),
    ("Nombre departamento", None),
    ("Nombre puesto", None),
    ("DESCRIPTOR DE PUESTO", None),
    ("FECHA DE INICIO (LETRA)", None),
    ("Fecha de alta", None),
    ("Fecha de antigüedad", None),
    ("FECHA DE TERMINO", None),
    ("FECHA DE TERMINO (2)", None),
    ("Tipo de contrato", None),
    ("Tipo de periodo", None),
    ("Salario mensual", None),
    ("Salario diario", None),
    ("LETRA", None),
    ("Factor de integración", None),
    ("Salario integrado fijo", None),
    ("Variable diaria", None),
    ("Salario base cotización", None),
    ("Días de aguinaldo", None),
    ("Días de vacaciones", None),
    ("Prima vacacional", None),
    ("HORAS", None),
    ("TIPO", None),
    ("JORNADA", None),
    ("UMA", None),
]

# Campos que deben exportarse como texto (number_format "@") para que Excel
# no les quite ceros a la izquierda ni los pase a notación científica.
_CAMPOS_DATOS_COMPLETOS_TEXTO = {
    "CURP", "RFC", "Número seguridad social", "RETENCIÓN INFONAVIT", "RETENCIÓN FONACOT",
    "C.P. fiscal", "CP", "Número de cuenta", "Número de cuenta Clabe", "Teléfono",
}


def _valores_datos_completos(candidato, filas):
    """Arma {encabezado: (valor, es_propuesta)} con lo que se puede derivar
    de los documentos ya procesados y de lo capturado en el formulario. Un
    valor None se deja en blanco -es de RH completarlo a mano, no es un
    error-. es_propuesta=True resalta la celda en amarillo (dato de OCR o
    derivado, igual que el resto del Excel: "a confirmar contra el
    documento"; False se usa para lo que el propio candidato/reclutador
    capturó tal cual en el formulario, que no necesita ese resaltado."""
    csf_fila = next((f for f in filas if f["categoria_clave"] == "CSF"), None)
    curp_fila = next((f for f in filas if f["categoria_clave"] == "CURP"), None)
    acta_fila = next((f for f in filas if f["categoria_clave"] == "ACTA_NACIMIENTO"), None)
    domicilio_fila = next((f for f in filas if f["categoria_clave"] == "COMPROBANTE_DOMICILIO"), None)
    ine_fila = next((f for f in filas if f["categoria_clave"] == "INE"), None)
    infonavit_fila = next((f for f in filas if f["categoria_clave"] == "INFONAVIT"), None)
    fonacot_fila = next((f for f in filas if f["categoria_clave"] == "FONACOT"), None)
    cuenta_fila = next((f for f in filas if f["categoria_clave"] == "CUENTA_BANCARIA"), None)
    nss_fila = next((f for f in filas if f["categoria_clave"] == "NSS"), None)

    curp = (curp_fila or {}).get("curp") or (csf_fila or {}).get("curp")
    derivados = datos_derivados_de_curp(curp) if curp else {}

    # Fecha de nacimiento: se prefiere la del acta de nacimiento (documento
    # oficial, ya extraída con su propia etiqueta) y solo si no se recibió
    # o no se pudo leer, se usa la que se deriva del CURP (AAMMDD + regla
    # del diferenciador para el siglo, ver datos_derivados_de_curp). La
    # edad se calcula igual en los dos casos, con una resta simple contra
    # hoy.
    fecha_nac = (acta_fila or {}).get("fecha_nacimiento")
    fecha_nac_texto = (acta_fila or {}).get("fecha_nacimiento_texto")
    if not fecha_nac and derivados.get("fecha_nacimiento"):
        fecha_nac = derivados["fecha_nacimiento"]
        fecha_nac_texto = fecha_nac.strftime("%d/%m/%Y")
    edad = None
    if fecha_nac:
        edad = HOY.year - fecha_nac.year - ((HOY.month, HOY.day) < (fecha_nac.month, fecha_nac.day))

    # "DIRECCION FISCAL" es el domicilio registrado ante el SAT (de la CSF),
    # no la dirección particular del candidato -son datos distintos, aunque
    # a veces coincidan-. Las columnas Calle/Número/Colonia/Municipio/
    # Estado/CP/País, en cambio, sí son la dirección particular: se prefiere
    # la del INE (más clara y homogénea) y, si no se recibió o no se pudo
    # reconocer nada en ella, se cae de vuelta a la del comprobante de
    # domicilio. Ver descomponer_direccion() para las limitaciones de este
    # reconocimiento (siempre una propuesta a confirmar).
    domicilio_ine = (ine_fila or {}).get("domicilio") or {}
    domicilio_comprobante = (domicilio_fila or {}).get("domicilio") or {}
    domicilio_particular = domicilio_ine if any(domicilio_ine.values()) else domicilio_comprobante
    nacionalidad_manual = candidato.get("nacionalidad")

    # Nombre(s)/Apellidos: se prefieren los que trae la CSF (documento
    # oficial, ya separados en sus 3 campos) para armar el "Nombre completo
    # en Worky" uniéndolos en ese orden (Nombres + Apellido paterno +
    # Apellido materno); si la CSF no se recibió o no se pudo leer ese
    # bloque, se cae de vuelta al nombre tal cual lo capturó el candidato en
    # el formulario -que es lo que hacía este campo antes de tener la CSF
    # separada en Nombres/Apellidos-.
    nombres_csf = (csf_fila or {}).get("nombres")
    apellido_paterno_csf = (csf_fila or {}).get("apellido_paterno")
    apellido_materno_csf = (csf_fila or {}).get("apellido_materno")
    if nombres_csf and apellido_paterno_csf:
        nombre_worky = " ".join(p for p in (nombres_csf, apellido_paterno_csf, apellido_materno_csf) if p)
        nombre_worky_es_propuesta = True
    else:
        nombre_worky = candidato.get("nombre")
        nombre_worky_es_propuesta = False

    v = {}
    v["Nombres"] = (nombres_csf, True)
    v["Apellido paterno"] = (apellido_paterno_csf, True)
    v["Apellido materno"] = (apellido_materno_csf, True)
    v["Nombre completo en Worky"] = (nombre_worky, nombre_worky_es_propuesta)
    v["Número seguridad social"] = ((nss_fila or {}).get("numero_seguridad_social"), True)
    v["CURP"] = (curp, True)
    v["RFC"] = ((csf_fila or {}).get("rfc") or candidato.get("rfc") or None, True)
    v["Nombre fiscal"] = (candidato.get("nombre"), True)
    v["RETENCIÓN INFONAVIT"] = ((infonavit_fila or {}).get("numero_credito"), True)
    v["RETENCIÓN FONACOT"] = ((fonacot_fila or {}).get("numero_credito"), True)
    v["Régimen fiscal"] = ((csf_fila or {}).get("regimen"), True)
    v["C.P. fiscal"] = ((csf_fila or {}).get("codigo_postal"), True)
    v["Género"] = (derivados.get("genero"), True)
    v["Estado de nacimiento"] = (derivados.get("estado_nacimiento"), True)
    v["Nacionalidad"] = (nacionalidad_manual or derivados.get("nacionalidad"), not nacionalidad_manual)
    v["Fecha de nacimiento"] = (fecha_nac_texto, True)
    v["Edad"] = (edad, True)
    v["Estado civil"] = (candidato.get("estado_civil"), False)
    v["Correo personal"] = (candidato.get("email"), False)
    v["Teléfono"] = (candidato.get("telefono"), False)
    v["Calle"] = (domicilio_particular.get("calle"), True)
    v["Número exterior"] = (domicilio_particular.get("numero_exterior"), True)
    v["Número interior"] = (domicilio_particular.get("numero_interior"), True)
    v["CP"] = (domicilio_particular.get("codigo_postal"), True)
    v["Estado"] = (domicilio_particular.get("estado"), True)
    v["Municipio"] = (domicilio_particular.get("municipio"), True)
    v["Colonia"] = (domicilio_particular.get("colonia"), True)
    v["País"] = ("México" if (ine_fila or domicilio_fila) else None, True)
    v["DIRECCION FISCAL"] = ((csf_fila or {}).get("direccion_fiscal"), True)
    v["Banco"] = ((cuenta_fila or {}).get("banco"), True)
    v["Número de cuenta"] = ((cuenta_fila or {}).get("numero_cuenta"), True)
    v["Número de cuenta Clabe"] = ((cuenta_fila or {}).get("clabe"), True)
    v["Nombre de la empresa"] = ("FITNESS PARA TODOS S. DE R.L. DE C.V.", True)
    v["Método de pago"] = ("Transferencia Electrónica", True)
    return v


def _llenar_hoja_datos_completos(ws, candidato, filas, fila_inicio=1, ajustar_anchos=True):
    """Escribe la hoja/bloque "Datos completos": encabezados de la plantilla
    de RH (fila 1), su nota de origen tal cual la trae esa plantilla (fila
    2, de referencia), y los valores propuestos para este candidato (fila
    3). Regresa la siguiente fila libre, igual que las otras _llenar_hoja_*,
    para poder reutilizarse tanto en el Excel individual (hoja propia) como
    debajo del detalle en el Excel de lote (misma hoja del candidato)."""
    r = fila_inicio
    ws.cell(row=r, column=1, value=(
        "Datos completos (mismos campos que la plantilla de alta en nómina/RH) — "
        "propuesta a confirmar, no sustituye la revisión de RH."
    )).font = Font(bold=True)
    r += 2

    _set_encabezados(ws, [c[0] for c in CAMPOS_DATOS_COMPLETOS], fila=r)
    fila_encabezados = r
    r += 1

    for col, (_campo, nota) in enumerate(CAMPOS_DATOS_COMPLETOS, start=1):
        if nota:
            c = ws.cell(row=r, column=col, value=nota)
            c.font = Font(italic=True, size=9, color="666666")
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 28
    r += 1

    valores = _valores_datos_completos(candidato, filas)
    for col, (campo, _nota) in enumerate(CAMPOS_DATOS_COMPLETOS, start=1):
        valor, es_propuesta = valores.get(campo, (None, False))
        if valor in (None, ""):
            continue
        c = ws.cell(row=r, column=col, value=valor)
        if campo in _CAMPOS_DATOS_COMPLETOS_TEXTO:
            c.number_format = "@"
        if es_propuesta:
            c.fill = AMARILLO
    r += 1

    if ajustar_anchos:
        _autoancho(ws, [18] * len(CAMPOS_DATOS_COMPLETOS))
    return r


def generar_excel(candidato, filas, checklist, extra, ruta_salida):
    """Un solo candidato -> 3 hojas (Resumen + Detalle por documento + Datos
    completos). Es el mismo formato de siempre; el único cambio interno es
    que ahora reutiliza _llenar_hoja_resumen / _llenar_hoja_detalle /
    _llenar_hoja_datos_completos, que también usa generar_excel_lote para la
    carga masiva."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    _llenar_hoja_resumen(ws, candidato, filas, checklist, extra, fila_inicio=1)
    _autoancho(ws, [42, 26, 20, 45])

    ws2 = wb.create_sheet("Detalle por documento")
    _llenar_hoja_detalle(ws2, filas, fila_inicio=1)
    _autoancho(ws2, [30, 26, 9, 10, 12, 20, 12, 24, 22, 18, 20, 70])

    ws3 = wb.create_sheet("Datos completos")
    _llenar_hoja_datos_completos(ws3, candidato, filas, fila_inicio=1)

    wb.save(ruta_salida)


# ---------------------------------------------------------------------------
# Carga masiva por ZIP: un Excel con una hoja por candidato
# ---------------------------------------------------------------------------

_CARACTERES_INVALIDOS_HOJA = re.compile(r"[\\/*?:\[\]]")


def _nombre_hoja_excel_seguro(nombre_candidato, nombres_usados):
    """Excel no permite \\ / * ? : [ ] en el nombre de una hoja, ni más de
    31 caracteres, ni dos hojas con el mismo nombre. Aquí se limpia el
    nombre del candidato para que sirva como nombre de hoja y, si ya existe
    (dos candidatos con el mismo nombre en el mismo lote, o un nombre que
    se trunca igual), se le agrega un sufijo numérico."""
    base = _CARACTERES_INVALIDOS_HOJA.sub(" ", nombre_candidato or "Candidato").strip()
    base = re.sub(r"\s+", " ", base) or "Candidato"
    base = base[:31]
    candidato_nombre_hoja = base
    sufijo = 2
    while candidato_nombre_hoja.lower() in nombres_usados:
        sufijo_txt = f" ({sufijo})"
        candidato_nombre_hoja = base[: 31 - len(sufijo_txt)] + sufijo_txt
        sufijo += 1
    nombres_usados.add(candidato_nombre_hoja.lower())
    return candidato_nombre_hoja


def _llenar_resumen_general(ws, resultados, nombres_hoja):
    """Hoja de portada del Excel de lote: una fila por candidato con su
    estado de completitud y un enlace directo a su hoja de detalle —así
    quien revise el lote no tiene que ir abriendo hoja por hoja para saber
    cuáles candidatos ya están completos."""
    ws.cell(row=1, column=1, value="Carga masiva de expedientes — validación automática").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value=f"Generado: {HOY.isoformat()}")
    ws.cell(row=3, column=1, value=f"Candidatos procesados: {len(resultados)}")

    _set_encabezados(ws, ["#", "Candidato", "Estado", "Documentos faltantes (obligatorios)", "Ir al detalle"], fila=5)
    r = 6
    for i, (res, nombre_hoja) in enumerate(zip(resultados, nombres_hoja), start=1):
        checklist = res["checklist"]
        faltantes = [c for c in checklist if c["obligatorio"] and not c["recibido"]]
        completo = len(faltantes) == 0
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=res["candidato"]["nombre"])
        c_estado = ws.cell(row=r, column=3, value="COMPLETO" if completo else f"INCOMPLETO — faltan {len(faltantes)}")
        c_estado.fill = VERDE if completo else ROJO
        c_estado.font = Font(bold=True)
        ws.cell(row=r, column=4, value=", ".join(c["categoria"] for c in faltantes) or "—")
        c_link = ws.cell(row=r, column=5, value="Ver detalle")
        c_link.hyperlink = f"#'{nombre_hoja}'!A1"
        c_link.font = Font(color="1155CC", underline="single")
        r += 1

    _autoancho(ws, [4, 32, 26, 60, 16])


def generar_excel_lote(resultados, ruta_salida):
    """Genera UN Excel para varios candidatos: una hoja 'Resumen general'
    (portada con enlaces) + una hoja POR CANDIDATO con su resumen y su
    detalle por documento combinados (para que sea un solo tab por
    empleado, no dos). 'resultados' es una lista de dicts:
    {"candidato": {...}, "filas": [...], "checklist": [...], "extra": [...]}."""
    wb = Workbook()
    ws_portada = wb.active
    ws_portada.title = "Resumen general"

    nombres_usados = set()
    nombres_hoja = [_nombre_hoja_excel_seguro(res["candidato"]["nombre"], nombres_usados) for res in resultados]

    _llenar_resumen_general(ws_portada, resultados, nombres_hoja)

    for res, nombre_hoja in zip(resultados, nombres_hoja):
        ws = wb.create_sheet(nombre_hoja)
        r_libre = _llenar_hoja_resumen(ws, res["candidato"], res["filas"], res["checklist"], res["extra"], fila_inicio=1)
        r_libre += 1  # una fila de separación visual antes del detalle
        r_libre = _llenar_hoja_detalle(ws, res["filas"], fila_inicio=r_libre)
        r_libre += 1  # una fila de separación visual antes de "Datos completos"
        _llenar_hoja_datos_completos(ws, res["candidato"], res["filas"], fila_inicio=r_libre, ajustar_anchos=False)
        # El resumen (arriba) solo usa las columnas A-D y el detalle usa 12;
        # "Datos completos" (abajo del todo) usa hasta 63. El ancho de
        # columna se define una sola vez para toda la hoja, combinando los
        # anchos ya afinados del detalle con un ancho parejo para las
        # columnas adicionales que solo usa "Datos completos".
        anchos_detalle = [30, 26, 9, 10, 12, 20, 12, 24, 22, 18, 20, 70]
        anchos = anchos_detalle + [18] * (len(CAMPOS_DATOS_COMPLETOS) - len(anchos_detalle))
        _autoancho(ws, anchos)

    wb.save(ruta_salida)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Valida el expediente de un candidato y genera un Excel.")
    ap.add_argument("--nombre", required=True, help="Nombre completo registrado del candidato")
    ap.add_argument("--rfc", default="", help="RFC capturado (opcional)")
    ap.add_argument("--curp", default="", help="CURP capturado (opcional)")
    ap.add_argument("--salida", default="expediente_validado.xlsx", help="Ruta del Excel de salida")
    ap.add_argument("pdfs", nargs="+", help="Rutas a los PDFs del candidato")
    args = ap.parse_args()

    candidato = {"nombre": args.nombre, "rfc": args.rfc, "curp": args.curp}

    filas = []
    for ruta in args.pdfs:
        print(f"Procesando: {os.path.basename(ruta)} ...")
        try:
            fila = procesar_documento(ruta, candidato["nombre"])
        except Exception as e:
            fila = {
                "archivo": os.path.basename(ruta), "categoria_clave": "ERROR",
                "categoria": "Error al procesar", "num_paginas": 0, "uso_ocr": False,
                "legible": False, "texto_muestra": "", "nombre_coincide": None,
                "detalle": f"Error: {e}",
            }
        filas.append(fila)

    checklist, extra = construir_reporte(candidato, filas)
    generar_excel(candidato, filas, checklist, extra, args.salida)
    print(f"\nListo. Reporte guardado en: {args.salida}")


if __name__ == "__main__":
    main()
