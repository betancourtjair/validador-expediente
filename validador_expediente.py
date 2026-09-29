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
import os
import re
import sys
import unicodedata
import zipfile

import pdfplumber
from pdf2image import convert_from_path
from PIL import ImageOps, ImageStat, ImageFilter
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

    # formato abreviado tipo recibo (ej. CFE): "03 MAY 26" o "03/MAY/26"
    patron_abrev = re.compile(r"(\d{1,2})\s*[/\s]\s*([A-Z]{3})\s*[/\s]\s*(\d{2})\b")
    for m in patron_abrev.finditer(normaliza(t)):
        dia, mes_txt, anio2 = m.group(1), m.group(2), m.group(3)
        mes = MESES_ABREV.get(mes_txt.upper())
        if mes:
            anio = 2000 + int(anio2)
            try:
                fechas.append((datetime.date(anio, mes, int(dia)), m.group(0)))
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


def _ocr_con_confianza(imagen, lang="spa", config=""):
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
            imagen, lang=lang, config=config, timeout=OCR_TIMEOUT_SEGUNDOS,
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
    with pdfplumber.open(ruta) as pdf:
        num_paginas = len(pdf.pages)
        for pagina in pdf.pages:
            texto = (pagina.extract_text() or "").strip()
            paginas_texto.append(texto)
            ocr_usado.append(False)

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
            try:
                imagenes_pagina_hd = convert_from_path(
                    ruta, dpi=300, first_page=i + 1, last_page=i + 1, grayscale=True
                )
            except Exception as e:
                imagenes_pagina_hd = []
                print(f"  [aviso] no se pudo rasterizar en HD la página {i+1} para OCR ({e})", file=sys.stderr)
            if not imagenes_pagina_hd:
                continue
            imagen_hd = _corrige_rotacion(imagenes_pagina_hd[0])
            del imagenes_pagina_hd

            mejor_texto, mejor_confianza = paginas_texto[i], -1.0
            mejoro = False
            terminar = False
            # Se prueban dos variantes de preprocesamiento (con y sin
            # binarizar — binarizar ayuda mucho con fondos de color pero en
            # documentos de fondo claro a veces borra más de lo que ayuda),
            # UNA A LA VEZ (no las dos en memoria simultáneamente), cada una
            # con dos configuraciones de segmentación de Tesseract, y se
            # elige la de mayor confianza reportada por el propio Tesseract
            # (no la primera que salga ni la más larga).
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
                    # Ya se ve confiable y con longitud razonable: no vale la
                    # pena seguir probando más configuraciones.
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
    "CV": ("CV", ["EXPERIENCIA LABORAL", "REFERENCIAS LABORALES", "PERFIL PROFESIONAL", "CURRICULUM"], True),
    "ACTA_NACIMIENTO": ("Acta de nacimiento", ["ACTA DE NACIMIENTO", "REGISTRO CIVIL", "OFICIALIA", "NACIMIENTOS"], True),
    "INE": ("INE", ["INSTITUTO NACIONAL ELECTORAL", "CREDENCIAL PARA VOTAR", "CLAVE DE ELECTOR"], True),
    "COMPROBANTE_DOMICILIO": ("Comprobante de domicilio", ["COMISION FEDERAL DE ELECTRICIDAD", "CFE", "TELMEX", "RECIBO", "TOTAL A PAGAR", "PERIODO FACTURADO", "IZZI", "TELEFONOS DE MEXICO", "AGUA"], True),
    "COMPROBANTE_ESTUDIOS": ("Certificado de estudios", ["CERTIFICADO DE ESTUDIOS", "CURSO Y ACREDITO", "UNIVERSIDAD", "LICENCIATURA", "CEDULA PROFESIONAL", "SECRETARIA DE EDUCACION", "BACHILLERATO", "PROMEDIO GENERAL"], True),
    "CURP": ("CURP", ["CLAVE UNICA DE REGISTRO DE POBLACION", "CURP CERTIFICADA", "CLAVE:"], True),
    "CSF": ("CSF", ["CONSTANCIA DE SITUACION FISCAL", "CEDULA DE IDENTIFICACION FISCAL", "REGISTRO FEDERAL DE CONTRIBUYENTES"], True),
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


def clasificar(texto_completo, nombre_archivo):
    t = normaliza(texto_completo)
    nombre_arch_norm = normaliza(nombre_archivo)
    mejor_clave, mejor_score = "DESCONOCIDO", 0
    for clave, (_, palabras, _) in CATEGORIAS.items():
        score = sum(1 for p in palabras if normaliza(p) in t)
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
    m = PATRON_CURP.search(normaliza(texto_completo))
    return m.group(1) if m else None


# Número de Seguridad Social (NSS): 11 dígitos, impreso en la constancia de
# "Asignación de Número de Seguridad Social" del IMSS, junto a la etiqueta
# "Número de Seguridad Social".
PATRON_NSS = re.compile(r"NUMERO\s*DE\s*SEGURIDAD\s*SOCIAL\s*:?\s*(\d{11})\b")


def extraer_nss(texto_completo):
    """Extrae el Número de Seguridad Social del documento de "Asignación de
    NSS" del IMSS, anclado a la etiqueta "Número de Seguridad Social".
    Regresa el NSS detectado (11 dígitos) o None si no se encontró."""
    m = PATRON_NSS.search(normaliza(texto_completo))
    return m.group(1) if m else None


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
_CSF_ETQ_ENTIDAD = r"NOMBRE\s*DE\s*LA\s*ENTIDAD\s*FEDERATIVA"
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
PATRON_NOMBRES_CSF = re.compile(r"NOMBRE\s*\(S\)\s*:?\s*(.+?)\s*PRIMER\s*APELLIDO\s*:")
PATRON_APELLIDO_PATERNO_CSF = re.compile(r"PRIMER\s*APELLIDO\s*:?\s*(.+?)\s*SEGUNDO\s*APELLIDO\s*:")
PATRON_APELLIDO_MATERNO_CSF = re.compile(r"SEGUNDO\s*APELLIDO\s*:?\s*(.+?)\s*FECHA\s*INICIO\s*DE\s*OPERACIONES\s*:")


def extraer_nombres_csf(texto_completo):
    """Extrae el/los Nombre(s) impresos en la CSF, anclado entre la etiqueta
    "Nombre (s):" y la etiqueta "Primer Apellido:" que le sigue. Regresa el
    texto detectado o None si no se encontró."""
    m = PATRON_NOMBRES_CSF.search(normaliza(texto_completo))
    return (m.group(1).strip() or None) if m else None


def extraer_apellido_paterno_csf(texto_completo):
    """Extrae el Primer Apellido impreso en la CSF, anclado entre las
    etiquetas "Primer Apellido:" y "Segundo Apellido:". Regresa el texto
    detectado o None si no se encontró."""
    m = PATRON_APELLIDO_PATERNO_CSF.search(normaliza(texto_completo))
    return (m.group(1).strip() or None) if m else None


def extraer_apellido_materno_csf(texto_completo):
    """Extrae el Segundo Apellido impreso en la CSF, anclado entre la
    etiqueta "Segundo Apellido:" y la etiqueta "Fecha inicio de
    operaciones:" que le sigue en la plantilla del SAT. Regresa el texto
    detectado o None si no se encontró (por ejemplo, si el contribuyente
    solo tiene un apellido, ese renglón viene vacío en la CSF)."""
    m = PATRON_APELLIDO_MATERNO_CSF.search(normaliza(texto_completo))
    return (m.group(1).strip() or None) if m else None


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
    return (m.group(1).strip() or None) if m else None


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
            linea_orig_limpia = re.sub(r"\s+", " ", linea_orig).strip(" :.-\t")
            linea_plano_limpia = linea_plano.strip()
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
    "CIUDAD DE MEXICO", "CDMX", "DISTRITO FEDERAL", "DF",
    "COAHUILA", "COAH", "COAHUILA DE ZARAGOZA",
    "COLIMA", "COL",
    "CHIAPAS", "CHIS",
    "CHIHUAHUA", "CHIH",
    "DURANGO", "DGO",
    "GUANAJUATO", "GTO",
    "GUERRERO", "GRO",
    "HIDALGO", "HGO",
    "JALISCO", "JAL",
    "ESTADO DE MEXICO", "EDOMEX", "EDO MEX", "EDO DE MEXICO",
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
    return texto_normalizado.rstrip(".").strip() in ESTADOS_MEXICO_ALIAS


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
    Postal en particular solo se reconoce si viene junto a la etiqueta
    "C.P."/"CP" (un INE normalmente no trae código postal impreso en el
    domicilio, así que ahí quedará en blanco: no se adivina de un número
    suelto de 5 dígitos para no confundirlo con un número exterior).

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
    texto = re.sub(r",?\s*M[EÉ]XICO\s*$", "", texto, flags=re.IGNORECASE).strip(" ,")

    partes = [p.strip() for p in texto.split(",") if p.strip()]
    if not partes:
        return resultado

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
        m_int = re.search(r"\bINT(?:ERIOR)?\.?\s*([A-Za-z0-9\-]+)$", primero, flags=re.IGNORECASE)
        if m_int:
            resultado["numero_interior"] = m_int.group(1)
            primero = primero[:m_int.start()].strip(" ,.-")
        palabras = primero.split()
        if len(palabras) >= 2 and re.match(r"^\d+[A-Za-z]?$", palabras[-1]):
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


def analiza_cuenta_bancaria(texto_completo, nombre_candidato):
    obs = []
    t_norm = normaliza(texto_completo)

    m_clabe = re.search(r"CLABE[^0-9]{0,15}(\d[\d\s]{16,22}\d)", t_norm)
    clabe = re.sub(r"\s", "", m_clabe.group(1)) if m_clabe else None
    if clabe and len(clabe) >= 18:
        clabe = clabe[:18]

    banco = None
    if clabe:
        banco = CODIGOS_CLABE_BANCOS.get(clabe[:3], f"Código de banco no identificado ({clabe[:3]})")

    m_cuenta = re.search(r"(?:NO\.?\s*DE\s*CUENTA|NUMERO\s*DE\s*CUENTA|CUENTA)\s*:?\s*(\d{6,20})", t_norm)
    numero_cuenta = m_cuenta.group(1) if m_cuenta else None
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
    if not tiene_cuenta:
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
    else:
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
    else:
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
        "num_paginas_ok": len(paginas_texto) >= 2,
        "observaciones": " ".join(obs),
    }


def analiza_ine(paginas_texto):
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
    vigente = None
    anio_fin = None
    idx = t_norm.find("VIGENCIA")
    if idx != -1:
        ventana = t_norm[idx: idx + 80]
        anios = re.findall(r"\b(19\d{2}|20\d{2})\b", ventana)
        if len(anios) >= 2:
            anio_inicio, anio_fin = int(anios[-2]), int(anios[-1])
            vigente = anio_fin >= HOY.year
            dias_para_vencer = (datetime.date(anio_fin, 12, 31) - HOY).days
            if vigente:
                obs.append(f"Vigencia impresa en la cara frontal: {anio_inicio}-{anio_fin} (vigente hoy {HOY.isoformat()}).")
            else:
                obs.append(f"Vigencia impresa en la cara frontal: {anio_inicio}-{anio_fin} — VENCIDA (venció hace {abs(dias_para_vencer)} días respecto a hoy {HOY.isoformat()}).")
        else:
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
        r = analiza_ine(paginas_texto)
        fila["vigencia_ok"] = r["vigente"]
        fila["vigencia_fecha_texto"] = f"Vigente hasta {r['vigencia_anio_fin']}" if r["vigencia_anio_fin"] else None
        fila["num_paginas_ok"] = r["dos_paginas"]
        detalles_extra.append(r["observaciones"])
        direccion_ine = extraer_direccion(texto_completo, ["DOMICILIO"])
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
        direccion_domicilio = extraer_direccion(texto_completo, [
            "DOMICILIO DEL SERVICIO", "DOMICILIO DEL USUARIO", "NOMBRE Y DOMICILIO DEL USUARIO",
            "DIRECCION DEL SERVICIO", "DOMICILIO DE INSTALACION", "DOMICILIO",
        ])
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


def construir_reporte(candidato, filas):
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
