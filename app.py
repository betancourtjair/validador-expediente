#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mini aplicación web para el Validador de Expediente de Reclutamiento.

A diferencia del validador_expediente.html de la primera versión (que corría
100% en el navegador), esta parte SÍ necesita un servidor con Python porque
hace OCR y lectura de PDFs en el servidor — eso no se puede hacer de forma
confiable solo con JavaScript en el navegador para documentos escaneados.

La app está dividida en una página de inicio y 3 secciones (cada una es su
propia página, no pestañas de JavaScript, para mantenerlo simple):

  /            Inicio: instrucciones en español, sin usuario ni contraseña.
  /individual  Carga de UN candidato (nombre + RFC + CURP + documentos).
               Sin usuario ni contraseña. El Excel generado NO se descarga
               aquí: se guarda para que el equipo de reclutamiento lo
               descargue después desde /descargas.
  /lote        Carga masiva por ZIP (hasta 10 candidatos). Solo para el
               equipo de reclutamiento -pide usuario y contraseña-. El Excel
               consolidado SÍ se descarga de inmediato aquí, y no se guarda.
  /descargas   Lista y descarga los expedientes generados por /individual.
               Solo para el equipo de reclutamiento -pide usuario y
               contraseña-. Los archivos se eliminan automáticamente a los
               7 días.

Cómo correrlo
--------------
    pip install flask pdfplumber pytesseract pdf2image openpyxl google-cloud-storage
    (y en el sistema: tesseract-ocr, tesseract-ocr-spa, poppler-utils)

    python3 app.py

Abre http://localhost:5000 en el navegador.

Para montarlo en tu página web necesitas un hosting que soporte Python
(Render, Cloud Run, un VPS con Gunicorn + Nginx, etc.) — no es un simple
archivo HTML que se sube a cualquier hosting estático.
"""

import base64
import datetime
import json
import os
import re
import secrets
import shutil
import sys
import tempfile
import time
from functools import wraps

from flask import Flask, request, render_template_string, Response, stream_with_context

import validador_expediente as ve

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Usuarios autorizados (equipo de reclutamiento), vía HTTP Basic Auth.
# ---------------------------------------------------------------------------
# Se pueden sobrescribir con la variable de entorno APP_USUARIOS, con el
# formato "correo1:clave1,correo2:clave2,...". Si no se configura nada, se
# usan estas cuentas (las del equipo de Fitness Para Todos) por default.
def _cargar_usuarios_autorizados():
    variable = os.environ.get("APP_USUARIOS", "").strip()
    usuarios = {}
    if variable:
        for par in variable.split(","):
            par = par.strip()
            if ":" in par:
                correo, clave = par.split(":", 1)
                correo = correo.strip().lower()
                if correo:
                    usuarios[correo] = clave
    if not usuarios:
        usuarios = {
            "carlos.diaz@fpt.com.mx": "PlanetFitness$01",
            "angelica.fuentes@fpt.com.mx": "PlanetFitness$01",
            "jessica.otamendi@fpt.com.mx": "PlanetFitness$01",
            "jair@fpt.com.mx": "PlanetFitness$01",
        }
    return usuarios


USUARIOS_AUTORIZADOS = _cargar_usuarios_autorizados()


def requiere_login(f):
    @wraps(f)
    def decorado(*args, **kwargs):
        auth = request.authorization
        correo = (auth.username or "").strip().lower() if auth else ""
        clave_esperada = USUARIOS_AUTORIZADOS.get(correo)
        credenciales_ok = (
            auth is not None
            and clave_esperada is not None
            and secrets.compare_digest(auth.password or "", clave_esperada)
        )
        if not credenciales_ok:
            return Response(
                "Acceso restringido. Ingresa el usuario y contraseña del equipo de reclutamiento.",
                401,
                {"WWW-Authenticate": 'Basic realm="Validador de expediente"'},
            )
        return f(*args, **kwargs)

    return decorado


CAMPOS_DOCUMENTOS = [
    ("cv", "CV", True),
    ("acta_nacimiento", "Acta de nacimiento", True),
    ("ine", "INE (ambos lados en 1 PDF)", True),
    ("comprobante_domicilio", "Comprobante de domicilio", True),
    ("comprobante_estudios", "Comprobante de estudios", True),
    ("curp", "CURP", True),
    ("csf", "CSF (ambos lados en 1 PDF)", True),
    ("nss", "NSS", True),
    ("cuenta_bancaria", "Cuenta bancaria (carátula)", True),
    ("infonavit_fonacot", "Aviso de retención Infonavit/Fonacot", False),
    ("certificado_medico", "Certificado médico", False),
    ("certificado_instructor", "Certificado de entrenador / barbero / estilista", False),
    ("constancia_laboral", "Constancia(s) laboral(es)", False),
]

# ---------------------------------------------------------------------------
# Almacenamiento de los expedientes individuales (para la página /descargas)
# ---------------------------------------------------------------------------
# En producción (Cloud Run) esto se guarda en un bucket de Google Cloud
# Storage -indispensable porque Cloud Run puede correr varias instancias (o
# reiniciar la única que haya) y el disco local NO se comparte entre ellas;
# si se guardara solo en disco, la página de Documentos a veces "no
# encontraría" un archivo que sí se generó, según qué instancia atendiera
# cada petición. El borrado automático a los 7 días se configura como una
# regla de "lifecycle" del propio bucket (más confiable que borrarlo desde
# el código: funciona aunque nadie visite la página en varios días).
#
# Para desarrollo/pruebas locales (sin GCS_BUCKET_EXPEDIENTES configurada)
# se cae de vuelta a guardar los archivos en disco, para poder probar todo
# el flujo sin necesitar credenciales de Google Cloud.
GCS_BUCKET_EXPEDIENTES = os.environ.get("GCS_BUCKET_EXPEDIENTES", "").strip()
_cliente_gcs_cache = None
ALMACEN_LOCAL_DIR = os.path.join(tempfile.gettempdir(), "expedientes_individuales")
os.makedirs(ALMACEN_LOCAL_DIR, exist_ok=True)


def _cliente_gcs():
    global _cliente_gcs_cache
    if _cliente_gcs_cache is None:
        from google.cloud import storage  # import perezoso: no hace falta si no hay bucket configurado
        _cliente_gcs_cache = storage.Client()
    return _cliente_gcs_cache


def _nombre_archivo_seguro(texto):
    base = re.sub(r"[^A-Za-z0-9]+", "_", texto or "candidato").strip("_")
    return (base or "candidato")[:60]


def guardar_expediente_individual(ruta_local_xlsx, nombre_candidato):
    """Guarda el Excel ya generado de un candidato (carga individual) para
    que el equipo de reclutamiento lo descargue después desde /descargas.
    Regresa el nombre interno con el que quedó guardado."""
    marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_interno = f"{marca}_{_nombre_archivo_seguro(nombre_candidato)}.xlsx"
    if GCS_BUCKET_EXPEDIENTES:
        bucket = _cliente_gcs().bucket(GCS_BUCKET_EXPEDIENTES)
        blob = bucket.blob(f"individuales/{nombre_interno}")
        blob.metadata = {"candidato": nombre_candidato}
        blob.upload_from_filename(
            ruta_local_xlsx,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        destino = os.path.join(ALMACEN_LOCAL_DIR, nombre_interno)
        shutil.copyfile(ruta_local_xlsx, destino)
        with open(destino + ".meta.json", "w", encoding="utf-8") as fh:
            json.dump({"candidato": nombre_candidato}, fh, ensure_ascii=False)
    return nombre_interno


def listar_expedientes_individuales():
    """Regresa la lista de expedientes individuales guardados (más reciente
    primero), como dicts {nombre_interno, candidato, fecha}."""
    resultados = []
    if GCS_BUCKET_EXPEDIENTES:
        bucket = _cliente_gcs().bucket(GCS_BUCKET_EXPEDIENTES)
        for blob in bucket.list_blobs(prefix="individuales/"):
            nombre_interno = blob.name.split("/", 1)[-1]
            if not nombre_interno:
                continue
            candidato = (blob.metadata or {}).get("candidato") or nombre_interno
            fecha = blob.time_created or datetime.datetime.now(datetime.timezone.utc)
            resultados.append({"nombre_interno": nombre_interno, "candidato": candidato, "fecha": fecha})
    else:
        for nombre_archivo in os.listdir(ALMACEN_LOCAL_DIR):
            if not nombre_archivo.endswith(".xlsx"):
                continue
            ruta = os.path.join(ALMACEN_LOCAL_DIR, nombre_archivo)
            candidato = nombre_archivo
            meta_path = ruta + ".meta.json"
            if os.path.exists(meta_path):
                try:
                    with open(meta_path, encoding="utf-8") as fh:
                        candidato = json.load(fh).get("candidato") or nombre_archivo
                except Exception:
                    pass
            resultados.append({
                "nombre_interno": nombre_archivo,
                "candidato": candidato,
                "fecha": datetime.datetime.fromtimestamp(os.path.getmtime(ruta)),
            })
    resultados.sort(key=lambda r: r["fecha"], reverse=True)
    return resultados


def leer_expediente_individual(nombre_interno):
    """Regresa los bytes del Excel guardado, o None si no existe (ya se
    borró, o el nombre no es válido)."""
    if not nombre_interno or "/" in nombre_interno or ".." in nombre_interno or not nombre_interno.endswith(".xlsx"):
        return None
    if GCS_BUCKET_EXPEDIENTES:
        bucket = _cliente_gcs().bucket(GCS_BUCKET_EXPEDIENTES)
        blob = bucket.blob(f"individuales/{nombre_interno}")
        if not blob.exists():
            return None
        return blob.download_as_bytes()
    ruta = os.path.join(ALMACEN_LOCAL_DIR, nombre_interno)
    if not os.path.isfile(ruta):
        return None
    with open(ruta, "rb") as fh:
        return fh.read()


# ---------------------------------------------------------------------------
# HTML compartido: estilos, barra de navegación y el JS de "subir con
# barra de progreso" que usan tanto /individual como /lote.
# ---------------------------------------------------------------------------
# Esto retoma el "look and feel" del sitio https://fpt.com.mx/ (morado de
# marca #592c82, morado oscuro #3c1053 en la barra superior, logo en blanco,
# tipografía Barlow Condensed en encabezados y Open Sans en el resto).
ESTILOS = """
  :root{--fpt-morado:#592c82;--fpt-morado-oscuro:#3c1053;}
  *{box-sizing:border-box;}
  body{font-family:"Open Sans",Arial,sans-serif;margin:0;color:#333333;background:#ffffff;}
  h1,h2{font-family:"Barlow Condensed","Arial Narrow",Arial,sans-serif;font-weight:700;color:var(--fpt-morado);}
  h1{font-size:1.9rem;margin-top:0;letter-spacing:.01em;}
  h2{font-size:1.35rem;margin-top:0;}
  a{color:var(--fpt-morado);}
  .contenido{max-width:760px;margin:0 auto;padding:30px 16px 40px;}
  .fpt-header{background:var(--fpt-morado-oscuro);display:flex;align-items:center;flex-wrap:wrap;gap:18px;padding:10px 20px;position:sticky;top:0;z-index:10;}
  .fpt-logo{display:flex;align-items:center;padding-right:18px;border-right:1px solid rgba(255,255,255,.25);}
  .fpt-logo img{display:block;height:40px;width:auto;}
  nav.tabs{display:flex;gap:2px;flex-wrap:wrap;}
  nav.tabs a{padding:10px 14px;text-decoration:none;color:#fff;font-weight:600;font-size:.82rem;text-transform:uppercase;letter-spacing:.03em;border-bottom:3px solid transparent;opacity:.85;}
  nav.tabs a.activo{opacity:1;border-bottom-color:#fff;}
  nav.tabs a:hover{opacity:1;}
  .campo{margin-bottom:14px;}
  label{display:block;font-weight:600;font-size:.85rem;margin-bottom:4px;color:#333;}
  input[type=text]{width:100%;padding:8px;border:1px solid #ccc;border-radius:6px;font-family:inherit;}
  input[type=text]:focus{outline:none;border-color:var(--fpt-morado);box-shadow:0 0 0 2px rgba(89,44,130,.15);}
  .doc{border:1px solid #e0e0e0;border-radius:8px;padding:10px 14px;margin-bottom:8px;}
  .doc.req{background:#f5f0fa;border-color:#d9c7ec;}
  .tag{font-size:.7rem;background:var(--fpt-morado);color:#fff;padding:2px 6px;border-radius:10px;margin-left:6px;}
  .oculto{display:none;}
  button{background:var(--fpt-morado);color:#fff;border:none;padding:12px 22px;border-radius:8px;font-size:1rem;font-weight:600;cursor:pointer;font-family:inherit;}
  button:hover:not(:disabled){background:var(--fpt-morado-oscuro);}
  button:disabled{opacity:.6;cursor:not-allowed;}
  #estado,#estadoLote{margin-top:14px;font-size:.9rem;font-weight:600;min-height:1.2em;}
  #barraContenedor,#barraContenedorLote{margin-top:12px;display:none;}
  progress{width:100%;height:16px;border-radius:8px;overflow:hidden;}
  progress::-webkit-progress-bar{background:#ece4f3;border-radius:8px;}
  progress::-webkit-progress-value{background:var(--fpt-morado);border-radius:8px;transition:width .25s ease;}
  progress::-moz-progress-bar{background:var(--fpt-morado);border-radius:8px;}
  #barraTexto,#barraTextoLote{font-size:.8rem;color:#555;margin-top:4px;}
  .pista{background:#f5f0fa;border:1px solid #d9c7ec;border-radius:8px;padding:10px 14px;font-size:.82rem;margin-bottom:16px;line-height:1.5;color:#3c1053;}
  .pista code{background:#e4d7ef;padding:1px 5px;border-radius:4px;}
  .lote-fila{display:flex;gap:10px;align-items:center;border:1px solid #e0e0e0;border-radius:8px;padding:8px 12px;margin-bottom:8px;}
  .lote-fila .num{font-weight:600;font-size:.85rem;color:#555;width:22px;flex:none;}
  .lote-fila input[type=text]{flex:1 1 auto;min-width:0;}
  .lote-fila input[type=file]{flex:1 1 auto;min-width:0;font-size:.82rem;}
  #avisoTamanoLote{display:none;background:#fff3cd;border:1px solid #ffe08a;color:#7a5b00;border-radius:8px;padding:10px 14px;font-size:.85rem;margin-bottom:12px;}
  .tarjeta{border:1px solid #e0e0e0;border-radius:10px;padding:16px 18px;margin-bottom:16px;}
  .tarjeta a.boton{display:inline-block;margin-top:10px;background:var(--fpt-morado);color:#fff;padding:9px 18px;border-radius:8px;text-decoration:none;font-size:.9rem;font-weight:600;}
  .tarjeta a.boton:hover{background:var(--fpt-morado-oscuro);}
  table.descargas{width:100%;border-collapse:collapse;margin-top:10px;font-size:.88rem;}
  table.descargas th,table.descargas td{text-align:left;padding:8px 10px;border-bottom:1px solid #e6e6e6;}
  table.descargas th{color:#555;font-size:.78rem;text-transform:uppercase;letter-spacing:.03em;}
  table.descargas a{font-weight:600;}
  .vacio{color:#777;font-size:.9rem;padding:10px 0;}
  .aviso-info{background:#f5f0fa;border:1px solid #d9c7ec;border-radius:8px;padding:10px 14px;font-size:.85rem;margin-bottom:16px;color:#3c1053;}
"""

_ENLACES_NAV = [
    ("inicio", "/", "Inicio"),
    ("individual", "/individual", "Carga individual"),
    ("lote", "/lote", "Carga masiva"),
    ("descargas", "/descargas", "Documentos"),
]

# Logo oficial (versión blanca, pensada para fondo morado oscuro) tomado
# directamente de https://fpt.com.mx/ para mantener el mismo look and feel.
_LOGO_FPT_URL = "https://fpt.com.mx/img/fpt-logo-blanco.png"


def _nav(activo):
    piezas = []
    for clave, url, etiqueta in _ENLACES_NAV:
        clase = "activo" if clave == activo else ""
        piezas.append(f'<a href="{url}" class="{clase}">{etiqueta}</a>')
    return '<nav class="tabs">' + "".join(piezas) + "</nav>"


def _envolver_pagina(activo, titulo, cuerpo_html):
    """Junta el <head>/estilos/nav con el contenido propio de cada página.
    El resultado todavía puede tener sintaxis de Jinja (p.ej. {{ }} o {% %})
    heredada de cuerpo_html -eso se procesa después, al pasar el resultado
    de esta función a render_template_string()-."""
    return f"""
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{titulo}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700&family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
<style>{ESTILOS}</style>
</head>
<body>
<header class="fpt-header">
  <a href="/" class="fpt-logo"><img src="{_LOGO_FPT_URL}" alt="Fitness Para Todos"></a>
  {_nav(activo)}
</header>
<main class="contenido">
{cuerpo_html}
</main>
</body>
</html>
"""


# El "núcleo" de JS (la función configurarFormularioConProgreso) es igual
# para /individual y /lote, así que se define una sola vez aquí y cada
# página solo agrega su propia llamada de configuración.
JS_NUCLEO = """
<script>
// Los formularios se mandan por fetch (en vez de un submit normal) para
// poder limpiarlos por completo -incluyendo los archivos ya seleccionados-
// justo después de terminar, y para poder mostrar una barra de progreso en
// vivo.
//
// El servidor va mandando el avance real (documento por documento, o
// candidato por candidato en la carga masiva) como un "stream" de eventos
// (Server-Sent Events) dentro de esta misma petición POST -no en peticiones
// aparte- para que, tanto en Render como en Cloud Run, el servidor siga
// teniendo CPU asignado mientras hace el OCR (si se regresara la respuesta
// de inmediato y el trabajo se hiciera "detrás" en otro lado, Cloud Run
// podría congelar el CPU del contenedor entre una consulta y otra).
//
// configurarFormularioConProgreso() cablea un formulario con su barra de
// progreso. opciones.descargar controla si al terminar se debe disparar la
// descarga del Excel en el navegador (carga masiva) o no (carga
// individual, donde el Excel solo se guarda en el servidor).
function configurarFormularioConProgreso(opciones) {
  var form = document.getElementById(opciones.formId);
  var boton = document.getElementById(opciones.botonId);
  var estado = document.getElementById(opciones.estadoId);
  var barraContenedor = document.getElementById(opciones.barraContenedorId);
  var barra = document.getElementById(opciones.barraId);
  var barraTexto = document.getElementById(opciones.barraTextoId);
  var textoBoton = boton.textContent;
  var yaTermino = false;

  function terminar(mensaje, esError) {
    yaTermino = true;
    estado.style.color = esError ? "#b3261e" : "#0f6b4c";
    estado.textContent = mensaje;
    barraContenedor.style.display = "none";
    boton.disabled = false;
    boton.textContent = textoBoton;
  }

  function base64ABlob(base64, tipoMime) {
    var texto = atob(base64);
    var numeros = new Array(texto.length);
    for (var i = 0; i < texto.length; i++) numeros[i] = texto.charCodeAt(i);
    return new Blob([new Uint8Array(numeros)], { type: tipoMime });
  }

  function manejarEvento(info) {
    if (info.tipo === "progreso") {
      barra.max = info.total || 1;
      barra.value = info.hecho || 0;
      var porcentaje = Math.round(100 * (info.hecho || 0) / (info.total || 1));
      var prefijo = info.candidato_total ? ("Candidato " + info.candidato_index + "/" + info.candidato_total + " (" + (info.candidato_actual || "") + ") — ") : "";
      barraTexto.textContent = porcentaje + "% - " + prefijo + (info.archivo_actual || "Procesando...");
    } else if (info.tipo === "listo") {
      barra.value = barra.max;
      barraTexto.textContent = "100%";
      if (opciones.descargar !== false && info.datos_base64) {
        var blob = base64ABlob(info.datos_base64, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        var url = window.URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = info.nombre_archivo || "expediente.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      }

      // Limpia todos los campos y archivos seleccionados para la siguiente carga.
      form.reset();
      terminar(info.mensaje || ("Listo: se descargó " + (info.nombre_archivo || "expediente.xlsx") + ". El formulario ya está limpio para la siguiente carga."), false);
    } else if (info.tipo === "error") {
      terminar("Ocurrió un error generando el expediente: " + (info.error || "error desconocido") + " (no se borró nada, puedes reintentar).", true);
    }
  }

  form.addEventListener("submit", function (ev) {
    ev.preventDefault();
    if (opciones.antesDeEnviar && !opciones.antesDeEnviar()) {
      return;
    }
    estado.style.color = "#0f6b4c";
    estado.textContent = opciones.mensajeInicial || "Procesando (puede tardar varios minutos)...";
    boton.disabled = true;
    boton.textContent = "Procesando...";
    barraContenedor.style.display = "block";
    barra.value = 0;
    barra.max = 1;
    barraTexto.textContent = "Subiendo archivos...";

    var datos = new FormData(form);

    fetch(form.getAttribute("action"), { method: "POST", body: datos })
      .then(function (resp) {
        if (!resp.ok) {
          return resp.text().then(function (txt) {
            throw new Error(txt || ("Error del servidor (" + resp.status + ")"));
          });
        }
        if (!resp.body || !resp.body.getReader) {
          // Navegador viejo sin soporte de streams: al menos no se rompe,
          // aunque no se vea el avance en tiempo real.
          return resp.text().then(function () {
            throw new Error("Tu navegador no soporta ver el avance en vivo; actualízalo o inténtalo desde otro navegador.");
          });
        }

        var lector = resp.body.getReader();
        var decodificador = new TextDecoder("utf-8");
        var buffer = "";

        function procesaBuffer() {
          var bloques = buffer.split("\\n\\n");
          buffer = bloques.pop();
          bloques.forEach(function (bloque) {
            bloque.split("\\n").forEach(function (linea) {
              linea = linea.trim();
              if (linea.indexOf("data:") !== 0) return;
              var jsonTexto = linea.slice(5).trim();
              if (!jsonTexto) return;
              var info;
              try {
                info = JSON.parse(jsonTexto);
              } catch (e) {
                return;
              }
              manejarEvento(info);
            });
          });
        }

        function leer() {
          return lector.read().then(function (resultado) {
            if (resultado.done) {
              buffer += decodificador.decode();
              procesaBuffer();
              return;
            }
            buffer += decodificador.decode(resultado.value, { stream: true });
            procesaBuffer();
            return leer();
          });
        }

        return leer();
      })
      .then(function () {
        if (!yaTermino) {
          terminar("El servidor cerró la conexión sin terminar el proceso. Verifica los documentos e inténtalo de nuevo.", true);
        }
      })
      .catch(function (err) {
        terminar("Ocurrió un error: " + err.message + " (no se borró nada, puedes reintentar).", true);
      });
  });
}
</script>
"""


# ---------------------------------------------------------------------------
# Página de inicio (/) — solo instrucciones, sin usuario ni contraseña.
# ---------------------------------------------------------------------------
CONTENIDO_INICIO = """
<h1>Validador de Expediente — Fitness Para Todos</h1>
<p>Esta herramienta revisa que el expediente de un candidato esté completo:
lee los documentos en PDF (incluso escaneados, usando OCR) y genera un
Excel con el resultado.</p>

<div class="tarjeta">
  <h2>1. Carga individual (para el candidato)</h2>
  <p>El candidato sube sus documentos uno por uno (CV, INE, comprobante de
  domicilio, etc.) junto con su nombre, RFC y CURP. <strong>No se necesita
  usuario ni contraseña.</strong></p>
  <p>Al terminar, el Excel del expediente se guarda automáticamente — el
  candidato NO lo descarga. El equipo de reclutamiento lo descarga después
  desde <strong>Documentos</strong>. El formulario queda listo de inmediato
  para el siguiente candidato.</p>
  <a class="boton" href="/individual">Ir a carga individual</a>
</div>

<div class="tarjeta">
  <h2>2. Carga masiva por ZIP (equipo de reclutamiento)</h2>
  <p>Permite subir hasta <strong>10 candidatos a la vez</strong>: un archivo
  ZIP por candidato, con sus documentos en PDF adentro nombrados según la
  convención (cv.pdf, ine.pdf, comprobante_domicilio.pdf, etc. — se explica
  con más detalle en esa página). <strong>Requiere iniciar sesión.</strong></p>
  <p>Al terminar se descarga de inmediato un solo Excel con una hoja por
  cada candidato del lote.</p>
  <a class="boton" href="/lote">Ir a carga masiva</a>
</div>

<div class="tarjeta">
  <h2>3. Documentos (equipo de reclutamiento)</h2>
  <p>Aquí se descargan los expedientes en Excel generados por la
  <strong>carga individual</strong> de los candidatos. <strong>Requiere
  iniciar sesión.</strong></p>
  <p>Los archivos se eliminan automáticamente 7 días después de haberse
  generado.</p>
  <a class="boton" href="/descargas">Ir a documentos</a>
</div>
"""


# ---------------------------------------------------------------------------
# Carga individual (/individual, /validar) — sin usuario ni contraseña.
# ---------------------------------------------------------------------------
CONTENIDO_INDIVIDUAL = """
<h1>Carga individual de documentos</h1>
<div class="aviso-info">
  Sube tus documentos en PDF. Al terminar, tu expediente queda guardado
  para que el equipo de reclutamiento lo revise — no se descarga nada en
  este paso.
</div>
<form method="post" action="/validar" enctype="multipart/form-data" id="formValidador">
  <div class="campo">
    <label>Nombre completo del candidato</label>
    <input type="text" name="nombre" required>
  </div>
  <div class="campo">
    <label>RFC (opcional)</label>
    <input type="text" name="rfc">
  </div>
  <div class="campo">
    <label>CURP (opcional)</label>
    <input type="text" name="curp">
  </div>
  {% for campo, etiqueta, obligatorio in documentos %}
  <div class="doc {{ 'req' if obligatorio else '' }}">
    <label>{{ etiqueta }} {% if not obligatorio %}<span class="tag">Opcional</span>{% endif %}</label>
    <input type="file" name="{{ campo }}" accept="application/pdf">
  </div>
  {% endfor %}
  <!-- "Otros documentos adicionales" oculto a petición del equipo de FPT
       (2026-08-05). El campo sigue existiendo en el formulario -por si se
       vuelve a necesitar solo hace falta quitar la clase "oculto"-, pero no
       se manda nada en este campo mientras esté oculto porque no hay forma
       de seleccionar archivos. -->
  <div class="campo oculto">
    <label>Otros documentos adicionales (puedes seleccionar varios)</label>
    <input type="file" name="otros" multiple accept="application/pdf">
  </div>
  <button type="submit" id="btnSubmit">Enviar documentos</button>
  <div id="estado"></div>
  <div id="barraContenedor">
    <progress id="barra" value="0" max="1"></progress>
    <div id="barraTexto"></div>
  </div>
</form>
""" + JS_NUCLEO + """
<script>
configurarFormularioConProgreso({
  formId: "formValidador",
  botonId: "btnSubmit",
  estadoId: "estado",
  barraContenedorId: "barraContenedor",
  barraId: "barra",
  barraTextoId: "barraTexto",
  mensajeInicial: "Procesando documentos (puede tardar uno o dos minutos)...",
  descargar: false,
});
</script>
"""


# ---------------------------------------------------------------------------
# Carga masiva por ZIP (/lote, /validar_lote) — requiere iniciar sesión.
# ---------------------------------------------------------------------------
CONTENIDO_LOTE = """
<h1>Carga masiva por ZIP (hasta 10 candidatos)</h1>
<div class="pista">
  Sube <strong>un archivo ZIP por candidato</strong>, con los PDFs de ese candidato
  adentro. El sistema identifica cada documento por el <strong>nombre del archivo</strong>
  dentro del ZIP (no hace falta acomodarlos en campos separados). Nombra los PDFs
  así (mayúsculas/minúsculas y guiones no importan):
  <br><br>
  {% for campo, etiqueta, obligatorio in documentos %}<code>{{ campo }}.pdf</code>{% if not loop.last %}, {% endif %}{% endfor %},
  y cualquier otro documento con el nombre que quieras (se reporta como "otros" o se
  intenta reconocer por su contenido).
  <br><br>
  Límite: 10 candidatos por tanda y <strong>~28&nbsp;MB en total</strong> entre todos los
  ZIP juntos (límite de la plataforma) — si pesan más, sube los candidatos en 2 tandas.
</div>
<div id="avisoTamanoLote"></div>
<form method="post" action="/validar_lote" enctype="multipart/form-data" id="formLote">
  {% for i in range(1, 11) %}
  <div class="lote-fila">
    <div class="num">{{ i }}.</div>
    <input type="text" name="nombre_{{ i }}" placeholder="Nombre completo del candidato {{ i }}">
    <input type="file" name="zip_{{ i }}" accept=".zip,application/zip">
  </div>
  {% endfor %}
  <button type="submit" id="btnSubmitLote">Procesar candidatos (ZIP)</button>
  <div id="estadoLote"></div>
  <div id="barraContenedorLote">
    <progress id="barraLote" value="0" max="1"></progress>
    <div id="barraTextoLote"></div>
  </div>
</form>
""" + JS_NUCLEO + """
<script>
// Límite de la plataforma (Cloud Run) para el tamaño de una sola petición
// HTTP: ~32 MiB. Se deja un margen (28 MB) para los demás campos del
// formulario y la sobrecarga propia de multipart/form-data. Si se detecta
// que la suma de los ZIP seleccionados pasa de ese margen, se avisa ANTES
// de mandar la petición -así el usuario no espera para nada a que falle-.
var LIMITE_TAMANO_LOTE_BYTES = 28 * 1024 * 1024;

configurarFormularioConProgreso({
  formId: "formLote",
  botonId: "btnSubmitLote",
  estadoId: "estadoLote",
  barraContenedorId: "barraContenedorLote",
  barraId: "barraLote",
  barraTextoId: "barraTextoLote",
  mensajeInicial: "Procesando candidatos (puede tardar varios minutos con 10 candidatos)...",
  antesDeEnviar: function () {
    var aviso = document.getElementById("avisoTamanoLote");
    var formLote = document.getElementById("formLote");
    var total = 0;
    var hayAlgunArchivo = false;
    for (var i = 1; i <= 10; i++) {
      var input = formLote.querySelector('input[name="zip_' + i + '"]');
      if (input && input.files && input.files[0]) {
        hayAlgunArchivo = true;
        total += input.files[0].size;
      }
    }
    if (!hayAlgunArchivo) {
      aviso.style.display = "block";
      aviso.textContent = "Llena al menos una fila con nombre y archivo ZIP.";
      return false;
    }
    if (total > LIMITE_TAMANO_LOTE_BYTES) {
      aviso.style.display = "block";
      aviso.textContent = "Los ZIP seleccionados suman " + (total / (1024 * 1024)).toFixed(1) +
        " MB, y el límite de la plataforma es de aproximadamente 28 MB por tanda. " +
        "Quita algún candidato de esta tanda y súbelo en una tanda aparte.";
      return false;
    }
    aviso.style.display = "none";
    return true;
  },
});
</script>
"""


# ---------------------------------------------------------------------------
# Documentos (/descargas) — requiere iniciar sesión.
# ---------------------------------------------------------------------------
CONTENIDO_DESCARGAS = """
<h1>Documentos — expedientes individuales</h1>
<div class="aviso-info">
  Aquí se descargan los expedientes en Excel generados por la <strong>carga
  individual</strong> de candidatos. Se eliminan automáticamente 7 días
  después de haberse generado.
</div>
<table class="descargas">
  <thead><tr><th>Candidato</th><th>Generado</th><th></th></tr></thead>
  <tbody>
  {% for archivo in archivos %}
    <tr>
      <td>{{ archivo.candidato }}</td>
      <td>{{ archivo.fecha_texto }}</td>
      <td><a href="/descargas/archivo/{{ archivo.nombre_interno }}">Descargar</a></td>
    </tr>
  {% else %}
    <tr><td colspan="3" class="vacio">Todavía no hay expedientes individuales guardados.</td></tr>
  {% endfor %}
  </tbody>
</table>
"""


@app.route("/", methods=["GET"])
def inicio():
    return render_template_string(_envolver_pagina("inicio", "Validador de Expediente — Fitness Para Todos", CONTENIDO_INICIO))


@app.route("/individual", methods=["GET"])
def formulario_individual():
    return render_template_string(
        _envolver_pagina("individual", "Carga individual — Validador de Expediente", CONTENIDO_INDIVIDUAL),
        documentos=CAMPOS_DOCUMENTOS,
    )


@app.route("/lote", methods=["GET"])
@requiere_login
def formulario_lote():
    return render_template_string(
        _envolver_pagina("lote", "Carga masiva — Validador de Expediente", CONTENIDO_LOTE),
        documentos=CAMPOS_DOCUMENTOS,
    )


@app.route("/descargas", methods=["GET"])
@requiere_login
def descargas():
    archivos = listar_expedientes_individuales()
    for item in archivos:
        fecha = item.get("fecha")
        item["fecha_texto"] = fecha.strftime("%d/%m/%Y %H:%M") if hasattr(fecha, "strftime") else str(fecha)
    return render_template_string(
        _envolver_pagina("descargas", "Documentos — Validador de Expediente", CONTENIDO_DESCARGAS),
        archivos=archivos,
    )


@app.route("/descargas/archivo/<nombre_interno>", methods=["GET"])
@requiere_login
def descargar_archivo(nombre_interno):
    datos = leer_expediente_individual(nombre_interno)
    if datos is None:
        return "Archivo no encontrado (puede que ya se haya eliminado automáticamente después de 7 días).", 404
    return Response(
        datos,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre_interno}"'},
    )


def _evento_sse(datos_dict):
    """Da formato a un diccionario como un evento de Server-Sent Events."""
    return "data: " + json.dumps(datos_dict, ensure_ascii=False) + "\n\n"


def _generar_eventos_validacion(tmp, rutas, candidato):
    """Generador que hace todo el trabajo (OCR + Excel) y va cediendo (yield)
    un evento de progreso por cada documento, todo DENTRO de la misma
    petición HTTP de /validar (por eso se manda como Server-Sent Events en
    vez de, por ejemplo, lanzar un hilo aparte y contestar de inmediato):
    tanto Render como Cloud Run le asignan CPU real al contenedor mientras
    haya una petición en curso, pero en Cloud Run en particular el CPU se
    puede congelar en cuanto ya no hay ninguna petición abierta — si el OCR
    seguidor corriera "detrás" en un hilo después de haber contestado ya la
    petición, se podría quedar sin CPU y tardar muchísimo más (o nunca
    terminar) hasta que llegara otra petición.

    A diferencia de antes, al terminar NO se manda el Excel de vuelta al
    navegador: se guarda con guardar_expediente_individual() para que el
    equipo de reclutamiento lo descargue después desde /descargas -este
    formulario es público (sin usuario/contraseña), así que no tiene sentido
    que el candidato se quede con una copia del archivo.

    Al terminar (bien o mal) borra el directorio temporal con los PDFs
    subidos, igual que antes hacía `tempfile.TemporaryDirectory()` — nomás
    que aquí se hace a mano en el "finally" porque ya no se usa ese context
    manager (el directorio tiene que seguir vivo mientras dura el generador).
    """
    nombre = candidato["nombre"]
    total = len(rutas) + 1  # +1 = paso final de generar el Excel
    hecho = 0
    try:
        print(f"[validar] iniciando expediente de '{nombre}' con {len(rutas)} documento(s)", file=sys.stderr, flush=True)
        yield _evento_sse({"tipo": "progreso", "hecho": hecho, "total": total, "archivo_actual": "Iniciando..."})

        filas = []
        for ruta in rutas:
            t0 = time.time()
            nombre_doc = os.path.basename(ruta)
            yield _evento_sse({"tipo": "progreso", "hecho": hecho, "total": total, "archivo_actual": nombre_doc})
            print(f"[validar]   procesando {nombre_doc} ...", file=sys.stderr, flush=True)
            try:
                fila = ve.procesar_documento(ruta, nombre)
            except Exception as e:
                fila = {
                    "archivo": nombre_doc, "categoria_clave": "ERROR",
                    "categoria": "Error al procesar", "num_paginas": 0, "uso_ocr": False,
                    "legible": False, "texto_muestra": "", "nombre_coincide": None,
                    "detalle": f"Error: {e}",
                }
            print(f"[validar]   listo {nombre_doc} ({time.time() - t0:.1f}s)", file=sys.stderr, flush=True)
            filas.append(fila)
            hecho += 1
            yield _evento_sse({"tipo": "progreso", "hecho": hecho, "total": total, "archivo_actual": nombre_doc})

        yield _evento_sse({"tipo": "progreso", "hecho": hecho, "total": total, "archivo_actual": "Generando Excel..."})

        checklist, extra = ve.construir_reporte(candidato, filas)
        salida = os.path.join(tmp, "expediente.xlsx")
        ve.generar_excel(candidato, filas, checklist, extra, salida)

        nombre_interno = guardar_expediente_individual(salida, nombre)
        print(f"[validar] expediente de '{nombre}' guardado como '{nombre_interno}'", file=sys.stderr, flush=True)

        hecho = total
        yield _evento_sse({
            "tipo": "listo",
            "hecho": hecho,
            "total": total,
            "mensaje": (
                f"Listo: se recibió el expediente de {nombre}. Ya puedes cerrar esta página; "
                "el equipo de reclutamiento lo revisará. El formulario está listo para el siguiente candidato."
            ),
        })

        print(f"[validar] expediente de '{nombre}' listo", file=sys.stderr, flush=True)
    except Exception as e:
        # Cualquier error inesperado (no solo los de un documento en
        # particular, que ya se capturan arriba) se reporta al navegador vía
        # el mismo stream en vez de dejar la barra de progreso trabada.
        print(f"[validar] ERROR inesperado con el expediente de '{nombre}': {e}", file=sys.stderr, flush=True)
        yield _evento_sse({"tipo": "error", "error": str(e)})
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.route("/validar", methods=["POST"])
def validar():
    nombre = request.form.get("nombre", "").strip()
    rfc = request.form.get("rfc", "").strip()
    curp = request.form.get("curp", "").strip()
    if not nombre:
        return "Falta el nombre del candidato", 400

    candidato = {"nombre": nombre, "rfc": rfc, "curp": curp}

    # A diferencia de antes, aquí NO se usa `with tempfile.TemporaryDirectory()`
    # porque el directorio tiene que seguir existiendo mientras el generador
    # de eventos (_generar_eventos_validacion) procesa los documentos, y ese
    # generador sigue corriendo después de que esta función ya regresó su
    # Response (Flask lo va consumiendo mientras manda el stream). Se limpia
    # a mano en el "finally" del generador, tanto si todo sale bien como si
    # truena.
    tmp = tempfile.mkdtemp(prefix="expediente_")
    try:
        rutas = []
        for campo, _etiqueta, _ob in CAMPOS_DOCUMENTOS:
            f = request.files.get(campo)
            if f and f.filename:
                ruta = os.path.join(tmp, f.filename)
                f.save(ruta)
                rutas.append(ruta)
        for f in request.files.getlist("otros"):
            if f and f.filename:
                ruta = os.path.join(tmp, f.filename)
                f.save(ruta)
                rutas.append(ruta)

        if not rutas:
            shutil.rmtree(tmp, ignore_errors=True)
            return "No se recibió ningún documento", 400
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return f"Ocurrió un error recibiendo los documentos: {e}", 500

    respuesta = Response(
        stream_with_context(_generar_eventos_validacion(tmp, rutas, candidato)),
        mimetype="text/event-stream",
    )
    # Evita que algún proxy intermedio (nginx, el balanceador de Render/Cloud
    # Run, etc.) guarde el stream en buffer y lo mande todo junto hasta el
    # final -eso volvería inútil la barra de progreso en tiempo real.
    respuesta.headers["Cache-Control"] = "no-cache"
    respuesta.headers["X-Accel-Buffering"] = "no"
    return respuesta


def _generar_eventos_validacion_lote(tmp_raiz, candidatos):
    """Generador de eventos SSE para /validar_lote (varios candidatos, un ZIP
    por candidato). 'candidatos' es una lista de dicts:
    {"nombre": str, "ruta_zip": str, "dir_extraccion": str}.

    Mismo motivo que _generar_eventos_validacion para usar streaming en vez
    de un hilo aparte: Cloud Run solo garantiza CPU mientras la petición HTTP
    esté abierta, y un lote de hasta 10 candidatos puede tardar varios
    minutos.

    Para no acumular en disco los PDFs extraídos de los 10 ZIPs al mismo
    tiempo, se procesa un candidato a la vez y se borra su carpeta de
    extracción antes de pasar al siguiente (no hasta el final del lote).
    """
    total_general = 0
    pdfs_por_candidato = {}
    try:
        # Primera pasada: contar cuántos PDFs trae cada ZIP para poder
        # mostrar un total global correcto desde el primer evento de avance
        # (en vez de ir descubriendo el total sobre la marcha).
        for cand in candidatos:
            try:
                nombres_pdf = ve.listar_pdfs_en_zip(cand["ruta_zip"])
            except Exception as e:
                nombres_pdf = None
                cand["error_zip"] = str(e)
            pdfs_por_candidato[cand["nombre"]] = nombres_pdf
            total_general += len(nombres_pdf) if nombres_pdf else 0
        total_general += len(candidatos)  # +1 paso por candidato: su Excel/consolidado

        hecho_general = 0
        resultados = []
        candidato_total = len(candidatos)

        for idx, cand in enumerate(candidatos, start=1):
            nombre = cand["nombre"]
            nombres_pdf = pdfs_por_candidato.get(nombre)

            if nombres_pdf is None:
                # El ZIP de este candidato venía corrupto o no era un ZIP
                # real: se reporta como error y se sigue con los demás en
                # vez de tronar todo el lote.
                hecho_general += 1
                yield _evento_sse({
                    "tipo": "progreso", "hecho": hecho_general, "total": total_general,
                    "candidato_index": idx, "candidato_total": candidato_total,
                    "candidato_actual": nombre,
                    "archivo_actual": f"Error: el ZIP no se pudo leer ({cand.get('error_zip', 'archivo dañado')})",
                })
                filas_error = [{
                    "archivo": "(ZIP)", "categoria_clave": "ERROR",
                    "categoria": "Error al procesar", "num_paginas": 0, "uso_ocr": False,
                    "legible": False, "texto_muestra": "", "nombre_coincide": None,
                    "detalle": f"No se pudo leer el ZIP: {cand.get('error_zip', 'archivo dañado')}",
                }]
                candidato_dict = {"nombre": nombre, "rfc": "", "curp": ""}
                # Se construye el checklist igual que con un candidato normal
                # (en vez de dejarlo vacío) para que el resumen general marque
                # correctamente a TODOS los documentos obligatorios como
                # faltantes -un checklist vacío se leería como "COMPLETO", que
                # sería incorrecto: en realidad no se pudo revisar nada.
                checklist_error, extra_error = ve.construir_reporte(candidato_dict, filas_error)
                resultados.append({
                    "candidato": candidato_dict,
                    "filas": filas_error,
                    "checklist": checklist_error, "extra": extra_error,
                })
                continue

            print(f"[validar_lote] candidato {idx}/{candidato_total} '{nombre}': {len(nombres_pdf)} documento(s)", file=sys.stderr, flush=True)
            yield _evento_sse({
                "tipo": "progreso", "hecho": hecho_general, "total": total_general,
                "candidato_index": idx, "candidato_total": candidato_total,
                "candidato_actual": nombre, "archivo_actual": "Iniciando...",
            })

            generador = ve.procesar_zip_candidato_progresivo(
                cand["ruta_zip"], nombre, cand["dir_extraccion"], nombres_pdf=nombres_pdf,
            )
            filas = None
            while True:
                try:
                    nombre_archivo = next(generador)
                except StopIteration as parada:
                    filas = parada.value
                    break
                yield _evento_sse({
                    "tipo": "progreso", "hecho": hecho_general, "total": total_general,
                    "candidato_index": idx, "candidato_total": candidato_total,
                    "candidato_actual": nombre, "archivo_actual": nombre_archivo,
                })
                hecho_general += 1
                yield _evento_sse({
                    "tipo": "progreso", "hecho": hecho_general, "total": total_general,
                    "candidato_index": idx, "candidato_total": candidato_total,
                    "candidato_actual": nombre, "archivo_actual": nombre_archivo,
                })

            checklist, extra = ve.construir_reporte({"nombre": nombre, "rfc": "", "curp": ""}, filas or [])
            resultados.append({
                "candidato": {"nombre": nombre, "rfc": "", "curp": ""},
                "filas": filas or [], "checklist": checklist, "extra": extra,
            })

            hecho_general += 1  # paso de "consolidado" de este candidato
            yield _evento_sse({
                "tipo": "progreso", "hecho": hecho_general, "total": total_general,
                "candidato_index": idx, "candidato_total": candidato_total,
                "candidato_actual": nombre, "archivo_actual": "Listo",
            })

            # Se borra la carpeta de extracción de ESTE candidato antes de
            # seguir con el siguiente para no acumular en disco los PDFs
            # extraídos de los 10 ZIPs a la vez.
            shutil.rmtree(cand["dir_extraccion"], ignore_errors=True)

        yield _evento_sse({
            "tipo": "progreso", "hecho": hecho_general, "total": total_general,
            "candidato_total": candidato_total, "archivo_actual": "Generando Excel consolidado...",
        })

        salida = os.path.join(tmp_raiz, "expedientes_lote.xlsx")
        ve.generar_excel_lote(resultados, salida)

        with open(salida, "rb") as fh:
            data = fh.read()

        nombre_archivo_final = f"expedientes_{len(candidatos)}_candidatos.xlsx"
        yield _evento_sse({
            "tipo": "listo",
            "hecho": total_general,
            "total": total_general,
            "nombre_archivo": nombre_archivo_final,
            "datos_base64": base64.b64encode(data).decode("ascii"),
        })

        print(f"[validar_lote] lote de {len(candidatos)} candidato(s) listo", file=sys.stderr, flush=True)
    except Exception as e:
        print(f"[validar_lote] ERROR inesperado en el lote: {e}", file=sys.stderr, flush=True)
        yield _evento_sse({"tipo": "error", "error": str(e)})
    finally:
        shutil.rmtree(tmp_raiz, ignore_errors=True)


@app.route("/validar_lote", methods=["POST"])
@requiere_login
def validar_lote():
    tmp_raiz = tempfile.mkdtemp(prefix="lote_")
    try:
        candidatos = []
        for i in range(1, 11):
            nombre = request.form.get(f"nombre_{i}", "").strip()
            archivo = request.files.get(f"zip_{i}")
            tiene_archivo = bool(archivo and archivo.filename)

            if not nombre and not tiene_archivo:
                continue  # fila vacía: se ignora sin quejarse

            if not nombre or not tiene_archivo:
                shutil.rmtree(tmp_raiz, ignore_errors=True)
                return (
                    f"La fila {i} está incompleta: hace falta "
                    + ("el nombre del candidato" if not nombre else "el archivo ZIP")
                    + ". Llena ambos campos de esa fila o déjala vacía por completo.",
                    400,
                )

            dir_candidato = os.path.join(tmp_raiz, f"candidato_{i}")
            os.makedirs(dir_candidato, exist_ok=True)
            ruta_zip = os.path.join(dir_candidato, "expediente.zip")
            archivo.save(ruta_zip)
            dir_extraccion = os.path.join(dir_candidato, "extraidos")
            os.makedirs(dir_extraccion, exist_ok=True)
            candidatos.append({"nombre": nombre, "ruta_zip": ruta_zip, "dir_extraccion": dir_extraccion})

        if not candidatos:
            shutil.rmtree(tmp_raiz, ignore_errors=True)
            return "No se recibió ningún candidato (nombre + ZIP)", 400

        if len(candidatos) > 10:
            shutil.rmtree(tmp_raiz, ignore_errors=True)
            return "Solo se permiten hasta 10 candidatos por tanda", 400
    except Exception as e:
        shutil.rmtree(tmp_raiz, ignore_errors=True)
        return f"Ocurrió un error recibiendo los archivos: {e}", 500

    respuesta = Response(
        stream_with_context(_generar_eventos_validacion_lote(tmp_raiz, candidatos)),
        mimetype="text/event-stream",
    )
    respuesta.headers["Cache-Control"] = "no-cache"
    respuesta.headers["X-Accel-Buffering"] = "no"
    return respuesta


if __name__ == "__main__":
    # En Render (y la mayoría de hostings) el puerto real llega en la
    # variable de entorno PORT; localmente usamos 5000 por default.
    puerto = int(os.environ.get("PORT", 5000))
    debug_local = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug_local, host="0.0.0.0", port=puerto)
