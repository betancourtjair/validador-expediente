#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mini aplicación web para el Validador de Expediente de Reclutamiento.

A diferencia del validador_expediente.html de la primera versión (que corría
100% en el navegador), esta parte SÍ necesita un servidor con Python porque
hace OCR y lectura de PDFs en el servidor — eso no se puede hacer de forma
confiable solo con JavaScript en el navegador para documentos escaneados.

Cómo correrlo
--------------
    pip install flask pdfplumber pytesseract pdf2image openpyxl
    (y en el sistema: tesseract-ocr, tesseract-ocr-spa, poppler-utils)

    python3 app.py

Abre http://localhost:5000 en el navegador. Sube los documentos del
candidato y descarga el Excel del expediente.

Para montarlo en tu página web necesitas un hosting que soporte Python
(Render, Railway, un VPS con Gunicorn + Nginx, etc.) — no es un simple
archivo HTML que se sube a cualquier hosting estático.
"""

import base64
import json
import os
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
# Usuario y contraseña (protección básica del formulario, vía HTTP Basic Auth)
# ---------------------------------------------------------------------------
# El usuario y la contraseña NO van escritos en el código: se configuran como
# variables de entorno en el panel de Render (Settings -> Environment).
# Si no se configuran, se usan estos valores por default SOLO para que la
# app no truene al desplegarla por primera vez — cámbialos de inmediato en
# Render con APP_USER y APP_PASSWORD.
APP_USER = os.environ.get("APP_USER", "reclutamiento")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "cambia-esta-clave")


def requiere_login(f):
    @wraps(f)
    def decorado(*args, **kwargs):
        auth = request.authorization
        credenciales_ok = (
            auth
            and secrets.compare_digest(auth.username or "", APP_USER)
            and secrets.compare_digest(auth.password or "", APP_PASSWORD)
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

PAGINA = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Validador de Expediente</title>
<style>
  body{font-family:Arial,sans-serif;max-width:760px;margin:30px auto;color:#1f2530;}
  h1{font-size:1.4rem;}
  .campo{margin-bottom:14px;}
  label{display:block;font-weight:600;font-size:.85rem;margin-bottom:4px;}
  input[type=text]{width:100%;padding:8px;border:1px solid #ccc;border-radius:6px;}
  .doc{border:1px solid #e0e0e0;border-radius:8px;padding:10px 14px;margin-bottom:8px;}
  .doc.req{background:#f0f8f4;border-color:#bfe4d3;}
  .tag{font-size:.7rem;background:#0f6b4c;color:#fff;padding:2px 6px;border-radius:10px;margin-left:6px;}
  button{background:#0f6b4c;color:#fff;border:none;padding:12px 22px;border-radius:8px;font-size:1rem;cursor:pointer;}
  button:disabled{opacity:.6;cursor:not-allowed;}
  #estado{margin-top:14px;font-size:.9rem;font-weight:600;min-height:1.2em;}
  #barraContenedor{margin-top:12px;display:none;}
  progress#barra{width:100%;height:16px;border-radius:8px;overflow:hidden;}
  progress#barra::-webkit-progress-bar{background:#e6ece9;border-radius:8px;}
  progress#barra::-webkit-progress-value{background:#0f6b4c;border-radius:8px;transition:width .25s ease;}
  progress#barra::-moz-progress-bar{background:#0f6b4c;border-radius:8px;}
  #barraTexto{font-size:.8rem;color:#555;margin-top:4px;}
  .separador{border:none;border-top:2px solid #e0e0e0;margin:34px 0 22px;}
  .pista{background:#f0f8f4;border:1px solid #bfe4d3;border-radius:8px;padding:10px 14px;font-size:.82rem;margin-bottom:16px;line-height:1.5;}
  .pista code{background:#e6ece9;padding:1px 5px;border-radius:4px;}
  .lote-fila{display:flex;gap:10px;align-items:center;border:1px solid #e0e0e0;border-radius:8px;padding:8px 12px;margin-bottom:8px;}
  .lote-fila .num{font-weight:600;font-size:.85rem;color:#555;width:22px;flex:none;}
  .lote-fila input[type=text]{flex:1 1 auto;min-width:0;}
  .lote-fila input[type=file]{flex:1 1 auto;min-width:0;font-size:.82rem;}
  #avisoTamanoLote{display:none;background:#fff3cd;border:1px solid #ffe08a;color:#7a5b00;border-radius:8px;padding:10px 14px;font-size:.85rem;margin-bottom:12px;}
</style>
</head>
<body>
<h1>Validador de Expediente de Reclutamiento</h1>
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
    <label>{{ etiqueta }} {% if not obligatorio %}<span class="tag">Condicional</span>{% endif %}</label>
    <input type="file" name="{{ campo }}" accept="application/pdf">
  </div>
  {% endfor %}
  <div class="campo">
    <label>Otros documentos adicionales (puedes seleccionar varios)</label>
    <input type="file" name="otros" multiple accept="application/pdf">
  </div>
  <button type="submit" id="btnSubmit">Validar y generar Excel</button>
  <div id="estado"></div>
  <div id="barraContenedor">
    <progress id="barra" value="0" max="1"></progress>
    <div id="barraTexto"></div>
  </div>
</form>

<hr class="separador">

<h1>Carga masiva por ZIP (hasta 10 candidatos)</h1>
<div class="pista">
  Sube <strong>un archivo ZIP por candidato</strong>, con los PDFs de ese candidato
  adentro. El sistema identifica cada documento por el <strong>nombre del archivo</strong>
  dentro del ZIP (no hace falta acomodarlos en campos como arriba). Nombra los PDFs
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

<script>
// Ambos formularios (un candidato / carga masiva por ZIP) se mandan por
// fetch (en vez de un submit normal) para poder limpiarlos por completo
// -incluyendo los archivos ya seleccionados- justo después de descargar el
// Excel, y para poder mostrar una barra de progreso en vivo.
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
// progreso; se usa una vez por formulario (abajo) para no repetir esta
// lógica dos veces.
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
      var blob = base64ABlob(info.datos_base64, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      var url = window.URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = info.nombre_archivo || "expediente.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      // Limpia todos los campos y archivos seleccionados.
      form.reset();
      terminar("Listo: se descargó " + (info.nombre_archivo || "expediente.xlsx") + ". El formulario ya está limpio para la siguiente carga.", false);
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

configurarFormularioConProgreso({
  formId: "formValidador",
  botonId: "btnSubmit",
  estadoId: "estado",
  barraContenedorId: "barraContenedor",
  barraId: "barra",
  barraTextoId: "barraTexto",
  mensajeInicial: "Procesando documentos (puede tardar uno o dos minutos)...",
});

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
</body>
</html>
"""


@app.route("/", methods=["GET"])
@requiere_login
def index():
    return render_template_string(PAGINA, documentos=CAMPOS_DOCUMENTOS)


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

        with open(salida, "rb") as fh:
            data = fh.read()

        hecho = total
        nombre_archivo = f"expediente_{nombre.strip().replace(' ', '_') or 'candidato'}.xlsx"
        yield _evento_sse({
            "tipo": "listo",
            "hecho": hecho,
            "total": total,
            "nombre_archivo": nombre_archivo,
            "datos_base64": base64.b64encode(data).decode("ascii"),
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
@requiere_login
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
