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
<script>
// Se manda el formulario por fetch (en vez de un submit normal) para poder
// limpiarlo por completo -incluyendo los archivos ya seleccionados- justo
// después de que se descargue el Excel, y así quede listo para capturar al
// siguiente candidato sin tener que quitar archivo por archivo a mano.
//
// El servidor va mandando el avance real (documento por documento) como un
// "stream" de eventos (Server-Sent Events) dentro de esta misma petición POST
// -no en peticiones aparte- para que, tanto en Render como en Cloud Run, el
// servidor siga teniendo CPU asignado mientras hace el OCR (si se regresara
// la respuesta de inmediato y el trabajo se hiciera "detrás" en otro lado,
// Cloud Run podría congelar el CPU del contenedor entre una consulta y otra).
(function () {
  var form = document.getElementById("formValidador");
  var boton = document.getElementById("btnSubmit");
  var estado = document.getElementById("estado");
  var barraContenedor = document.getElementById("barraContenedor");
  var barra = document.getElementById("barra");
  var barraTexto = document.getElementById("barraTexto");
  var yaTermino = false;

  function terminar(mensaje, esError) {
    yaTermino = true;
    estado.style.color = esError ? "#b3261e" : "#0f6b4c";
    estado.textContent = mensaje;
    barraContenedor.style.display = "none";
    boton.disabled = false;
    boton.textContent = "Validar y generar Excel";
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
      barraTexto.textContent = porcentaje + "% - " + (info.archivo_actual || "Procesando...");
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

      // Limpia nombre/RFC/CURP y TODOS los archivos seleccionados.
      form.reset();
      terminar("Listo: se descargó " + (info.nombre_archivo || "expediente.xlsx") + ". El formulario ya está limpio para el siguiente candidato.", false);
    } else if (info.tipo === "error") {
      terminar("Ocurrió un error generando el expediente: " + (info.error || "error desconocido") + " (no se borró nada, puedes reintentar).", true);
    }
  }

  form.addEventListener("submit", function (ev) {
    ev.preventDefault();
    estado.style.color = "#0f6b4c";
    estado.textContent = "Procesando documentos (puede tardar uno o dos minutos)...";
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
})();
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


if __name__ == "__main__":
    # En Render (y la mayoría de hostings) el puerto real llega en la
    # variable de entorno PORT; localmente usamos 5000 por default.
    puerto = int(os.environ.get("PORT", 5000))
    debug_local = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug_local, host="0.0.0.0", port=puerto)
