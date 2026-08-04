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

import io
import os
import secrets
import sys
import tempfile
import time
from functools import wraps

from flask import Flask, request, send_file, render_template_string, Response

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
</form>
<script>
// Se manda el formulario por fetch (en vez de un submit normal) para poder
// limpiarlo por completo -incluyendo los archivos ya seleccionados- justo
// después de que se descargue el Excel, y así quede listo para capturar al
// siguiente candidato sin tener que quitar archivo por archivo a mano.
(function () {
  var form = document.getElementById("formValidador");
  var boton = document.getElementById("btnSubmit");
  var estado = document.getElementById("estado");

  form.addEventListener("submit", function (ev) {
    ev.preventDefault();
    estado.style.color = "#0f6b4c";
    estado.textContent = "Procesando documentos (puede tardar uno o dos minutos)...";
    boton.disabled = true;
    boton.textContent = "Procesando...";

    var datos = new FormData(form);

    fetch(form.getAttribute("action"), { method: "POST", body: datos })
      .then(function (resp) {
        if (!resp.ok) {
          return resp.text().then(function (txt) {
            throw new Error(txt || ("Error del servidor (" + resp.status + ")"));
          });
        }
        var disposicion = resp.headers.get("Content-Disposition") || "";
        var m = disposicion.match(/filename="?([^"]+)"?/);
        var nombreArchivo = m ? m[1] : "expediente.xlsx";
        return resp.blob().then(function (blob) {
          return { blob: blob, nombre: nombreArchivo };
        });
      })
      .then(function (resultado) {
        var url = window.URL.createObjectURL(resultado.blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = resultado.nombre;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);

        // Limpia nombre/RFC/CURP y TODOS los archivos seleccionados.
        form.reset();
        estado.style.color = "#0f6b4c";
        estado.textContent = "Listo: se descargó " + resultado.nombre + ". El formulario ya está limpio para el siguiente candidato.";
      })
      .catch(function (err) {
        estado.style.color = "#b3261e";
        estado.textContent = "Ocurrió un error: " + err.message + " (no se borró nada, puedes reintentar).";
      })
      .finally(function () {
        boton.disabled = false;
        boton.textContent = "Validar y generar Excel";
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


@app.route("/validar", methods=["POST"])
@requiere_login
def validar():
    nombre = request.form.get("nombre", "").strip()
    rfc = request.form.get("rfc", "").strip()
    curp = request.form.get("curp", "").strip()
    if not nombre:
        return "Falta el nombre del candidato", 400

    candidato = {"nombre": nombre, "rfc": rfc, "curp": curp}

    # Todo lo que se sube (PDFs) y el Excel generado viven SOLO dentro de este
    # directorio temporal: al salir del "with" (ya sea que todo salga bien o
    # truene una excepción) Python lo borra del disco del servidor por
    # completo. O sea que en el servidor nunca queda nada guardado de un
    # candidato al siguiente; lo que faltaba para "empezar de cero" era
    # limpiar el FORMULARIO en el navegador después de descargar el Excel
    # (ver el <script> en PAGINA, que hace form.reset() tras la descarga).
    try:
        with tempfile.TemporaryDirectory() as tmp:
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
                return "No se recibió ningún documento", 400

            # Log de progreso a stdout: en el plan gratuito de Render el OCR
            # puede tardar bastante (poco CPU disponible). Sin esto, si algo se
            # traba no hay ninguna pista en los logs de qué archivo fue.
            print(f"[validar] iniciando expediente de '{nombre}' con {len(rutas)} documento(s)", file=sys.stderr, flush=True)

            filas = []
            for ruta in rutas:
                t0 = time.time()
                print(f"[validar]   procesando {os.path.basename(ruta)} ...", file=sys.stderr, flush=True)
                try:
                    fila = ve.procesar_documento(ruta, candidato["nombre"])
                except Exception as e:
                    fila = {
                        "archivo": os.path.basename(ruta), "categoria_clave": "ERROR",
                        "categoria": "Error al procesar", "num_paginas": 0, "uso_ocr": False,
                        "legible": False, "texto_muestra": "", "nombre_coincide": None,
                        "detalle": f"Error: {e}",
                    }
                print(f"[validar]   listo {os.path.basename(ruta)} ({time.time() - t0:.1f}s)", file=sys.stderr, flush=True)
                filas.append(fila)

            checklist, extra = ve.construir_reporte(candidato, filas)

            salida = os.path.join(tmp, "expediente.xlsx")
            ve.generar_excel(candidato, filas, checklist, extra, salida)

            with open(salida, "rb") as fh:
                data = fh.read()

            print(f"[validar] expediente de '{nombre}' listo", file=sys.stderr, flush=True)
    except Exception as e:
        # Cualquier error inesperado (no solo los de un documento en
        # particular, que ya se capturan arriba) se reporta como texto plano
        # en vez de dejar pasar la página de error genérica de Flask — así el
        # aviso que ve la persona en el formulario es entendible.
        print(f"[validar] ERROR inesperado con el expediente de '{nombre}': {e}", file=sys.stderr, flush=True)
        return f"Ocurrió un error inesperado generando el expediente: {e}", 500

    nombre_archivo = f"expediente_{nombre.strip().replace(' ', '_') or 'candidato'}.xlsx"
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    # En Render (y la mayoría de hostings) el puerto real llega en la
    # variable de entorno PORT; localmente usamos 5000 por default.
    puerto = int(os.environ.get("PORT", 5000))
    debug_local = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug_local, host="0.0.0.0", port=puerto)
