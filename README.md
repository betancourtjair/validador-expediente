# Validador de Expediente — Fitness Para Todos

Aplicación web (Flask) que valida los documentos de un candidato y genera un
Excel para el expediente. Necesita un hosting con soporte para Python +
Docker porque usa OCR (tesseract) para leer documentos escaneados.

## Contenido de esta carpeta

- `app.py` — la aplicación web (formulario de carga + generación del Excel).
- `validador_expediente.py` — el motor de validación (reglas, OCR, Excel).
- `requirements.txt` — librerías de Python necesarias.
- `Dockerfile` — instrucciones para construir la aplicación con todo lo que
  necesita (incluye tesseract-ocr y poppler-utils a nivel de sistema).
- `.gitignore` — archivos que no deben subirse al repositorio.

## Desplegar en Render (paso a paso)

### 1. Sube este código a GitHub

1. Entra a [github.com](https://github.com) y crea una cuenta gratuita si no tienes.
2. Da clic en el botón verde **"New"** (o el ícono "+" arriba a la derecha → "New repository").
3. Ponle un nombre, por ejemplo `validador-expediente`. Puede quedar como **privado** (recomendado, ya que el código no debe ser público).
4. Crea el repositorio. En la página del repo vacío, usa la opción **"uploading an existing file"** y arrastra los archivos de esta carpeta (`app.py`, `validador_expediente.py`, `requirements.txt`, `Dockerfile`, `.gitignore`).
5. Da clic en **"Commit changes"**.

### 2. Crea la cuenta en Render y conecta el repositorio

1. Entra a [render.com](https://render.com) y crea una cuenta — puedes usar **"Sign up with GitHub"** para conectar ambas cuentas en un solo paso.
2. En el panel, da clic en **"New +"** → **"Web Service"**.
3. Elige el repositorio `validador-expediente` que acabas de subir (si no aparece, dale permiso a Render para verlo).
4. Render detecta automáticamente el `Dockerfile` y lo usa para construir la app — no necesitas configurar nada de "build command" o "start command".
5. Elige un nombre para el servicio (será parte de la URL, ej. `fpt-expedientes.onrender.com`).
6. En **"Instance Type"** elige el plan gratuito para empezar (puedes subir de plan después si el tráfico crece).

### 3. Configura el usuario y contraseña

Antes de darle "Deploy", ve a la sección **"Environment"** y agrega estas dos variables:

| Key | Value |
|---|---|
| `APP_USER` | el usuario que quieras (ej. `reclutamiento`) |
| `APP_PASSWORD` | una contraseña segura — **cámbiala**, no dejes la de ejemplo |

### 4. Despliega

1. Da clic en **"Create Web Service"**.
2. Render construye la imagen (tarda 3-5 minutos la primera vez, porque instala tesseract).
3. Cuando termine, te da una URL tipo `https://fpt-expedientes.onrender.com`.
4. Ábrela: el navegador te pedirá el usuario y contraseña que configuraste en el paso 3.

### 5. Úsala desde tu página web

Puedes simplemente compartir el enlace con el equipo de reclutamiento, o poner
un botón/enlace en tu sitio actual que lleve a esa URL (por ejemplo
"Validar expediente de candidato" → `https://fpt-expedientes.onrender.com`).

## Notas importantes

- **Plan gratuito de Render:** el servicio "se duerme" tras ~15 minutos sin
  uso y tarda unos segundos en despertar la siguiente vez que alguien lo
  abre. Para uso interno de un equipo de reclutamiento esto normalmente no es
  problema; si más adelante quieres que esté siempre despierto, se necesita
  un plan pago (unos $7 USD/mes).
- **Cambia la contraseña de ejemplo.** Si despliegas sin configurar
  `APP_PASSWORD`, la app usa una contraseña por default que cualquiera que
  vea el código podría adivinar.
- **Datos sensibles:** los documentos se procesan en memoria/temporalmente y
  no se guardan en ningún lado — pero el Excel resultante sí queda en el
  navegador de quien lo descargue, así que trátalo como el archivo
  confidencial que es.
- Si más adelante quieres agregar más reglas o categorías de documentos,
  todo lo que hay que tocar está en `validador_expediente.py`
  (diccionario `CATEGORIAS` y funciones `analiza_*`).
