"""
App web para generar contratos a partir de plantillas Word.
Para añadir un nuevo modelo de contrato, solo hay que:
  1. Añadir una entrada al diccionario MODELOS de abajo.
  2. Crear la plantilla .docx en la carpeta /modelos/ con los marcadores {{ campo }}.
  3. Crear el formulario HTML en /templates/.
"""

from flask import Flask, render_template, request, send_file, abort, session, redirect, url_for
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime, timedelta
from functools import wraps
import os

app = Flask(__name__)

# -----------------------------------------------------------------------------
# AUTENTICACIÓN (contraseña compartida)
# -----------------------------------------------------------------------------
# En Render se configuran SECRET_KEY y APP_PASSWORD como variables de entorno
# (Settings → Environment). En local se usan los valores por defecto, así
# puedes probar la app sin tener que configurar nada.
# -----------------------------------------------------------------------------
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-key-change-in-production")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "dev")

# La sesión dura 7 días (luego pide login otra vez).
app.permanent_session_lifetime = timedelta(days=7)


def login_required(f):
    """Decorador: si no hay sesión activa, redirige a /login."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE MODELOS DE CONTRATO
# -----------------------------------------------------------------------------
# Para añadir un nuevo modelo, copia el bloque de "compraventa_inmueble" y
# adáptalo. Los nombres de los campos deben coincidir EXACTAMENTE con los
# marcadores {{ campo }} usados en la plantilla .docx y en el formulario HTML.
# -----------------------------------------------------------------------------
MODELOS = {
    "compraventa_inmueble": {
        "nombre": "Compraventa de inmueble",
        "descripcion": "Contrato privado de compraventa de bien inmueble (legislación española).",
        "plantilla": "compraventa_inmueble.docx",
        "formulario": "compraventa_inmueble.html",
        "campos": [
            "lugar_firma", "fecha_firma",
            "vendedor_nombre", "vendedor_dni", "vendedor_domicilio",
            "comprador_nombre", "comprador_dni", "comprador_domicilio",
            "inmueble_descripcion", "inmueble_direccion",
            "registro_propiedad", "registro_tomo", "registro_libro",
            "registro_folio", "registro_finca", "referencia_catastral",
            "precio_compra_numero", "precio_compra_letras",
            "forma_pago", "fecha_entrega", "jurisdiccion",
        ],
    },
    "nda_confidencialidad": {
        "nombre": "Acuerdo de confidencialidad (NDA)",
        "descripcion": "Acuerdo bilateral de confidencialidad entre dos partes (legislación española).",
        "plantilla": "nda_confidencialidad.docx",
        "formulario": "nda_confidencialidad.html",
        "campos": [
            "lugar_firma", "fecha_firma",
            "parte_a_nombre", "parte_a_dni", "parte_a_domicilio", "parte_a_representante",
            "parte_b_nombre", "parte_b_dni", "parte_b_domicilio", "parte_b_representante",
            "objeto_relacion", "duracion_acuerdo", "plazo_confidencialidad",
            "penalizacion", "jurisdiccion",
        ],
    },
    "prestacion_servicios": {
        "nombre": "Prestación de servicios profesionales (autónomo)",
        "descripcion": "Contrato mercantil de prestación de servicios profesionales para freelance / autónomos (legislación española).",
        "plantilla": "prestacion_servicios.docx",
        "formulario": "prestacion_servicios.html",
        "campos": [
            "lugar_firma", "fecha_firma",
            "prestador_nombre", "prestador_dni", "prestador_domicilio", "prestador_actividad",
            "cliente_nombre", "cliente_dni", "cliente_domicilio", "cliente_representante",
            "objeto_servicios", "entregables",
            "fecha_inicio", "fecha_fin",
            "precio_total_numero", "precio_total_letras", "iva_tipo", "forma_pago",
            "plazo_confidencialidad", "preaviso_resolucion", "jurisdiccion",
        ],
    },
}

CARPETA_MODELOS = os.path.join(os.path.dirname(__file__), "modelos")

# Texto que aparecerá en el contrato cuando un campo se deje vacío.
# Aplica a TODOS los modelos, presentes y futuros.
PLACEHOLDER_VACIO = "[pendiente de completar]"


@app.route("/login", methods=["GET", "POST"])
def login():
    """Pantalla de login con contraseña compartida."""
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["logged_in"] = True
            session.permanent = True
            return redirect(url_for("index"))
        return render_template("login.html", error=True)
    return render_template("login.html", error=False)


@app.route("/logout")
def logout():
    """Cierra la sesión y vuelve al login."""
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    """Página principal: muestra la lista de modelos disponibles."""
    return render_template("index.html", modelos=MODELOS)


@app.route("/contrato/<modelo_id>")
@login_required
def formulario(modelo_id):
    """Muestra el formulario de un modelo concreto."""
    if modelo_id not in MODELOS:
        abort(404)
    modelo = MODELOS[modelo_id]
    return render_template(modelo["formulario"], modelo=modelo, modelo_id=modelo_id)


@app.route("/generar/<modelo_id>", methods=["POST"])
@login_required
def generar(modelo_id):
    """Toma los datos del formulario, rellena la plantilla y devuelve el .docx."""
    if modelo_id not in MODELOS:
        abort(404)

    modelo = MODELOS[modelo_id]
    # Si un campo se deja vacío, se sustituye por "[pendiente de completar]"
    datos = {
        campo: (request.form.get(campo, "").strip() or PLACEHOLDER_VACIO)
        for campo in modelo["campos"]
    }

    # Cargar y rellenar la plantilla
    plantilla_path = os.path.join(CARPETA_MODELOS, modelo["plantilla"])
    if not os.path.exists(plantilla_path):
        return (
            f"No se encuentra la plantilla {modelo['plantilla']}. "
            f"Asegúrate de haber ejecutado primero crear_plantilla.py.",
            500,
        )

    doc = DocxTemplate(plantilla_path)
    doc.render(datos)

    # Guardar en memoria y enviar al navegador
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    nombre_archivo = (
        f"contrato_{modelo_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    )
    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    # En local: abrir http://127.0.0.1:5000 en el navegador
    app.run(debug=True, host="0.0.0.0", port=5000)
