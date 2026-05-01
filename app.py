"""
App web para generar contratos a partir de plantillas Word.
Para añadir un nuevo modelo de contrato, solo hay que:
  1. Añadir una entrada al diccionario MODELOS de abajo.
  2. Crear la plantilla .docx en la carpeta /modelos/ con los marcadores {{ campo }}.
  3. Crear el formulario HTML en /templates/.
"""

from flask import Flask, render_template, request, send_file, abort
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

app = Flask(__name__)

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
    # Ejemplo de cómo añadir un nuevo modelo en el futuro:
    # "arrendamiento_vivienda": {
    #     "nombre": "Arrendamiento de vivienda",
    #     "descripcion": "Contrato de alquiler de vivienda habitual.",
    #     "plantilla": "arrendamiento_vivienda.docx",
    #     "formulario": "arrendamiento_vivienda.html",
    #     "campos": ["arrendador_nombre", "arrendatario_nombre", ...],
    # },
}

CARPETA_MODELOS = os.path.join(os.path.dirname(__file__), "modelos")


@app.route("/")
def index():
    """Página principal: muestra la lista de modelos disponibles."""
    return render_template("index.html", modelos=MODELOS)


@app.route("/contrato/<modelo_id>")
def formulario(modelo_id):
    """Muestra el formulario de un modelo concreto."""
    if modelo_id not in MODELOS:
        abort(404)
    modelo = MODELOS[modelo_id]
    return render_template(modelo["formulario"], modelo=modelo, modelo_id=modelo_id)


@app.route("/generar/<modelo_id>", methods=["POST"])
def generar(modelo_id):
    """Toma los datos del formulario, rellena la plantilla y devuelve el .docx."""
    if modelo_id not in MODELOS:
        abort(404)

    modelo = MODELOS[modelo_id]
    datos = {campo: request.form.get(campo, "").strip() for campo in modelo["campos"]}

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
