# Generador de Contratos

App web sencilla para generar contratos a partir de plantillas Word. Pensada para uso interno entre compañeros: se rellena un formulario en el navegador y se descarga el contrato listo en formato `.docx`.

---

## 1. Cómo arrancarla en tu ordenador (paso a paso)

### Requisitos
- Python 3.10 o superior instalado.
- VS Code (opcional, para editar cómodamente).

### Pasos

1. **Abre la carpeta del proyecto en VS Code**
   `Archivo → Abrir carpeta → contratos-app`

2. **Abre la terminal integrada**
   `Terminal → Nueva terminal` (o pulsa `` Ctrl+` ``)

3. **Crea un entorno virtual** (aísla las librerías de este proyecto)

   En Windows:
   ```
   python -m venv .venv
   .venv\Scripts\activate
   ```
   En Mac/Linux:
   ```
   python3 -m venv .venv
   source .venv/bin/activate
   ```

4. **Instala las librerías**
   ```
   pip install -r requirements.txt
   ```

5. **Genera la plantilla Word del contrato** (solo la primera vez)
   ```
   python crear_plantilla.py
   ```
   Esto creará el archivo `modelos/compraventa_inmueble.docx`.

6. **Arranca la app**
   ```
   python app.py
   ```

7. **Abre el navegador** en `http://127.0.0.1:5000`

¡Listo! Verás la lista de modelos. Pincha en "Crear contrato", rellena los datos, y se descargará el Word generado.

---

## 2. Estructura del proyecto

```
contratos-app/
├── app.py                  # Servidor web (lógica principal)
├── crear_plantilla.py      # Script para generar la plantilla inicial
├── requirements.txt        # Lista de librerías de Python
├── Procfile                # Para desplegar en Render
├── modelos/                # Plantillas Word (.docx)
│   └── compraventa_inmueble.docx
└── templates/              # Páginas HTML
    ├── index.html
    └── compraventa_inmueble.html
```

---

## 3. Cómo añadir un nuevo modelo de contrato

Todo está pensado para que añadir más contratos sea sencillo. Tres pasos:

### Paso A · Crear la plantilla Word

Abre Word y crea un documento con el texto del contrato. En los puntos donde quieras que se rellenen datos, pon marcadores con el formato `{{ nombre_del_campo }}`.

Ejemplo:
> En **{{ lugar_firma }}**, a **{{ fecha_firma }}**, comparecen…

Guárdalo en la carpeta `modelos/` con un nombre claro, p. ej. `arrendamiento_vivienda.docx`.

### Paso B · Añadir el modelo en `app.py`

Abre `app.py` y, dentro del diccionario `MODELOS`, añade un bloque siguiendo el ejemplo del comentario. Lista todos los nombres de campo que has usado en la plantilla.

### Paso C · Crear el formulario HTML

Copia `templates/compraventa_inmueble.html`, renómbralo (p. ej. `arrendamiento_vivienda.html`) y adapta los campos del formulario para que los `name="..."` coincidan con los nombres usados en la plantilla.

¡Y ya está! Reinicia la app y aparecerá el nuevo modelo en la página principal.

> **Truco:** si no te ves capaz de hacer estos pasos solo, abre el archivo en VS Code, selecciona el código y pídele a Claude que lo adapte. Indícale qué campos tiene tu nueva plantilla y él hará la mayor parte del trabajo.

---

## 4. Subirlo a Render (para que tus compañeros lo usen desde Internet)

Render es un servicio gratuito en la nube. **No se descarga nada** — todo se hace desde su web.

### Paso 1 · Subir el código a GitHub
1. Crea una cuenta gratis en [github.com](https://github.com).
2. Crea un repositorio nuevo (privado si quieres) y sube los archivos del proyecto. Si no has usado Git nunca, GitHub Desktop es la opción más fácil.

### Paso 2 · Crear cuenta en Render
1. Entra en [render.com](https://render.com) y regístrate (puedes usar tu cuenta de GitHub).
2. Pulsa **New → Web Service**.
3. Conecta tu repositorio de GitHub.
4. Render detectará Python automáticamente. Verifica:
   - **Build Command:** `pip install -r requirements.txt && python crear_plantilla.py`
   - **Start Command:** `gunicorn app:app`
5. Elige el plan **Free** y pulsa **Create Web Service**.

En unos minutos tendrás una URL pública (algo como `https://contratos-app.onrender.com`) que puedes pasar a tus compañeros.

> **Aviso del plan gratuito:** la app se "duerme" tras 15 minutos sin uso. La primera petición tarda ~30 s en despertarla; las siguientes son instantáneas. Para uso interno ocasional sobra; si pasáis a uso intensivo, hay planes de pago a partir de unos 7 €/mes.

---

## 5. Recomendación importante

Esta app sirve para automatizar la **redacción** de contratos a partir de plantillas. **No sustituye el criterio profesional**: revisa siempre el contrato generado antes de usarlo y consulta a un abogado si la operación lo requiere.
