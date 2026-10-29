# constructor | microsaas

Un panel de gestión de negocio ligero construido con Python, Flask y SQLite. Sin frameworks, sin npm, sin pasos de compilación — solo ejecútalo.

## Qué hace

Gestiona las tres entidades que cualquier negocio pequeño necesita de verdad:

- **Clientes** — nombre, email, teléfono, dirección
- **Productos** — descripción, precio, stock
- **Pedidos** — vinculados a un cliente, con múltiples líneas de pedido y totales calculados automáticamente

Cada entidad tiene CRUD completo. Los pedidos y sus líneas recalculan el total en cada guardado. Puedes exportar cualquier tabla a **Excel (.xlsx)** o **PDF** directamente desde la vista de listado.

Incluye pantalla de login con gestión de sesiones y un dashboard con resumen de conteos y facturación total.

## Stack

- **Backend:** Python 3 + Flask + SQLite (con el módulo estándar `sqlite3`)
- **Frontend:** HTML + CSS puro, plantillas Jinja2 — cero JavaScript
- **Exportaciones:** `openpyxl` para Excel, `reportlab` para PDF

## Cómo empezar

```bash
git clone https://github.com/ConstructorQP/flask-microsaas
cd flask-microsaas
pip install -r requirements.txt
python app.py
```

Abre `http://127.0.0.1:5000` en el navegador.

Credenciales de la cuenta demo (se crean al arrancar por primera vez):
```
usuario: constructor
contraseña: tryme
```

> Son credenciales de demostración para probar la app en local. No uses esta configuración en producción.

## Estructura del proyecto

```
app.py              # Todas las rutas y la lógica de base de datos
templates/          # Plantillas HTML con Jinja2
  base.html         # Layout con sidebar y topbar
  login.html        # Login con pantalla dividida
  dashboard.html    # Vista resumen con estadísticas
  clientes.html / cliente_form.html
  productos.html / producto_form.html
  pedidos.html / pedido_form.html / pedido_ver.html
static/
  css/style.css     # Todos los estilos en un único archivo
  logo.png          # Pon aquí tu logotipo
requirements.txt
```

## Logotipo

Coloca un archivo `logo.png` dentro de `static/`. La pantalla de login y el sidebar lo mostrarán automáticamente. Si no existe, se muestra el nombre de la marca en texto.

## Notas

- El archivo de base de datos (`database.db`) se crea automáticamente al arrancar por primera vez
- Los totales de pedido se recalculan en cada guardado — sin datos desactualizados
- El formulario de pedido tiene 5 filas de línea estáticas por diseño; las filas vacías se ignoran
- No hay migraciones para este alcance; si cambias el esquema, borra `database.db` y reinicia

## Seguridad

- Las contraseñas se guardan con hash (`werkzeug.security`), nunca en texto plano.
- La clave de sesión se lee de la variable de entorno `SECRET_KEY`. Si no está definida, se genera una aleatoria en cada arranque.
- Las consultas a la base de datos usan parámetros (`?`) para evitar inyección SQL.
- El modo debug solo se activa con `FLASK_DEBUG=1`.

```bash
export SECRET_KEY="una-clave-larga-y-aleatoria"
python app.py
```

## Licencia

MIT
