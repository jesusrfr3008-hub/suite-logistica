# Suite Logística — Prototipo (Módulos: Proveedores + Compras)

Prototipo funcional de la suite de control logístico. Esta primera entrega
incluye dos módulos completos y operativos, con el resto de la suite
mapeado en el dashboard como "próximamente":

1. **Mantenimiento de Proveedores**: crear, editar, desactivar/reactivar
   proveedores, con datos de contacto, dirección, correo y tipo
   (Nacional/Extranjero). Incluye el catálogo de productos de cada
   proveedor (código, descripción, empaque, moneda, precio por caja y por
   unidad), editable desde la ficha del proveedor.
2. **Compras / Órdenes de Compra**: crear una PO seleccionando un
   proveedor — la orden queda **personalizada** porque solo puedes agregar
   productos del catálogo de ese proveedor. Cada línea registra cantidad
   en cajas, precio pactado y fecha estimada de despacho. En el detalle de
   la orden puedes actualizar por línea la fecha de disponibilidad y de
   despacho **confirmadas por el proveedor**, cambiar el estado de la
   orden, imprimir/generar PDF y enviar la PO por correo (abre tu cliente
   de correo con el mensaje prellenado).

El modelo de datos ya incluye los campos que van a necesitar los próximos
módulos (despacho parcial por línea, órdenes consolidadas, transporte y
tracking) para no tener que rehacer el esquema más adelante — solo falta
construir sus pantallas.

## Origen de los datos

Los proveedores y productos se cargan automáticamente la primera vez que
arrancas la app, desde `Master_Lista_precios_Proveedores.xlsx` (hoja
`2026`, que ya consolida los 8 proveedores: DORC, VOLK, ELLEX, BVI BEAVER,
BVI PHYSIOL, BVI OPTIKON, QUANTEL y MEDICONTUR — 1514 productos en total).
Las hojas `QM` y `MEDICONTUR` del Excel no se importaron porque son
detalle/borrador de esos mismos dos proveedores, ya reflejado en la hoja
`2026`.

Los datos de contacto, dirección y correo de cada proveedor **no venían en
el Excel**, así que quedan en blanco: complétalos desde "Mantenimiento de
Proveedores" → Editar.

## Cómo ejecutar el prototipo en tu computador

Requisitos: Python 3.10 o superior.

```bash
cd suite_logistica
python3 -m venv venv
source venv/bin/activate        # en Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abre tu navegador en **http://127.0.0.1:5000**

La primera vez que arranca, crea la base de datos `data/logistica.db` y la
llena con el catálogo del Excel. Las siguientes veces reutiliza esa misma
base, así que tus proveedores y órdenes quedan guardados entre sesiones.

Si quieres reiniciar todo desde cero (volver a importar el Excel), borra
el archivo `data/logistica.db` y vuelve a ejecutar `python app.py`.

## Estructura del proyecto

```
suite_logistica/
├── app.py                # Rutas y lógica de la aplicación (Flask)
├── models.py              # Modelos de datos (SQLAlchemy)
├── seed_data.py            # Importación inicial desde el Excel
├── requirements.txt
├── Master_Lista_precios_Proveedores.xlsx   # Fuente de datos inicial
├── data/
│   └── logistica.db        # Base de datos SQLite (se crea al arrancar)
├── static/css/style.css
└── templates/              # Vistas (Proveedores, Compras, Dashboard)
```

## Próximos módulos (siguientes iteraciones)

- Despacho parcial de órdenes y control de saldos pendientes por línea
  (el modelo ya tiene `cantidad_despachada_cajas` listo para esto).
- Despachos consolidados / saldos de órdenes anteriores agrupados con
  nuevas (el modelo ya tiene `orden_consolidada_id`).
- Transporte: vía (aéreo/marítimo/terrestre), N° de tracking/AWB, e
  integración de seguimiento online.
- Asignación de costos asociados y gastos de aduana.
- Reportes de compras, costos de importación por concepto.
- Análisis de ventas/rotación para proyección de compras y estimación de
  venta por ejecutivo.
- Dashboards de control con notificaciones automáticas.

## Notas técnicas

- Backend: Flask + Flask-SQLAlchemy + SQLite (fácil de migrar a
  PostgreSQL/MySQL más adelante si el proyecto crece).
- Frontend: plantillas Jinja2 + Bootstrap 5 (vía CDN) — sin build step,
  fácil de ajustar.
- El envío de correo usa un enlace `mailto:` (abre tu cliente de correo
  con el mensaje prellenado). Si más adelante quieres envío automático
  real (módulo 9), se puede integrar SMTP o un proveedor como
  SendGrid/Gmail API.
- Este es un servidor de desarrollo (`debug=True`); no usar tal cual en
  producción — para eso se recomienda un servidor WSGI (gunicorn/uwsgi)
  detrás de un proxy.
