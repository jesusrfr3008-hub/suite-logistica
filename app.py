import csv
import glob
import io
import os
import re
import shutil
import uuid
from datetime import datetime, date
from urllib.parse import quote

import openpyxl
from flask import (
    Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, abort
)

from models import (
    db, Empresa, Proveedor, Producto, ProductoVariante, OrdenCompra, OrdenCompraLinea, OrdenDocumento,
    ESTADOS_OC, ETAPAS_LINEA, TIPOS_DOCUMENTO_ORDEN,
    Importacion, Parcial, ParcialLinea, ParcialLineaLote, GastoImportacion, GastoDocumento,
    ImportacionDocumento,
    REGIMENES_PARCIAL, VIAS_EMBARQUE, CONDICIONES_COMPRA, CONCEPTOS_GASTO,
    CONCEPTOS_ITEM_FACTURA, TIPOS_DOCUMENTO_GASTO,
    Despacho, ESTADOS_DESPACHO,
    CargoAdicionalImportacion, TipoCambioMensual,
)
from seed_data import seed_from_excel
import costing

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCUMENTOS_DIR = os.path.join(BASE_DIR, "data", "documentos")
EXTENSIONES_PERMITIDAS = {".pdf", ".jpg", ".jpeg", ".png"}
# Documentos de Gastos (ronda G, punto 7): ademas de PDF/imagen, el usuario
# pidio poder adjuntar Word y Excel (suele recibir el respaldo de un gasto
# en esos formatos, no solo PDF).
DOCUMENTOS_GASTOS_DIR = os.path.join(BASE_DIR, "data", "documentos_gastos")
EXTENSIONES_PERMITIDAS_GASTO = {".pdf", ".jpg", ".jpeg", ".png", ".doc", ".docx", ".xls", ".xlsx"}
# "Legajo": archivo consolidado a nivel de toda la importacion (ronda H,
# punto 3, opcion "Adjuntar > Legajo") -- mismas extensiones que los
# documentos de Gasto.
DOCUMENTOS_LEGAJO_DIR = os.path.join(BASE_DIR, "data", "documentos_legajo")
# Logos de las Empresas compradoras (ronda K, 2026-09-07, punto 1/7).
EMPRESAS_LOGOS_DIR = os.path.join(BASE_DIR, "data", "logos_empresas")
# Listado de variantes de lentes MEDICONTUR (ronda M, 2026-09-10, punto 1):
# mapea cada "codigo padre" del catalogo (ej. "677ADY") a sus codigos de
# dioptria/variante reales -- ver seed_variantes_lentes_medicontur() abajo.
VARIANTES_LENTES_MEDICONTUR_EXCEL = os.path.join(BASE_DIR, "Listado codigos lentes medicontur.xlsx")
EXTENSIONES_LOGO_PERMITIDAS = {".png", ".jpg", ".jpeg", ".svg"}
# PDF de la Orden de Compra generado en disco para poder adjuntarlo a un
# correo (ronda L, punto 1, 2026-09-09) -- se sobrescribe cada vez que se
# genera, no se acumulan versiones viejas.
PDF_OC_DIR = os.path.join(BASE_DIR, "data", "pdf_oc")

app = Flask(__name__)
# Soporte Postgres (ronda M, 2026-09-10): si existe la variable de entorno
# DATABASE_URL (la inyecta Railway automaticamente al referenciar
# ${{Postgres.DATABASE_URL}} en el servicio), se usa esa base en vez de la
# SQLite local -- asi el mismo codigo corre igual en la PC del usuario (sin
# la variable, sigue usando data/logistica.db como siempre) y en Railway.
# Algunos proveedores entregan el URL con el esquema viejo "postgres://",
# que SQLAlchemy 1.4+ ya no acepta -- se normaliza a "postgresql://".
_database_url = os.environ.get("DATABASE_URL")
if _database_url:
    if _database_url.startswith("postgres://"):
        _database_url = _database_url.replace("postgres://", "postgresql://", 1)
    app.config["SQLALCHEMY_DATABASE_URI"] = _database_url
else:
    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(BASE_DIR, "data", "logistica.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB por archivo subido
# SECRET_KEY tambien viene de variable de entorno en produccion (Railway) --
# localmente, sin la variable, sigue usando la misma llave de siempre.
app.secret_key = os.environ.get("SECRET_KEY", "suite-logistica-dev-key")  # cambiar en un despliegue real

db.init_app(app)


# Formato de numeros (2026-09-01, punto 5/7, a pedido del usuario): en Chile
# los pesos (CLP) se usan sin decimales (no hay centavos en circulacion) y
# el separador de miles es "." con "," para decimales -- igual que en la
# Invoice del proveedor. Se registran como filtros de Jinja para usarlos en
# toda la pantalla de Costeo: {{ valor|clp }} para pesos (entero) y
# {{ valor|moneda }} para USD/EUR/otras monedas (2 decimales).
def _formatear_numero(valor, decimales):
    try:
        numero = float(valor or 0)
    except (TypeError, ValueError):
        return valor
    texto = f"{numero:,.{decimales}f}"
    # f-string usa la convencion EEUU (miles ",", decimales ".") -- se
    # invierte a la convencion chilena (miles ".", decimales ",").
    return texto.replace(",", "§").replace(".", ",").replace("§", ".")


@app.template_filter("clp")
def filtro_clp(valor):
    return _formatear_numero(valor, 0)


@app.template_filter("moneda")
def filtro_moneda(valor, decimales=2):
    return _formatear_numero(valor, decimales)


@app.template_filter("numero_flex")
def filtro_numero_flex(valor):
    """Entero (sin decimales) si el valor no tiene parte decimal, o con 2
    decimales si la tiene -- mismo separador de miles "." y decimales ","
    de los filtros clp/moneda (ronda K, punto 3, para la Orden de Compra
    en PDF: "los valores deben ser enteros si no tienen decimal")."""
    try:
        numero = float(valor or 0)
    except (TypeError, ValueError):
        return valor
    redondeado = round(numero, 2)
    if redondeado == int(redondeado):
        return _formatear_numero(redondeado, 0)
    return _formatear_numero(redondeado, 2)


def ensure_schema_migrations():
    """Agrega columnas nuevas a tablas existentes sin borrar datos.
    SQLite soporta ALTER TABLE ... ADD COLUMN de forma sencilla, asi que
    evitamos pedirle al usuario que borre la base de datos cada vez que
    agregamos un campo nuevo a un modelo ya desplegado."""
    inspector = db.inspect(db.engine)
    tablas = inspector.get_table_names()
    # Ronda M (2026-09-10): estas migraciones se escribieron pensando solo en
    # SQLite (la base local del usuario). Al agregar soporte Postgres para el
    # despliegue en Railway, dos cosas necesitan ajustarse por dialecto: 1)
    # "BOOLEAN DEFAULT 0/1" es valido en SQLite pero Postgres exige
    # DEFAULT FALSE/TRUE para columnas boolean; 2) la migracion que quita el
    # UNIQUE de numero_po lee la tabla interna "sqlite_master", que no existe
    # en Postgres. En una base Postgres nueva (recien creada por
    # db.create_all(), ya con el esquema actual de models.py) ninguna de las
    # dos deberia hacer falta -- se dejan protegidas por dialecto para que el
    # arranque nunca falle, no para que se ejecuten de verdad ahi.
    es_sqlite = db.engine.dialect.name == "sqlite"

    migraciones = {
        "ordenes_compra_lineas": [
            ("anulada", "BOOLEAN DEFAULT 0"),
            ("variante_codigo", "VARCHAR(120)"),
            ("variante_descripcion", "VARCHAR(500)"),
        ],
        "ordenes_compra": [
            ("despacho_id", "INTEGER"),
            ("empresa_id", "INTEGER"),
        ],
        "importaciones": [
            ("despacho_id", "INTEGER"),
            ("condicion_compra", "VARCHAR(10) DEFAULT 'EXW'"),
            ("flete_total_moneda", "FLOAT DEFAULT 0"),
            ("seguro_total_moneda", "FLOAT DEFAULT 0"),
        ],
        "parcial_lineas": [
            ("orden_compra_linea_id", "INTEGER"),
        ],
        "parciales": [
            ("referencia", "VARCHAR(120)"),
        ],
        "gastos_importacion": [
            ("moneda", "VARCHAR(10) DEFAULT 'CLP'"),
            ("tipo_cambio", "FLOAT DEFAULT 1"),
            ("monto_original", "FLOAT DEFAULT 0"),
        ],
        "cargos_adicionales_importacion": [
            ("moneda", "VARCHAR(10) DEFAULT 'USD'"),
            ("tipo_cambio", "FLOAT DEFAULT 1"),
        ],
        "despachos": [
            ("url_tracking_manual", "VARCHAR(500)"),
        ],
        "empresas": [
            ("plantilla_asunto_correo", "VARCHAR(200)"),
            ("plantilla_cuerpo_correo", "TEXT"),
        ],
    }

    for tabla, columnas_nuevas in migraciones.items():
        if tabla not in tablas:
            continue
        columnas_actuales = {c["name"] for c in inspector.get_columns(tabla)}
        for nombre, ddl in columnas_nuevas:
            if nombre not in columnas_actuales:
                ddl_final = ddl
                if not es_sqlite and "BOOLEAN" in ddl.upper():
                    ddl_final = ddl_final.replace("DEFAULT 0", "DEFAULT FALSE").replace("DEFAULT 1", "DEFAULT TRUE")
                with db.engine.connect() as conn:
                    conn.execute(db.text(f"ALTER TABLE {tabla} ADD COLUMN {nombre} {ddl_final}"))
                    conn.commit()

    # Reparacion de datos (2026-08-27): "FOB" paso a llamarse "EXW" en
    # condicion_compra (mas ajustado al Incoterm real, punto 3) -- se
    # actualizan las filas ya guardadas con el nombre viejo. Y las
    # importaciones viejas que tenian flete/seguro cargados en USD
    # (flete_total_usd/seguro_total_usd, ya deprecados) se migran a los
    # campos nuevos en moneda_factura SOLO si la factura ya estaba en USD
    # (ahi el numero es el mismo); si la factura estaba en otra moneda no
    # se puede reconstruir el valor original con certeza, se deja en 0 --
    # el usuario los vuelve a cargar si hace falta.
    if "importaciones" in tablas:
        with db.engine.begin() as conn:
            conn.execute(db.text(
                "UPDATE importaciones SET condicion_compra = 'EXW' WHERE condicion_compra = 'FOB'"
            ))
            conn.execute(db.text(
                """
                UPDATE importaciones
                SET flete_total_moneda = COALESCE(flete_total_usd, 0),
                    seguro_total_moneda = COALESCE(seguro_total_usd, 0)
                WHERE moneda_factura = 'USD'
                  AND (COALESCE(flete_total_moneda, 0) = 0 AND COALESCE(seguro_total_moneda, 0) = 0)
                  AND (COALESCE(flete_total_usd, 0) != 0 OR COALESCE(seguro_total_usd, 0) != 0)
                """
            ))

    # Reparacion de datos (2026-09-01, punto 6): Flete y Seguro dejaron de
    # ser campos fijos de Importacion (flete_total_moneda / seguro_total_moneda,
    # ya deprecados) y pasaron a ser filas de CargoAdicionalImportacion, con
    # concepto "Flete"/"Seguro", igual que Handling Fee y Otros. Se
    # traspasa el valor viejo UNA SOLA VEZ por importacion (si ya existe una
    # fila con ese concepto para esa importacion, se asume que ya se
    # traspaso o que el usuario ya la cargo de nuevo a mano, y no se toca).
    if "importaciones" in tablas and "cargos_adicionales_importacion" in tablas:
        with db.engine.begin() as conn:
            filas = conn.execute(db.text(
                "SELECT id, moneda_factura, tipo_cambio_moneda_usd, flete_total_moneda, seguro_total_moneda "
                "FROM importaciones"
            )).fetchall()
            for imp_id, moneda_factura, tc_moneda_usd, flete_m, seguro_m in filas:
                for concepto, monto in (("Flete", flete_m), ("Seguro", seguro_m)):
                    if not monto:
                        continue
                    ya_existe = conn.execute(
                        db.text(
                            "SELECT 1 FROM cargos_adicionales_importacion "
                            "WHERE importacion_id = :iid AND concepto = :concepto"
                        ),
                        {"iid": imp_id, "concepto": concepto},
                    ).first()
                    if ya_existe:
                        continue
                    conn.execute(
                        db.text(
                            "INSERT INTO cargos_adicionales_importacion "
                            "(importacion_id, concepto, monto_moneda, moneda, tipo_cambio, referencia) "
                            "VALUES (:iid, :concepto, :monto, :moneda, :tc, :ref)"
                        ),
                        {
                            "iid": imp_id, "concepto": concepto, "monto": monto,
                            "moneda": moneda_factura or "USD", "tc": tc_moneda_usd or 1,
                            "ref": f"Migrado automáticamente desde el campo {concepto} total (deprecado)",
                        },
                    )

    # Quitar la restriccion UNIQUE de numero_po: desde que las ordenes se
    # pueden dividir automaticamente al confirmar solo una parte de los
    # productos, dos ordenes distintas pueden compartir el mismo numero de
    # PO (asi lo maneja el proveedor del usuario). SQLite no permite quitar
    # una restriccion UNIQUE con un simple ALTER TABLE, hay que reconstruir
    # la tabla preservando todos los datos (incluyendo los id, de los que
    # dependen las lineas y documentos ya guardados).
    if es_sqlite and "ordenes_compra" in tablas:
        with db.engine.connect() as conn:
            sql_tabla = conn.execute(
                db.text("SELECT sql FROM sqlite_master WHERE type='table' AND name='ordenes_compra'")
            ).scalar()
        if sql_tabla and "UNIQUE (numero_po)" in sql_tabla:
            with db.engine.begin() as conn:
                conn.execute(db.text("ALTER TABLE ordenes_compra RENAME TO ordenes_compra_old_migracion"))
                conn.execute(db.text(
                    """
                    CREATE TABLE ordenes_compra (
                        id INTEGER NOT NULL,
                        numero_po VARCHAR(40) NOT NULL,
                        proveedor_id INTEGER NOT NULL,
                        fecha_emision DATE,
                        moneda VARCHAR(10),
                        estado VARCHAR(30),
                        notas TEXT,
                        orden_consolidada_id INTEGER,
                        despacho_id INTEGER,
                        creado_en DATETIME,
                        PRIMARY KEY (id),
                        FOREIGN KEY(proveedor_id) REFERENCES proveedores (id),
                        FOREIGN KEY(orden_consolidada_id) REFERENCES ordenes_compra (id),
                        FOREIGN KEY(despacho_id) REFERENCES despachos (id)
                    )
                    """
                ))
                conn.execute(db.text(
                    """
                    INSERT INTO ordenes_compra
                        (id, numero_po, proveedor_id, fecha_emision, moneda, estado, notas, orden_consolidada_id, creado_en)
                    SELECT id, numero_po, proveedor_id, fecha_emision, moneda, estado, notas, orden_consolidada_id, creado_en
                    FROM ordenes_compra_old_migracion
                    """
                ))
                conn.execute(db.text("DROP TABLE ordenes_compra_old_migracion"))

    # Las carpetas de documentos por orden estaban nombradas con el
    # numero_po; como ahora dos ordenes pueden compartir numero_po, se
    # renombran a usar el id de la orden (siempre unico) para no mezclar
    # documentos de ordenes distintas.
    if "ordenes_compra" in tablas:
        with db.engine.connect() as conn:
            filas = conn.execute(db.text("SELECT id, numero_po FROM ordenes_compra")).fetchall()
        for oid, numero_po in filas:
            origen = os.path.join(DOCUMENTOS_DIR, numero_po)
            destino = os.path.join(DOCUMENTOS_DIR, str(oid))
            if os.path.isdir(origen) and not os.path.isdir(destino):
                os.rename(origen, destino)


def seed_empresas_compradoras():
    """Precarga las 2 empresas compradoras del usuario (Accuvision,
    Accumedical -- ronda K, 2026-09-07, punto 7) con los logos que
    adjuntó, para que no tenga que darlas de alta ni subir los archivos a
    mano. Gateado (igual que seed_from_excel): si ya existe alguna
    Empresa, no hace nada -- así el usuario puede editarlas o borrarlas
    libremente despues sin que se vuelvan a crear al reiniciar."""
    if Empresa.query.count() > 0:
        return
    seed = [
        ("Accuvision", "accuvision.png"),
        ("Accumedical", "accumedical.png"),
    ]
    for nombre, logo in seed:
        ruta_logo = os.path.join(EMPRESAS_LOGOS_DIR, logo)
        db.session.add(Empresa(
            nombre=nombre,
            logo_nombre_archivo=logo if os.path.isfile(ruta_logo) else None,
        ))
    db.session.commit()


def seed_variantes_lentes_medicontur():
    """Precarga las variantes de dioptria de los lentes MEDICONTUR (ronda M,
    2026-09-10, punto 1) desde 'Listado codigos lentes medicontur.xlsx'
    (columnas: codigo padre, Codigo Producto, DESCRIPCION). Cada fila se
    asocia al Producto ya cargado en el catalogo cuyo `codigo` coincide con
    el "codigo padre" (proveedor MEDICONTUR) -- ese Producto ya tiene su
    propio precio en el catalogo, que sigue siendo el que se usa siempre
    (las variantes NUNCA traen precio propio). Gateado igual que los demas
    seeds: si ya existe alguna ProductoVariante, no hace nada -- para que el
    usuario pueda agregar/editar/borrar variantes libremente despues sin que
    se sobrescriban solas al reiniciar."""
    if ProductoVariante.query.count() > 0:
        return
    if not os.path.isfile(VARIANTES_LENTES_MEDICONTUR_EXCEL):
        print(f"[seed] No se encontró {VARIANTES_LENTES_MEDICONTUR_EXCEL}, se omite la carga de variantes.")
        return

    import openpyxl
    medicontur = Proveedor.query.filter(db.func.lower(Proveedor.nombre) == "medicontur").first()
    if not medicontur:
        print("[seed] No existe el proveedor MEDICONTUR todavía, se omite la carga de variantes de lentes.")
        return

    productos_padre = {
        p.codigo.strip().upper(): p
        for p in Producto.query.filter_by(proveedor_id=medicontur.id).all()
    }

    wb = openpyxl.load_workbook(VARIANTES_LENTES_MEDICONTUR_EXCEL, data_only=True)
    ws = wb.worksheets[0]
    creadas = 0
    padres_no_encontrados = set()
    # Detecta la fila de encabezado ("codigo padre", "Codigo Producto",
    # "DESCRIPCION") en vez de asumir una fila fija -- el archivo real trae
    # una fila en blanco antes del encabezado.
    fila_inicio = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        primera = str(row[0]).strip().lower() if row and row[0] else ""
        if primera == "codigo padre":
            fila_inicio = i + 1
            break
    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        codigo_padre, codigo_variante, descripcion = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
        codigo_padre = str(codigo_padre).strip() if codigo_padre else ""
        codigo_variante = str(codigo_variante).strip() if codigo_variante else ""
        descripcion = str(descripcion).strip() if descripcion else ""
        if not codigo_padre or not codigo_variante:
            continue
        producto = productos_padre.get(codigo_padre.upper())
        if not producto:
            padres_no_encontrados.add(codigo_padre)
            continue
        db.session.add(ProductoVariante(
            producto_id=producto.id,
            codigo=codigo_variante,
            descripcion=descripcion or codigo_variante,
        ))
        creadas += 1
    db.session.commit()
    print(
        f"[seed] Importadas {creadas} variantes de lentes MEDICONTUR "
        f"({len(padres_no_encontrados)} código(s) padre no encontrados en el catálogo: "
        f"{sorted(padres_no_encontrados)})."
    )


with app.app_context():
    os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)
    os.makedirs(DOCUMENTOS_DIR, exist_ok=True)
    os.makedirs(EMPRESAS_LOGOS_DIR, exist_ok=True)
    os.makedirs(PDF_OC_DIR, exist_ok=True)
    db.create_all()
    ensure_schema_migrations()
    seed_from_excel(app)
    seed_empresas_compradoras()
    seed_variantes_lentes_medicontur()


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def parse_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def parse_int(value, default=0):
    if value is None or value == "":
        return default
    try:
        # Acepta "3.0" o "3" y siempre redondea a entero (no se permiten cajas fraccionadas)
        return int(round(float(value)))
    except (ValueError, TypeError):
        return default


def siguiente_numero_po(empresa_id=None):
    anio = datetime.utcnow().year
    prefijo = f"PO-{anio}-"
    # Ronda K (2026-09-07, punto 6): el correlativo es UNICO POR EMPRESA
    # compradora (Accuvision, Accumedical, ...) -- cada una sigue su propia
    # numeracion, sin importar a que proveedor le compre (confirmado
    # explicitamente por el usuario). Las ordenes "Sin asignar" (empresa_id
    # NULL, de antes de esta ronda) quedan fuera del calculo para no
    # interferir con el correlativo de ninguna empresa real.
    # Recorre TODAS las ordenes con ese prefijo (no solo la ultima creada):
    # como una orden dividida crea una fila nueva que REUTILIZA el mismo
    # numero_po de la orden original, la fila mas reciente por id ya no es
    # confiable para saber cual es el correlativo mas alto usado.
    ordenes = OrdenCompra.query.filter(
        OrdenCompra.numero_po.like(f"{prefijo}%"), OrdenCompra.empresa_id == empresa_id
    ).all()
    max_correlativo = 0
    for o in ordenes:
        try:
            max_correlativo = max(max_correlativo, int(o.numero_po.split("-")[-1]))
        except ValueError:
            continue
    return f"{prefijo}{max_correlativo + 1:04d}"


def recalcular_estado_orden(orden):
    """Sincroniza orden.estado con las etapas de sus lineas activas
    (no anuladas). Si la orden esta Cancelada, se respeta y no se
    recalcula."""
    if orden.estado == "Cancelada":
        return
    etapas_lineas = [l.etapa for l in orden.lineas if not l.anulada]
    if not etapas_lineas:
        # Sin lineas, o todas las lineas fueron anuladas.
        orden.estado = "Anulada" if orden.lineas.count() > 0 else ETAPAS_LINEA[0]
        return
    distintas = set(etapas_lineas)
    if len(distintas) == 1:
        orden.estado = distintas.pop()
    else:
        orden.estado = "En proceso (mixto)"


def _cohorte_etapa(etapa):
    """Agrupa las 5 etapas en 3 'cohortes' de negocio: 0 = todavia no
    confirmado por el proveedor, 1 = confirmado pero aun no despachado
    (fisicamente puede cambiar), 2 = ya despachado o mas adelante (una vez
    despachado, el usuario indico que todo el lote avanza junto por
    aduana/recepcion, asi que Despachada/Aduana/Recibido cuentan como el
    mismo grupo)."""
    try:
        idx = ETAPAS_LINEA.index(etapa)
    except ValueError:
        return 0
    if idx <= 0:
        return 0
    if idx == 1:
        return 1
    return 2


def dividir_orden_si_corresponde(orden):
    """Mantiene la regla de negocio: una orden nunca debe mezclar lineas
    de mas de un 'cohorte' (ver _cohorte_etapa). Si al confirmar, marcar
    despachada, o retroceder una linea la orden queda con lineas activas
    en mas de un cohorte, separa cada cohorte MENOS avanzado en una orden
    NUEVA que conserva el MISMO numero de PO -- asi es como el proveedor
    del usuario identifica los despachos parciales de una misma orden de
    compra. La orden original conserva su id, su cohorte MAS avanzado, y
    sus documentos ya cargados; cada orden nueva arranca sin documentos
    (tendra los suyos propios cuando a su vez avance).

    Las lineas anuladas no se mueven (se quedan en la orden en la que
    estaban, independientemente de en que etapa hayan quedado anuladas).

    Antes de crear una orden nueva para un cohorte, busca si ya existe una
    orden hermana (mismo numero de PO) que represente exactamente ese mismo
    cohorte y, si la hay, acumula ahi en vez de crear una orden distinta --
    esto evita que agregar productos nuevos de a uno (punto 1, 2026-08-26)
    genere una orden separada por cada producto agregado.

    Devuelve la lista de ordenes DESTINO afectadas (vacia si no hizo falta
    dividir); cada item es una tupla (orden, es_nueva)."""
    if orden.estado == "Cancelada":
        return []

    lineas_activas = [l for l in orden.lineas if not l.anulada]
    if not lineas_activas:
        return []

    cohortes_presentes = sorted(set(_cohorte_etapa(l.etapa) for l in lineas_activas))
    if len(cohortes_presentes) <= 1:
        return []  # todas las lineas activas estan en el mismo cohorte, nada que dividir

    cohorte_max = cohortes_presentes[-1]
    ordenes_destino = []
    for cohorte in cohortes_presentes[:-1]:
        lineas_cohorte = [l for l in lineas_activas if _cohorte_etapa(l.etapa) == cohorte]
        if not lineas_cohorte:
            continue
        destino = None
        for candidata in OrdenCompra.query.filter_by(numero_po=orden.numero_po).filter(OrdenCompra.id != orden.id):
            if candidata.estado == "Cancelada":
                continue
            lineas_candidata = [l for l in candidata.lineas if not l.anulada]
            if lineas_candidata and all(_cohorte_etapa(l.etapa) == cohorte for l in lineas_candidata):
                destino = candidata
                break
        es_nueva = destino is None
        if destino is None:
            destino = OrdenCompra(
                numero_po=orden.numero_po,
                proveedor_id=orden.proveedor_id,
                empresa_id=orden.empresa_id,
                fecha_emision=orden.fecha_emision,
                moneda=orden.moneda,
                estado=lineas_cohorte[0].etapa,
                notas=orden.notas,
            )
            db.session.add(destino)
            db.session.flush()
        for linea in lineas_cohorte:
            linea.orden_id = destino.id
        recalcular_estado_orden(destino)
        ordenes_destino.append((destino, es_nueva))

    recalcular_estado_orden(orden)
    return ordenes_destino


def reparar_ordenes_mezcladas():
    """Pasada de reparacion, idempotente: aplica dividir_orden_si_corresponde
    a TODAS las ordenes existentes. Corrige ordenes que hayan quedado
    mezcladas por acciones anteriores a que existiera esta regla (por
    ejemplo, lineas avanzadas individualmente antes del 2026-08-25). Se
    ejecuta en cada arranque; una orden ya bien separada no se vuelve a
    tocar."""
    for orden in OrdenCompra.query.all():
        dividir_orden_si_corresponde(orden)
    db.session.commit()


def _borrar_carpeta_documentos_orden(orden_id):
    """Elimina del disco la carpeta de documentos de una orden (si existe),
    ademas de las filas en la base -- para que no queden archivos huerfanos
    despues de borrar la orden."""
    carpeta = os.path.join(DOCUMENTOS_DIR, str(orden_id))
    if os.path.isdir(carpeta):
        shutil.rmtree(carpeta, ignore_errors=True)


def limpiar_ordenes_canceladas():
    """Borra definitivamente las ordenes 'Cancelada' (a peticion explicita
    del usuario del 2026-08-26: una orden cancelada no aporta nada
    operativo, no hace falta conservarla). Se excluyen las que ya tienen
    un despacho_id asociado (pasaron por el modulo Despachos y pueden
    tener costeo generado a partir de sus lineas) para no perder
    trazabilidad real por una cancelacion posterior -- esas se dejan tal
    cual, solo en estado 'Cancelada'. Idempotente, se ejecuta en cada
    arranque (limpia tanto las que ya estaban Cancelada de antes de este
    cambio como cualquiera nueva que se haya colado sin pasar por la ruta
    de cancelar, por ejemplo por un reinicio a mitad de una operacion)."""
    candidatas = OrdenCompra.query.filter(OrdenCompra.estado == "Cancelada").all()
    borradas = 0
    for orden in candidatas:
        if orden.despacho_id is not None:
            continue
        orden_id = orden.id
        db.session.delete(orden)
        borradas += 1
        db.session.flush()
        _borrar_carpeta_documentos_orden(orden_id)
    if borradas:
        db.session.commit()


def reparar_lotes_legacy():
    """Pasada de reparacion, idempotente: toda ParcialLinea que ya tenga
    unidades pero todavia no tenga ningun ParcialLineaLote (lineas creadas
    antes de la ronda J, o cargadas por seed_data) recibe un lote inicial
    unico que reusa su codigo_lote/fecha_vencimiento de siempre, para que
    el desglose por lotes nunca empiece "vacio" en datos ya existentes."""
    lineas = ParcialLinea.query.filter(
        ParcialLinea.cantidad_unidades > 0, ~ParcialLinea.lotes.any()
    ).all()
    for linea in lineas:
        db.session.add(ParcialLineaLote(
            linea_id=linea.id,
            codigo_lote=linea.codigo_lote or "",
            fecha_vencimiento=linea.fecha_vencimiento,
            cantidad_unidades=linea.cantidad_unidades,
        ))
    if lineas:
        db.session.commit()


with app.app_context():
    reparar_ordenes_mezcladas()
    limpiar_ordenes_canceladas()
    reparar_lotes_legacy()


def _mensaje_division(ordenes_destino):
    """Frase corta para agregar al flash cuando dividir_orden_si_corresponde
    movio lineas a otro grupo (orden nueva, o una orden hermana ya existente
    que representaba ese mismo cohorte)."""
    if not ordenes_destino:
        return ""
    nuevas = sum(1 for _, es_nueva in ordenes_destino if es_nueva)
    reutilizadas = len(ordenes_destino) - nuevas
    partes = []
    if nuevas == 1:
        partes.append("se separaron en una orden nueva con el mismo número de PO")
    elif nuevas > 1:
        partes.append(f"se separaron en {nuevas} órdenes nuevas con el mismo número de PO")
    if reutilizadas == 1:
        partes.append("se agregaron a otra orden ya existente del mismo grupo")
    elif reutilizadas > 1:
        partes.append(f"se agregaron a {reutilizadas} órdenes ya existentes del mismo grupo")
    return " Los productos que quedaron en otro(s) grupo(s) " + " y ".join(partes) + "."


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

@app.route("/")
def dashboard():
    total_proveedores = Proveedor.query.filter_by(activo=True).count()
    total_productos = Producto.query.filter_by(activo=True).count()
    total_ordenes = OrdenCompra.query.count()
    ordenes_pendientes = OrdenCompra.query.filter(
        OrdenCompra.estado.notin_(["Recibido", "Cancelada"])
    ).count()
    ultimas_ordenes = OrdenCompra.query.order_by(OrdenCompra.id.desc()).limit(6).all()
    return render_template(
        "dashboard.html",
        total_proveedores=total_proveedores,
        total_productos=total_productos,
        total_ordenes=total_ordenes,
        ordenes_pendientes=ordenes_pendientes,
        ultimas_ordenes=ultimas_ordenes,
    )


# ---------------------------------------------------------------------------
# Modulo: Mantenimiento de Proveedores
# ---------------------------------------------------------------------------

@app.route("/proveedores")
def proveedores_list():
    q = request.args.get("q", "").strip()
    query = Proveedor.query
    if q:
        query = query.filter(Proveedor.nombre.ilike(f"%{q}%"))
    proveedores = query.order_by(Proveedor.nombre).all()
    return render_template("proveedores/list.html", proveedores=proveedores, q=q)


@app.route("/proveedores/nuevo", methods=["GET", "POST"])
def proveedores_nuevo():
    if request.method == "POST":
        prov = Proveedor(
            nombre=request.form["nombre"].strip(),
            tipo=request.form.get("tipo", "Extranjero"),
            pais=request.form.get("pais", "").strip(),
            contacto_nombre=request.form.get("contacto_nombre", "").strip(),
            contacto_telefono=request.form.get("contacto_telefono", "").strip(),
            contacto_email=request.form.get("contacto_email", "").strip(),
            direccion=request.form.get("direccion", "").strip(),
            moneda_default=request.form.get("moneda_default", "USD"),
            notas=request.form.get("notas", "").strip(),
            activo=True,
        )
        db.session.add(prov)
        db.session.commit()
        flash(f"Proveedor '{prov.nombre}' creado correctamente.", "success")
        return redirect(url_for("proveedores_list"))
    return render_template("proveedores/form.html", proveedor=None)


@app.route("/proveedores/<int:proveedor_id>/editar", methods=["GET", "POST"])
def proveedores_editar(proveedor_id):
    prov = Proveedor.query.get_or_404(proveedor_id)
    if request.method == "POST":
        prov.nombre = request.form["nombre"].strip()
        prov.tipo = request.form.get("tipo", "Extranjero")
        prov.pais = request.form.get("pais", "").strip()
        prov.contacto_nombre = request.form.get("contacto_nombre", "").strip()
        prov.contacto_telefono = request.form.get("contacto_telefono", "").strip()
        prov.contacto_email = request.form.get("contacto_email", "").strip()
        prov.direccion = request.form.get("direccion", "").strip()
        prov.moneda_default = request.form.get("moneda_default", "USD")
        prov.notas = request.form.get("notas", "").strip()
        db.session.commit()
        flash(f"Proveedor '{prov.nombre}' actualizado.", "success")
        return redirect(url_for("proveedores_list"))
    return render_template("proveedores/form.html", proveedor=prov)


@app.route("/proveedores/<int:proveedor_id>/eliminar", methods=["POST"])
def proveedores_eliminar(proveedor_id):
    prov = Proveedor.query.get_or_404(proveedor_id)
    if prov.ordenes.count() > 0:
        prov.activo = False
        db.session.commit()
        flash(
            f"'{prov.nombre}' tiene ordenes de compra asociadas: se desactivo en vez de eliminarse.",
            "warning",
        )
    else:
        db.session.delete(prov)
        db.session.commit()
        flash(f"Proveedor eliminado.", "success")
    return redirect(url_for("proveedores_list"))


@app.route("/proveedores/<int:proveedor_id>/reactivar", methods=["POST"])
def proveedores_reactivar(proveedor_id):
    prov = Proveedor.query.get_or_404(proveedor_id)
    prov.activo = True
    db.session.commit()
    flash(f"Proveedor '{prov.nombre}' reactivado.", "success")
    return redirect(url_for("proveedores_list"))


@app.route("/proveedores/<int:proveedor_id>")
def proveedores_detalle(proveedor_id):
    prov = Proveedor.query.get_or_404(proveedor_id)
    q = request.args.get("q", "").strip()
    productos_query = prov.productos
    if q:
        productos_query = productos_query.filter(
            db.or_(Producto.codigo.ilike(f"%{q}%"), Producto.descripcion.ilike(f"%{q}%"))
        )
    productos = productos_query.order_by(Producto.codigo).all()
    ordenes = prov.ordenes.order_by(OrdenCompra.id.desc()).limit(10).all()
    return render_template(
        "proveedores/detalle.html", proveedor=prov, productos=productos, q=q, ordenes=ordenes
    )


# --- Productos (catalogo por proveedor) ---

@app.route("/proveedores/<int:proveedor_id>/productos/nuevo", methods=["POST"])
def productos_nuevo(proveedor_id):
    prov = Proveedor.query.get_or_404(proveedor_id)
    producto = Producto(
        proveedor_id=prov.id,
        codigo=request.form["codigo"].strip(),
        descripcion=request.form["descripcion"].strip(),
        empaque=int(request.form.get("empaque") or 1),
        moneda=request.form.get("moneda", prov.moneda_default or "USD"),
        precio_caja=float(request.form.get("precio_caja") or 0),
        precio_unitario=float(request.form.get("precio_unitario") or 0),
        activo=True,
    )
    db.session.add(producto)
    db.session.commit()
    flash(f"Producto '{producto.codigo}' agregado al catalogo de {prov.nombre}.", "success")
    return redirect(url_for("proveedores_detalle", proveedor_id=prov.id))


@app.route("/productos/<int:producto_id>/editar", methods=["POST"])
def productos_editar(producto_id):
    producto = Producto.query.get_or_404(producto_id)
    producto.codigo = request.form["codigo"].strip()
    producto.descripcion = request.form["descripcion"].strip()
    producto.empaque = int(request.form.get("empaque") or 1)
    producto.moneda = request.form.get("moneda", "USD")
    producto.precio_caja = float(request.form.get("precio_caja") or 0)
    producto.precio_unitario = float(request.form.get("precio_unitario") or 0)
    producto.activo = "activo" in request.form
    db.session.commit()
    flash(f"Producto '{producto.codigo}' actualizado.", "success")
    return redirect(url_for("proveedores_detalle", proveedor_id=producto.proveedor_id))


@app.route("/productos/<int:producto_id>/eliminar", methods=["POST"])
def productos_eliminar(producto_id):
    producto = Producto.query.get_or_404(producto_id)
    proveedor_id = producto.proveedor_id
    producto.activo = False
    db.session.commit()
    flash("Producto desactivado del catalogo.", "success")
    return redirect(url_for("proveedores_detalle", proveedor_id=proveedor_id))


# ---------------------------------------------------------------------------
# Modulo: Compras (Ordenes de Compra)
# ---------------------------------------------------------------------------

@app.route("/ordenes")
def ordenes_list():
    estado = request.args.get("estado", "")
    proveedor_id = request.args.get("proveedor_id", "")
    empresa_id = request.args.get("empresa_id", "")
    mostrar_despachadas = request.args.get("mostrar_despachadas") == "1"
    query = OrdenCompra.query
    if empresa_id:
        if empresa_id == "sin-asignar":
            query = query.filter(OrdenCompra.empresa_id.is_(None))
        else:
            query = query.filter_by(empresa_id=empresa_id)
    if estado:
        if estado in ETAPAS_LINEA:
            # Una orden "En proceso (mixto)" no tiene ese estado exacto, pero
            # puede tener lineas activas en la etapa buscada: la mostramos
            # igual para que el filtro refleje el avance real por producto.
            query = query.filter(
                OrdenCompra.lineas.any(
                    db.and_(OrdenCompraLinea.etapa == estado, OrdenCompraLinea.anulada == False)  # noqa: E712
                )
            )
        else:
            query = query.filter_by(estado=estado)
    if proveedor_id:
        query = query.filter_by(proveedor_id=proveedor_id)
    # Una vez que una orden pasa a "Orden Despachada" (o mas adelante), deja
    # de gestionarse aqui: el control pasa al modulo Despachos (courier,
    # tracking, aduana, recibido). Por defecto ya no aparece en este listado
    # -- solo compras que todavia estan en emision/confirmacion. El toggle
    # 'mostrar_despachadas' es una salida de emergencia para ubicarlas igual
    # si hace falta consultarlas desde aqui.
    if not mostrar_despachadas and estado not in ETAPAS_ORDEN_DESPACHABLE:
        query = query.filter(~OrdenCompra.estado.in_(ETAPAS_ORDEN_DESPACHABLE))
    ordenes = query.order_by(OrdenCompra.id.desc()).all()
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    empresas = Empresa.query.order_by(Empresa.nombre).all()
    return render_template(
        "ordenes/list.html",
        ordenes=ordenes,
        proveedores=proveedores,
        empresas=empresas,
        estados=ESTADOS_OC,
        estado_sel=estado,
        proveedor_sel=proveedor_id,
        empresa_sel=empresa_id,
        mostrar_despachadas=mostrar_despachadas,
    )


@app.route("/ordenes/nueva", methods=["GET", "POST"])
def ordenes_nueva():
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    empresas = Empresa.query.filter_by(activo=True).order_by(Empresa.nombre).all()

    if request.method == "POST":
        proveedor_id = int(request.form["proveedor_id"])
        prov = Proveedor.query.get_or_404(proveedor_id)
        empresa_id = request.form.get("empresa_id") or None
        if empresa_id:
            Empresa.query.get_or_404(int(empresa_id))

        orden = OrdenCompra(
            numero_po=siguiente_numero_po(int(empresa_id) if empresa_id else None),
            proveedor_id=prov.id,
            empresa_id=int(empresa_id) if empresa_id else None,
            fecha_emision=parse_date(request.form.get("fecha_emision")) or date.today(),
            moneda=request.form.get("moneda", prov.moneda_default or "USD"),
            estado=ETAPAS_LINEA[0],
            notas=request.form.get("notas", "").strip(),
        )
        db.session.add(orden)
        db.session.flush()

        producto_ids = request.form.getlist("producto_id")
        cantidades = request.form.getlist("cantidad_cajas")
        precios = request.form.getlist("precio_unitario_pactado")
        fechas = request.form.getlist("fecha_estimada_despacho")
        # Variante elegida por linea (ronda M, 2026-09-10, punto 1) -- listas
        # paralelas a producto_ids, vacías ("") cuando el producto no tiene
        # variantes o no se eligió ninguna.
        variante_codigos = request.form.getlist("variante_codigo")
        variante_descripciones = request.form.getlist("variante_descripcion")

        lineas_creadas = 0
        for i, (pid, cant, precio, fecha) in enumerate(zip(producto_ids, cantidades, precios, fechas)):
            if not pid or not cant:
                continue
            cant_val = parse_int(cant, default=0)
            if cant_val <= 0:
                continue
            linea = OrdenCompraLinea(
                orden_id=orden.id,
                producto_id=int(pid),
                cantidad_cajas=cant_val,
                precio_unitario_pactado=float(precio) if precio else 0,
                fecha_estimada_despacho=parse_date(fecha),
                etapa=ETAPAS_LINEA[0],
                variante_codigo=(variante_codigos[i].strip() if i < len(variante_codigos) and variante_codigos[i].strip() else None),
                variante_descripcion=(variante_descripciones[i].strip() if i < len(variante_descripciones) and variante_descripciones[i].strip() else None),
            )
            db.session.add(linea)
            lineas_creadas += 1

        if lineas_creadas == 0:
            db.session.rollback()
            flash("Debes agregar al menos una linea con cantidad mayor a 0.", "danger")
            return redirect(url_for("ordenes_nueva", proveedor_id=proveedor_id))

        db.session.commit()
        flash(f"Orden {orden.numero_po} creada con {lineas_creadas} lineas.", "success")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    proveedor_id = request.args.get("proveedor_id", type=int)
    return render_template(
        "ordenes/form.html", proveedores=proveedores, empresas=empresas,
        proveedor_id=proveedor_id, now=date.today()
    )


@app.route("/api/proveedores/<int:proveedor_id>/productos")
def api_productos_por_proveedor(proveedor_id):
    q = request.args.get("q", "").strip()
    query = Producto.query.filter_by(proveedor_id=proveedor_id, activo=True)
    if q:
        query = query.filter(
            db.or_(Producto.codigo.ilike(f"%{q}%"), Producto.descripcion.ilike(f"%{q}%"))
        )
    productos = query.order_by(Producto.codigo).limit(50).all()
    return jsonify(
        [
            {
                "id": p.id,
                "codigo": p.codigo,
                "descripcion": p.descripcion,
                "empaque": p.empaque,
                "moneda": p.moneda,
                "precio_caja": p.precio_caja,
                "precio_unitario": p.precio_unitario,
                # Ronda M (2026-09-10): si tiene variantes (ej. dioptrias de
                # lentes MEDICONTUR), el front-end abre un modal para elegir
                # cual en vez de agregar la linea directo con el codigo padre.
                "tiene_variantes": p.variantes.count() > 0,
            }
            for p in productos
        ]
    )


@app.route("/api/productos/<int:producto_id>/variantes")
def api_producto_variantes(producto_id):
    """Variantes de un producto padre (ronda M, 2026-09-10, punto 1) --
    usado por el modal de "elegir variante" al agregar una linea a la
    orden. El precio NUNCA viene de aquí, siempre es el del producto
    padre."""
    producto = Producto.query.get_or_404(producto_id)
    variantes = producto.variantes.order_by(ProductoVariante.codigo).all()
    return jsonify(
        [{"id": v.id, "codigo": v.codigo, "descripcion": v.descripcion} for v in variantes]
    )


@app.route("/ordenes/<int:orden_id>")
def ordenes_detalle(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    lineas = orden.lineas.all()
    documentos = orden.documentos.all()
    # Otras ordenes que comparten el mismo numero de PO (por una division
    # automatica al confirmar o despachar solo una parte de los productos).
    # Se excluyen las 'Cancelada' sin lineas: una orden cancelada y vacia no
    # es una parte real de la division (nuestra logica de division nunca
    # crea una orden nueva sin productos), asi que mostrarla como "parte"
    # solo confunde -- probablemente quedo de una orden creada por error y
    # cancelada antes de agregarle productos.
    otras_partes = (
        OrdenCompra.query.filter(
            OrdenCompra.numero_po == orden.numero_po, OrdenCompra.id != orden.id
        )
        .order_by(OrdenCompra.id)
        .all()
    )
    otras_partes = [
        o for o in otras_partes if not (o.estado == "Cancelada" and o.lineas.count() == 0)
    ]
    empresas = Empresa.query.filter_by(activo=True).order_by(Empresa.nombre).all()
    return render_template(
        "ordenes/detalle.html",
        orden=orden,
        lineas=lineas,
        estados=ESTADOS_OC,
        etapas=ETAPAS_LINEA,
        documentos=documentos,
        tipos_documento=TIPOS_DOCUMENTO_ORDEN,
        otras_partes=otras_partes,
        empresas=empresas,
    )


@app.route("/ordenes/<int:orden_id>/editar", methods=["POST"])
def ordenes_editar(orden_id):
    """Edicion de los datos de cabecera de la orden -- ronda K (2026-09-07):
    nace de la necesidad de poder asignarle una Empresa compradora a
    ordenes ya existentes (creadas antes de esta mejora, sin ninguna
    asignada) y de poder corregir el N° de PO a mano cuando el usuario ya
    venia numerando sus OC fuera del sistema antes de implementarlo (punto
    5) -- todo a traves de la propia aplicacion, sin tocar la base de
    datos directamente."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    nuevo_numero = request.form.get("numero_po", "").strip()
    if not nuevo_numero:
        flash("El número de OC no puede quedar vacío.", "danger")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))
    empresa_id = request.form.get("empresa_id") or None
    if empresa_id:
        Empresa.query.get_or_404(int(empresa_id))
    orden.numero_po = nuevo_numero
    orden.empresa_id = int(empresa_id) if empresa_id else None
    orden.fecha_emision = parse_date(request.form.get("fecha_emision")) or orden.fecha_emision
    orden.moneda = request.form.get("moneda", orden.moneda)
    orden.notas = request.form.get("notas", "").strip()
    db.session.commit()
    flash(f"Orden actualizada: {orden.numero_po}.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/cancelar", methods=["POST"])
def ordenes_cancelar(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    numero_po = orden.numero_po
    if orden.despacho_id is not None:
        # Ya paso por el modulo Despachos (tiene tracking/guia y puede tener
        # costeo generado a partir de sus lineas): no se borra para no
        # perder esa trazabilidad, solo se marca Cancelada como antes.
        orden.estado = "Cancelada"
        db.session.commit()
        flash(
            f"Orden {numero_po} marcada como Cancelada. No se borro porque ya esta "
            "asociada a un despacho (puede tener costeo generado).",
            "warning",
        )
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    db.session.delete(orden)
    db.session.commit()
    _borrar_carpeta_documentos_orden(orden_id)
    flash(f"Orden {numero_po} cancelada y eliminada del sistema.", "warning")
    return redirect(url_for("ordenes_list"))


@app.route("/ordenes/<int:orden_id>/reactivar", methods=["POST"])
def ordenes_reactivar(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    orden.estado = "Emisión de Orden"  # se recalcula abajo segun las lineas
    recalcular_estado_orden(orden)
    db.session.commit()
    flash(f"Orden {orden.numero_po} reactivada.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/lineas/nueva", methods=["POST"])
def ordenes_linea_nueva(orden_id):
    """Agrega un producto nuevo a una orden YA CREADA -- antes solo se
    podian agregar productos al momento de crear la orden; una vez emitida
    solo se podia modificar o eliminar lo existente (pedido del usuario,
    2026-08-26)."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.estado in ("Cancelada", "Anulada"):
        flash("No se pueden agregar productos a una orden cancelada o anulada.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    producto_id = request.form.get("producto_id")
    if not producto_id:
        flash("Selecciona un producto del catálogo antes de agregar.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))
    producto_id = int(producto_id)
    # Variante elegida (ronda M, 2026-09-10, punto 1) -- si el producto tiene
    # variantes (ej. dioptrias de lentes MEDICONTUR), el modal del front-end
    # manda aca el codigo/descripcion especifico elegido.
    variante_codigo = (request.form.get("variante_codigo") or "").strip() or None
    variante_descripcion = (request.form.get("variante_descripcion") or "").strip() or None

    # Evitar duplicar la misma linea activa en la orden (punto 9): si ya
    # esta como linea activa CON LA MISMA VARIANTE (o ninguna variante en
    # ambos casos), se avisa y no se crea una segunda linea -- hay que
    # ajustar la cantidad en la linea existente en su lugar. Dos variantes
    # DISTINTAS del mismo producto padre (ej. dos dioptrias distintas de un
    # mismo lente) SI pueden convivir como lineas separadas.
    filtro_duplicado = [OrdenCompraLinea.producto_id == producto_id, OrdenCompraLinea.anulada == False]  # noqa: E712
    if variante_codigo:
        filtro_duplicado.append(OrdenCompraLinea.variante_codigo == variante_codigo)
    else:
        filtro_duplicado.append(OrdenCompraLinea.variante_codigo.is_(None))
    ya_existe = orden.lineas.filter(*filtro_duplicado).first()
    if ya_existe:
        flash(
            f"'{ya_existe.codigo_mostrar}' ya está en esta orden — ajusta la cantidad en esa línea "
            "en vez de agregarla de nuevo.",
            "warning",
        )
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    cant = parse_int(request.form.get("cantidad_cajas"), default=0)
    if cant <= 0:
        flash("Ingresa una cantidad de cajas mayor a 0.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    linea = OrdenCompraLinea(
        orden_id=orden.id,
        producto_id=producto_id,
        cantidad_cajas=cant,
        precio_unitario_pactado=float(request.form.get("precio_unitario_pactado") or 0),
        fecha_estimada_despacho=parse_date(request.form.get("fecha_estimada_despacho")),
        etapa=ETAPAS_LINEA[0],
        variante_codigo=variante_codigo,
        variante_descripcion=variante_descripcion,
    )
    db.session.add(linea)
    db.session.flush()
    recalcular_estado_orden(orden)
    # Si la orden ya tenia lineas mas avanzadas (ej. despachada) esta linea
    # nueva en "Emision de Orden" quedaria mezclando cohortes -- se separa
    # igual que al confirmar/despachar.
    nuevas_ordenes = dividir_orden_si_corresponde(orden)
    db.session.commit()
    flash(f"'{linea.codigo_mostrar}' agregado a la orden.{_mensaje_division(nuevas_ordenes)}", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/actualizar", methods=["POST"])
def ordenes_linea_actualizar(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    if linea.bloqueada:
        flash("Esta línea ya fue despachada: no se puede editar. Gestiona sus documentos en la sección de Documentos de la orden.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    linea.cantidad_cajas = parse_int(request.form.get("cantidad_cajas"), default=linea.cantidad_cajas)
    linea.precio_unitario_pactado = float(
        request.form.get("precio_unitario_pactado") or linea.precio_unitario_pactado
    )
    linea.fecha_estimada_despacho = (
        parse_date(request.form.get("fecha_estimada_despacho")) or linea.fecha_estimada_despacho
    )
    linea.cantidad_despachada_cajas = parse_int(
        request.form.get("cantidad_despachada_cajas"), default=linea.cantidad_despachada_cajas
    )
    fecha_disp = request.form.get("fecha_disponibilidad_confirmada")
    fecha_desp_conf = request.form.get("fecha_estimada_despacho_confirmada")
    if fecha_disp:
        linea.fecha_disponibilidad_confirmada = parse_date(fecha_disp)
    if fecha_desp_conf:
        linea.fecha_estimada_despacho_confirmada = parse_date(fecha_desp_conf)
    db.session.commit()
    flash("Linea de la orden actualizada.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/eliminar", methods=["POST"])
def ordenes_linea_eliminar(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    if linea.bloqueada:
        flash("Esta línea ya fue despachada: no se puede eliminar.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    orden = linea.orden
    db.session.delete(linea)
    db.session.flush()
    recalcular_estado_orden(orden)
    db.session.commit()
    flash("Linea eliminada de la orden.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/lineas/fecha-masiva", methods=["POST"])
def ordenes_lineas_fecha_masiva(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    linea_ids = request.form.getlist("linea_ids")
    nueva_fecha = parse_date(request.form.get("fecha_estimada_despacho"))
    if not linea_ids:
        flash("Selecciona al menos una linea para aplicar la fecha.", "warning")
    elif not nueva_fecha:
        flash("Ingresa una fecha valida.", "warning")
    else:
        actualizadas = 0
        omitidas = 0
        for lid in linea_ids:
            linea = OrdenCompraLinea.query.get(int(lid))
            if not linea or linea.orden_id != orden.id:
                continue
            if linea.bloqueada:
                omitidas += 1
                continue
            linea.fecha_estimada_despacho = nueva_fecha
            actualizadas += 1
        db.session.commit()
        mensaje = f"Fecha estimada de despacho aplicada a {actualizadas} linea(s)."
        if omitidas:
            mensaje += f" {omitidas} se omitieron por estar ya despachadas."
        flash(mensaje, "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


# --- Etapas por linea ---

@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/confirmar", methods=["POST"])
def ordenes_linea_confirmar(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    orden = linea.orden
    if linea.anulada:
        flash("Esta linea esta anulada. Reactivala primero para poder confirmarla.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    fecha_desp_conf = parse_date(request.form.get("fecha_estimada_despacho_confirmada"))
    if not fecha_desp_conf:
        flash("Debes indicar la fecha de despacho confirmada por el proveedor.", "danger")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    linea.fecha_estimada_despacho_confirmada = fecha_desp_conf
    linea.etapa = "Orden Confirmada"
    recalcular_estado_orden(orden)
    nuevas_ordenes = dividir_orden_si_corresponde(orden)
    db.session.commit()
    flash(f"'{linea.producto.codigo}' confirmado por el proveedor.{_mensaje_division(nuevas_ordenes)}", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/lineas/confirmar-masivo", methods=["POST"])
def ordenes_lineas_confirmar_masivo(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    linea_ids = request.form.getlist("linea_ids")
    fecha_desp_conf = parse_date(request.form.get("fecha_estimada_despacho_confirmada"))

    if not linea_ids:
        flash("Selecciona al menos una linea para confirmar.", "warning")
    elif not fecha_desp_conf:
        flash("Ingresa la fecha de despacho confirmada.", "warning")
    else:
        confirmadas = 0
        omitidas = 0
        for lid in linea_ids:
            linea = OrdenCompraLinea.query.get(int(lid))
            if not linea or linea.orden_id != orden.id:
                continue
            if linea.anulada or linea.etapa != "Emisión de Orden":
                omitidas += 1
                continue
            linea.fecha_estimada_despacho_confirmada = fecha_desp_conf
            linea.etapa = "Orden Confirmada"
            confirmadas += 1
        recalcular_estado_orden(orden)
        nuevas_ordenes = dividir_orden_si_corresponde(orden) if confirmadas else []
        db.session.commit()
        mensaje = f"{confirmadas} línea(s) confirmada(s)."
        if omitidas:
            mensaje += f" {omitidas} se omitieron por no estar en 'Emisión de Orden'."
        mensaje += _mensaje_division(nuevas_ordenes)
        flash(mensaje, "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


# NOTA (2026-08-26): "ordenes_linea_avanzar" (marcar despachada/aduana/recibido
# linea por linea desde Ordenes) fue retirado a pedido del usuario. Desde ahora
# Ordenes solo gestiona Confirmar / Fecha de despacho / Cambiar estado / Anular;
# el paso a "Orden Despachada" ocurre unicamente al asociar la orden a un
# Despacho (ver _marcar_orden_despachada), y aduana/recibido se gestionan en
# bloque desde el modulo Despachos (ver despachos_aduana_masivo / recibido_masivo).


@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/retroceder", methods=["POST"])
def ordenes_linea_retroceder(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    orden = linea.orden
    if linea.bloqueada:
        flash("Esta línea ya fue despachada: no se puede devolver a una etapa anterior.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    anterior = linea.etapa_anterior
    if not anterior:
        flash("Esta linea ya esta en la primera etapa.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    linea.etapa = anterior
    recalcular_estado_orden(orden)
    db.session.commit()
    flash(f"'{linea.producto.codigo}' devuelto a la etapa '{anterior}'.", "info")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


def _lineas_seleccionadas(orden):
    linea_ids = request.form.getlist("linea_ids")
    lineas = []
    for lid in linea_ids:
        try:
            linea = OrdenCompraLinea.query.get(int(lid))
        except (TypeError, ValueError):
            continue
        if linea and linea.orden_id == orden.id:
            lineas.append(linea)
    return lineas


# NOTA (2026-08-26): "ordenes_lineas_avanzar_masivo" (Despachado),
# "ordenes_lineas_aduana_masivo" y "ordenes_lineas_recibido_masivo" fueron
# retiradas del modulo Ordenes a pedido del usuario -- esas transiciones
# ahora viven exclusivamente en el modulo Despachos (asociar orden = marca
# despachada; ver despachos_aduana_masivo / despachos_recibido_masivo para
# el resto del ciclo).


@app.route("/ordenes/<int:orden_id>/lineas/retroceder-masivo", methods=["POST"])
def ordenes_lineas_retroceder_masivo(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    lineas = _lineas_seleccionadas(orden)
    if not lineas:
        flash("Selecciona al menos una linea para devolver a la etapa anterior.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    retrocedidas = 0
    omitidas = 0
    for linea in lineas:
        if linea.anulada or linea.bloqueada:
            omitidas += 1
            continue
        anterior = linea.etapa_anterior
        if not anterior:
            omitidas += 1
            continue
        linea.etapa = anterior
        retrocedidas += 1

    recalcular_estado_orden(orden)
    db.session.commit()
    mensaje = f"{retrocedidas} línea(s) devuelta(s) a su etapa anterior."
    if omitidas:
        mensaje += f" {omitidas} se omitieron (anuladas, ya despachadas o ya en la primera etapa)."
    flash(mensaje, "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/lineas/anular-masivo", methods=["POST"])
def ordenes_lineas_anular_masivo(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    lineas = _lineas_seleccionadas(orden)
    if not lineas:
        flash("Selecciona al menos una linea para anular.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    anuladas = 0
    omitidas = 0
    for linea in lineas:
        if linea.anulada or linea.bloqueada:
            omitidas += 1
            continue
        linea.anulada = True
        anuladas += 1

    recalcular_estado_orden(orden)
    db.session.commit()
    mensaje = f"{anuladas} línea(s) anulada(s)."
    if omitidas:
        mensaje += f" {omitidas} se omitieron (ya anuladas o ya despachadas)."
    flash(mensaje, "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/anular", methods=["POST"])
def ordenes_linea_anular(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    orden = linea.orden
    if linea.bloqueada:
        flash("Esta línea ya fue despachada, no se puede anular.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    linea.anulada = True
    recalcular_estado_orden(orden)
    db.session.commit()
    flash(f"'{linea.producto.codigo}' anulada.", "info")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/reactivar-linea", methods=["POST"])
def ordenes_linea_reactivar(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    orden = linea.orden
    linea.anulada = False
    recalcular_estado_orden(orden)
    db.session.commit()
    flash(f"'{linea.producto.codigo}' reactivada. Sigue en la etapa '{linea.etapa}'.", "info")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


# --- Documentos de la orden ---

def _extension_valida(nombre_archivo):
    _, ext = os.path.splitext(nombre_archivo.lower())
    return ext in EXTENSIONES_PERMITIDAS


@app.route("/ordenes/<int:orden_id>/documentos/subir", methods=["POST"])
def ordenes_documento_subir(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    archivo = request.files.get("archivo")
    tipo = request.form.get("tipo", "Otro")

    if not archivo or archivo.filename == "":
        flash("Selecciona un archivo para subir.", "danger")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    if not _extension_valida(archivo.filename):
        flash("Solo se permiten archivos PDF o imagenes (JPG, PNG).", "danger")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(orden.id))
    os.makedirs(carpeta_orden, exist_ok=True)

    _, ext = os.path.splitext(archivo.filename)
    nombre_disco = f"{uuid.uuid4().hex}{ext.lower()}"
    archivo.save(os.path.join(carpeta_orden, nombre_disco))

    doc = OrdenDocumento(
        orden_id=orden.id,
        tipo=tipo if tipo in TIPOS_DOCUMENTO_ORDEN else "Otro",
        nombre_original=archivo.filename,
        nombre_archivo=nombre_disco,
    )
    db.session.add(doc)
    db.session.commit()
    flash(f"Documento '{archivo.filename}' subido correctamente.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/documentos/<int:doc_id>/ver")
def ordenes_documento_ver(orden_id, doc_id):
    """Abre el documento en el navegador (PDF/imagen inline) en vez de forzar
    la descarga -- para que el usuario pueda visualizarlo con un clic."""
    doc = OrdenDocumento.query.get_or_404(doc_id)
    if doc.orden_id != orden_id:
        abort(404)
    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(doc.orden_id))
    return send_from_directory(
        carpeta_orden, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=False
    )


@app.route("/ordenes/<int:orden_id>/documentos/<int:doc_id>/descargar")
def ordenes_documento_descargar(orden_id, doc_id):
    doc = OrdenDocumento.query.get_or_404(doc_id)
    if doc.orden_id != orden_id:
        abort(404)
    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(doc.orden_id))
    return send_from_directory(
        carpeta_orden, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=True
    )


@app.route("/ordenes/<int:orden_id>/documentos/<int:doc_id>/eliminar", methods=["POST"])
def ordenes_documento_eliminar(orden_id, doc_id):
    doc = OrdenDocumento.query.get_or_404(doc_id)
    if doc.orden_id != orden_id:
        abort(404)
    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(doc.orden_id))
    ruta = os.path.join(carpeta_orden, doc.nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)
    db.session.delete(doc)
    db.session.commit()
    flash("Documento eliminado.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/imprimir")
def ordenes_imprimir(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    lineas = orden.lineas.all()
    return render_template("ordenes/imprimir.html", orden=orden, lineas=lineas)


# Plantilla de correo por defecto (ronda L, punto 1, 2026-09-09): se usa
# cuando la empresa compradora de la orden no tiene una plantilla propia
# configurada en Configuración > Empresas. Usa los mismos marcadores
# [ENTRE_CORCHETES] que puede usar el usuario en su plantilla personalizada.
DEFAULT_ASUNTO_CORREO = "Orden de Compra [NUMERO_OC] - [EMPRESA]"
DEFAULT_CUERPO_CORREO = (
    "Estimados [PROVEEDOR],\n\n"
    "Adjuntamos la Orden de Compra [NUMERO_OC].\n\n"
    "[NOTAS]\n\n"
    "Quedamos atentos a la confirmación de disponibilidad y fecha de despacho.\n\n"
    "Saludos."
)


def construir_correo_oc(orden):
    """Arma el asunto y cuerpo del correo para enviar la OC al proveedor,
    reemplazando marcadores [ENTRE_CORCHETES] por los datos reales de la
    orden (ronda L, punto 1, 2026-09-09). Usa la plantilla propia de la
    empresa compradora si está configurada (Configuración > Empresas); si
    no, cae en la plantilla por defecto. Así el usuario no necesita editar
    nada a mano en Outlook -- solo escribe lo que quiera agregar en el campo
    "Notas" de la orden y eso queda en el marcador [NOTAS]."""
    empresa = orden.empresa
    asunto_plantilla = (empresa.plantilla_asunto_correo if empresa else None) or DEFAULT_ASUNTO_CORREO
    cuerpo_plantilla = (empresa.plantilla_cuerpo_correo if empresa else None) or DEFAULT_CUERPO_CORREO

    marcadores = {
        "[PROVEEDOR]": orden.proveedor.nombre if orden.proveedor else "",
        "[NUMERO_OC]": orden.numero_po or "",
        "[EMPRESA]": empresa.nombre if empresa else "",
        "[FECHA_EMISION]": str(orden.fecha_emision) if orden.fecha_emision else "",
        "[MONEDA]": orden.moneda or "",
        "[TOTAL]": f"{orden.total:.2f}" if orden.total is not None else "",
        "[NOTAS]": (orden.notas or "").strip(),
    }

    def _reemplazar(texto):
        for marcador, valor in marcadores.items():
            texto = texto.replace(marcador, valor)
        return texto

    asunto = _reemplazar(asunto_plantilla)
    cuerpo = _reemplazar(cuerpo_plantilla)
    # Si no hay notas, evita dejar un párrafo en blanco feo en medio del correo.
    cuerpo = re.sub(r"\n{3,}", "\n\n", cuerpo).strip()
    return asunto, cuerpo


def generar_pdf_orden(orden):
    """Genera el PDF de una Orden de Compra como ARCHIVO en disco (ronda L,
    punto 1, 2026-09-09: hace falta un archivo real para poder adjuntarlo a
    un correo de Outlook). Usa la libreria xhtml2pdf (HTML/CSS -> PDF, pura
    Python, sin depender de un navegador ni de un binario externo) sobre
    una plantilla dedicada (ordenes/pdf_oc.html) -- DISTINTA de
    ordenes/imprimir.html, que sigue siendo la que se usa para "Ver /
    Imprimir" desde el navegador con window.print() (esa sigue igual, sin
    cambios). El PDF se sobrescribe cada vez que se genera: no se
    acumulan versiones viejas de la misma orden.

    Devuelve la ruta absoluta del archivo generado. Lanza una excepcion si
    xhtml2pdf no esta instalado o si reporta un error al generar."""
    try:
        from xhtml2pdf import pisa
    except ImportError as exc:
        raise RuntimeError(
            "Falta instalar la librería 'xhtml2pdf' (agregar "
            "'py -m pip install -r requirements.txt' y reintentar)."
        ) from exc

    lineas = orden.lineas.all()
    logo_abs_path = None
    if orden.empresa and orden.empresa.logo_nombre_archivo:
        candidato = os.path.join(EMPRESAS_LOGOS_DIR, orden.empresa.logo_nombre_archivo)
        if os.path.isfile(candidato):
            logo_abs_path = candidato.replace("\\", "/")

    html = render_template(
        "ordenes/pdf_oc.html",
        orden=orden,
        lineas=lineas,
        logo_abs_path=logo_abs_path,
        fecha_generacion=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )

    os.makedirs(PDF_OC_DIR, exist_ok=True)
    nombre_archivo = f"OC_{orden.numero_po}_{orden.id}.pdf".replace("/", "-").replace("\\", "-")
    ruta_pdf = os.path.join(PDF_OC_DIR, nombre_archivo)
    with open(ruta_pdf, "wb") as archivo_salida:
        resultado = pisa.CreatePDF(src=html, dest=archivo_salida)
    if resultado.err:
        raise RuntimeError("xhtml2pdf reportó errores generando el PDF de la orden.")
    return ruta_pdf


@app.route("/ordenes/<int:orden_id>/pdf")
def ordenes_pdf(orden_id):
    """Descarga directa del PDF de la orden (ronda L, punto 1) -- sirve
    tanto como fin en si mismo (el usuario quiere el archivo) como
    respaldo cuando 'Enviar por Outlook' no puede abrir Outlook (Outlook
    no instalado/no es Windows, etc.): el PDF queda generado igual y el
    usuario lo adjunta a mano."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.lineas.count() == 0:
        flash("La orden no tiene productos todavía, no se puede generar el PDF.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    try:
        ruta_pdf = generar_pdf_orden(orden)
    except Exception as exc:
        flash(f"No se pudo generar el PDF de la orden: {exc}", "danger")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
    directorio, nombre = os.path.split(ruta_pdf)
    return send_from_directory(
        directorio, nombre, as_attachment=True,
        download_name=f"OC_{orden.numero_po}.pdf".replace("/", "-"),
    )


@app.route("/ordenes/<int:orden_id>/enviar-outlook", methods=["POST"])
def ordenes_enviar_outlook(orden_id):
    """Envío de la OC al proveedor directo desde Outlook de escritorio
    (ronda L, punto 1, 2026-09-09): genera el PDF de la orden y abre un
    borrador de correo en el Outlook INSTALADO EN ESTE COMPUTADOR (via
    automatización COM con pywin32), con el proveedor como destinatario y
    el PDF ya adjunto. A propósito NO se envía solo (.Display(), no
    .Send()) -- el usuario revisa el borrador en Outlook y lo envía él
    mismo, para no arriesgarse a mandarle algo mal a un proveedor sin que
    nadie lo revise antes.

    Nota importante para el futuro (ver Pendiente / doc del proyecto):
    esto solo funciona porque hoy el programa corre LOCAL en este mismo
    computador donde está instalado Outlook. Si más adelante el programa
    se muda a un servidor en internet (ronda L, punto 1.1), esta
    automatización deja de poder alcanzar el Outlook del usuario y habría
    que rehacerla contra Microsoft Graph (Outlook / Microsoft 365 en la
    nube) en su lugar."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.lineas.count() == 0:
        flash("La orden no tiene productos todavía, no se puede generar el PDF.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    try:
        ruta_pdf = generar_pdf_orden(orden)
    except Exception as exc:
        flash(f"No se pudo generar el PDF de la orden: {exc}", "danger")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    destinatario = (orden.proveedor.contacto_email or "").strip()
    asunto, cuerpo = construir_correo_oc(orden)

    try:
        import win32com.client
    except ImportError:
        flash(
            "No se encontró la librería para conectar con Outlook (pywin32). Instálala con "
            "'py -m pip install -r requirements.txt' y vuelve a intentar (recuerda cerrar y "
            "volver a abrir el programa después). Mientras tanto, el PDF de la orden ya quedó "
            "generado: descárgalo con el botón 'Descargar PDF' y adjúntalo a mano.", "warning",
        )
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        if destinatario:
            mail.To = destinatario
        mail.Subject = asunto
        mail.Body = cuerpo
        mail.Attachments.Add(ruta_pdf)
        mail.Display()  # abre el borrador -- el envio lo hace el usuario desde Outlook
    except Exception as exc:
        flash(
            f"No se pudo abrir Outlook automáticamente ({exc}). Verifica que Outlook esté "
            "instalado en este computador. El PDF de la orden ya quedó generado: descárgalo con "
            "el botón 'Descargar PDF' y adjúntalo a mano en tu correo.", "warning",
        )
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))

    if destinatario:
        flash(
            f"Se abrió un borrador en Outlook para {destinatario} con la OC {orden.numero_po} "
            "adjunta en PDF. Revisa el correo y presiona Enviar desde Outlook.", "success",
        )
    else:
        flash(
            f"Se abrió un borrador en Outlook con la OC {orden.numero_po} adjunta en PDF, pero "
            "este proveedor no tiene correo de contacto cargado -- completa el destinatario a "
            "mano en el borrador antes de enviarlo.", "warning",
        )
    return redirect(url_for("ordenes_detalle", orden_id=orden_id))


@app.route("/ordenes/<int:orden_id>/mailto")
def ordenes_mailto(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    prov = orden.proveedor
    # Usa la misma plantilla configurable que el envío por Outlook (ronda L,
    # punto 1, 2026-09-09) para que ambas vías manden el mismo mensaje.
    asunto_txt, cuerpo_txt = construir_correo_oc(orden)
    asunto = quote(asunto_txt)
    cuerpo = quote(cuerpo_txt)
    destinatario = prov.contacto_email or ""
    mailto = f"mailto:{destinatario}?subject={asunto}&body={cuerpo}"
    return redirect(mailto)


# ---------------------------------------------------------------------------
# Modulo: Despachos (control de envio bajo tracking, una o varias ordenes)
# ---------------------------------------------------------------------------
# Cambio de flujo (2026-08-26, a pedido del usuario): "marcar despachada"
# dejo de ser una accion del modulo de Compras/Ordenes. Ahora ocurre al
# ASOCIAR una orden (ya CONFIRMADA) a un Despacho: ese es el momento en que
# se le asigna courier/tracking, y es tambien el momento en que sus lineas
# confirmadas pasan a 'Orden Despachada' (cantidad completa). Por eso:
#   - Las candidatas para asociar a un despacho son ordenes en estado
#     'Orden Confirmada' (ya no 'Orden Despachada' -- esa transicion la
#     hace esta misma pantalla).
#   - Una vez asociada, la orden sale del listado de Compras/Ordenes (ver
#     ordenes_list) porque ya paso a control de despacho.
#   - 'Generar costeo' desde un despacho exige que TODAS sus lineas esten
#     'Recibido' (ver _despacho_completo_recibido), no solo despachadas.

ETAPAS_ORDEN_DESPACHABLE = ["Orden Despachada", "Internación Aduanas", "Recibido"]


def _ordenes_listas_para_despachar():
    """Ordenes totalmente confirmadas (todas sus lineas activas en 'Orden
    Confirmada') que todavia no estan asociadas a ningun despacho -- son las
    candidatas para asociar a un Despacho nuevo o existente."""
    return (
        OrdenCompra.query.filter(
            OrdenCompra.estado == "Orden Confirmada",
            OrdenCompra.despacho_id.is_(None),
        )
        .order_by(OrdenCompra.id.desc())
        .all()
    )


def _marcar_orden_despachada(orden):
    """Avanza TODAS las lineas activas 'Orden Confirmada' de esta orden a
    'Orden Despachada' (cantidad completa). Se llama al asociar la orden a
    un Despacho -- ese es ahora el unico lugar donde una orden pasa a
    despachada."""
    hoy = date.today()
    for linea in orden.lineas:
        if linea.anulada or linea.etapa != "Orden Confirmada":
            continue
        linea.cantidad_despachada_cajas = linea.cantidad_cajas
        linea.fecha_despacho_real = hoy
        linea.etapa = "Orden Despachada"
    recalcular_estado_orden(orden)


@app.route("/despachos")
def despachos_list():
    despachos = Despacho.query.order_by(Despacho.id.desc()).all()
    hay_candidatas = bool(_ordenes_listas_para_despachar())
    return render_template("despachos/list.html", despachos=despachos, hay_candidatas=hay_candidatas)


@app.route("/despachos/nuevo", methods=["GET", "POST"])
def despachos_nuevo():
    candidatas = _ordenes_listas_para_despachar()
    if request.method == "POST":
        despacho = Despacho(
            numero_tracking=request.form.get("numero_tracking", "").strip(),
            courier=request.form.get("courier", "").strip(),
            status=request.form.get("status") or ESTADOS_DESPACHO[0],
            fecha_envio=parse_date(request.form.get("fecha_envio")),
            fecha_estimada_llegada=parse_date(request.form.get("fecha_estimada_llegada")),
            notas=request.form.get("notas", "").strip(),
        )
        db.session.add(despacho)
        db.session.flush()

        orden_ids = request.form.getlist("orden_ids")
        asociadas = 0
        if orden_ids:
            ordenes = OrdenCompra.query.filter(
                OrdenCompra.id.in_(orden_ids), OrdenCompra.despacho_id.is_(None)
            ).all()
            for orden in ordenes:
                orden.despacho_id = despacho.id
                _marcar_orden_despachada(orden)
                asociadas += 1
        db.session.commit()
        flash(f"Despacho creado con {asociadas} orden(es) asociada(s) y marcada(s) como despachadas.", "success")
        return redirect(url_for("despachos_detalle", despacho_id=despacho.id))
    return render_template(
        "despachos/form.html", candidatas=candidatas, estados=ESTADOS_DESPACHO, despacho=None
    )


@app.route("/despachos/<int:despacho_id>")
def despachos_detalle(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    ordenes = despacho.ordenes.order_by(OrdenCompra.id).all()
    candidatas = _ordenes_listas_para_despachar()
    return render_template(
        "despachos/detalle.html",
        despacho=despacho,
        ordenes=ordenes,
        candidatas=candidatas,
        estados=ESTADOS_DESPACHO,
        listo_para_costeo=_despacho_completo_recibido(despacho),
    )


@app.route("/despachos/<int:despacho_id>/editar", methods=["POST"])
def despachos_editar(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    despacho.numero_tracking = request.form.get("numero_tracking", "").strip()
    despacho.courier = request.form.get("courier", "").strip()
    despacho.status = request.form.get("status") or ESTADOS_DESPACHO[0]
    despacho.fecha_envio = parse_date(request.form.get("fecha_envio"))
    despacho.fecha_estimada_llegada = parse_date(request.form.get("fecha_estimada_llegada"))
    despacho.notas = request.form.get("notas", "").strip()
    despacho.url_tracking_manual = request.form.get("url_tracking_manual", "").strip() or None
    db.session.commit()
    flash("Datos del despacho actualizados.", "success")
    return redirect(url_for("despachos_detalle", despacho_id=despacho.id))


@app.route("/despachos/<int:despacho_id>/asociar", methods=["POST"])
def despachos_asociar(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    orden_ids = request.form.getlist("orden_ids")
    if not orden_ids:
        flash("Selecciona al menos una orden para asociar a este despacho.", "warning")
        return redirect(url_for("despachos_detalle", despacho_id=despacho.id))
    ordenes = OrdenCompra.query.filter(
        OrdenCompra.id.in_(orden_ids), OrdenCompra.despacho_id.is_(None)
    ).all()
    for orden in ordenes:
        orden.despacho_id = despacho.id
        _marcar_orden_despachada(orden)
    db.session.commit()
    flash(f"{len(ordenes)} orden(es) asociada(s) al despacho y marcada(s) como despachadas.", "success")
    return redirect(url_for("despachos_detalle", despacho_id=despacho.id))


@app.route("/despachos/<int:despacho_id>/orden/<int:orden_id>/quitar", methods=["POST"])
def despachos_quitar_orden(despacho_id, orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.despacho_id != despacho_id:
        abort(404)
    orden.despacho_id = None
    # Si ninguna linea avanzo mas alla de 'Orden Despachada' todavia (no
    # entro a aduana ni fue recibida), se deshace tambien el despacho de
    # sus lineas -- as vuelve a quedar 'Orden Confirmada', visible de nuevo
    # en Compras/Ordenes y disponible como candidata para otro despacho. Si
    # ya avanzo mas (aduana/recibido), no se toca -- fisicamente ya salio.
    if orden.estado == "Orden Despachada":
        for linea in orden.lineas:
            if linea.anulada or linea.etapa != "Orden Despachada":
                continue
            linea.etapa = "Orden Confirmada"
            linea.cantidad_despachada_cajas = 0
            linea.fecha_despacho_real = None
        recalcular_estado_orden(orden)
    db.session.commit()
    flash(f"Orden {orden.numero_po} desasociada del despacho.", "info")
    return redirect(url_for("despachos_detalle", despacho_id=despacho_id))


def _avanzar_lineas_despacho(despacho, etapa_desde, etapa_hasta, campo_fecha):
    hoy = date.today()
    marcadas = 0
    for orden in despacho.ordenes:
        cambiaron = False
        for linea in orden.lineas:
            if linea.anulada or linea.etapa != etapa_desde:
                continue
            setattr(linea, campo_fecha, hoy)
            linea.etapa = etapa_hasta
            marcadas += 1
            cambiaron = True
        if cambiaron:
            recalcular_estado_orden(orden)
    return marcadas


@app.route("/despachos/<int:despacho_id>/aduana-masivo", methods=["POST"])
def despachos_aduana_masivo(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    marcadas = _avanzar_lineas_despacho(
        despacho, "Orden Despachada", "Internación Aduanas", "fecha_llegada_aduana"
    )
    db.session.commit()
    flash(f"{marcadas} línea(s), de todas las órdenes de este despacho, marcadas en internación de aduanas.", "success")
    return redirect(url_for("despachos_detalle", despacho_id=despacho.id))


@app.route("/despachos/<int:despacho_id>/recibido-masivo", methods=["POST"])
def despachos_recibido_masivo(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    marcadas = _avanzar_lineas_despacho(
        despacho, "Internación Aduanas", "Recibido", "fecha_recepcion_bodega"
    )
    db.session.commit()
    flash(f"{marcadas} línea(s), de todas las órdenes de este despacho, marcadas como recibidas en bodega.", "success")
    return redirect(url_for("despachos_detalle", despacho_id=despacho.id))


@app.route("/despachos/<int:despacho_id>/eliminar", methods=["POST"])
def despachos_eliminar(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    for orden in despacho.ordenes.all():
        orden.despacho_id = None
    db.session.delete(despacho)
    db.session.commit()
    flash("Despacho eliminado (las órdenes asociadas quedan sueltas, sin perder datos).", "success")
    return redirect(url_for("despachos_list"))


def _despacho_completo_recibido(despacho):
    """True si TODAS las lineas activas (no anuladas) de TODAS las ordenes
    de este despacho ya estan 'Recibido'. El costeo solo tiene sentido una
    vez que la mercaderia llego fisicamente a bodega -- antes de eso puede
    seguir cambiando (ej. una diferencia detectada al abrir las cajas)."""
    lineas_activas = [
        l for o in despacho.ordenes for l in o.lineas if not l.anulada
    ]
    if not lineas_activas:
        return False
    return all(l.etapa == "Recibido" for l in lineas_activas)


@app.route("/despachos/<int:despacho_id>/generar-importacion", methods=["POST"])
def despachos_generar_importacion(despacho_id):
    """Crea una Importacion (Costeo) precargada con los datos de este
    despacho: un Parcial nuevo con una ParcialLinea por cada linea
    despachada (no anulada) de cada orden asociada, usando la cantidad
    REALMENTE despachada (no la solicitada) y el precio pactado convertido
    a valor por unidad. Asi el usuario no tiene que retipear los productos
    -- solo completa los datos que no vienen de la orden (tipo de cambio,
    flete, seguro, gastos). Requiere que TODAS las lineas del despacho
    esten 'Recibido' (ver _despacho_completo_recibido)."""
    despacho = Despacho.query.get_or_404(despacho_id)
    ordenes = despacho.ordenes.all()
    if not ordenes:
        flash("Este despacho todavía no tiene órdenes asociadas.", "warning")
        return redirect(url_for("despachos_detalle", despacho_id=despacho.id))

    if despacho.importaciones_generadas:
        # Punto 2 (2026-08-26): evitar generar el costeo mas de una vez
        # desde el mismo despacho -- generaria lineas/importaciones
        # duplicadas. Si de verdad hace falta otra, se borra la existente
        # primero (boton "Eliminar" en Costeo) y se vuelve a generar.
        imp_existente = despacho.importaciones_generadas[0]
        flash(
            f"Este despacho ya tiene un costeo generado (Importación #{imp_existente.id}). "
            "Para evitar duplicados no se puede generar otra vez desde aquí -- si necesitas "
            "rehacerla, elimina esa importación primero.",
            "warning",
        )
        return redirect(url_for("importaciones_detalle", importacion_id=imp_existente.id))

    if not _despacho_completo_recibido(despacho):
        flash(
            "Para generar el costeo, primero todas las líneas de este despacho deben estar "
            "'Recibido' en bodega (usa 'Marcar todo en aduana' / 'Marcar todo recibido' arriba).",
            "warning",
        )
        return redirect(url_for("despachos_detalle", despacho_id=despacho.id))

    proveedores_distintos = despacho.proveedores
    proveedor_id = proveedores_distintos[0].id if proveedores_distintos else ordenes[0].proveedor_id
    moneda = ordenes[0].moneda or "USD"

    imp = Importacion(
        proveedor_id=proveedor_id,
        despacho_id=despacho.id,
        numero_factura="",
        moneda_factura=moneda,
        tipo_cambio_aduanero=0,
        tipo_cambio_moneda_usd=1,
        notas=f"Generada automáticamente desde el despacho {despacho.numero_tracking or ('#' + str(despacho.id))}.",
    )
    db.session.add(imp)
    db.session.flush()

    parcial = Parcial(
        importacion_id=imp.id,
        numero_parcial=despacho.numero_tracking or "",
        referencia=despacho.courier or "",
    )
    db.session.add(parcial)
    db.session.flush()

    lineas_creadas = 0
    for orden in ordenes:
        for linea in orden.lineas:
            if linea.anulada:
                continue
            cantidad_despachada = linea.cantidad_despachada_cajas or 0
            if cantidad_despachada <= 0:
                continue
            empaque = linea.producto.empaque or 1
            cantidad_unidades = cantidad_despachada * empaque
            valor_unitario = (linea.precio_unitario_pactado or 0) / empaque if empaque else 0
            db.session.add(ParcialLinea(
                parcial_id=parcial.id,
                producto_id=linea.producto_id,
                orden_compra_linea_id=linea.id,
                # Ronda M (2026-09-10): si la linea de la orden tenia una
                # variante elegida (ej. dioptria de un lente MEDICONTUR), esa
                # es la que se traspasa al Costeo -- no el codigo padre.
                codigo=linea.codigo_mostrar,
                descripcion=linea.descripcion_mostrar,
                cantidad_unidades=cantidad_unidades,
                valor_unitario_moneda=valor_unitario,
            ))
            lineas_creadas += 1

    db.session.commit()

    mensaje = (
        f"Importación creada desde el despacho, con {lineas_creadas} línea(s) traídas automáticamente "
        f"de {len(ordenes)} orden(es). Completa el tipo de cambio, flete/seguro y gastos para terminar el costeo."
    )
    if len(proveedores_distintos) > 1:
        mensaje += (
            f" Ojo: este despacho consolida {len(proveedores_distintos)} proveedores distintos "
            f"({', '.join(p.nombre for p in proveedores_distintos)}) — se asignó "
            f"'{proveedores_distintos[0].nombre}' por defecto, verifica si corresponde ajustarlo."
        )
    flash(mensaje, "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


# ---------------------------------------------------------------------------
# Modulo: Costeo de Importaciones
# ---------------------------------------------------------------------------

def _despachos_pendientes_de_costeo():
    """Despachos con TODAS sus lineas activas 'Recibido' que todavia no
    generaron ninguna Importacion -- son los que justifican habilitar
    'Nueva importacion' (punto 7, 2026-08-26): si no hay ninguno, crear una
    importacion a mano no tiene de donde salir (no hay nada despachado y
    recibido esperando costeo)."""
    return [
        d for d in Despacho.query.all()
        if not d.importaciones_generadas and _despacho_completo_recibido(d)
    ]


@app.route("/importaciones")
def importaciones_list():
    proveedor_id = request.args.get("proveedor_id", "")
    query = Importacion.query
    if proveedor_id:
        query = query.filter_by(proveedor_id=proveedor_id)
    importaciones = query.order_by(Importacion.id.desc()).all()
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    resumenes = {}
    for imp in importaciones:
        resultado = costing.calcular_costeo(imp)
        resumenes[imp.id] = resultado["totales"]
    hay_despachos_pendientes = bool(_despachos_pendientes_de_costeo())
    return render_template(
        "importaciones/list.html", importaciones=importaciones, proveedores=proveedores,
        proveedor_sel=proveedor_id, resumenes=resumenes,
        hay_despachos_pendientes=hay_despachos_pendientes,
    )


@app.route("/importaciones/comparativo")
def importaciones_comparativo():
    """Vista comparativa (punto 8, 2026-08-26): todas las importaciones
    valoradas en la misma moneda para poder compararlas entre si, sin
    importar en que moneda vino cada factura. Se ofrece Pesos (CLP) y
    Dolares (USD) -- ambos siempre calculables con los tipos de cambio que
    ya se cargan por importacion. No se ofrece Euros: segun confirmo el
    usuario, la DIN declara el CIF en USD aunque la factura venga en
    Euros, asi que las comparaciones reales que necesita son CLP/USD; para
    el registro y visualizacion de cada importacion en particular se sigue
    mostrando su moneda de factura original (ver importaciones_detalle)."""
    moneda_vista = request.args.get("moneda", "CLP")
    if moneda_vista not in ("CLP", "USD"):
        moneda_vista = "CLP"
    proveedor_id = request.args.get("proveedor_id", "")
    query = Importacion.query
    if proveedor_id:
        query = query.filter_by(proveedor_id=proveedor_id)
    importaciones = query.order_by(Importacion.id.desc()).all()
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    filas = []
    for imp in importaciones:
        totales = costing.calcular_costeo(imp)["totales"]
        tc = imp.tipo_cambio_aduanero or 0
        if moneda_vista == "USD":
            fila = {
                "fob": totales["fob_usd"],
                "cif": totales["cif_usd"],
                "gastos": (totales["gastos_clp"] / tc) if tc else 0,
                "costo_total": totales["costo_total_usd"],
            }
        else:
            fila = {
                "fob": totales["fob_usd"] * tc,
                "cif": totales["cif_clp"],
                "gastos": totales["gastos_clp"],
                "costo_total": totales["costo_total_clp"],
            }
        fila["imp"] = imp
        filas.append(fila)
    total_general = sum(f["costo_total"] for f in filas)
    return render_template(
        "importaciones/comparativo.html", filas=filas, proveedores=proveedores,
        proveedor_sel=proveedor_id, moneda_vista=moneda_vista, total_general=total_general,
    )


def _tipo_cambio_mes(anio, mes):
    return TipoCambioMensual.query.filter_by(anio=anio, mes=mes).first()


@app.route("/importaciones/nueva", methods=["GET", "POST"])
def importaciones_nueva():
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    if request.method == "POST":
        imp = Importacion(
            proveedor_id=int(request.form["proveedor_id"]),
            numero_factura=request.form.get("numero_factura", "").strip(),
            fecha_factura=parse_date(request.form.get("fecha_factura")),
            moneda_factura=request.form.get("moneda_factura", "EURO"),
            condicion_compra=request.form.get("condicion_compra", "EXW"),
            tipo_cambio_aduanero=float(request.form.get("tipo_cambio_aduanero") or 0),
            tipo_cambio_moneda_usd=float(request.form.get("tipo_cambio_moneda_usd") or 1),
            notas=request.form.get("notas", "").strip(),
        )
        db.session.add(imp)
        db.session.commit()
        flash(
            f"Importación creada para {imp.proveedor.nombre}. Agrega los parciales/productos y, si "
            "corresponde, Flete/Seguro/Handling Fee desde esta misma pantalla.", "success",
        )
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))
    hoy = date.today()
    tc_mes = _tipo_cambio_mes(hoy.year, hoy.month)
    return render_template(
        "importaciones/form.html", proveedores=proveedores, condiciones=CONDICIONES_COMPRA,
        tc_mes=tc_mes,
    )


@app.route("/importaciones/<int:importacion_id>")
def importaciones_detalle(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    resultado = costing.calcular_costeo(imp)
    tipos_cambio = TipoCambioMensual.query.order_by(
        TipoCambioMensual.anio.desc(), TipoCambioMensual.mes.desc()
    ).all()
    return render_template(
        "importaciones/detalle.html", imp=imp, resultado=resultado,
        regimenes=REGIMENES_PARCIAL, vias=VIAS_EMBARQUE,
        condiciones=CONDICIONES_COMPRA, conceptos_gasto=CONCEPTOS_GASTO,
        conceptos_item_factura=CONCEPTOS_ITEM_FACTURA, tipos_cambio=tipos_cambio,
        tipos_documento_gasto=TIPOS_DOCUMENTO_GASTO,
    )


@app.route("/importaciones/<int:importacion_id>/editar", methods=["POST"])
def importaciones_editar(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    imp.numero_factura = request.form.get("numero_factura", "").strip()
    imp.fecha_factura = parse_date(request.form.get("fecha_factura"))
    imp.moneda_factura = request.form.get("moneda_factura", "EURO")
    imp.condicion_compra = request.form.get("condicion_compra", "EXW")
    imp.tipo_cambio_aduanero = float(request.form.get("tipo_cambio_aduanero") or 0)
    imp.tipo_cambio_moneda_usd = float(request.form.get("tipo_cambio_moneda_usd") or 1)
    imp.notas = request.form.get("notas", "").strip()
    db.session.commit()
    flash("Datos de la importación actualizados.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


@app.route("/importaciones/<int:importacion_id>/eliminar", methods=["POST"])
def importaciones_eliminar(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    carpeta_legajo = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(imp.id))
    if os.path.isdir(carpeta_legajo):
        shutil.rmtree(carpeta_legajo, ignore_errors=True)
    db.session.delete(imp)
    db.session.commit()
    flash("Importación eliminada.", "success")
    return redirect(url_for("importaciones_list"))


# --- Parciales ---

@app.route("/importaciones/<int:importacion_id>/parciales/nuevo", methods=["POST"])
def parciales_nuevo(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    parcial = Parcial(
        importacion_id=imp.id,
        numero_parcial=request.form.get("numero_parcial", "").strip(),
        referencia=request.form.get("referencia", "").strip(),
        tipo_regimen=request.form.get("tipo_regimen", "General"),
        via_embarque=request.form.get("via_embarque", "Marítimo"),
        notas=request.form.get("notas", "").strip(),
    )
    db.session.add(parcial)
    db.session.commit()
    flash(f"Parcial '{parcial.numero_parcial or parcial.referencia}' agregado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


@app.route("/parciales/<int:parcial_id>/editar", methods=["POST"])
def parciales_editar(parcial_id):
    parcial = Parcial.query.get_or_404(parcial_id)
    parcial.numero_parcial = request.form.get("numero_parcial", "").strip()
    parcial.referencia = request.form.get("referencia", "").strip()
    parcial.tipo_regimen = request.form.get("tipo_regimen", "General")
    parcial.via_embarque = request.form.get("via_embarque", "Marítimo")
    parcial.notas = request.form.get("notas", "").strip()
    db.session.commit()
    flash(f"Parcial '{parcial.numero_parcial or parcial.referencia}' actualizado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))


@app.route("/parciales/<int:parcial_id>/eliminar", methods=["POST"])
def parciales_eliminar(parcial_id):
    parcial = Parcial.query.get_or_404(parcial_id)
    importacion_id = parcial.importacion_id
    db.session.delete(parcial)
    db.session.commit()
    flash("Parcial eliminado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


# --- Lineas de un parcial ---

def _guardar_lotes_linea(linea, form):
    """Reemplaza el desglose de lotes de una linea a partir de las listas
    paralelas lote_codigo/lote_fecha/lote_cantidad enviadas por el modal de
    edicion (boton '+' para agregar filas, ronda J punto 2). Filas
    completamente vacias se ignoran. codigo_lote y fecha_vencimiento de la
    linea se actualizan como resumen (compatibilidad hacia atras) con los
    datos del lote de vencimiento mas proximo."""
    codigos = form.getlist("lote_codigo")
    fechas = form.getlist("lote_fecha")
    cantidades = form.getlist("lote_cantidad")

    nuevos = []
    for i in range(max(len(codigos), len(fechas), len(cantidades))):
        codigo = codigos[i].strip() if i < len(codigos) else ""
        fecha = parse_date(fechas[i]) if i < len(fechas) else None
        cantidad = parse_int(cantidades[i] if i < len(cantidades) else "", default=0)
        if not codigo and not fecha and not cantidad:
            continue
        nuevos.append(ParcialLineaLote(codigo_lote=codigo, fecha_vencimiento=fecha, cantidad_unidades=cantidad))

    linea.lotes = nuevos
    if nuevos:
        primero = min(nuevos, key=lambda lo: (lo.fecha_vencimiento is None, lo.fecha_vencimiento or date.max))
        linea.codigo_lote = ", ".join(sorted({lo.codigo_lote for lo in nuevos if lo.codigo_lote}))
        linea.fecha_vencimiento = primero.fecha_vencimiento


@app.route("/parciales/<int:parcial_id>/lineas/nueva", methods=["POST"])
def parcial_lineas_nueva(parcial_id):
    parcial = Parcial.query.get_or_404(parcial_id)
    producto_id = request.form.get("producto_id") or None
    linea = ParcialLinea(
        parcial_id=parcial.id,
        producto_id=int(producto_id) if producto_id else None,
        codigo=request.form.get("codigo", "").strip(),
        descripcion=request.form.get("descripcion", "").strip(),
        codigo_lote=request.form.get("codigo_lote", "").strip(),
        fecha_vencimiento=parse_date(request.form.get("fecha_vencimiento")),
        cantidad_unidades=parse_int(request.form.get("cantidad_unidades"), default=0),
        valor_unitario_moneda=float(request.form.get("valor_unitario_moneda") or 0),
    )
    # El desglose por lotes (ronda J) arranca con un unico lote que refleja
    # el codigo/fecha/cantidad ya ingresados en este mismo formulario -- asi
    # una linea recien creada nunca queda "sin lotes"; si el usuario necesita
    # mas de uno, los agrega despues desde "Editar".
    if linea.cantidad_unidades:
        linea.lotes.append(ParcialLineaLote(
            codigo_lote=linea.codigo_lote,
            fecha_vencimiento=linea.fecha_vencimiento,
            cantidad_unidades=linea.cantidad_unidades,
        ))
    db.session.add(linea)
    db.session.commit()
    flash(f"Línea '{linea.codigo}' agregada al parcial {parcial.numero_parcial}.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))


@app.route("/parcial_lineas/<int:linea_id>/editar", methods=["POST"])
def parcial_lineas_editar(linea_id):
    linea = ParcialLinea.query.get_or_404(linea_id)
    linea.codigo = request.form.get("codigo", "").strip()
    linea.descripcion = request.form.get("descripcion", "").strip()
    linea.cantidad_unidades = parse_int(request.form.get("cantidad_unidades"), default=linea.cantidad_unidades)
    linea.valor_unitario_moneda = float(request.form.get("valor_unitario_moneda") or 0)
    _guardar_lotes_linea(linea, request.form)
    db.session.commit()
    flash("Línea actualizada.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=linea.parcial.importacion_id))


@app.route("/parcial_lineas/<int:linea_id>/eliminar", methods=["POST"])
def parcial_lineas_eliminar(linea_id):
    linea = ParcialLinea.query.get_or_404(linea_id)
    importacion_id = linea.parcial.importacion_id
    db.session.delete(linea)
    db.session.commit()
    flash("Línea eliminada.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


@app.route("/parcial_lineas/mover-masivo", methods=["POST"])
def parcial_lineas_mover_masivo():
    """Reasigna en bloque un grupo de productos a otro Parcial de la MISMA
    importacion -- util cuando un despacho se genero de una vez en un solo
    parcial y despues hay que repartir sus productos entre los DIN reales
    (ej. 2 o mas parciales por regimen aduanero distinto)."""
    destino_id = request.form.get("destino_parcial_id")
    linea_ids = request.form.getlist("linea_ids")
    destino = Parcial.query.get_or_404(destino_id)

    movidas = 0
    for lid in linea_ids:
        try:
            linea = ParcialLinea.query.get(int(lid))
        except (TypeError, ValueError):
            continue
        if not linea or linea.parcial.importacion_id != destino.importacion_id:
            continue
        linea.parcial_id = destino.id
        movidas += 1

    db.session.commit()
    flash(f"{movidas} producto(s) movido(s) al parcial '{destino.referencia or destino.numero_parcial or destino.id}'.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=destino.importacion_id))


# --- Carga masiva de Lotes (ronda J, punto 3) ---

def _parsear_fecha_archivo_lotes(valor):
    """Acepta fecha ya como date/datetime (celda Excel con formato fecha) o
    como texto en varios formatos comunes."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def _leer_filas_archivo_lotes(file_storage):
    """Lee un archivo .xlsx/.xls o .csv de carga de lotes. Espera columnas
    con encabezados flexibles (Código de Producto / Código, Lote, Fecha
    vencimiento) y UNA FILA POR CADA UNIDAD FISICA (sin columna de
    cantidad) -- se agrupan despues por (código, lote, fecha) para obtener
    la cantidad de cada lote, de forma que siempre "cuadre" con el total
    de unidades. Devuelve una lista de tuplas (codigo_producto, lote,
    fecha_o_None), o None si no se reconocieron los encabezados."""
    nombre = (file_storage.filename or "").lower()
    if nombre.endswith(".csv"):
        contenido = file_storage.read().decode("utf-8-sig", errors="ignore")
        filas_crudas = list(csv.reader(io.StringIO(contenido)))
    else:
        wb = openpyxl.load_workbook(file_storage, read_only=True, data_only=True)
        ws = wb.active
        filas_crudas = [list(fila) for fila in ws.iter_rows(values_only=True)]

    if not filas_crudas:
        return []

    encabezado = [str(c or "").strip().lower() for c in filas_crudas[0]]

    def _buscar_columna(*claves):
        for i, nombre_col in enumerate(encabezado):
            for clave in claves:
                if clave in nombre_col:
                    return i
        return None

    idx_codigo = _buscar_columna("codigo de producto", "código de producto", "codigo", "código")
    idx_lote = _buscar_columna("lote")
    idx_fecha = _buscar_columna("fecha vencimiento", "fecha de vencimiento", "vencimiento", "fecha")

    if idx_codigo is None or idx_lote is None:
        return None

    filas = []
    for fila in filas_crudas[1:]:
        if idx_codigo >= len(fila):
            continue
        codigo = str(fila[idx_codigo] or "").strip()
        if not codigo:
            continue
        lote = str(fila[idx_lote] or "").strip() if idx_lote < len(fila) else ""
        fecha_cruda = fila[idx_fecha] if (idx_fecha is not None and idx_fecha < len(fila)) else None
        filas.append((codigo, lote, _parsear_fecha_archivo_lotes(fecha_cruda)))
    return filas


@app.route("/parciales/<int:parcial_id>/lotes/cargar", methods=["POST"])
def parcial_lotes_cargar(parcial_id):
    """Ronda J, punto 3: carga masiva de lotes desde un archivo con Código
    de Producto / Lote / Fecha vencimiento (una fila por unidad). Reemplaza
    el desglose de lotes de cada producto encontrado en el archivo; los
    productos del parcial que no aparecen en el archivo no se tocan."""
    parcial = Parcial.query.get_or_404(parcial_id)
    archivo = request.files.get("archivo_lotes")
    if not archivo or not archivo.filename:
        flash("Selecciona un archivo para cargar los lotes.", "warning")
        return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))

    extension = os.path.splitext(archivo.filename)[1].lower()
    if extension not in (".xlsx", ".xls", ".csv"):
        flash("Formato no soportado. Sube un archivo .xlsx o .csv.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))

    try:
        filas = _leer_filas_archivo_lotes(archivo)
    except Exception:
        flash("No se pudo leer el archivo. Verifica que no esté dañado o abierto en otro programa.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))

    if filas is None:
        flash("No se reconocieron las columnas del archivo. Debe incluir 'Código de Producto', 'Lote' y 'Fecha vencimiento'.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))
    if not filas:
        flash("El archivo no tiene filas para cargar.", "warning")
        return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))

    lineas_por_codigo = {}
    for linea in parcial.lineas:
        lineas_por_codigo.setdefault((linea.codigo or "").strip().lower(), linea)

    conteo = {}
    for codigo, lote, fecha in filas:
        grupo = (codigo.strip().lower(), lote, fecha)
        conteo[grupo] = conteo.get(grupo, 0) + 1

    nuevos_por_linea = {}
    codigos_no_encontrados = set()
    for (codigo_key, lote, fecha), cantidad in conteo.items():
        linea = lineas_por_codigo.get(codigo_key)
        if not linea:
            codigos_no_encontrados.add(codigo_key)
            continue
        nuevos_por_linea.setdefault(linea.id, []).append(
            ParcialLineaLote(codigo_lote=lote, fecha_vencimiento=fecha, cantidad_unidades=cantidad)
        )

    for linea_id, nuevos in nuevos_por_linea.items():
        linea = ParcialLinea.query.get(linea_id)
        linea.lotes = nuevos
        primero = min(nuevos, key=lambda lo: (lo.fecha_vencimiento is None, lo.fecha_vencimiento or date.max))
        linea.codigo_lote = ", ".join(sorted({lo.codigo_lote for lo in nuevos if lo.codigo_lote}))
        linea.fecha_vencimiento = primero.fecha_vencimiento

    db.session.commit()

    mensaje = f"Lotes cargados para {len(nuevos_por_linea)} producto(s)."
    if codigos_no_encontrados:
        mensaje += f" {len(codigos_no_encontrados)} código(s) del archivo no se encontraron en este parcial y se ignoraron."
    flash(mensaje, "success" if nuevos_por_linea else "warning")
    return redirect(url_for("importaciones_detalle", importacion_id=parcial.importacion_id))


# --- Gastos compartidos ---

def _monto_clp_gasto(request_form):
    """Calcula el monto en CLP a partir de moneda/tipo_cambio, para que un
    gasto negociado en otra moneda (ej. Flete Int. facturado en USD con su
    propio tipo de cambio, distinto del Dolar Aduanero de la importacion)
    quede igual convertido a CLP para el prorrateo (punto 4, 2026-08-26).
    Si la moneda es CLP, tipo_cambio se fuerza a 1 y monto_clp = monto_original."""
    moneda = request_form.get("moneda", "CLP") or "CLP"
    monto_original = float(request_form.get("monto_original") or 0)
    tipo_cambio = float(request_form.get("tipo_cambio") or 1) if moneda != "CLP" else 1
    monto_clp = monto_original if moneda == "CLP" else monto_original * tipo_cambio
    return moneda, tipo_cambio, monto_original, monto_clp


@app.route("/importaciones/<int:importacion_id>/gastos/nuevo", methods=["POST"])
def gastos_nuevo(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    moneda, tipo_cambio, monto_original, monto_clp = _monto_clp_gasto(request.form)
    gasto = GastoImportacion(
        importacion_id=imp.id,
        concepto=request.form.get("concepto", "").strip() or "Otros",
        moneda=moneda,
        tipo_cambio=tipo_cambio,
        monto_original=monto_original,
        monto_clp=monto_clp,
        referencia=request.form.get("referencia", "").strip(),
    )
    parcial_ids = request.form.getlist("parcial_ids")
    if parcial_ids:
        gasto.parciales_aplicables = Parcial.query.filter(Parcial.id.in_(parcial_ids)).all()
    db.session.add(gasto)
    db.session.commit()
    flash(f"Gasto '{gasto.concepto}' agregado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


@app.route("/gastos/<int:gasto_id>/editar", methods=["POST"])
def gastos_editar(gasto_id):
    gasto = GastoImportacion.query.get_or_404(gasto_id)
    moneda, tipo_cambio, monto_original, monto_clp = _monto_clp_gasto(request.form)
    gasto.concepto = request.form.get("concepto", "").strip() or "Otros"
    gasto.moneda = moneda
    gasto.tipo_cambio = tipo_cambio
    gasto.monto_original = monto_original
    gasto.monto_clp = monto_clp
    gasto.referencia = request.form.get("referencia", "").strip()
    parcial_ids = request.form.getlist("parcial_ids")
    gasto.parciales_aplicables = (
        Parcial.query.filter(Parcial.id.in_(parcial_ids)).all() if parcial_ids else []
    )
    db.session.commit()
    flash("Gasto actualizado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=gasto.importacion_id))


@app.route("/gastos/<int:gasto_id>/eliminar", methods=["POST"])
def gastos_eliminar(gasto_id):
    gasto = GastoImportacion.query.get_or_404(gasto_id)
    importacion_id = gasto.importacion_id
    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(gasto.id))
    if os.path.isdir(carpeta_gasto):
        shutil.rmtree(carpeta_gasto, ignore_errors=True)
    db.session.delete(gasto)
    db.session.commit()
    flash("Gasto eliminado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


@app.route("/importaciones/<int:importacion_id>/gastos/eliminar-multiple", methods=["POST"])
def gastos_eliminar_multiple(importacion_id):
    """Eliminacion masiva de Gastos compartidos -- ronda I, 2026-09-01,
    mismo patron que cargos_eliminar_multiple. Borra tambien la carpeta de
    documentos de cada gasto eliminado (mismo cuidado que gastos_eliminar
    de a uno)."""
    imp = Importacion.query.get_or_404(importacion_id)
    ids = request.form.getlist("gasto_ids")
    gastos = GastoImportacion.query.filter(
        GastoImportacion.id.in_(ids), GastoImportacion.importacion_id == imp.id
    ).all()
    for gasto in gastos:
        carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(gasto.id))
        if os.path.isdir(carpeta_gasto):
            shutil.rmtree(carpeta_gasto, ignore_errors=True)
        db.session.delete(gasto)
    if gastos:
        db.session.commit()
        flash(f"{len(gastos)} gasto(s) eliminados.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


# --- Documentos de un Gasto compartido (comprobante de pago, factura del
# agente, etc.) -- ronda G, punto 7. Mismo patron que los documentos de
# Orden, pero admite ademas Word/Excel. ---

def _extension_valida_gasto(nombre_archivo):
    _, ext = os.path.splitext(nombre_archivo.lower())
    return ext in EXTENSIONES_PERMITIDAS_GASTO


@app.route("/gastos/<int:gasto_id>/documentos/subir", methods=["POST"])
def gastos_documento_subir(gasto_id):
    gasto = GastoImportacion.query.get_or_404(gasto_id)
    archivo = request.files.get("archivo")
    tipo = request.form.get("tipo", "Otro")

    if not archivo or archivo.filename == "":
        flash("Selecciona un archivo para subir.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=gasto.importacion_id))

    if not _extension_valida_gasto(archivo.filename):
        flash("Solo se permiten archivos PDF, Word, Excel o imagenes (JPG, PNG).", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=gasto.importacion_id))

    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(gasto.id))
    os.makedirs(carpeta_gasto, exist_ok=True)

    _, ext = os.path.splitext(archivo.filename)
    nombre_disco = f"{uuid.uuid4().hex}{ext.lower()}"
    archivo.save(os.path.join(carpeta_gasto, nombre_disco))

    doc = GastoDocumento(
        gasto_id=gasto.id,
        tipo=tipo if tipo in TIPOS_DOCUMENTO_GASTO else "Otro",
        nombre_original=archivo.filename,
        nombre_archivo=nombre_disco,
    )
    db.session.add(doc)
    db.session.commit()
    flash(f"Documento '{archivo.filename}' subido correctamente.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=gasto.importacion_id))


@app.route("/gastos/<int:gasto_id>/documentos/<int:doc_id>/ver")
def gastos_documento_ver(gasto_id, doc_id):
    doc = GastoDocumento.query.get_or_404(doc_id)
    if doc.gasto_id != gasto_id:
        abort(404)
    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(doc.gasto_id))
    return send_from_directory(
        carpeta_gasto, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=False
    )


@app.route("/gastos/<int:gasto_id>/documentos/<int:doc_id>/descargar")
def gastos_documento_descargar(gasto_id, doc_id):
    doc = GastoDocumento.query.get_or_404(doc_id)
    if doc.gasto_id != gasto_id:
        abort(404)
    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(doc.gasto_id))
    return send_from_directory(
        carpeta_gasto, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=True
    )


@app.route("/gastos/<int:gasto_id>/documentos/<int:doc_id>/eliminar", methods=["POST"])
def gastos_documento_eliminar(gasto_id, doc_id):
    doc = GastoDocumento.query.get_or_404(doc_id)
    if doc.gasto_id != gasto_id:
        abort(404)
    gasto = GastoImportacion.query.get_or_404(gasto_id)
    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(doc.gasto_id))
    ruta = os.path.join(carpeta_gasto, doc.nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)
    db.session.delete(doc)
    db.session.commit()
    flash("Documento eliminado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=gasto.importacion_id))


# --- "Legajo": archivo consolidado a nivel de toda la importacion (ronda H,
# punto 3 -- opcion "Adjuntar > Legajo" del menu de engranaje de Gastos
# compartidos). A diferencia de GastoDocumento (que cuelga de un Gasto
# puntual), el Legajo cuelga directo de la Importacion -- pensado para un
# solo archivo que junte todas las facturas/documentos del embarque. ---

def _extension_valida_legajo(nombre_archivo):
    _, ext = os.path.splitext(nombre_archivo.lower())
    return ext in EXTENSIONES_PERMITIDAS_GASTO


@app.route("/importaciones/<int:importacion_id>/legajo/subir", methods=["POST"])
def importacion_legajo_subir(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    archivo = request.files.get("archivo")
    descripcion = request.form.get("descripcion", "").strip()

    if not archivo or archivo.filename == "":
        flash("Selecciona un archivo para subir.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))

    if not _extension_valida_legajo(archivo.filename):
        flash("Solo se permiten archivos PDF, Word, Excel o imagenes (JPG, PNG).", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))

    carpeta = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(imp.id))
    os.makedirs(carpeta, exist_ok=True)

    _, ext = os.path.splitext(archivo.filename)
    nombre_disco = f"{uuid.uuid4().hex}{ext.lower()}"
    archivo.save(os.path.join(carpeta, nombre_disco))

    doc = ImportacionDocumento(
        importacion_id=imp.id,
        descripcion=descripcion,
        nombre_original=archivo.filename,
        nombre_archivo=nombre_disco,
    )
    db.session.add(doc)
    db.session.commit()
    flash(f"Legajo '{archivo.filename}' subido correctamente.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


@app.route("/importaciones/<int:importacion_id>/legajo/<int:doc_id>/ver")
def importacion_legajo_ver(importacion_id, doc_id):
    doc = ImportacionDocumento.query.get_or_404(doc_id)
    if doc.importacion_id != importacion_id:
        abort(404)
    carpeta = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(doc.importacion_id))
    return send_from_directory(
        carpeta, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=False
    )


@app.route("/importaciones/<int:importacion_id>/legajo/<int:doc_id>/descargar")
def importacion_legajo_descargar(importacion_id, doc_id):
    doc = ImportacionDocumento.query.get_or_404(doc_id)
    if doc.importacion_id != importacion_id:
        abort(404)
    carpeta = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(doc.importacion_id))
    return send_from_directory(
        carpeta, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=True
    )


@app.route("/importaciones/<int:importacion_id>/legajo/<int:doc_id>/eliminar", methods=["POST"])
def importacion_legajo_eliminar(importacion_id, doc_id):
    doc = ImportacionDocumento.query.get_or_404(doc_id)
    if doc.importacion_id != importacion_id:
        abort(404)
    carpeta = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(doc.importacion_id))
    ruta = os.path.join(carpeta, doc.nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)
    db.session.delete(doc)
    db.session.commit()
    flash("Documento eliminado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


# --- Items de la factura: Flete, Seguro, Handling Fee, Otros (2026-09-01,
# punto 6 -- unificados en un solo listado con selector de concepto; antes
# Flete y Seguro eran campos fijos de Importacion) ---
# A diferencia de Gastos compartidos (prorrateados por CIF entre parciales
# marcados), un item de este listado se suma directo al CIF de cada linea,
# prorrateado por FOB -- ver costing.py. Cada item tiene su propia
# moneda/tipo de cambio (punto 3): por defecto se asume la moneda de la
# factura, pero se puede cargar en otra si el proveedor lo factura distinto.

@app.route("/importaciones/<int:importacion_id>/cargos/nuevo", methods=["POST"])
def cargos_nuevo(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    moneda = request.form.get("moneda", "").strip() or imp.moneda_factura
    cargo = CargoAdicionalImportacion(
        importacion_id=imp.id,
        concepto=request.form.get("concepto", "").strip() or "Otros",
        monto_moneda=float(request.form.get("monto_moneda") or 0),
        moneda=moneda,
        tipo_cambio=1 if moneda == "USD" else float(request.form.get("tipo_cambio") or 1),
        referencia=request.form.get("referencia", "").strip(),
    )
    db.session.add(cargo)
    db.session.commit()
    flash(f"'{cargo.concepto}' agregado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


@app.route("/cargos/<int:cargo_id>/editar", methods=["POST"])
def cargos_editar(cargo_id):
    cargo = CargoAdicionalImportacion.query.get_or_404(cargo_id)
    moneda = request.form.get("moneda", "").strip() or cargo.importacion.moneda_factura
    cargo.concepto = request.form.get("concepto", "").strip() or "Otros"
    cargo.monto_moneda = float(request.form.get("monto_moneda") or 0)
    cargo.moneda = moneda
    cargo.tipo_cambio = 1 if moneda == "USD" else float(request.form.get("tipo_cambio") or 1)
    cargo.referencia = request.form.get("referencia", "").strip()
    db.session.commit()
    flash("Actualizado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=cargo.importacion_id))


@app.route("/cargos/<int:cargo_id>/eliminar", methods=["POST"])
def cargos_eliminar(cargo_id):
    cargo = CargoAdicionalImportacion.query.get_or_404(cargo_id)
    importacion_id = cargo.importacion_id
    db.session.delete(cargo)
    db.session.commit()
    flash("Cargo eliminado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


@app.route("/importaciones/<int:importacion_id>/cargos/eliminar-multiple", methods=["POST"])
def cargos_eliminar_multiple(importacion_id):
    """Eliminacion masiva de items de factura (Flete/Seguro/Handling Fee/
    Otros) -- ronda I, 2026-09-01: el engranaje unico del panel activa
    checkboxes por fila y este endpoint borra todos los marcados de una
    vez. Se filtra por importacion_id ademas del id de cada cargo, para no
    poder borrar items de otra importacion pasando ids a mano."""
    imp = Importacion.query.get_or_404(importacion_id)
    ids = request.form.getlist("cargo_ids")
    borrados = CargoAdicionalImportacion.query.filter(
        CargoAdicionalImportacion.id.in_(ids),
        CargoAdicionalImportacion.importacion_id == imp.id,
    ).delete(synchronize_session=False)
    if borrados:
        db.session.commit()
        flash(f"{borrados} ítem(s) eliminados.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


# ---------------------------------------------------------------------------
# Modulo: Configuracion -- Empresas compradoras (2026-09-07, ronda K)
# ---------------------------------------------------------------------------

@app.route("/configuracion/empresas")
def empresas_list():
    empresas = Empresa.query.order_by(Empresa.nombre).all()
    return render_template(
        "configuracion/empresas.html", empresas=empresas, hoy_anio=date.today().year,
        default_asunto_correo=DEFAULT_ASUNTO_CORREO, default_cuerpo_correo=DEFAULT_CUERPO_CORREO,
    )


def _guardar_logo_empresa(file_storage, empresa_id):
    """Guarda el logo subido para una Empresa compradora y devuelve el
    nombre de archivo en disco (o None si no se subio nada valido).
    Reemplaza cualquier logo anterior de esa empresa (nombre fijo
    logo_<id>.<ext>) para no acumular archivos huerfanos."""
    if not file_storage or not file_storage.filename:
        return None
    extension = os.path.splitext(file_storage.filename)[1].lower()
    if extension not in EXTENSIONES_LOGO_PERMITIDAS:
        flash("Logo no guardado: solo se permiten imágenes PNG, JPG o SVG.", "warning")
        return None
    for antiguo in glob.glob(os.path.join(EMPRESAS_LOGOS_DIR, f"logo_{empresa_id}.*")):
        os.remove(antiguo)
    nombre_archivo = f"logo_{empresa_id}{extension}"
    file_storage.save(os.path.join(EMPRESAS_LOGOS_DIR, nombre_archivo))
    return nombre_archivo


@app.route("/configuracion/empresas/nueva", methods=["POST"])
def empresas_nueva():
    nombre = request.form.get("nombre", "").strip()
    if not nombre:
        flash("El nombre de la empresa es obligatorio.", "danger")
        return redirect(url_for("empresas_list"))
    if Empresa.query.filter(db.func.lower(Empresa.nombre) == nombre.lower()).first():
        flash(f"Ya existe una empresa compradora llamada '{nombre}'.", "warning")
        return redirect(url_for("empresas_list"))
    empresa = Empresa(
        nombre=nombre,
        direccion=request.form.get("direccion", "").strip(),
        rut=request.form.get("rut", "").strip(),
        telefono=request.form.get("telefono", "").strip(),
        email=request.form.get("email", "").strip(),
    )
    db.session.add(empresa)
    db.session.flush()
    logo = _guardar_logo_empresa(request.files.get("logo"), empresa.id)
    if logo:
        empresa.logo_nombre_archivo = logo
    db.session.commit()
    flash(f"Empresa '{empresa.nombre}' creada.", "success")
    return redirect(url_for("empresas_list"))


@app.route("/configuracion/empresas/<int:empresa_id>/editar", methods=["POST"])
def empresas_editar(empresa_id):
    empresa = Empresa.query.get_or_404(empresa_id)
    nombre = request.form.get("nombre", "").strip()
    if not nombre:
        flash("El nombre de la empresa es obligatorio.", "danger")
        return redirect(url_for("empresas_list"))
    empresa.nombre = nombre
    empresa.direccion = request.form.get("direccion", "").strip()
    empresa.rut = request.form.get("rut", "").strip()
    empresa.telefono = request.form.get("telefono", "").strip()
    empresa.email = request.form.get("email", "").strip()
    empresa.plantilla_asunto_correo = request.form.get("plantilla_asunto_correo", "").strip() or None
    empresa.plantilla_cuerpo_correo = request.form.get("plantilla_cuerpo_correo", "").strip() or None
    logo = _guardar_logo_empresa(request.files.get("logo"), empresa.id)
    if logo:
        empresa.logo_nombre_archivo = logo
    db.session.commit()
    flash(f"Empresa '{empresa.nombre}' actualizada.", "success")
    return redirect(url_for("empresas_list"))


@app.route("/configuracion/empresas/<int:empresa_id>/eliminar", methods=["POST"])
def empresas_eliminar(empresa_id):
    empresa = Empresa.query.get_or_404(empresa_id)
    total_ordenes = empresa.ordenes.count()
    if total_ordenes > 0:
        flash(
            f"No se puede eliminar '{empresa.nombre}': tiene {total_ordenes} orden(es) de "
            "compra asociadas. Reasígnalas a otra empresa (o déjalas 'Sin asignar') antes de "
            "eliminarla.", "danger",
        )
        return redirect(url_for("empresas_list"))
    if empresa.logo_nombre_archivo:
        ruta_logo = os.path.join(EMPRESAS_LOGOS_DIR, empresa.logo_nombre_archivo)
        if os.path.isfile(ruta_logo):
            os.remove(ruta_logo)
    nombre = empresa.nombre
    db.session.delete(empresa)
    db.session.commit()
    flash(f"Empresa '{nombre}' eliminada.", "success")
    return redirect(url_for("empresas_list"))


@app.route("/empresas/<int:empresa_id>/logo")
def empresas_logo(empresa_id):
    empresa = Empresa.query.get_or_404(empresa_id)
    if not empresa.logo_nombre_archivo:
        abort(404)
    return send_from_directory(EMPRESAS_LOGOS_DIR, empresa.logo_nombre_archivo)


# ---------------------------------------------------------------------------
# Modulo: Configuracion -- Tipo de cambio mensual (2026-08-27, punto 6)
# ---------------------------------------------------------------------------

@app.route("/configuracion/tipo-cambio")
def tipo_cambio_list():
    tipos = TipoCambioMensual.query.order_by(
        TipoCambioMensual.anio.desc(), TipoCambioMensual.mes.desc()
    ).all()
    return render_template("configuracion/tipo_cambio.html", tipos=tipos, hoy_anio=date.today().year)


@app.route("/configuracion/tipo-cambio/nuevo", methods=["POST"])
def tipo_cambio_nuevo():
    anio = parse_int(request.form.get("anio"), default=date.today().year)
    mes = parse_int(request.form.get("mes"), default=date.today().month)
    existente = _tipo_cambio_mes(anio, mes)
    if existente:
        flash(
            f"Ya existe un tipo de cambio cargado para {existente.nombre_mes} {existente.anio} -- "
            "edítalo directamente en vez de crear otro.", "warning",
        )
        return redirect(url_for("tipo_cambio_list"))
    tc = TipoCambioMensual(
        anio=anio, mes=mes,
        tc_aduanero=float(request.form.get("tc_aduanero") or 0),
        paridad_eur_usd=float(request.form.get("paridad_eur_usd") or 1),
        notas=request.form.get("notas", "").strip(),
    )
    db.session.add(tc)
    db.session.commit()
    flash(f"Tipo de cambio de {tc.nombre_mes} {tc.anio} cargado.", "success")
    return redirect(url_for("tipo_cambio_list"))


@app.route("/configuracion/tipo-cambio/<int:tc_id>/editar", methods=["POST"])
def tipo_cambio_editar(tc_id):
    tc = TipoCambioMensual.query.get_or_404(tc_id)
    tc.tc_aduanero = float(request.form.get("tc_aduanero") or 0)
    tc.paridad_eur_usd = float(request.form.get("paridad_eur_usd") or 1)
    tc.notas = request.form.get("notas", "").strip()
    db.session.commit()
    flash(f"Tipo de cambio de {tc.nombre_mes} {tc.anio} actualizado.", "success")
    return redirect(url_for("tipo_cambio_list"))


@app.route("/configuracion/tipo-cambio/<int:tc_id>/eliminar", methods=["POST"])
def tipo_cambio_eliminar(tc_id):
    tc = TipoCambioMensual.query.get_or_404(tc_id)
    db.session.delete(tc)
    db.session.commit()
    flash("Tipo de cambio eliminado.", "success")
    return redirect(url_for("tipo_cambio_list"))


# ---------------------------------------------------------------------------
# Modulo: Reinicio de datos (2026-08-27, punto 7 -- "estamos en modulo de
# prueba", el usuario pidio poder limpiar datos de prueba por categoria
# antes de empezar a usar el sistema en serio.
# ---------------------------------------------------------------------------
# Reglas de dependencia (para no dejar la base en un estado roto): borrar
# Proveedores obliga a borrar tambien Ordenes+Despachos y Costeos, porque
# Importacion.proveedor_id y OrdenCompra.proveedor_id NO admiten nulos --
# si el proveedor desaparece esas filas quedarian rotas. Borrar
# Ordenes+Despachos NO obliga a borrar Costeos: una Importacion generada
# desde un Despacho solo guarda una referencia opcional (despacho_id
# admite nulo), asi que sigue funcionando aunque el despacho de origen ya
# no exista -- solo deja de mostrar el link "generado desde este despacho".

def _resumen_datos_reset():
    """Conteos actuales, para mostrar en la pantalla de confirmacion antes
    de borrar nada."""
    return {
        "proveedores": Proveedor.query.count(),
        "productos": Producto.query.count(),
        "ordenes": OrdenCompra.query.count(),
        "despachos": Despacho.query.count(),
        "importaciones": Importacion.query.count(),
    }


def _ejecutar_reset(categorias):
    categorias = set(categorias)
    if "proveedores" in categorias:
        categorias.add("ordenes")
        categorias.add("costeos")

    partes = []

    if "costeos" in categorias:
        n = Importacion.query.count()
        for imp in Importacion.query.all():
            db.session.delete(imp)
        db.session.commit()
        partes.append(f"{n} importación(es)/costeo(s)")

    if "ordenes" in categorias:
        n_ordenes = OrdenCompra.query.count()
        for orden in OrdenCompra.query.all():
            _borrar_carpeta_documentos_orden(orden.id)
            db.session.delete(orden)
        db.session.commit()
        n_despachos = Despacho.query.count()
        for despacho in Despacho.query.all():
            db.session.delete(despacho)
        db.session.commit()
        partes.append(f"{n_ordenes} orden(es) de compra y {n_despachos} despacho(s)")

    if "proveedores" in categorias:
        n_prov = Proveedor.query.count()
        n_prod = Producto.query.count()
        for prov in Proveedor.query.all():
            db.session.delete(prov)
        db.session.commit()
        seed_from_excel(app)
        partes.append(f"{n_prov} proveedor(es) y {n_prod} producto(s) (catálogo recargado desde el Excel maestro)")

    return partes


@app.route("/configuracion/reset")
def admin_reset():
    return render_template("configuracion/reset.html", resumen=_resumen_datos_reset())


@app.route("/configuracion/reset/ejecutar", methods=["POST"])
def admin_reset_ejecutar():
    categorias = set(request.form.getlist("categorias"))
    categorias &= {"proveedores", "ordenes", "costeos"}
    if not categorias:
        flash("No marcaste ninguna categoría para borrar.", "warning")
        return redirect(url_for("admin_reset"))
    partes = _ejecutar_reset(categorias)
    flash("Borrado: " + "; ".join(partes) + ".", "success")
    return redirect(url_for("dashboard"))


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
