import csv
import glob
import io
import os
import re
import shutil
import uuid
from collections import defaultdict, Counter
from datetime import datetime, date
from functools import wraps
from urllib.parse import quote

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from flask import (
    Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory,
    send_file, abort
)
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user, current_user,
)
from markupsafe import escape

from models import (
    db, Empresa, Proveedor, Producto, ProductoVariante, OrdenCompra, OrdenCompraLinea, OrdenDocumento,
    ESTADOS_OC, ETAPAS_LINEA, TIPOS_DOCUMENTO_ORDEN,
    Importacion, Parcial, ParcialLinea, ParcialLineaLote, GastoImportacion, GastoDocumento,
    ImportacionDocumento,
    REGIMENES_PARCIAL, VIAS_EMBARQUE, CONDICIONES_COMPRA, CONCEPTOS_GASTO,
    CONCEPTOS_ITEM_FACTURA, TIPOS_DOCUMENTO_GASTO,
    Despacho, ESTADOS_DESPACHO,
    CargoAdicionalImportacion, TipoCambioMensual,
    Usuario, Rol, PERMISOS_DISPONIBLES,
    HomologacionStock, StockExistencia, PedidoComprometido,
    CodigoErgopyme, CompraHistorica,
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
# Ronda V (2026-09-12): archivo que el usuario preparo a mano para el cruce
# UNA SOLA VEZ entre el codigo interno del sistema de Inventarios ("Ergopyme")
# y nuestro catalogo -- hoja "INVENTARIO ACTUAL" (columnas I/J/K con el
# cruce) y hoja "LENTES PHYSIOL" (mismo formato que los lentes MEDICONTUR:
# codigo padre / codigo variante / descripcion). Ver
# seed_variantes_lentes_physiol() y seed_homologacion_y_stock_inicial() abajo.
INVENTARIOS_CODIGOS_INTERNOS_EXCEL = os.path.join(BASE_DIR, "Inventarios y codigos interno sistema Inventarios.xlsx")
EXTENSIONES_LOGO_PERMITIDAS = {".png", ".jpg", ".jpeg", ".svg"}
# PDF de la Orden de Compra generado en disco para poder adjuntarlo a un
# correo (ronda L, punto 1, 2026-09-09) -- se sobrescribe cada vez que se
# genera, no se acumulan versiones viejas.
PDF_OC_DIR = os.path.join(BASE_DIR, "data", "pdf_oc")
# Reportes cargados para consulta posterior (ronda X, 2026-09-13, punto 3C):
# el archivo original "Notas de pedido" que el usuario sube en Consulta de
# Stock se guarda aca, uno por empresa (se sobrescribe solo el de esa
# empresa cada vez que se sube uno nuevo) -- asi queda disponible para
# volver a consultarlo despues, sin depender de que el usuario guarde su
# propia copia.
REPORTES_DIR = os.path.join(BASE_DIR, "data", "reportes")
# Ronda AA (2026-09-13): archivos para el reporte "Compras Proveedor" (menu
# Reportes) -- ver seed_codigos_ergopyme() y seed_compras_historicas() abajo.
# 1) mapeo maestro Proveedor + Codigo Proveedor + Descripcion para cada
#    "codigo interno" del sistema de Inventarios (Ergopyme), mas completo
#    que HomologacionStock (esa solo cubre codigos vistos en Stock).
CODIGOS_ERGOPYME_EXCEL = os.path.join(BASE_DIR, "codigos_ergopyme_homologacion.xlsx")
# 2) historico de compras/importaciones a proveedores anterior a este
#    sistema (una fila por linea de producto de cada factura, 2023-2026),
#    homologado contra (1) al cargarse.
HISTORICO_COMPRAS_EXCEL = os.path.join(BASE_DIR, "historico_compras_proveedores.xlsx")
# 3) Ronda AC (2026-09-14): planilla de referencia que mantiene el usuario
#    con la moneda HABITUAL de cada proveedor extranjero (USD/EURO) -- se
#    usa solo como DESEMPATE al elegir en qué moneda expresar un resumen
#    (ver _moneda_dominante), nunca como fuente principal: la moneda de
#    cada línea/factura se sigue determinando por su propia paridad.
MONEDA_PROVEEDOR_EXCEL = os.path.join(BASE_DIR, "moneda_proveedor.xlsx")
# 4) Ronda AE (2026-09-14): listado de código padre/variante de lentes BVI
#    PHYSIOL que mantiene el usuario aparte (superset corregido de la hoja
#    "LENTES PHYSIOL" que trae INVENTARIOS_CODIGOS_INTERNOS_EXCEL -- agrega
#    códigos que antes no estaban referenciados a ningún padre). Si existe,
#    tiene prioridad sobre la hoja embebida -- ver seed_variantes_lentes_
#    physiol() y reparar_variantes_physiol_ronda_ae().
AJUSTE_PADRE_PHYSIOL_EXCEL = os.path.join(BASE_DIR, "Ajuste de cuentas padre Physiol.xlsx")

# Ronda AE (2026-09-14): algunos proveedores extranjeros aparecen en el
# histórico de compras (columna "Proveedor" original del archivo) con un
# nombre corto que no trae el prefijo "BVI" que sí usa el catálogo de este
# sistema (Proveedor.nombre) -- ej. "OPTIKON" en vez de "BVI OPTIKON". Se
# homologa acá, en un solo lugar, para que el reporte de Compras Proveedor
# agrupe cada proveedor bajo su nombre real de catálogo (ver
# _proveedor_canonico_historico() y su uso en _fila_historica_dict() /
# _proveedores_reporte_compras()).
ALIAS_PROVEEDOR_HISTORICO = {
    "OPTIKON": "BVI OPTIKON",
    "PHYSIOL": "BVI PHYSIOL",
    "BEAVER": "BVI BEAVER",
}

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

# Sistema de usuarios/login (ronda R, 2026-09-12) -- ver Usuario/Rol en
# models.py. login_view redirige aca cuando alguien sin sesion intenta
# entrar a una pagina protegida (ver _requerir_login mas abajo).
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message = "Inicia sesión para continuar."
login_manager.login_message_category = "warning"


@login_manager.user_loader
def _cargar_usuario(user_id):
    return Usuario.query.get(int(user_id))


# Rutas que NO requieren sesion iniciada -- todo lo demas la exige (ronda R,
# 2026-09-12): antes de esta ronda la app no tenia ningun control de acceso.
_ENDPOINTS_PUBLICOS = {"login", "static"}


@app.before_request
def _requerir_login():
    if request.endpoint is None or request.endpoint in _ENDPOINTS_PUBLICOS:
        return None
    if not current_user.is_authenticated:
        return redirect(url_for("login", next=request.path))
    return None


@app.context_processor
def _inyectar_ordenes_por_aprobar():
    """Ronda T (2026-09-12): cantidad para el badge de la pestaña 'Por
    Aprobar' en la barra de navegacion -- las ordenes que ese usuario puede
    ver ahi mismo (ver ordenes_por_aprobar): todas las 'Por Aprobar' si
    tiene permiso de Aprobacion, o solo las suyas si no."""
    if not current_user.is_authenticated:
        return {}
    if not (current_user.tiene_permiso("crear_orden") or current_user.tiene_permiso("aprobar_orden")):
        return {}
    q = OrdenCompra.query.filter(OrdenCompra.estado_aprobacion == "Por Aprobar")
    if not current_user.tiene_permiso("aprobar_orden"):
        q = q.filter(OrdenCompra.creado_por_usuario_id == current_user.id)
    return {"ordenes_por_aprobar_count": q.count()}


def requiere_permiso(*permisos):
    """Decorador para una ruta: exige que el usuario logueado sea
    Administrador o tenga AL MENOS UNO de los `permisos` indicados (los
    codigos de PERMISOS_DISPONIBLES en models.py). El login en si ya lo
    exige _requerir_login() arriba para TODA la app -- este decorador solo
    agrega el chequeo de permiso especifico de cada pantalla."""
    def decorador(vista):
        @wraps(vista)
        def envoltura(*args, **kwargs):
            if not any(current_user.tiene_permiso(p) for p in permisos):
                flash("No tienes permiso para acceder a esta sección.", "danger")
                return redirect(url_for("dashboard"))
            return vista(*args, **kwargs)
        return envoltura
    return decorador


def requiere_admin(vista):
    """Como requiere_permiso, pero solo para el Administrador (gestion de
    usuarios/roles y pantallas de Configuración)."""
    @wraps(vista)
    def envoltura(*args, **kwargs):
        if not (current_user.is_authenticated and current_user.rol and current_user.rol.es_administrador):
            flash("Solo un Administrador puede acceder a esta sección.", "danger")
            return redirect(url_for("dashboard"))
        return vista(*args, **kwargs)
    return envoltura


def _puede_gestionar_orden_simple(orden):
    """Ronda S (2026-09-12): True si el usuario actual puede ver/gestionar
    esta orden a traves de las acciones COMPARTIDAS entre el flujo completo
    y el flujo 'Creación de Orden Simple' (documentos, notas) -- o bien
    tiene acceso completo (crear_orden/aprobar_orden, que ven CUALQUIER
    orden) o bien es SU PROPIA orden, creada con el perfil orden_simple."""
    if current_user.tiene_permiso("crear_orden") or current_user.tiene_permiso("aprobar_orden"):
        return True
    return orden.creado_por_usuario_id == current_user.id


def _orden_aprobada(orden):
    """Ronda T (2026-09-12): True si la orden ya es una Orden de Compra
    definitiva (estado_aprobacion 'Aprobada', o NULL en ordenes de antes de
    esta ronda). Las acciones de aprobacion por linea (Confirmar, Cambiar
    estado) solo tienen sentido una vez aprobada la orden completa."""
    return orden.estado_aprobacion in (None, "Aprobada")


def _destino_detalle_orden(orden):
    """A donde volver despues de una accion compartida (documentos, notas):
    quien tiene acceso completo vuelve al detalle normal, quien solo tiene
    orden_simple vuelve a su propia vista simplificada."""
    if current_user.tiene_permiso("crear_orden") or current_user.tiene_permiso("aprobar_orden"):
        return url_for("ordenes_detalle", orden_id=orden.id)
    return url_for("ordenes_simple_detalle", orden_id=orden.id)


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        usuario = Usuario.query.filter(db.func.lower(Usuario.email) == email).first()
        if usuario and usuario.check_password(password):
            if not usuario.activo:
                flash("Este usuario está desactivado. Contacta a un Administrador.", "danger")
            else:
                login_user(usuario)
                usuario.ultimo_acceso = datetime.utcnow()
                db.session.commit()
                destino = request.args.get("next") or url_for("dashboard")
                return redirect(destino)
        else:
            flash("Correo o contraseña incorrectos.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    logout_user()
    flash("Sesión cerrada.", "success")
    return redirect(url_for("login"))


@app.route("/mi-cuenta", methods=["GET", "POST"])
def mi_cuenta():
    if request.method == "POST":
        actual = request.form.get("password_actual", "")
        nueva = request.form.get("password_nueva", "")
        confirmar = request.form.get("password_confirmar", "")
        if not current_user.check_password(actual):
            flash("La contraseña actual no es correcta.", "danger")
        elif len(nueva) < 6:
            flash("La nueva contraseña debe tener al menos 6 caracteres.", "warning")
        elif nueva != confirmar:
            flash("La confirmación no coincide con la nueva contraseña.", "warning")
        else:
            current_user.set_password(nueva)
            db.session.commit()
            flash("Contraseña actualizada correctamente.", "success")
            return redirect(url_for("dashboard"))
    return render_template("mi_cuenta.html")


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
            ("precio_catalogo_oculto", "BOOLEAN DEFAULT 0"),
            ("producto_creado_por_esta_orden", "BOOLEAN DEFAULT 0"),
        ],
        "ordenes_compra": [
            ("despacho_id", "INTEGER"),
            ("empresa_id", "INTEGER"),
            ("creado_por_usuario_id", "INTEGER"),
            ("estado_aprobacion", "VARCHAR(20) DEFAULT 'Aprobada'"),
        ],
        "importaciones": [
            ("despacho_id", "INTEGER"),
            ("condicion_compra", "VARCHAR(10) DEFAULT 'EXW'"),
            ("flete_total_moneda", "FLOAT DEFAULT 0"),
            ("seguro_total_moneda", "FLOAT DEFAULT 0"),
            ("empresa_id", "INTEGER"),
            ("numero_correlativo_inventario", "INTEGER"),
        ],
        "proveedores": [
            ("codigo_sistema_inventario", "VARCHAR(50)"),
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
        "productos": [
            ("codigo_interno_inventario", "VARCHAR(40)"),
        ],
        "producto_variantes": [
            ("codigo_interno_inventario", "VARCHAR(40)"),
        ],
        "roles": [
            ("permiso_consultar_stock", "BOOLEAN DEFAULT 0"),
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

    # Reparacion de datos (ronda W, 2026-09-13, punto 5): todo perfil que ya
    # tenia el permiso "Inventarios" marcado, pero no "Consultar Stock"
    # (creado antes de que existiera este ultimo, ronda V), pasa a tener los
    # 2 -- a pedido explicito del usuario, para no dejar a nadie que ya
    # gestionaba Inventarios sin poder ver la pantalla de Consulta de Stock
    # nueva. De aca en mas _guardar_permisos_rol ya los mantiene alineados
    # solo (ver ese comentario), asi que esto no vuelve a hacer falta salvo
    # para perfiles de antes de esta ronda.
    if "roles" in tablas:
        columnas_roles = {c["name"] for c in inspector.get_columns("roles")}
        if "permiso_inventarios" in columnas_roles and "permiso_consultar_stock" in columnas_roles:
            with db.engine.begin() as conn:
                conn.execute(db.text(
                    "UPDATE roles SET permiso_consultar_stock = "
                    + ("1" if es_sqlite else "TRUE")
                    + " WHERE permiso_inventarios = "
                    + ("1" if es_sqlite else "TRUE")
                    + " AND (permiso_consultar_stock IS NULL OR permiso_consultar_stock = "
                    + ("0" if es_sqlite else "FALSE") + ")"
                ))

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


def seed_variantes_lentes_physiol():
    """Ronda V (2026-09-12): mismo mecanismo que seed_variantes_lentes_
    medicontur de arriba, pero para los lentes BVI PHYSIOL -- lee la hoja
    'LENTES PHYSIOL' de INVENTARIOS_CODIGOS_INTERNOS_EXCEL (mismas 3
    columnas: codigo padre, Codigo Producto/variante, DESCRIPCION). El
    archivo real trae varias filas repetidas para la misma variante (una
    por lote visto en el reporte de stock de origen), asi que se deduplica
    por (codigo padre, codigo variante) antes de crear cada ProductoVariante.
    Gateado de forma INDEPENDIENTE del gate de seed_variantes_lentes_
    medicontur (que en produccion ya tiene sus variantes cargadas de una
    ronda anterior) -- revisa si ya existen variantes especificamente para
    productos de BVI PHYSIOL, no el conteo global de la tabla."""
    physiol = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == "BVI PHYSIOL").first()
    if not physiol:
        print("[seed] No existe el proveedor BVI PHYSIOL todavía, se omite la carga de variantes de lentes Physiol.")
        return
    ya_existen = ProductoVariante.query.join(Producto).filter(Producto.proveedor_id == physiol.id).count() > 0
    if ya_existen:
        return
    if not os.path.isfile(INVENTARIOS_CODIGOS_INTERNOS_EXCEL):
        print(f"[seed] No se encontró {INVENTARIOS_CODIGOS_INTERNOS_EXCEL}, se omite la carga de variantes Physiol.")
        return

    import openpyxl
    wb = openpyxl.load_workbook(INVENTARIOS_CODIGOS_INTERNOS_EXCEL, data_only=True)
    if "LENTES PHYSIOL" not in wb.sheetnames:
        print("[seed] La hoja 'LENTES PHYSIOL' no existe en el archivo de homologación, se omite.")
        return
    ws = wb["LENTES PHYSIOL"]

    productos_padre = {
        p.codigo.strip().upper(): p
        for p in Producto.query.filter_by(proveedor_id=physiol.id).all()
    }

    fila_inicio = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        primera = str(row[0]).strip().lower() if row and row[0] else ""
        if primera == "codigo padre":
            fila_inicio = i + 1
            break

    creadas = 0
    padres_creados = set()
    vistos = set()
    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        codigo_padre, codigo_variante, descripcion = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
        codigo_padre = str(codigo_padre).strip() if codigo_padre else ""
        codigo_variante = str(codigo_variante).strip() if codigo_variante else ""
        descripcion = str(descripcion).strip() if descripcion else ""
        if not codigo_padre or not codigo_variante:
            continue
        clave = (codigo_padre.upper(), codigo_variante.upper())
        if clave in vistos:
            continue
        vistos.add(clave)
        producto = productos_padre.get(codigo_padre.upper())
        if not producto:
            # Ronda W (2026-09-13): antes esto se reportaba como "código
            # padre no encontrado" y se descartaba la fila entera -- el
            # usuario confirmó (con "Serenity Toric PODS49P" y "Podeye
            # Toric") que esos códigos padre SÍ deberían existir, solo que
            # el catálogo (importado del listado de precios) nunca tuvo esa
            # familia como Producto propio. Se crea el Producto padre nuevo
            # bajo BVI PHYSIOL -- si el nombre es la version " Toric" de una
            # familia que ya existe (ej. "Podeye Toric" -> "Podeye"), se le
            # copia el precio/moneda/empaque como punto de partida (a
            # verificar por el usuario); si no hay una familia "hermana"
            # reconocible, nace en 0 igual que cualquier producto nuevo de
            # la homologación.
            candidato_hermano = re.sub(r"\btoric\b", "", codigo_padre, flags=re.IGNORECASE)
            candidato_hermano = re.sub(r"\s+", " ", candidato_hermano).strip()
            hermano = productos_padre.get(candidato_hermano.upper()) if candidato_hermano else None
            producto = Producto(
                proveedor_id=physiol.id,
                codigo=codigo_padre,
                descripcion=codigo_padre,
                empaque=(hermano.empaque if hermano else 1),
                moneda=(hermano.moneda if hermano else (physiol.moneda_default or "USD")),
                precio_caja=(hermano.precio_caja if hermano else 0),
                precio_unitario=(hermano.precio_unitario if hermano else 0),
                activo=True,
            )
            db.session.add(producto)
            db.session.flush()
            productos_padre[codigo_padre.upper()] = producto
            padres_creados.add((codigo_padre, hermano.codigo if hermano else None))
        db.session.add(ProductoVariante(
            producto_id=producto.id,
            codigo=codigo_variante,
            descripcion=descripcion or codigo_variante,
        ))
        creadas += 1
    db.session.commit()
    print(
        f"[seed] Importadas {creadas} variantes de lentes PHYSIOL "
        f"({len(padres_creados)} código(s) padre nuevo(s) creados en el catálogo -- "
        f"verificar precio: {sorted(padres_creados)})."
    )


def _normalizar_codigo_interno(valor):
    """Ronda V (2026-09-12): limpia un codigo interno del sistema de
    Inventarios tal como viene en el reporte original -- trae una comilla
    simple (') adelante (para que el otro sistema no le borre los ceros a
    la izquierda) y espacios de relleno al final."""
    if valor is None:
        return ""
    s = str(valor).strip()
    if s.startswith("'"):
        s = s[1:]
    s = s.strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def _empresa_por_nombre_reporte(nombre_reporte, cache):
    """Ronda V (2026-09-12): el reporte de Stock trae el nombre de la
    empresa como 'ACCUVISION SPA' / 'ACCUMEDICAL SPA' -- nuestras Empresa
    ya cargadas se llaman 'Accuvision' / 'Accumedical' (sin el SPA). Se
    normaliza sacando el sufijo ' SPA' y comparando sin distinguir
    mayusculas/minusculas."""
    if not nombre_reporte:
        return None
    clave = re.sub(r"\s+SPA$", "", nombre_reporte.strip(), flags=re.IGNORECASE).strip().upper()
    if clave in cache:
        return cache[clave]
    empresa = Empresa.query.filter(db.func.upper(Empresa.nombre) == clave).first()
    cache[clave] = empresa
    return empresa


def _leer_filas_reporte_stock(ws):
    """Ronda V (2026-09-12): lee una hoja con el formato ORIGINAL del
    reporte de Stock del sistema de Inventarios ('Ergopyme') -- columnas
    A-H: Cód.Bod (se ignora), Cód.Producto, Denominacion, Uni, Cód.Lote,
    Vencimiento, F.Compra (se ignora), Stock físico. El archivo trae el
    stock de VARIAS empresas seguidas, cada bloque separado por una fila
    con solo el nombre de la empresa en la columna A (ej. 'ACCUVISION
    SPA'), seguida del titulo, la fecha de emision, una fila en blanco, el
    encabezado de la tabla y una fila de guiones -- todo eso se reconoce y
    se salta solo, sin asumir un numero de fila fijo (la cantidad de filas
    de cada empresa cambia en cada carga segun compras/ventas)."""
    filas = []
    empresa_actual = None
    for row in ws.iter_rows(min_row=1, values_only=True):
        a = row[0] if len(row) > 0 else None
        b = row[1] if len(row) > 1 else None
        if a is None and b is None:
            continue
        if b is None:
            texto = str(a).strip() if a is not None else ""
            if not texto or texto.lower().startswith("fecha") or texto.lower() == "cód.bod" or set(texto) <= {"-"}:
                continue
            empresa_actual = texto
            continue
        if not (isinstance(b, str) and b.strip().startswith("'")):
            continue
        codigo_interno = _normalizar_codigo_interno(b)
        if not codigo_interno:
            continue
        descripcion = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        codigo_lote = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
        fecha_venc = row[5] if len(row) > 5 else None
        if isinstance(fecha_venc, datetime):
            fecha_venc = fecha_venc.date()
        stock = row[7] if len(row) > 7 else 0
        try:
            stock = int(stock) if stock is not None else 0
        except (TypeError, ValueError):
            stock = 0
        filas.append({
            "empresa_texto": empresa_actual,
            "codigo_interno": codigo_interno,
            "descripcion": descripcion,
            "codigo_lote": codigo_lote,
            "fecha_vencimiento": fecha_venc,
            "stock_fisico": stock,
        })
    return filas


def _clasificar_y_cargar_stock(filas):
    """Ronda V (2026-09-12): convierte filas ya leidas (_leer_filas_reporte_
    stock) en registros StockExistencia, usando la tabla HomologacionStock
    para saber a que Producto/Variante ligar cada codigo interno (o si hay
    que excluirlo / dejarlo sin marca). Un codigo interno que nunca se vio
    antes queda 'pendiente' -- se carga igual al Stock (sin producto/
    variante ligado, para no perder el dato) y aparece en /stock/
    homologacion para resolverlo a mano. NO crea Productos nuevos aca (eso
    solo ocurre en el cruce inicial, ver seed_homologacion_y_stock_inicial)
    -- una carga de stock del dia a dia nunca inventa productos solos."""
    resumen = {
        "cargados": 0, "vinculados": 0, "sin_marca": 0, "excluidos": 0,
        "pendientes_nuevos": 0, "empresas_no_encontradas": set(),
    }
    cache_empresas = {}
    homologaciones = {h.codigo_interno: h for h in HomologacionStock.query.all()}
    for fila in filas:
        codigo = fila["codigo_interno"]
        homolog = homologaciones.get(codigo)
        if homolog is None:
            homolog = HomologacionStock(
                codigo_interno=codigo,
                estado="pendiente",
                descripcion_referencia=fila["descripcion"],
            )
            db.session.add(homolog)
            homologaciones[codigo] = homolog
            resumen["pendientes_nuevos"] += 1
        elif fila["descripcion"] and not homolog.producto_id and not homolog.variante_id:
            homolog.descripcion_referencia = fila["descripcion"]

        if homolog.estado == "excluido":
            resumen["excluidos"] += 1
            continue

        empresa = _empresa_por_nombre_reporte(fila["empresa_texto"], cache_empresas)
        if not empresa:
            resumen["empresas_no_encontradas"].add(fila["empresa_texto"] or "(sin identificar)")
            continue

        db.session.add(StockExistencia(
            empresa_id=empresa.id,
            producto_id=homolog.producto_id,
            variante_id=homolog.variante_id,
            codigo_interno=codigo,
            descripcion=fila["descripcion"],
            codigo_lote=fila["codigo_lote"],
            fecha_vencimiento=fila["fecha_vencimiento"],
            stock_fisico=fila["stock_fisico"],
        ))
        resumen["cargados"] += 1
        if homolog.estado == "vinculado":
            resumen["vinculados"] += 1
        elif homolog.estado == "sin_marca":
            resumen["sin_marca"] += 1
    return resumen


def seed_homologacion_y_stock_inicial():
    """Ronda V (2026-09-12, punto 1): cruce UNA SOLA VEZ entre el codigo
    interno del sistema de Inventarios y nuestro catalogo, a partir del
    archivo que el usuario preparo a mano ('Inventarios y codigos interno
    sistema Inventarios.xlsx', hoja 'INVENTARIO ACTUAL', columnas I/J/K).
    Clasifica cada codigo interno distinto segun la columna K (PROVEEDOR):
    - 'NO INCLUIR' (o vacio) -> queda 'excluido': nunca se carga al Stock.
    - 'SIN MARCA' -> queda 'sin_marca': se carga al Stock para consulta,
      pero no se crea ni se liga a ningun Producto del catalogo.
    - cualquier otro proveedor -> se busca ese Producto (o esa
      ProductoVariante, para familias con dioptrias como MEDICONTUR/BVI
      PHYSIOL) por el codigo de la columna J; si no existe todavia en
      nuestro catalogo, se CREA (a pedido explicito del usuario).
    Gateado: si ya existe alguna fila en HomologacionStock, no hace nada
    (para que el usuario pueda ajustar homologaciones a mano despues sin
    que se pisen solas en cada arranque). Debe correr DESPUES de
    seed_variantes_lentes_medicontur() y seed_variantes_lentes_physiol(),
    para que las variantes de esas 2 familias ya existan al momento del
    cruce."""
    if HomologacionStock.query.count() > 0:
        return
    if not os.path.isfile(INVENTARIOS_CODIGOS_INTERNOS_EXCEL):
        print(f"[seed] No se encontró {INVENTARIOS_CODIGOS_INTERNOS_EXCEL}, se omite la homologación inicial de Stock.")
        return

    import openpyxl
    wb = openpyxl.load_workbook(INVENTARIOS_CODIGOS_INTERNOS_EXCEL, data_only=True)
    if "INVENTARIO ACTUAL" not in wb.sheetnames:
        print("[seed] La hoja 'INVENTARIO ACTUAL' no existe en el archivo de homologación, se omite.")
        return
    ws = wb["INVENTARIO ACTUAL"]

    clasificacion = {}
    filas_crudas = []
    empresa_actual = None
    for row in ws.iter_rows(min_row=1, values_only=True):
        a = row[0] if len(row) > 0 else None
        b = row[1] if len(row) > 1 else None
        if a is None and b is None:
            continue
        if b is None:
            texto = str(a).strip() if a is not None else ""
            if not texto or texto.lower().startswith("fecha") or texto.lower() == "cód.bod" or set(texto) <= {"-"}:
                continue
            empresa_actual = texto
            continue
        if not (isinstance(b, str) and b.strip().startswith("'")):
            continue
        codigo_interno = _normalizar_codigo_interno(b)
        if not codigo_interno:
            continue
        descripcion = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        codigo_lote = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
        fecha_venc = row[5] if len(row) > 5 else None
        if isinstance(fecha_venc, datetime):
            fecha_venc = fecha_venc.date()
        stock = row[7] if len(row) > 7 else 0
        try:
            stock = int(stock) if stock is not None else 0
        except (TypeError, ValueError):
            stock = 0
        filas_crudas.append({
            "empresa_texto": empresa_actual, "codigo_interno": codigo_interno,
            "descripcion": descripcion, "codigo_lote": codigo_lote,
            "fecha_vencimiento": fecha_venc, "stock_fisico": stock,
        })
        if codigo_interno not in clasificacion:
            j = row[9] if len(row) > 9 else None
            k = row[10] if len(row) > 10 else None
            clasificacion[codigo_interno] = {
                "j": str(j).strip() if j is not None else "",
                "k": str(k).strip() if k is not None else "",
                "descripcion": descripcion,
            }

    proveedores_cache = {}
    productos_creados = 0
    proveedores_creados = set()
    variantes_ligadas = 0
    variantes_ligadas_por_descripcion = 0
    productos_ligados = 0
    sin_marca = 0
    excluidos = 0
    pendientes_manual = 0

    for codigo_interno, info in clasificacion.items():
        k_upper = info["k"].strip().upper()
        j_valor = info["j"].strip()
        if not k_upper or k_upper == "NO INCLUIR":
            db.session.add(HomologacionStock(
                codigo_interno=codigo_interno, estado="excluido",
                descripcion_referencia=info["descripcion"],
            ))
            excluidos += 1
            continue
        if k_upper == "SIN MARCA":
            db.session.add(HomologacionStock(
                codigo_interno=codigo_interno, estado="sin_marca",
                descripcion_referencia=info["descripcion"],
            ))
            sin_marca += 1
            continue

        if k_upper not in proveedores_cache:
            proveedores_cache[k_upper] = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == k_upper).first()
        proveedor = proveedores_cache[k_upper]
        if not proveedor and j_valor:
            # Ronda W (2026-09-13, punto 3): antes un proveedor no
            # reconocido (nombre en la columna K que no calzaba con ningun
            # Proveedor ya cargado) dejaba el codigo "pendiente" para
            # resolverlo a mano -- a pedido explicito del usuario, ahora se
            # CREA el proveedor que falte (algunos venian con un "*" al
            # final en el archivo, se saca por no ser parte del nombre
            # real) para poder seguir e ingresar sus productos igual que
            # con cualquier otro proveedor ya conocido.
            nombre_nuevo = info["k"].strip().rstrip("*").strip() or info["k"].strip()
            proveedor = Proveedor(nombre=nombre_nuevo, tipo="Extranjero", moneda_default="USD", activo=True)
            db.session.add(proveedor)
            db.session.flush()
            proveedores_cache[k_upper] = proveedor
            proveedores_creados.add(nombre_nuevo)
        if not proveedor or not j_valor:
            db.session.add(HomologacionStock(
                codigo_interno=codigo_interno, estado="pendiente",
                descripcion_referencia=info["descripcion"],
            ))
            pendientes_manual += 1
            continue

        j_upper = j_valor.upper()
        producto_match = Producto.query.filter(
            Producto.proveedor_id == proveedor.id, db.func.upper(Producto.codigo) == j_upper
        ).first()
        if producto_match:
            producto_match.codigo_interno_inventario = codigo_interno
            db.session.add(HomologacionStock(
                codigo_interno=codigo_interno, estado="vinculado",
                producto_id=producto_match.id, descripcion_referencia=info["descripcion"],
            ))
            productos_ligados += 1
            continue

        variante_match = ProductoVariante.query.join(Producto).filter(
            Producto.proveedor_id == proveedor.id, db.func.upper(ProductoVariante.codigo) == j_upper
        ).first()
        if variante_match:
            variante_match.codigo_interno_inventario = codigo_interno
            db.session.add(HomologacionStock(
                codigo_interno=codigo_interno, estado="vinculado",
                producto_id=variante_match.producto_id, variante_id=variante_match.id,
                descripcion_referencia=info["descripcion"],
            ))
            variantes_ligadas += 1
            continue

        # Ronda W (2026-09-13, punto 1): a veces el código de la columna J
        # (escrito a mano por el usuario al armar el cruce) no coincide
        # exactamente con ningún código de variante ya cargado, pero la
        # Denominación (columna C, la descripción real del producto en el
        # reporte de origen) SÍ coincide -- ej. "MICROPURE 07.00D" en J vs.
        # la variante ya cargada "MICROPURE 07.0D", cuya propia
        # descripción es justamente "MICROPURE 07.0D". Antes de crear un
        # Producto nuevo (y terminar con un duplicado sombra a precio 0),
        # se intenta este último cruce por descripción dentro del mismo
        # proveedor.
        descripcion_upper = (info["descripcion"] or "").strip().upper()
        variante_por_desc = None
        if descripcion_upper:
            variante_por_desc = ProductoVariante.query.join(Producto).filter(
                Producto.proveedor_id == proveedor.id,
                db.or_(
                    db.func.upper(ProductoVariante.codigo) == descripcion_upper,
                    db.func.upper(ProductoVariante.descripcion) == descripcion_upper,
                ),
            ).first()
        if variante_por_desc:
            variante_por_desc.codigo_interno_inventario = codigo_interno
            db.session.add(HomologacionStock(
                codigo_interno=codigo_interno, estado="vinculado",
                producto_id=variante_por_desc.producto_id, variante_id=variante_por_desc.id,
                descripcion_referencia=info["descripcion"],
            ))
            variantes_ligadas_por_descripcion += 1
            continue

        nuevo = Producto(
            proveedor_id=proveedor.id,
            codigo=j_valor,
            descripcion=info["descripcion"] or j_valor,
            empaque=1,
            moneda=proveedor.moneda_default or "USD",
            precio_caja=0,
            precio_unitario=0,
            activo=True,
            codigo_interno_inventario=codigo_interno,
        )
        db.session.add(nuevo)
        db.session.flush()
        db.session.add(HomologacionStock(
            codigo_interno=codigo_interno, estado="vinculado",
            producto_id=nuevo.id, descripcion_referencia=info["descripcion"],
        ))
        productos_creados += 1

    db.session.commit()
    print(
        f"[seed] Homologación inicial de Stock: {productos_ligados} código(s) ligados a producto ya existente, "
        f"{variantes_ligadas} ligados a variante ya existente ({variantes_ligadas_por_descripcion} de ellos por "
        f"descripción, no por código), {productos_creados} producto(s) NUEVOS creados en el catálogo, "
        f"{len(proveedores_creados)} proveedor(es) NUEVO(s) creados ({sorted(proveedores_creados)}), "
        f"{sin_marca} sin marca, {excluidos} excluidos, {pendientes_manual} pendiente(s) por resolver a mano."
    )

    resumen = _clasificar_y_cargar_stock(filas_crudas)
    db.session.commit()
    print(
        f"[seed] Stock inicial cargado: {resumen['cargados']} fila(s) "
        f"({resumen['vinculados']} vinculadas, {resumen['sin_marca']} sin marca), "
        f"{resumen['pendientes_nuevos']} código(s) nuevos quedaron pendientes, "
        f"empresas no encontradas: {sorted(resumen['empresas_no_encontradas'])}."
    )


def reparar_datos_ronda_w():
    """Ronda W (2026-09-13): repara datos que YA quedaron mal creados en una
    base real donde seed_homologacion_y_stock_inicial() y seed_variantes_
    lentes_physiol() ya habian corrido (con la logica vieja, antes de los
    arreglos de esta ronda) -- en ese caso los gates de esas dos funciones
    (que solo corren "si la tabla esta vacia") quedan cerrados para siempre
    y las correcciones de Ronda W nunca llegan a aplicarse solas. Esta
    funcion SI puede correr en cada arranque sin problema: no depende de
    ninguna bandera, revisa el estado real de los datos cada vez y no hace
    nada si ya no encuentra nada por reparar (idempotente).

    Hace 3 cosas, siempre migrando las referencias ya guardadas
    (HomologacionStock y StockExistencia) al lugar correcto antes de borrar
    cualquier producto "placeholder" que haya quedado mal creado, igual que
    ya hace a mano /stock/homologacion/<id>/resolver:

    A) Familias de lentes PHYSIOL (Serenity Toric PODS49P, Podeye Toric):
       las variantes que quedaron creadas como Producto suelto (uno por
       dioptria, a precio 0) se migran a ser ProductoVariante bajo su
       cuenta padre real, creando esa cuenta padre si todavia no existe (y
       reconociendo -- ver ALIAS_PADRES_PHYSIOL más abajo -- si el usuario
       ya la habia creado el mismo a mano con un nombre levemente distinto
       mientras esperaba esta correccion, para reusar ESE producto en vez
       de crear una cuenta padre duplicada).
    B) Cualquier producto (de cualquier proveedor) que haya quedado creado
       como placeholder a precio 0 pero que en realidad corresponde a una
       variante ya existente segun su propia Denominacion (mismo mecanismo
       que el cruce por descripcion agregado en seed_homologacion_y_stock_
       inicial, aplicado aqui retroactivamente).
    C) Codigos internos que hayan quedado "pendiente" (proveedor no
       reconocido en su momento): crea el proveedor que falte y liga/crea
       su producto, igual que ya hace seed_homologacion_y_stock_inicial()
       para una base nueva."""
    if not os.path.isfile(INVENTARIOS_CODIGOS_INTERNOS_EXCEL):
        return

    import openpyxl

    resumen = {
        "variantes_migradas": 0, "padres_reparados": set(),
        "denominacion_migradas": 0, "proveedores_creados": set(),
        "pendientes_resueltos": 0,
    }

    # ---- Parte A: familias PHYSIOL sueltas -> bajo su cuenta padre ----
    physiol = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == "BVI PHYSIOL").first()
    if physiol and os.path.isfile(INVENTARIOS_CODIGOS_INTERNOS_EXCEL):
        wb = openpyxl.load_workbook(INVENTARIOS_CODIGOS_INTERNOS_EXCEL, data_only=True)
        if "LENTES PHYSIOL" in wb.sheetnames:
            ws = wb["LENTES PHYSIOL"]
            fila_inicio = 2
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
                primera = str(row[0]).strip().lower() if row and row[0] else ""
                if primera == "codigo padre":
                    fila_inicio = i + 1
                    break

            # Nombres alternativos bajo los que el usuario ya habia creado a
            # mano una cuenta padre (mientras esperaba esta correccion),
            # levemente distintos del nombre oficial de este archivo -- si
            # se encuentra uno, se reusa ese Producto (conservando su id y
            # su precio, que suele ser mas confiable que el "precio copiado
            # de la familia hermana" que se usa como respaldo) y solo se le
            # corrige el nombre al oficial.
            ALIAS_PADRES_PHYSIOL = {
                "SERENITY TORIC PODS49P": "SERENITY TORIC PODST49P",
            }

            productos_por_codigo = {
                p.codigo.strip().upper(): p
                for p in Producto.query.filter_by(proveedor_id=physiol.id).all()
            }
            variantes_por_clave = {
                (v.producto_id, v.codigo.strip().upper())
                for v in ProductoVariante.query.join(Producto).filter(Producto.proveedor_id == physiol.id).all()
            }

            def obtener_o_crear_padre(codigo_padre):
                clave = codigo_padre.upper()
                padre = productos_por_codigo.get(clave)
                if not padre:
                    alias = ALIAS_PADRES_PHYSIOL.get(clave)
                    if alias and alias in productos_por_codigo:
                        padre = productos_por_codigo.pop(alias)
                        padre.codigo = codigo_padre
                        padre.descripcion = codigo_padre
                        productos_por_codigo[clave] = padre
                if not padre:
                    hermano_codigo = re.sub(r"\btoric\b", "", codigo_padre, flags=re.IGNORECASE)
                    hermano_codigo = re.sub(r"\s+", " ", hermano_codigo).strip()
                    hermano = productos_por_codigo.get(hermano_codigo.upper())
                    padre = Producto(
                        proveedor_id=physiol.id, codigo=codigo_padre, descripcion=codigo_padre,
                        empaque=(hermano.empaque if hermano else 1),
                        moneda=(hermano.moneda if hermano else (physiol.moneda_default or "USD")),
                        precio_caja=(hermano.precio_caja if hermano else 0),
                        precio_unitario=(hermano.precio_unitario if hermano else 0),
                        activo=True,
                    )
                    db.session.add(padre)
                    db.session.flush()
                    productos_por_codigo[clave] = padre
                    resumen["padres_reparados"].add(codigo_padre)
                return padre

            vistos = set()
            for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
                codigo_padre, codigo_variante, descripcion = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
                codigo_padre = str(codigo_padre).strip() if codigo_padre else ""
                codigo_variante = str(codigo_variante).strip() if codigo_variante else ""
                descripcion = str(descripcion).strip() if descripcion else ""
                if not codigo_padre or not codigo_variante:
                    continue
                clave_vista = (codigo_padre.upper(), codigo_variante.upper())
                if clave_vista in vistos:
                    continue
                vistos.add(clave_vista)

                placeholder = productos_por_codigo.get(codigo_variante.upper())
                if not placeholder:
                    continue  # nunca quedo mal creado, o ya se reparo antes

                padre = obtener_o_crear_padre(codigo_padre)
                if placeholder.id == padre.id:
                    continue

                clave_variante = (padre.id, codigo_variante.upper())
                if clave_variante in variantes_por_clave:
                    variante_final = ProductoVariante.query.filter_by(producto_id=padre.id, codigo=codigo_variante).first()
                else:
                    variante_final = ProductoVariante(producto_id=padre.id, codigo=codigo_variante, descripcion=descripcion or codigo_variante)
                    db.session.add(variante_final)
                    db.session.flush()
                    variantes_por_clave.add(clave_variante)

                codigo_interno = placeholder.codigo_interno_inventario
                if codigo_interno:
                    variante_final.codigo_interno_inventario = codigo_interno
                    HomologacionStock.query.filter_by(producto_id=placeholder.id).update({
                        "producto_id": padre.id, "variante_id": variante_final.id,
                    })
                    StockExistencia.query.filter_by(producto_id=placeholder.id).update({
                        "producto_id": padre.id, "variante_id": variante_final.id,
                    })
                db.session.delete(placeholder)
                del productos_por_codigo[codigo_variante.upper()]
                resumen["variantes_migradas"] += 1

            db.session.commit()

    # ---- Parte B: cruce por Denominacion para placeholders ya creados ----
    candidatos = Producto.query.filter(
        Producto.codigo_interno_inventario.isnot(None),
        Producto.precio_unitario == 0, Producto.precio_caja == 0,
    ).all()
    for placeholder in candidatos:
        homolog = HomologacionStock.query.filter_by(producto_id=placeholder.id, variante_id=None).first()
        if not homolog or not homolog.descripcion_referencia:
            continue
        descripcion_upper = homolog.descripcion_referencia.strip().upper()
        variante_match = ProductoVariante.query.join(Producto).filter(
            Producto.proveedor_id == placeholder.proveedor_id,
            db.or_(
                db.func.upper(ProductoVariante.codigo) == descripcion_upper,
                db.func.upper(ProductoVariante.descripcion) == descripcion_upper,
            ),
        ).first()
        if not variante_match:
            continue
        variante_match.codigo_interno_inventario = placeholder.codigo_interno_inventario
        homolog.producto_id = variante_match.producto_id
        homolog.variante_id = variante_match.id
        StockExistencia.query.filter_by(producto_id=placeholder.id).update({
            "producto_id": variante_match.producto_id, "variante_id": variante_match.id,
        })
        db.session.delete(placeholder)
        resumen["denominacion_migradas"] += 1
    db.session.commit()

    # ---- Parte C: proveedores faltantes + codigos que quedaron pendientes ----
    pendientes = HomologacionStock.query.filter_by(estado="pendiente").all()
    if pendientes and os.path.isfile(INVENTARIOS_CODIGOS_INTERNOS_EXCEL):
        wb2 = openpyxl.load_workbook(INVENTARIOS_CODIGOS_INTERNOS_EXCEL, data_only=True)
        if "INVENTARIO ACTUAL" in wb2.sheetnames:
            ws2 = wb2["INVENTARIO ACTUAL"]
            clasificacion = {}
            for row in ws2.iter_rows(min_row=1, values_only=True):
                a = row[0] if len(row) > 0 else None
                b = row[1] if len(row) > 1 else None
                if a is None and b is None:
                    continue
                if b is None or not (isinstance(b, str) and b.strip().startswith("'")):
                    continue
                codigo_interno = _normalizar_codigo_interno(b)
                if not codigo_interno or codigo_interno in clasificacion:
                    continue
                j = row[9] if len(row) > 9 else None
                k = row[10] if len(row) > 10 else None
                descripcion = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                clasificacion[codigo_interno] = {
                    "j": str(j).strip() if j is not None else "",
                    "k": str(k).strip() if k is not None else "",
                    "descripcion": descripcion,
                }

            proveedores_cache = {}
            for homolog in pendientes:
                info = clasificacion.get(homolog.codigo_interno)
                if not info:
                    continue
                k_upper = info["k"].strip().upper()
                j_valor = info["j"].strip()
                if not k_upper or k_upper == "NO INCLUIR":
                    homolog.estado = "excluido"
                    StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).delete()
                    continue
                if k_upper == "SIN MARCA":
                    homolog.estado = "sin_marca"
                    continue

                if k_upper not in proveedores_cache:
                    proveedores_cache[k_upper] = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == k_upper).first()
                proveedor = proveedores_cache[k_upper]
                if not proveedor and j_valor:
                    nombre_nuevo = info["k"].strip().rstrip("*").strip() or info["k"].strip()
                    proveedor = Proveedor(nombre=nombre_nuevo, tipo="Extranjero", moneda_default="USD", activo=True)
                    db.session.add(proveedor)
                    db.session.flush()
                    proveedores_cache[k_upper] = proveedor
                    resumen["proveedores_creados"].add(nombre_nuevo)
                if not proveedor or not j_valor:
                    continue  # sigue sin info suficiente, se deja pendiente

                j_upper = j_valor.upper()
                producto_match = Producto.query.filter(
                    Producto.proveedor_id == proveedor.id, db.func.upper(Producto.codigo) == j_upper
                ).first()
                variante_match = None
                if not producto_match:
                    variante_match = ProductoVariante.query.join(Producto).filter(
                        Producto.proveedor_id == proveedor.id, db.func.upper(ProductoVariante.codigo) == j_upper
                    ).first()
                if not producto_match and not variante_match:
                    descripcion_upper = (info["descripcion"] or "").strip().upper()
                    if descripcion_upper:
                        variante_match = ProductoVariante.query.join(Producto).filter(
                            Producto.proveedor_id == proveedor.id,
                            db.or_(
                                db.func.upper(ProductoVariante.codigo) == descripcion_upper,
                                db.func.upper(ProductoVariante.descripcion) == descripcion_upper,
                            ),
                        ).first()

                if producto_match:
                    producto_match.codigo_interno_inventario = homolog.codigo_interno
                    homolog.estado = "vinculado"
                    homolog.producto_id = producto_match.id
                    StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).update({"producto_id": producto_match.id})
                elif variante_match:
                    variante_match.codigo_interno_inventario = homolog.codigo_interno
                    homolog.estado = "vinculado"
                    homolog.producto_id = variante_match.producto_id
                    homolog.variante_id = variante_match.id
                    StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).update({
                        "producto_id": variante_match.producto_id, "variante_id": variante_match.id,
                    })
                else:
                    nuevo = Producto(
                        proveedor_id=proveedor.id, codigo=j_valor, descripcion=info["descripcion"] or j_valor,
                        empaque=1, moneda=proveedor.moneda_default or "USD", precio_caja=0, precio_unitario=0,
                        activo=True, codigo_interno_inventario=homolog.codigo_interno,
                    )
                    db.session.add(nuevo)
                    db.session.flush()
                    homolog.estado = "vinculado"
                    homolog.producto_id = nuevo.id
                    StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).update({"producto_id": nuevo.id})
                resumen["pendientes_resueltos"] += 1
            db.session.commit()

    if any(resumen.values()):
        print(
            f"[reparación ronda W] {resumen['variantes_migradas']} variante(s) PHYSIOL migradas bajo su cuenta "
            f"padre ({sorted(resumen['padres_reparados'])} cuenta(s) padre nueva(s) creada(s)), "
            f"{resumen['denominacion_migradas']} producto(s) migrados por coincidencia de Denominación, "
            f"{len(resumen['proveedores_creados'])} proveedor(es) nuevo(s) creados ({sorted(resumen['proveedores_creados'])}), "
            f"{resumen['pendientes_resueltos']} código(s) pendiente(s) resueltos."
        )


def reparar_variantes_medicontur_ronda_ae():
    """Ronda AE (2026-09-14, punto 1): repara los códigos de variante
    MEDICONTUR que quedaron con la nomenclatura VIEJA de proveedor -- el
    usuario corrigió 'Listado codigos lentes medicontur.xlsx' para que los
    productos cuyo código padre es 877PETY y cuya descripción termina en
    "CYL 1" usen la letra "O" al final del código de proveedor (ej.
    877PETYP210O), en vez del "0" (cero) que se había usado antes. De
    paso se detectó que 2 códigos (677MTYP230A y 877PETYP1200) habían
    quedado compartiendo el MISMO texto de código entre 2 dioptrías
    distintas -- el archivo nuevo los separa en códigos propios.

    Como seed_variantes_lentes_medicontur() está gateada (corre una sola
    vez, para no pisar ediciones manuales), estas correcciones NUNCA
    llegan solas a una base ya sembrada -- de ahí esta función, que SÍ
    corre en cada arranque (no gateada, idempotente: si ya no encuentra
    nada por renombrar, no hace nada). NUNCA borra ni crea
    ProductoVariante -- solo RENOMBRA pv.codigo (y ajusta la descripción
    si cambió), comparando para cada código padre lo que ya existe en la
    base contra lo que dice el archivo actual: primero empareja por
    código EXACTO cuando ese código no está duplicado (para el caso de
    "solo cambió la descripción"), y con lo que sobra empareja por
    descripción (el caso real de esta ronda: un código viejo desaparece y
    otro nuevo -- con la misma descripción -- lo reemplaza). Los códigos
    duplicados en la base (2 filas con el mismo texto de código) se
    excluyen a propósito del emparejamiento por código exacto y se
    resuelven siempre por descripción, para no arriesgar pisar la
    descripción de la fila equivocada por el orden en que vuelva la
    consulta. HomologacionStock/StockExistencia enlazan por variante_id
    (no por el texto del código), así que renombrar es seguro y se
    refleja solo, tanto en Consulta de Stock como en los reportes."""
    medicontur = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == "MEDICONTUR").first()
    if not medicontur or not os.path.isfile(VARIANTES_LENTES_MEDICONTUR_EXCEL):
        return

    import openpyxl
    wb = openpyxl.load_workbook(VARIANTES_LENTES_MEDICONTUR_EXCEL, data_only=True)
    ws = wb.worksheets[0]
    fila_inicio = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        primera = str(row[0]).strip().lower() if row and row[0] else ""
        if primera == "codigo padre":
            fila_inicio = i + 1
            break

    nuevas_por_padre = defaultdict(list)
    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        codigo_padre, codigo_variante, descripcion = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
        codigo_padre = str(codigo_padre).strip() if codigo_padre else ""
        codigo_variante = str(codigo_variante).strip() if codigo_variante else ""
        descripcion = str(descripcion).strip() if descripcion else ""
        if not codigo_padre or not codigo_variante:
            continue
        nuevas_por_padre[codigo_padre.upper()].append({"codigo": codigo_variante, "desc": descripcion, "usada": False})

    renombres = []
    for padre in Producto.query.filter_by(proveedor_id=medicontur.id).all():
        nuevas = nuevas_por_padre.get(padre.codigo.strip().upper())
        if not nuevas:
            continue
        actuales = ProductoVariante.query.filter_by(producto_id=padre.id).all()
        conteo_codigo_actual = Counter(v.codigo.strip().upper() for v in actuales)

        # Paso 1: mismo código EXACTO (y sin ambigüedad) -- solo puede
        # haber cambiado la descripción.
        for v in actuales:
            if conteo_codigo_actual[v.codigo.strip().upper()] != 1:
                continue  # código duplicado en la base -- se resuelve en el paso 2
            for n in nuevas:
                if not n["usada"] and n["codigo"].upper() == v.codigo.strip().upper():
                    n["usada"] = True
                    if n["desc"] and v.descripcion != n["desc"]:
                        v.descripcion = n["desc"]
                    break

        # Paso 2: lo que sobra de cada lado se empareja por descripción.
        usados_codigo = {n["codigo"].upper() for n in nuevas if n["usada"]}
        sobran_variantes = [v for v in actuales if v.codigo.strip().upper() not in usados_codigo]
        por_desc = defaultdict(list)
        for v in sobran_variantes:
            por_desc[(v.descripcion or "").strip().upper()].append(v)
        for n in nuevas:
            if n["usada"]:
                continue
            candidatos = por_desc.get(n["desc"].strip().upper())
            if not candidatos:
                continue
            v = candidatos.pop(0)
            if v.codigo.strip().upper() != n["codigo"].upper():
                renombres.append(f"{v.codigo}->{n['codigo']}")
                v.codigo = n["codigo"]
            if n["desc"]:
                v.descripcion = n["desc"]
            n["usada"] = True

    if renombres:
        db.session.commit()
        print(f"[reparar_ronda_ae] MEDICONTUR: {len(renombres)} código(s) de variante renombrado(s): {renombres}")


def reparar_variantes_medicontur_ronda_ag():
    """Ronda AG (2026-09-15, punto 3): el usuario reportó 6 códigos MEDICONTUR
    (lentes "ADDON" -- familias A45DT, A45RD2, A46R, 690MY, 860PAY, 860PEY)
    que nunca habían existido en el catálogo ni como código padre ni como
    variante -- por eso el reporte de Compras Proveedor los mostraba sueltos
    con el número de código interno en vez de su descripción, y no agrupaban
    con nada. Se agregaron al final de 'Listado codigos lentes medicontur.
    xlsx' -- mismo mecanismo que reparar_variantes_physiol_ronda_ae() (esta
    función corre en cada arranque, a diferencia de seed_variantes_lentes_
    medicontur() que está gateada y no vuelve a leer el archivo una vez
    sembrada la base): crea el Producto padre si no existe (precio 0 EUR,
    a pedido explícito del usuario -- se corrige después desde el catálogo)
    y cada ProductoVariante que falte. Puramente ADITIVA -- si el padre o la
    variante ya existen, no los toca."""
    medicontur = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == "MEDICONTUR").first()
    if not medicontur or not os.path.isfile(VARIANTES_LENTES_MEDICONTUR_EXCEL):
        return

    import openpyxl
    wb = openpyxl.load_workbook(VARIANTES_LENTES_MEDICONTUR_EXCEL, data_only=True)
    ws = wb.worksheets[0]
    fila_inicio = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        primera = str(row[0]).strip().lower() if row and row[0] else ""
        if primera == "codigo padre":
            fila_inicio = i + 1
            break

    productos_padre = {
        p.codigo.strip().upper(): p
        for p in Producto.query.filter_by(proveedor_id=medicontur.id).all()
    }
    variantes_existentes = {
        (v.producto_id, v.codigo.strip().upper())
        for v in ProductoVariante.query.join(Producto).filter(Producto.proveedor_id == medicontur.id).all()
    }

    creadas = 0
    padres_creados = set()
    vistos = set()
    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        codigo_padre, codigo_variante, descripcion = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
        codigo_padre = str(codigo_padre).strip() if codigo_padre else ""
        codigo_variante = str(codigo_variante).strip() if codigo_variante else ""
        descripcion = str(descripcion).strip() if descripcion else ""
        if not codigo_padre or not codigo_variante:
            continue
        clave = (codigo_padre.upper(), codigo_variante.upper())
        if clave in vistos:
            continue
        vistos.add(clave)

        producto = productos_padre.get(codigo_padre.upper())
        if not producto:
            producto = Producto(
                proveedor_id=medicontur.id, codigo=codigo_padre, descripcion=descripcion or codigo_padre,
                empaque=1, moneda="EUR", precio_caja=0, precio_unitario=0, activo=True,
            )
            db.session.add(producto)
            db.session.flush()
            productos_padre[codigo_padre.upper()] = producto
            padres_creados.add(codigo_padre)

        if (producto.id, codigo_variante.upper()) in variantes_existentes:
            continue
        db.session.add(ProductoVariante(
            producto_id=producto.id, codigo=codigo_variante, descripcion=descripcion or codigo_variante,
        ))
        variantes_existentes.add((producto.id, codigo_variante.upper()))
        creadas += 1

    if creadas or padres_creados:
        db.session.commit()
        print(
            f"[reparar_ronda_ag] MEDICONTUR: {creadas} variante(s) nueva(s) agregada(s) "
            f"({len(padres_creados)} código(s) padre nuevo(s) creados, precio 0 a completar por el usuario: "
            f"{sorted(padres_creados)})."
        )


def reparar_variantes_physiol_ronda_ae():
    """Ronda AE (2026-09-14, punto 2): el usuario aportó un archivo nuevo,
    más completo, de variantes BVI PHYSIOL ('Ajuste de cuentas padre
    Physiol.xlsx' -- ver AJUSTE_PADRE_PHYSIOL_EXCEL) porque el histórico de
    compras traía códigos que no estaban referenciados con su código padre
    correcto. Mismo mecanismo/misma tabla que seed_variantes_lentes_
    physiol() (3 columnas: codigo padre, Codigo Producto, DESCRIPCION,
    encabezado detectado igual), pero esta función SÍ corre en cada
    arranque (no gateada) para que las variantes que falten en una base
    donde seed_variantes_lentes_physiol() ya corrió (con el archivo viejo,
    antes de este ajuste) se agreguen igual -- sin gate no hay forma de que
    ese seed original vuelva a correr. Es puramente ADITIVA: si la
    variante (código padre + código de variante) ya existe, no la toca; si
    el código padre no existe todavía como Producto, lo crea igual que
    hace seed_variantes_lentes_physiol() (copiando precio de una familia
    "hermana" si el nombre es la versión " Toric" de una ya existente)."""
    physiol = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == "BVI PHYSIOL").first()
    if not physiol or not os.path.isfile(AJUSTE_PADRE_PHYSIOL_EXCEL):
        return

    import openpyxl
    wb = openpyxl.load_workbook(AJUSTE_PADRE_PHYSIOL_EXCEL, data_only=True)
    if "LENTES PHYSIOL" not in wb.sheetnames:
        return
    ws = wb["LENTES PHYSIOL"]
    fila_inicio = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        primera = str(row[0]).strip().lower() if row and row[0] else ""
        if primera == "codigo padre":
            fila_inicio = i + 1
            break

    productos_padre = {
        p.codigo.strip().upper(): p
        for p in Producto.query.filter_by(proveedor_id=physiol.id).all()
    }
    variantes_existentes = {
        (v.producto_id, v.codigo.strip().upper())
        for v in ProductoVariante.query.join(Producto).filter(Producto.proveedor_id == physiol.id).all()
    }

    creadas = 0
    padres_creados = set()
    vistos = set()
    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        codigo_padre, codigo_variante, descripcion = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
        codigo_padre = str(codigo_padre).strip() if codigo_padre else ""
        codigo_variante = str(codigo_variante).strip() if codigo_variante else ""
        descripcion = str(descripcion).strip() if descripcion else ""
        if not codigo_padre or not codigo_variante:
            continue
        clave = (codigo_padre.upper(), codigo_variante.upper())
        if clave in vistos:
            continue
        vistos.add(clave)

        producto = productos_padre.get(codigo_padre.upper())
        if not producto:
            candidato_hermano = re.sub(r"\btoric\b", "", codigo_padre, flags=re.IGNORECASE)
            candidato_hermano = re.sub(r"\s+", " ", candidato_hermano).strip()
            hermano = productos_padre.get(candidato_hermano.upper()) if candidato_hermano else None
            producto = Producto(
                proveedor_id=physiol.id, codigo=codigo_padre, descripcion=codigo_padre,
                empaque=(hermano.empaque if hermano else 1),
                moneda=(hermano.moneda if hermano else (physiol.moneda_default or "USD")),
                precio_caja=(hermano.precio_caja if hermano else 0),
                precio_unitario=(hermano.precio_unitario if hermano else 0),
                activo=True,
            )
            db.session.add(producto)
            db.session.flush()
            productos_padre[codigo_padre.upper()] = producto
            padres_creados.add(codigo_padre)

        if (producto.id, codigo_variante.upper()) in variantes_existentes:
            continue
        db.session.add(ProductoVariante(
            producto_id=producto.id, codigo=codigo_variante, descripcion=descripcion or codigo_variante,
        ))
        variantes_existentes.add((producto.id, codigo_variante.upper()))
        creadas += 1

    if creadas or padres_creados:
        db.session.commit()
        print(
            f"[reparar_ronda_ae] BVI PHYSIOL: {creadas} variante(s) nueva(s) agregada(s) desde el archivo de "
            f"ajuste ({len(padres_creados)} código(s) padre nuevo(s) creados: {sorted(padres_creados)})."
        )


def _normalizar_codigo_ergopyme(valor):
    """Ronda AA (2026-09-13): normaliza un código interno del sistema de
    Inventarios (Ergopyme) tal como viene en los archivos de Reportes --
    ahí llegan como número (int o float, sin comilla ni ceros a la
    izquierda que preservar), a diferencia de _normalizar_codigo_interno
    (usado para Stock/Notas de pedido, que sí trae la comilla delante)."""
    if valor is None:
        return ""
    if isinstance(valor, float):
        valor = int(valor) if valor.is_integer() else valor
    s = str(valor).strip()
    if s.startswith("'"):
        s = s[1:].strip()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    return s


def seed_codigos_ergopyme():
    """Ronda AA (2026-09-13), convertido a upsert en Ronda AB (2026-09-14):
    lee el mapeo maestro Proveedor + Código Proveedor + Descripción para
    cada código interno de Ergopyme, desde 'codigos_ergopyme_homologacion.
    xlsx' ("2da Revisión códigos ergopyme" que mantiene el usuario) -- se
    usa para homologar en vivo cualquier archivo de Ergopyme (histórico de
    compras, Notas de pedido, Consulta de Stock, etc.) por código interno.

    A diferencia de Ronda AA, esto YA NO es "una sola vez": cada arranque
    vuelve a leer el archivo completo y actualiza (o crea) cada código por
    su codigo_interno, para que una corrección en el archivo de mapeo (ej.
    el usuario resuelve un '#N/A' de Proveedor y sube una versión nueva) se
    refleje con solo reiniciar la app, sin tener que vaciar la base de
    datos. No borra códigos que ya no aparezcan en el archivo, por si se
    sube por error una versión parcial."""
    if not os.path.isfile(CODIGOS_ERGOPYME_EXCEL):
        print(f"[seed] No se encontró {CODIGOS_ERGOPYME_EXCEL}, se omite la carga de códigos Ergopyme.")
        return

    wb = openpyxl.load_workbook(CODIGOS_ERGOPYME_EXCEL, data_only=True)
    ws = wb.worksheets[0]
    fila_inicio = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        primera = str(row[0]).strip().upper() if row and row[0] is not None else ""
        if primera == "CODIGO ITEM":
            fila_inicio = i + 1
            break
    if fila_inicio is None:
        print("[seed] No se encontró el encabezado 'CODIGO ITEM' en codigos_ergopyme_homologacion.xlsx, se omite.")
        return

    existentes = {c.codigo_interno: c for c in CodigoErgopyme.query.all()}
    creados = 0
    actualizados = 0
    vistos = set()
    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        if not row or row[0] is None:
            continue
        codigo = _normalizar_codigo_ergopyme(row[0])
        if not codigo or codigo in vistos:
            continue
        vistos.add(codigo)
        proveedor = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        codigo_proveedor = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        descripcion = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
        existente = existentes.get(codigo)
        if existente:
            if (existente.proveedor_nombre != proveedor
                    or existente.codigo_proveedor != codigo_proveedor
                    or existente.descripcion != descripcion):
                existente.proveedor_nombre = proveedor
                existente.codigo_proveedor = codigo_proveedor
                existente.descripcion = descripcion
                actualizados += 1
        else:
            nuevo = CodigoErgopyme(
                codigo_interno=codigo, proveedor_nombre=proveedor,
                codigo_proveedor=codigo_proveedor, descripcion=descripcion,
            )
            db.session.add(nuevo)
            existentes[codigo] = nuevo
            creados += 1
    db.session.commit()
    print(
        f"[seed] Códigos Ergopyme: {creados} nuevo(s), {actualizados} "
        "actualizado(s) para homologación (Stock, Reportes)."
    )


def seed_compras_historicas():
    """Ronda AA (2026-09-13): carga UNA VEZ el histórico de compras/
    importaciones a proveedores anterior a este sistema, desde
    'historico_compras_proveedores.xlsx' ("Data Histórica Compra
    proveedores, costo fletes y gastos Importación"), homologando cada
    fila contra CodigoErgopyme para saber el proveedor "real" (mismo
    nombre que usa el catálogo de Proveedor de este sistema). El reporte
    Compras Proveedor (ver app.py) combina estas filas "congeladas" con
    las compras hechas DESDE la plataforma (calculadas en vivo a partir
    de Importacion/Parcial/ParcialLinea) -- gateado igual que los demás
    seeds, si ya hay datos no hace nada."""
    if CompraHistorica.query.count() > 0:
        return
    if not os.path.isfile(HISTORICO_COMPRAS_EXCEL):
        print(f"[seed] No se encontró {HISTORICO_COMPRAS_EXCEL}, se omite la carga de compras históricas.")
        return

    seed_codigos_ergopyme()
    mapeo = {c.codigo_interno: c for c in CodigoErgopyme.query.all()}

    wb = openpyxl.load_workbook(HISTORICO_COMPRAS_EXCEL, data_only=True)
    ws = wb.worksheets[0]

    # Detecta la fila de encabezado buscando "CODIGO" y "UNIDADES" (en vez
    # de asumir una fila fija) -- el archivo real trae una fila en blanco
    # antes del encabezado. La columna "Compradora" (ACCUVISION/
    # ACCUMEDICAL) no trae texto de encabezado -- viene justo despues de
    # la ultima columna con nombre ("COSTO"), se ubica por posicion
    # relativa en vez de un numero de columna fijo.
    fila_inicio = None
    columnas = {}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        nombres = [str(v).strip().upper() if v is not None else "" for v in row]
        if "CODIGO" in nombres and "UNIDADES" in nombres:
            fila_inicio = i + 1
            for idx, nombre in enumerate(nombres):
                if nombre:
                    columnas[nombre] = idx
            break
    if fila_inicio is None:
        print("[seed] No se encontró el encabezado esperado en historico_compras_proveedores.xlsx, se omite.")
        return

    def col(nombre, row):
        idx = columnas.get(nombre)
        return row[idx] if idx is not None and idx < len(row) else None

    idx_costo = columnas.get("COSTO")
    idx_compradora = (idx_costo + 1) if idx_costo is not None else None

    creados = 0
    homologados = 0
    codigos_sin_homologar = set()
    filas_sin_codigo = 0

    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        if not row or all(v is None for v in row):
            continue
        codigo_crudo = col("CODIGO", row)
        if codigo_crudo is None:
            filas_sin_codigo += 1
            continue
        codigo = _normalizar_codigo_ergopyme(codigo_crudo)
        if not codigo:
            filas_sin_codigo += 1
            continue

        fecha_factura = col("FECHA", row)
        fecha_factura = fecha_factura.date() if isinstance(fecha_factura, datetime) else None

        homolog = mapeo.get(codigo)
        proveedor_valor = col("PROVEEDOR", row)
        proveedor_original = str(proveedor_valor).strip() if proveedor_valor is not None else ""
        # El archivo de mapeo trae 59 códigos con Proveedor "#N/A" (una
        # fórmula de búsqueda sin resolver en el Excel original) -- eso NO
        # cuenta como una homologación real: se usa el proveedor tal como
        # viene en el propio histórico (más confiable que "#N/A").
        proveedor_mapeo = (homolog.proveedor_nombre or "").strip() if homolog else ""
        if homolog and proveedor_mapeo and proveedor_mapeo.upper() != "#N/A":
            proveedor_homologado = proveedor_mapeo
            codigo_proveedor = homolog.codigo_proveedor or ""
            homologados += 1
        else:
            proveedor_homologado = proveedor_original
            codigo_proveedor = (homolog.codigo_proveedor or "") if homolog else ""
            codigos_sin_homologar.add(codigo)

        def num(nombre):
            # Ronda AD (2026-09-14): corrige la causa RAÍZ de un bug
            # encontrado en la ronda AC -- 54 filas de DORC (facturas
            # CD100058713/CD100059451) traían la columna "Paridad EUR" como
            # TEXTO con coma decimal ("0,8687", formato chileno/europeo) en
            # vez de un número real de Excel. `float("0,8687")` lanza
            # ValueError, y el código anterior lo coercionaba en silencio a
            # 0.0 -- eso se interpretó (ronda AC) como "paridad faltante",
            # cuando en realidad el dato SÍ estaba, solo mal formateado. El
            # usuario confirmó revisando el archivo que ninguna celda real
            # está vacía/en 0. Ahora se detecta el formato de texto y se
            # convierte antes de fallar (soporta también miles con "." si
            # los hubiera, ej. "1.234,56" -> 1234.56).
            v = col(nombre, row)
            if v is None:
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            texto = str(v).strip()
            if not texto:
                return 0.0
            if "," in texto:
                texto = texto.replace(".", "").replace(",", ".")
            try:
                return float(texto)
            except (TypeError, ValueError):
                return 0.0

        compradora_valor = row[idx_compradora] if idx_compradora is not None and idx_compradora < len(row) else None

        db.session.add(CompraHistorica(
            fecha_factura=fecha_factura,
            mes_anio=str(col("MES AÑO", row) or "").strip(),
            proveedor_original=proveedor_original,
            proveedor_homologado=proveedor_homologado,
            factura=str(col("FACTURA", row) or "").strip(),
            tipo_cambio=num("USD TIPO CAMBIO"),
            paridad_eur=num("PARIDAD EUR"),
            transporte=str(col("TRANSPORTE", row) or "").strip(),
            codigo_interno=codigo,
            codigo_proveedor=codigo_proveedor,
            descripcion=str(col("DESCRIPCION", row) or "").strip(),
            tipo_flete=str(col("TIPO FLETE", row) or "").strip(),
            unidades=num("UNIDADES"),
            total_invoice=num("TOTAL IVOICE EU"),
            total_usd=num("US$"),
            flete_usd=num("FLETE"),
            seguro_usd=num("SEGURO"),
            cif_usd=num("C.I.F."),
            cif_clp=num("C.I.F. $"),
            derechos_clp=num("DERECHOS  $"),
            otros_gastos_clp=num("OTROS GASTOS $"),
            costo_total_clp=num("TOTAL COSTO $"),
            costo_unitario_clp=num("COSTO UNITARIO $"),
            categoria=str(col("LINEA", row) or "").strip(),
            otros_costos_usd=num("OTROS COSTOS USD$"),
            empresa_compradora=str(compradora_valor or "").strip(),
            homologado=bool(homolog and proveedor_mapeo and proveedor_mapeo.upper() != "#N/A"),
        ))
        creados += 1

    db.session.commit()
    print(
        f"[seed] Importadas {creados} línea(s) de compras históricas "
        f"({homologados} homologadas contra el mapeo Ergopyme, "
        f"{len(codigos_sin_homologar)} código(s) distintos sin homologar, "
        f"{filas_sin_codigo} fila(s) sin código omitidas)."
    )
    if codigos_sin_homologar:
        print(f"[seed] Códigos históricos sin homologar: {sorted(codigos_sin_homologar)}")


# Contraseña temporal del Administrador inicial (ronda R, 2026-09-12) --
# ver seed_administrador_inicial() abajo. Puramente informativa aca (el
# usuario la cambia desde "Mi cuenta" apenas entra la primera vez); no es
# un secreto de produccion real, es solo para no dejar la app sin forma de
# entrar la primera vez que se activa el login.
PASSWORD_TEMPORAL_ADMIN_INICIAL = "CambiaEsta123!"


def seed_administrador_inicial():
    """Crea el primer Rol Administrador + el primer Usuario la primera vez
    que la app arranca con el sistema de login (ronda R, 2026-09-12) -- sin
    esto, al activar el login nadie podria entrar nunca mas. Gateado: si ya
    existe algun Usuario, no hace nada (para no resetear la contraseña de
    nadie en arranques posteriores)."""
    if Usuario.query.count() > 0:
        return
    rol_admin = Rol.query.filter_by(es_administrador=True).first()
    if not rol_admin:
        rol_admin = Rol(nombre="Administrador", es_administrador=True)
        db.session.add(rol_admin)
        db.session.flush()
    usuario = Usuario(
        nombre_completo="Jesus Rafael",
        email="jesusrfr3008@gmail.com",
        rol_id=rol_admin.id,
        activo=True,
    )
    usuario.set_password(PASSWORD_TEMPORAL_ADMIN_INICIAL)
    db.session.add(usuario)
    db.session.commit()
    print(
        f"[seed] Usuario Administrador inicial creado ({usuario.email}) -- "
        "cambiar la contraseña temporal desde 'Mi cuenta' cuanto antes."
    )


with app.app_context():
    os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)
    os.makedirs(DOCUMENTOS_DIR, exist_ok=True)
    os.makedirs(EMPRESAS_LOGOS_DIR, exist_ok=True)
    os.makedirs(PDF_OC_DIR, exist_ok=True)
    os.makedirs(REPORTES_DIR, exist_ok=True)
    db.create_all()
    ensure_schema_migrations()
    seed_from_excel(app)
    seed_empresas_compradoras()
    seed_variantes_lentes_medicontur()
    seed_variantes_lentes_physiol()
    seed_homologacion_y_stock_inicial()
    seed_administrador_inicial()
    reparar_datos_ronda_w()
    reparar_variantes_medicontur_ronda_ae()
    reparar_variantes_medicontur_ronda_ag()
    reparar_variantes_physiol_ronda_ae()
    seed_codigos_ergopyme()
    seed_compras_historicas()


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


def siguiente_correlativo_inventario(empresa_id):
    """Sugiere el siguiente correlativo de importacion del sistema de
    Inventarios para una empresa compradora (ronda O, 2026-09-12) -- ese
    correlativo vive en el OTRO sistema, no en este, asi que no hay forma de
    generarlo desde cero: se calcula como el maximo ya guardado en
    Importacion.numero_correlativo_inventario para esa empresa + 1. Si
    todavia no hay ninguno cargado para esa empresa, devuelve None (no se
    inventa un numero de partida) -- el usuario debe indicar el primero a
    mano, tal como avisó."""
    if not empresa_id:
        return None
    maximo = db.session.query(db.func.max(Importacion.numero_correlativo_inventario)).filter(
        Importacion.empresa_id == empresa_id
    ).scalar()
    return (maximo + 1) if maximo else None


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
                # Ronda T (2026-09-12): la orden nueva hereda quien creo la
                # original y su estado de aprobacion -- si no se copiaran, la
                # parte separada quedaria "huerfana" (sin creado_por_usuario_id)
                # y el usuario que la creo dejaria de poder verla como suya en
                # "Mis ordenes"/"Por Aprobar".
                creado_por_usuario_id=orden.creado_por_usuario_id,
                estado_aprobacion=orden.estado_aprobacion,
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


def reparar_paridad_eur_historica():
    """Ronda AD (2026-09-14): repara los datos YA CARGADOS de
    compras_historicas (tabla insert-only/gated, no se vuelve a sembrar)
    afectados por el bug de parseo corregido en seed_compras_historicas
    (ver num() ahí): las 54 filas de DORC (facturas CD100058713 y
    CD100059451) que quedaron con paridad_eur=0.0 porque la celda original
    del Excel era el texto "0,8687" (coma decimal), no un dato faltante.
    Idempotente: si ya no queda ninguna fila en 0.0 para esas facturas, no
    hace nada."""
    filas = CompraHistorica.query.filter(
        CompraHistorica.factura.in_(["CD100058713", "CD100059451"]),
        CompraHistorica.proveedor_original == "DORC",
        CompraHistorica.paridad_eur == 0.0,
    ).all()
    if not filas:
        return
    for f in filas:
        f.paridad_eur = 0.8687
    db.session.commit()
    print(
        f"[reparar] Corregida paridad_eur (0.8687) en {len(filas)} fila(s) históricas de DORC "
        "(facturas CD100058713/CD100059451, dato de texto con coma decimal mal parseado)."
    )


with app.app_context():
    reparar_ordenes_mezcladas()
    limpiar_ordenes_canceladas()
    reparar_lotes_legacy()
    reparar_paridad_eur_historica()


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
    # Ronda Y (2026-09-13): la tabla "Ultimas ordenes de compra" muestra
    # proveedor y montos -- los mismos datos sensibles que protege el
    # acceso a /ordenes (ver ordenes_list, y el mismo permiso que la
    # navegacion usa para mostrar el menu "Compras / Ordenes"). Antes se
    # calculaba y mostraba a CUALQUIER usuario logueado sin chequear
    # permiso, por lo que un perfil de solo Consulta de Stock (ej. Felipe
    # Bravo, perfil COMERCIALES) la veia igual aunque no pueda entrar a
    # Ordenes. Se calcula solo si corresponde.
    puede_ver_ordenes = current_user.tiene_permiso("crear_orden") or current_user.tiene_permiso("aprobar_orden")
    ultimas_ordenes = (
        OrdenCompra.query.order_by(OrdenCompra.id.desc()).limit(6).all()
        if puede_ver_ordenes else []
    )
    return render_template(
        "dashboard.html",
        total_proveedores=total_proveedores,
        total_productos=total_productos,
        total_ordenes=total_ordenes,
        ordenes_pendientes=ordenes_pendientes,
        ultimas_ordenes=ultimas_ordenes,
        puede_ver_ordenes=puede_ver_ordenes,
    )


# ---------------------------------------------------------------------------
# Modulo: Mantenimiento de Proveedores
# ---------------------------------------------------------------------------

@app.route("/proveedores")
@requiere_permiso("crear_orden", "generar_costeo")
def proveedores_list():
    q = request.args.get("q", "").strip()
    query = Proveedor.query
    if q:
        query = query.filter(Proveedor.nombre.ilike(f"%{q}%"))
    proveedores = query.order_by(Proveedor.nombre).all()
    return render_template("proveedores/list.html", proveedores=proveedores, q=q)


@app.route("/proveedores/nuevo", methods=["GET", "POST"])
@requiere_permiso("crear_orden", "generar_costeo")
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
            codigo_sistema_inventario=request.form.get("codigo_sistema_inventario", "").strip(),
            notas=request.form.get("notas", "").strip(),
            activo=True,
        )
        db.session.add(prov)
        db.session.commit()
        flash(f"Proveedor '{prov.nombre}' creado correctamente.", "success")
        return redirect(url_for("proveedores_list"))
    return render_template("proveedores/form.html", proveedor=None)


@app.route("/proveedores/<int:proveedor_id>/editar", methods=["GET", "POST"])
@requiere_permiso("crear_orden", "generar_costeo")
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
        prov.codigo_sistema_inventario = request.form.get("codigo_sistema_inventario", "").strip()
        prov.notas = request.form.get("notas", "").strip()
        db.session.commit()
        flash(f"Proveedor '{prov.nombre}' actualizado.", "success")
        return redirect(url_for("proveedores_list"))
    return render_template("proveedores/form.html", proveedor=prov)


@app.route("/proveedores/<int:proveedor_id>/eliminar", methods=["POST"])
@requiere_permiso("crear_orden", "generar_costeo")
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
@requiere_permiso("crear_orden", "generar_costeo")
def proveedores_reactivar(proveedor_id):
    prov = Proveedor.query.get_or_404(proveedor_id)
    prov.activo = True
    db.session.commit()
    flash(f"Proveedor '{prov.nombre}' reactivado.", "success")
    return redirect(url_for("proveedores_list"))


@app.route("/proveedores/<int:proveedor_id>")
@requiere_permiso("crear_orden", "generar_costeo")
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
@requiere_permiso("crear_orden", "generar_costeo")
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
@requiere_permiso("crear_orden", "generar_costeo")
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
@requiere_permiso("crear_orden", "generar_costeo")
def productos_eliminar(producto_id):
    """Ronda U (2026-09-12, punto 2): elimina el producto DE VERDAD del
    catálogo (antes esta ruta existía pero no estaba conectada a ningún
    botón, y solo desactivaba). Solo se permite si el producto nunca se usó
    en ninguna Orden de Compra ni en ningún Parcial de Costeo -- si tiene
    historial, se desactiva en su lugar (igual que el checkbox 'Activo en
    catálogo' del modal Editar) para no romper ningún dato ya guardado."""
    producto = Producto.query.get_or_404(producto_id)
    proveedor_id = producto.proveedor_id
    en_uso = (
        OrdenCompraLinea.query.filter_by(producto_id=producto.id).count() > 0
        or ParcialLinea.query.filter_by(producto_id=producto.id).count() > 0
    )
    if en_uso:
        producto.activo = False
        db.session.commit()
        flash(
            f"'{producto.codigo}' ya se usó en alguna orden o costeo: no se puede eliminar sin perder "
            "ese historial, así que se desactivó en su lugar.",
            "warning",
        )
    else:
        ProductoVariante.query.filter_by(producto_id=producto.id).delete()
        codigo = producto.codigo
        db.session.delete(producto)
        db.session.commit()
        flash(f"Producto '{codigo}' eliminado del catálogo.", "success")
    return redirect(url_for("proveedores_detalle", proveedor_id=proveedor_id))


# ---------------------------------------------------------------------------
# Modulo: Compras (Ordenes de Compra)
# ---------------------------------------------------------------------------

@app.route("/ordenes")
@requiere_permiso("crear_orden", "aprobar_orden")
def ordenes_list():
    estado = request.args.get("estado", "")
    proveedor_id = request.args.get("proveedor_id", "")
    empresa_id = request.args.get("empresa_id", "")
    mostrar_despachadas = request.args.get("mostrar_despachadas") == "1"
    # Ronda T (2026-09-12): "los registros" (este listado) solo muestran
    # ordenes ya Aprobadas -- las 'Por Aprobar' y 'Sin Emitir' se gestionan
    # aparte, en /ordenes/por-aprobar (ver ordenes_por_aprobar).
    query = OrdenCompra.query.filter(
        db.or_(OrdenCompra.estado_aprobacion.is_(None), OrdenCompra.estado_aprobacion == "Aprobada")
    )
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


@app.route("/ordenes/por-aprobar")
@requiere_permiso("crear_orden", "aprobar_orden")
def ordenes_por_aprobar():
    """Ronda T (2026-09-12): bandeja de ordenes que todavia NO son una
    Orden de Compra definitiva. Quien tiene permiso de Aprobacion ve TODAS
    las 'Por Aprobar' (de cualquier creador) para poder actuar sobre ellas;
    quien solo tiene Creacion de Orden (sin Aprobacion) ve unicamente las
    SUYAS -- tanto 'Por Aprobar' (a la espera) como 'Sin Emitir'
    (devueltas, para corregir y reenviar) -- funciona como su propio 'Mis
    ordenes', igual que ya existia para el perfil Orden Simple."""
    puede_aprobar = current_user.tiene_permiso("aprobar_orden")
    if puede_aprobar:
        query = OrdenCompra.query.filter(
            db.or_(OrdenCompra.estado_aprobacion == "Por Aprobar", db.and_(
                OrdenCompra.estado_aprobacion == "Sin Emitir",
                OrdenCompra.creado_por_usuario_id == current_user.id,
            ))
        )
    else:
        query = OrdenCompra.query.filter(
            OrdenCompra.estado_aprobacion.in_(["Por Aprobar", "Sin Emitir"]),
            OrdenCompra.creado_por_usuario_id == current_user.id,
        )
    ordenes = query.order_by(OrdenCompra.id.desc()).all()
    return render_template("ordenes/por_aprobar.html", ordenes=ordenes, puede_aprobar=puede_aprobar)


@app.route("/ordenes/<int:orden_id>/aprobar", methods=["POST"])
@requiere_permiso("aprobar_orden")
def ordenes_aprobar(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.estado_aprobacion != "Por Aprobar":
        flash("Esta orden ya no está 'Por Aprobar'.", "warning")
        return redirect(url_for("ordenes_por_aprobar"))
    orden.estado_aprobacion = "Aprobada"
    # Ronda U (2026-09-12, punto 1): los productos nuevos que dio de alta
    # esta orden (perfil "Creación de Orden Simple") nacieron inactivos para
    # no ensuciar el catálogo con datos sin confirmar -- recien ahora, al
    # aprobarse, se activan de verdad.
    activados = 0
    for linea in orden.lineas:
        if linea.producto_creado_por_esta_orden and linea.producto and not linea.producto.activo:
            linea.producto.activo = True
            activados += 1
    db.session.commit()
    mensaje = f"Orden {orden.numero_po} aprobada -- ya es una Orden de Compra definitiva y aparece en Compras/Órdenes."
    if activados:
        mensaje += f" Se activaron {activados} producto(s) nuevo(s) en el catálogo del proveedor."
    flash(mensaje, "success")
    return redirect(url_for("ordenes_por_aprobar"))


@app.route("/ordenes/<int:orden_id>/devolver-aprobacion", methods=["POST"])
@requiere_permiso("aprobar_orden")
def ordenes_devolver_aprobacion(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.estado_aprobacion != "Por Aprobar":
        flash("Esta orden ya no está 'Por Aprobar'.", "warning")
        return redirect(url_for("ordenes_por_aprobar"))
    orden.estado_aprobacion = "Sin Emitir"
    motivo = (request.form.get("motivo") or "").strip()
    sello = f"[{datetime.utcnow().strftime('%Y-%m-%d %H:%M')} · {current_user.nombre_completo}]"
    nota = f"{sello} Devolvió la orden (queda 'Sin Emitir')."
    if motivo:
        nota += f" Motivo: {motivo}"
    orden.notas = f"{orden.notas}\n{nota}" if orden.notas else nota
    db.session.commit()
    flash(f"Orden {orden.numero_po} devuelta -- queda 'Sin Emitir', visible solo para quien la creó.", "info")
    return redirect(url_for("ordenes_por_aprobar"))


@app.route("/ordenes/<int:orden_id>/reenviar-aprobacion", methods=["POST"])
@requiere_permiso("crear_orden", "orden_simple")
def ordenes_reenviar_aprobacion(orden_id):
    """Quien creó la orden la reenvía a aprobación despues de corregirla:
    quien tiene acceso completo (crear_orden) puede reenviar cualquier
    orden (mismo criterio que el resto del flujo completo); quien solo
    tiene el perfil orden_simple unicamente puede reenviar la SUYA."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _puede_gestionar_orden_simple(orden):
        flash("Solo puedes reenviar las órdenes que tú mismo creaste.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    if orden.estado_aprobacion != "Sin Emitir":
        flash("Esta orden no está 'Sin Emitir'.", "warning")
        return redirect(_destino_detalle_orden(orden))
    orden.estado_aprobacion = "Por Aprobar"
    db.session.commit()
    flash(f"Orden {orden.numero_po} reenviada -- vuelve a quedar 'Por Aprobar'.", "success")
    return redirect(_destino_detalle_orden(orden))


@app.route("/ordenes/nueva", methods=["GET", "POST"])
@requiere_permiso("crear_orden")
def ordenes_nueva():
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    empresas = Empresa.query.filter_by(activo=True).order_by(Empresa.nombre).all()

    if request.method == "POST":
        proveedor_id = int(request.form["proveedor_id"])
        prov = Proveedor.query.get_or_404(proveedor_id)
        empresa_id = request.form.get("empresa_id") or None
        if not empresa_id:
            # Ronda U (2026-09-12, punto 3): una orden no se puede crear sin
            # Empresa compradora identificada (el <select> ya lo exige en el
            # HTML, pero se valida tambien acá por si llega sin ese campo).
            flash("Selecciona la Empresa compradora antes de crear la orden.", "danger")
            return redirect(url_for("ordenes_nueva", proveedor_id=proveedor_id))
        Empresa.query.get_or_404(int(empresa_id))

        orden = OrdenCompra(
            numero_po=siguiente_numero_po(int(empresa_id) if empresa_id else None),
            proveedor_id=prov.id,
            empresa_id=int(empresa_id) if empresa_id else None,
            fecha_emision=parse_date(request.form.get("fecha_emision")) or date.today(),
            moneda=request.form.get("moneda", prov.moneda_default or "USD"),
            estado=ETAPAS_LINEA[0],
            notas=request.form.get("notas", "").strip(),
            creado_por_usuario_id=current_user.id,
            estado_aprobacion="Por Aprobar",
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

        db.session.flush()
        # Ronda U (2026-09-12, punto 4): si se eligio del catalogo un
        # producto inactivo (el front-end lo muestra sombreado y pide
        # confirmar antes), se reactiva al quedar agregado a la orden.
        reactivados = 0
        for l in orden.lineas:
            if l.producto and not l.producto.activo:
                l.producto.activo = True
                reactivados += 1

        db.session.commit()
        mensaje = f"Orden {orden.numero_po} creada con {lineas_creadas} lineas. Queda 'Por Aprobar' -- no aparecera en el listado general hasta que alguien con permiso de Aprobacion la apruebe."
        if reactivados:
            mensaje += f" Se reactivaron {reactivados} producto(s) que estaban inactivos en el catálogo."
        flash(mensaje, "success")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    proveedor_id = request.args.get("proveedor_id", type=int)
    return render_template(
        "ordenes/form.html", proveedores=proveedores, empresas=empresas,
        proveedor_id=proveedor_id, now=date.today()
    )


@app.route("/ordenes/simple/nueva", methods=["GET", "POST"])
def ordenes_simple_nueva():
    """Creación de Orden 'simple' (ronda R, 2026-09-12, permiso
    orden_simple): pensada para un usuario que NO debe ver los precios ya
    cargados en el catálogo de ningún proveedor -- por eso este formulario
    nunca busca ni muestra el catálogo, solo pide escribir el código/
    descripción/cantidad y el precio que el proveedor le cotizó a ÉL para
    esta compra puntual. Si ese código YA existe en el catálogo de ese
    proveedor, se usa el producto y precio YA cargados (el precio tecleado
    acá se descarta sin mostrarlo) -- para no pisar el dato que mantiene
    quien sí tiene acceso a Creación de Orden completa. Si no existe, se
    crea un Producto nuevo en el catálogo de ese proveedor con el precio
    indicado. Ronda T (2026-09-12): la orden resultante nace 'Por Aprobar'
    (ver ESTADOS_APROBACION_OC en models.py) -- no aparece en el listado
    general de Compras/Órdenes hasta que alguien con permiso de Aprobación
    la apruebe desde la pestaña 'Por Aprobar'; si la devuelve, queda 'Sin
    Emitir' y solo la ve quien la creó, en 'Mis órdenes', para corregirla y
    reenviarla. Ronda S (2026-09-12, puntos 1 y 2): ahora puede adjuntar
    Cotización/Orden de Compra del proveedor desde este mismo formulario, y
    queda registrado como creador (creado_por_usuario_id) para que después
    pueda ver/editar/anular ESTA orden puntual desde "Mis órdenes" -- ver
    ordenes_simple_list/detalle mas abajo."""
    if not (current_user.tiene_permiso("orden_simple") or current_user.tiene_permiso("crear_orden")):
        flash("No tienes permiso para acceder a esta sección.", "danger")
        return redirect(url_for("dashboard"))

    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    empresas = Empresa.query.filter_by(activo=True).order_by(Empresa.nombre).all()

    if request.method == "POST":
        proveedor_id = request.form.get("proveedor_id") or None
        empresa_id = request.form.get("empresa_id") or None
        if not proveedor_id:
            flash("Selecciona un proveedor.", "warning")
            return redirect(url_for("ordenes_simple_nueva"))
        if not empresa_id:
            # Ronda U (2026-09-12, punto 3): una orden no se puede crear sin
            # Empresa compradora identificada.
            flash("Selecciona la Empresa compradora antes de crear la orden.", "danger")
            return redirect(url_for("ordenes_simple_nueva"))
        prov = Proveedor.query.get_or_404(int(proveedor_id))

        nota_sistema = (
            f"Orden creada por {current_user.nombre_completo} con el perfil "
            "'Creación de Orden Simple' (sin acceso al catálogo de precios)."
        )
        nota_usuario = request.form.get("notas", "").strip()
        if nota_usuario:
            nota_sistema += f"\nNota de {current_user.nombre_completo}: {nota_usuario}"

        orden = OrdenCompra(
            numero_po=siguiente_numero_po(int(empresa_id) if empresa_id else None),
            proveedor_id=prov.id,
            empresa_id=int(empresa_id) if empresa_id else None,
            fecha_emision=date.today(),
            moneda=prov.moneda_default or "USD",
            estado=ETAPAS_LINEA[0],
            notas=nota_sistema,
            creado_por_usuario_id=current_user.id,
            estado_aprobacion="Por Aprobar",
        )
        db.session.add(orden)
        db.session.flush()

        codigos = request.form.getlist("codigo")
        descripciones = request.form.getlist("descripcion")
        cantidades = request.form.getlist("cantidad_cajas")
        precios = request.form.getlist("precio_cotizado")
        fechas = request.form.getlist("fecha_estimada_despacho")

        lineas_creadas = 0
        productos_nuevos = 0
        for codigo, descripcion, cant, precio, fecha in zip(codigos, descripciones, cantidades, precios, fechas):
            codigo = (codigo or "").strip()
            descripcion = (descripcion or "").strip()
            cant_val = parse_int(cant, default=0)
            if not codigo or cant_val <= 0:
                continue

            producto = Producto.query.filter(
                Producto.proveedor_id == prov.id, db.func.lower(Producto.codigo) == codigo.lower()
            ).first()
            # Si el producto YA existia en el catalogo, el precio que se
            # use es el que YA tenia cargado (el tecleado aca se descarta) y
            # esta linea queda marcada para que su precio siga oculto para
            # este mismo usuario mas adelante (ver precio_catalogo_oculto en
            # models.py). Si es nuevo, el precio es el que el tecleo -- no
            # hace falta ocultarselo a si mismo.
            precio_venia_de_catalogo = producto is not None
            producto_nuevo_esta_linea = not precio_venia_de_catalogo
            if not producto:
                precio_caja = float(precio) if precio else 0
                producto = Producto(
                    proveedor_id=prov.id,
                    codigo=codigo,
                    descripcion=descripcion or codigo,
                    empaque=1,
                    moneda=prov.moneda_default or "USD",
                    precio_caja=precio_caja,
                    precio_unitario=precio_caja,
                    # Ronda U (2026-09-12, punto 1): nace INACTIVO -- recien
                    # se activa (se suma de verdad al catalogo) si esta orden
                    # llega a Aprobarse (ver ordenes_aprobar). Asi una orden
                    # que nunca se aprueba no deja "basura" visible en el
                    # catalogo del proveedor.
                    activo=False,
                )
                db.session.add(producto)
                db.session.flush()
                productos_nuevos += 1

            linea = OrdenCompraLinea(
                orden_id=orden.id,
                producto_id=producto.id,
                cantidad_cajas=cant_val,
                precio_unitario_pactado=producto.precio_caja,
                fecha_estimada_despacho=parse_date(fecha),
                etapa=ETAPAS_LINEA[0],
                precio_catalogo_oculto=precio_venia_de_catalogo,
                producto_creado_por_esta_orden=producto_nuevo_esta_linea,
            )
            db.session.add(linea)
            lineas_creadas += 1

        if lineas_creadas == 0:
            db.session.rollback()
            flash("Agrega al menos un producto con código y cantidad mayor a 0.", "danger")
            return redirect(url_for("ordenes_simple_nueva"))

        # Documentos opcionales adjuntados desde el mismo formulario (ronda
        # S, 2026-09-12, punto 1) -- para que quien aprueba pueda corroborar
        # por que se esta pidiendo esta compra.
        docs_adjuntados = 0
        for campo, tipo_doc in (
            ("archivo_cotizacion", "Cotización"),
            ("archivo_orden_compra", "Orden de Compra (proveedor)"),
        ):
            archivo = request.files.get(campo)
            if archivo and archivo.filename:
                if _extension_valida(archivo.filename):
                    if _guardar_documento_orden(orden, archivo, tipo_doc):
                        docs_adjuntados += 1
                else:
                    flash(f"'{archivo.filename}' no se adjuntó: solo se permiten PDF o imágenes (JPG, PNG).", "warning")

        db.session.commit()
        mensaje = (
            f"Orden {orden.numero_po} creada para {prov.nombre} con {lineas_creadas} producto(s) "
            f"({productos_nuevos} nuevo(s) en el catálogo)."
        )
        if docs_adjuntados:
            mensaje += f" Se adjuntaron {docs_adjuntados} documento(s)."
        mensaje += (
            " Queda 'Por Aprobar': no aparecerá en el listado general hasta que alguien con permiso "
            "de Aprobación la apruebe. Mientras tanto la puedes ver, editar o anular desde "
            "'Mis órdenes'."
        )
        if productos_nuevos:
            mensaje += (
                f" Los {productos_nuevos} producto(s) nuevo(s) quedan INACTIVOS en el catálogo de "
                f"{prov.nombre} hasta que la orden se apruebe -- así no se llena el catálogo con "
                "datos sin confirmar."
            )
        flash(mensaje, "success")
        return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))

    return render_template(
        "ordenes/simple_form.html", proveedores=proveedores, empresas=empresas,
        tipos_documento=TIPOS_DOCUMENTO_ORDEN,
    )


def _resumen_precio_para_simple(orden):
    """Ronda S (2026-09-12): total 'seguro' de mostrarle a quien tiene el
    perfil orden_simple sobre SU PROPIA orden. Si se mostrara el total real
    de la orden completa, alguien podria deducir el precio oculto de una
    linea de catalogo restando el resto de subtotales (que si conoce,
    porque el escribio cantidad y a veces precio) -- por eso el total que
    se muestra excluye las lineas con precio oculto, y se avisa aparte."""
    lineas_activas = [l for l in orden.lineas if not l.anulada]
    tiene_ocultos = any(l.precio_catalogo_oculto for l in lineas_activas)
    total_visible = sum(l.subtotal for l in lineas_activas if not l.precio_catalogo_oculto)
    return total_visible, tiene_ocultos


@app.route("/ordenes/simple")
@requiere_permiso("orden_simple")
def ordenes_simple_list():
    """'Mis órdenes' (ronda S, 2026-09-12, punto 2): a diferencia del
    listado completo de /ordenes (que requiere crear_orden/aprobar_orden),
    este perfil solo ve las ordenes que EL MISMO creo -- nunca las de otro
    usuario ni el listado global."""
    ordenes = (
        OrdenCompra.query.filter_by(creado_por_usuario_id=current_user.id)
        .order_by(OrdenCompra.id.desc())
        .all()
    )
    for o in ordenes:
        o.total_visible, o.tiene_precios_ocultos = _resumen_precio_para_simple(o)
    return render_template("ordenes/simple_list.html", ordenes=ordenes)


@app.route("/ordenes/simple/<int:orden_id>")
@requiere_permiso("orden_simple")
def ordenes_simple_detalle(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.creado_por_usuario_id != current_user.id:
        flash("Solo puedes ver las órdenes que tú mismo creaste con este perfil.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    lineas = orden.lineas.order_by(OrdenCompraLinea.id).all()
    documentos = orden.documentos.order_by(OrdenDocumento.fecha_subida.desc()).all()
    total_visible, tiene_precios_ocultos = _resumen_precio_para_simple(orden)
    return render_template(
        "ordenes/simple_detalle.html",
        orden=orden, lineas=lineas, documentos=documentos, tipos_documento=TIPOS_DOCUMENTO_ORDEN,
        total_visible=total_visible, tiene_precios_ocultos=tiene_precios_ocultos,
    )


@app.route("/ordenes/simple/<int:orden_id>/lineas/nueva", methods=["POST"])
@requiere_permiso("orden_simple")
def ordenes_simple_linea_nueva(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.creado_por_usuario_id != current_user.id:
        flash("Solo puedes editar las órdenes que tú mismo creaste con este perfil.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    if orden.estado in ("Cancelada", "Anulada"):
        flash("No se pueden agregar productos a una orden cancelada o anulada.", "warning")
        return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))

    codigo = (request.form.get("codigo") or "").strip()
    descripcion = (request.form.get("descripcion") or "").strip()
    cant_val = parse_int(request.form.get("cantidad_cajas"), default=0)
    precio = request.form.get("precio_cotizado")
    fecha = request.form.get("fecha_estimada_despacho")
    if not codigo or cant_val <= 0:
        flash("Ingresa un código y una cantidad mayor a 0.", "warning")
        return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))

    producto = Producto.query.filter(
        Producto.proveedor_id == orden.proveedor_id, db.func.lower(Producto.codigo) == codigo.lower()
    ).first()
    precio_venia_de_catalogo = producto is not None
    producto_nuevo_esta_linea = not precio_venia_de_catalogo
    # Ronda U (2026-09-12, punto 1): si la orden todavia NO esta Aprobada, un
    # producto nuevo nace inactivo hasta que se apruebe (ver ordenes_aprobar).
    # Si la orden ya esta Aprobada (esta linea se agrega despues), no hay
    # ningun paso de aprobacion pendiente que lo vaya a activar mas
    # adelante -- nace activo directamente, igual que antes de esta ronda.
    pendiente_aprobacion = orden.estado_aprobacion in ("Por Aprobar", "Sin Emitir")
    if not producto:
        precio_caja = float(precio) if precio else 0
        producto = Producto(
            proveedor_id=orden.proveedor_id,
            codigo=codigo,
            descripcion=descripcion or codigo,
            empaque=1,
            moneda=orden.proveedor.moneda_default or "USD",
            precio_caja=precio_caja,
            precio_unitario=precio_caja,
            activo=not pendiente_aprobacion,
        )
        db.session.add(producto)
        db.session.flush()

    linea = OrdenCompraLinea(
        orden_id=orden.id,
        producto_id=producto.id,
        cantidad_cajas=cant_val,
        precio_unitario_pactado=producto.precio_caja,
        fecha_estimada_despacho=parse_date(fecha),
        etapa=ETAPAS_LINEA[0],
        precio_catalogo_oculto=precio_venia_de_catalogo,
        producto_creado_por_esta_orden=producto_nuevo_esta_linea and pendiente_aprobacion,
    )
    db.session.add(linea)
    db.session.flush()
    recalcular_estado_orden(orden)
    db.session.commit()
    mensaje = f"Producto agregado a la orden {orden.numero_po}."
    if producto_nuevo_esta_linea and pendiente_aprobacion:
        mensaje += " Queda INACTIVO en el catálogo hasta que la orden se apruebe."
    flash(mensaje, "success")
    return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))


@app.route("/ordenes/simple/<int:orden_id>/lineas/<int:linea_id>/actualizar", methods=["POST"])
@requiere_permiso("orden_simple")
def ordenes_simple_linea_actualizar(orden_id, linea_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.creado_por_usuario_id != current_user.id:
        flash("Solo puedes editar las órdenes que tú mismo creaste con este perfil.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    if linea.orden_id != orden.id:
        abort(404)
    if linea.bloqueada:
        flash("Esta línea ya fue despachada: no se puede editar.", "warning")
        return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))

    linea.cantidad_cajas = parse_int(request.form.get("cantidad_cajas"), default=linea.cantidad_cajas)
    linea.fecha_estimada_despacho = (
        parse_date(request.form.get("fecha_estimada_despacho")) or linea.fecha_estimada_despacho
    )
    # El precio solo se puede tocar si NO vino del catalogo -- si vino del
    # catalogo (precio_catalogo_oculto=True) sigue oculto e intocable desde
    # este perfil, ni siquiera para "corregirlo".
    if not linea.precio_catalogo_oculto:
        nuevo_precio = request.form.get("precio_cotizado")
        if nuevo_precio:
            nuevo_precio_val = float(nuevo_precio)
            linea.precio_unitario_pactado = nuevo_precio_val
            linea.producto.precio_caja = nuevo_precio_val
            linea.producto.precio_unitario = nuevo_precio_val
    db.session.commit()
    flash("Línea actualizada.", "success")
    return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))


@app.route("/ordenes/simple/<int:orden_id>/lineas/<int:linea_id>/anular", methods=["POST"])
@requiere_permiso("orden_simple")
def ordenes_simple_linea_anular(orden_id, linea_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.creado_por_usuario_id != current_user.id:
        flash("Solo puedes editar las órdenes que tú mismo creaste con este perfil.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    if linea.orden_id != orden.id:
        abort(404)
    if linea.bloqueada:
        flash("Esta línea ya fue despachada: no se puede anular.", "warning")
    elif linea.etapa != ETAPAS_LINEA[0]:
        flash(
            "Esta línea ya fue confirmada por quien aprueba: pídele a un usuario con permiso de "
            "Aprobación o Creación de Orden que la anule.",
            "warning",
        )
    else:
        linea.anulada = True
        db.session.flush()
        recalcular_estado_orden(orden)
        db.session.commit()
        flash("Línea anulada.", "success")
    return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))


@app.route("/ordenes/simple/<int:orden_id>/cancelar", methods=["POST"])
@requiere_permiso("orden_simple")
def ordenes_simple_cancelar(orden_id):
    """Anular la orden completa (ronda S, 2026-09-12, punto 2) -- solo
    mientras NADA se haya confirmado todavia (orden.estado sigue en
    'Emisión de Orden'). Una vez que alguien con permiso de Aprobación
    empezo a procesarla, este perfil ya no puede revertir esa decision por
    su cuenta -- tiene que pedirselo a quien aprueba/crea ordenes."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    if orden.creado_por_usuario_id != current_user.id:
        flash("Solo puedes anular las órdenes que tú mismo creaste con este perfil.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    if orden.estado != ETAPAS_LINEA[0]:
        flash(
            "Esta orden ya tiene líneas confirmadas: pídele a un usuario con permiso de Aprobación "
            "o Creación de Orden que la anule.",
            "warning",
        )
        return redirect(url_for("ordenes_simple_detalle", orden_id=orden.id))
    numero_po = orden.numero_po
    db.session.delete(orden)
    db.session.commit()
    _borrar_carpeta_documentos_orden(orden_id)
    flash(f"Orden {numero_po} anulada.", "warning")
    return redirect(url_for("ordenes_simple_list"))


@app.route("/api/proveedores/<int:proveedor_id>/productos")
@requiere_permiso("crear_orden", "generar_costeo")
def api_productos_por_proveedor(proveedor_id):
    """Ronda U (2026-09-12, punto 4): antes solo devolvía productos activos
    -- uno inactivo simplemente no aparecía en la búsqueda del catálogo al
    armar una orden. Ahora se incluyen también los inactivos (con
    "activo": false) para que el front-end los muestre sombreados y pida
    confirmación antes de agregarlos -- ver ordenes/form.html y
    ordenes/detalle.html. Si se selecciona uno, el backend lo reactiva
    automáticamente al agregar la línea (ver ordenes_nueva/
    ordenes_linea_nueva)."""
    q = request.args.get("q", "").strip()
    query = Producto.query.filter_by(proveedor_id=proveedor_id)
    if q:
        query = query.filter(
            db.or_(Producto.codigo.ilike(f"%{q}%"), Producto.descripcion.ilike(f"%{q}%"))
        )
    productos = query.order_by(Producto.activo.desc(), Producto.codigo).limit(50).all()
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
                "activo": p.activo,
            }
            for p in productos
        ]
    )


@app.route("/api/productos/<int:producto_id>/variantes")
@requiere_permiso("crear_orden", "generar_costeo")
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
@requiere_permiso("crear_orden", "aprobar_orden")
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
@requiere_permiso("crear_orden")
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
    nuevo_empresa_id = int(empresa_id) if empresa_id else None

    # Ronda S (2026-09-12): el N° de PO tiene que ser unico POR EMPRESA
    # compradora -- si no se valida esto, dos ordenes SIN relacion pueden
    # terminar con el mismo numero (por coincidencia al tipearlo a mano) y
    # el sistema las trata como si fueran partes de una misma division
    # automatica (ver dividir_orden_si_corresponde), lo que confunde a
    # quien aprueba. Solo se valida cuando el numero CAMBIA -- una orden ya
    # dividida (que comparte numero de PO a proposito con sus hermanas) se
    # puede volver a guardar sin tocar nada mas sin que esto la bloquee.
    if nuevo_empresa_id is not None and nuevo_numero.lower() != (orden.numero_po or "").lower():
        conflicto = OrdenCompra.query.filter(
            OrdenCompra.id != orden.id,
            OrdenCompra.empresa_id == nuevo_empresa_id,
            db.func.lower(OrdenCompra.numero_po) == nuevo_numero.lower(),
        ).first()
        if conflicto:
            flash(
                f"Ya existe otra orden con el número '{nuevo_numero}' para esta empresa compradora "
                f"(proveedor {conflicto.proveedor.nombre}, estado '{conflicto.estado}') -- el número de "
                "PO no puede repetirse dentro de la misma empresa. Usa uno distinto.",
                "danger",
            )
            return redirect(url_for("ordenes_detalle", orden_id=orden.id))

    orden.numero_po = nuevo_numero
    orden.empresa_id = nuevo_empresa_id
    orden.fecha_emision = parse_date(request.form.get("fecha_emision")) or orden.fecha_emision
    orden.moneda = request.form.get("moneda", orden.moneda)
    orden.notas = request.form.get("notas", "").strip()
    db.session.commit()
    flash(f"Orden actualizada: {orden.numero_po}.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/cancelar", methods=["POST"])
@requiere_permiso("crear_orden")
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
@requiere_permiso("crear_orden")
def ordenes_reactivar(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    orden.estado = "Emisión de Orden"  # se recalcula abajo segun las lineas
    recalcular_estado_orden(orden)
    db.session.commit()
    flash(f"Orden {orden.numero_po} reactivada.", "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/lineas/nueva", methods=["POST"])
@requiere_permiso("crear_orden")
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
    # Ronda U (2026-09-12, punto 4): el catálogo ahora tambien deja elegir un
    # producto inactivo (el front-end lo muestra sombreado y pide
    # confirmacion antes) -- si se confirmo y se llego hasta aca, se
    # reactiva de una.
    reactivado = False
    if linea.producto and not linea.producto.activo:
        linea.producto.activo = True
        reactivado = True
    recalcular_estado_orden(orden)
    # Si la orden ya tenia lineas mas avanzadas (ej. despachada) esta linea
    # nueva en "Emision de Orden" quedaria mezclando cohortes -- se separa
    # igual que al confirmar/despachar.
    nuevas_ordenes = dividir_orden_si_corresponde(orden)
    db.session.commit()
    mensaje = f"'{linea.codigo_mostrar}' agregado a la orden.{_mensaje_division(nuevas_ordenes)}"
    if reactivado:
        mensaje += " Estaba inactivo en el catálogo -- se reactivó."
    flash(mensaje, "success")
    return redirect(url_for("ordenes_detalle", orden_id=orden.id))


@app.route("/ordenes/<int:orden_id>/linea/<int:linea_id>/actualizar", methods=["POST"])
@requiere_permiso("crear_orden")
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
@requiere_permiso("crear_orden")
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
@requiere_permiso("crear_orden")
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
@requiere_permiso("aprobar_orden")
def ordenes_linea_confirmar(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    orden = linea.orden
    if not _orden_aprobada(orden):
        flash("Esta orden todavía no está Aprobada -- apruébala primero.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
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
@requiere_permiso("aprobar_orden")
def ordenes_lineas_confirmar_masivo(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _orden_aprobada(orden):
        flash("Esta orden todavía no está Aprobada -- apruébala primero.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))
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
@requiere_permiso("aprobar_orden")
def ordenes_linea_retroceder(orden_id, linea_id):
    linea = OrdenCompraLinea.query.get_or_404(linea_id)
    orden = linea.orden
    if not _orden_aprobada(orden):
        flash("Esta orden todavía no está Aprobada.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden_id))
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
@requiere_permiso("aprobar_orden")
def ordenes_lineas_retroceder_masivo(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _orden_aprobada(orden):
        flash("Esta orden todavía no está Aprobada.", "warning")
        return redirect(url_for("ordenes_detalle", orden_id=orden.id))
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
@requiere_permiso("crear_orden")
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
@requiere_permiso("crear_orden")
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
@requiere_permiso("crear_orden")
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


def _guardar_documento_orden(orden, archivo, tipo):
    """Guarda un archivo adjunto de una orden en disco + su fila
    OrdenDocumento (sin hacer commit -- lo hace el caller). Devuelve None si
    no hay archivo o la extension no es valida, para que el caller decida
    que flash mostrar. Reutilizada por la subida manual (crear_orden/
    orden_simple) y por la creacion de una orden 'simple' con Cotización/
    Orden de Compra adjuntas desde el mismo formulario (ronda S,
    2026-09-12)."""
    if not archivo or not archivo.filename:
        return None
    if not _extension_valida(archivo.filename):
        return None
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
    return doc


@app.route("/ordenes/<int:orden_id>/documentos/subir", methods=["POST"])
@requiere_permiso("crear_orden", "orden_simple")
def ordenes_documento_subir(orden_id):
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _puede_gestionar_orden_simple(orden):
        flash("Solo puedes subir documentos a las órdenes que tú mismo creaste.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    archivo = request.files.get("archivo")
    tipo = request.form.get("tipo", "Otro")

    if not archivo or archivo.filename == "":
        flash("Selecciona un archivo para subir.", "danger")
        return redirect(_destino_detalle_orden(orden))

    if not _extension_valida(archivo.filename):
        flash("Solo se permiten archivos PDF o imagenes (JPG, PNG).", "danger")
        return redirect(_destino_detalle_orden(orden))

    _guardar_documento_orden(orden, archivo, tipo)
    db.session.commit()
    flash(f"Documento '{archivo.filename}' subido correctamente.", "success")
    return redirect(_destino_detalle_orden(orden))


@app.route("/ordenes/<int:orden_id>/documentos/<int:doc_id>/ver")
@requiere_permiso("crear_orden", "aprobar_orden", "orden_simple")
def ordenes_documento_ver(orden_id, doc_id):
    """Abre el documento en el navegador (PDF/imagen inline) en vez de forzar
    la descarga -- para que el usuario pueda visualizarlo con un clic."""
    doc = OrdenDocumento.query.get_or_404(doc_id)
    if doc.orden_id != orden_id:
        abort(404)
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _puede_gestionar_orden_simple(orden):
        flash("No tienes acceso a los documentos de esta orden.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(doc.orden_id))
    return send_from_directory(
        carpeta_orden, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=False
    )


@app.route("/ordenes/<int:orden_id>/documentos/<int:doc_id>/descargar")
@requiere_permiso("crear_orden", "aprobar_orden", "orden_simple")
def ordenes_documento_descargar(orden_id, doc_id):
    doc = OrdenDocumento.query.get_or_404(doc_id)
    if doc.orden_id != orden_id:
        abort(404)
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _puede_gestionar_orden_simple(orden):
        flash("No tienes acceso a los documentos de esta orden.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(doc.orden_id))
    return send_from_directory(
        carpeta_orden, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=True
    )


@app.route("/ordenes/<int:orden_id>/documentos/<int:doc_id>/eliminar", methods=["POST"])
@requiere_permiso("crear_orden", "orden_simple")
def ordenes_documento_eliminar(orden_id, doc_id):
    doc = OrdenDocumento.query.get_or_404(doc_id)
    if doc.orden_id != orden_id:
        abort(404)
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _puede_gestionar_orden_simple(orden):
        flash("Solo puedes eliminar documentos de las órdenes que tú mismo creaste.", "danger")
        return redirect(url_for("ordenes_simple_list"))
    carpeta_orden = os.path.join(DOCUMENTOS_DIR, str(doc.orden_id))
    ruta = os.path.join(carpeta_orden, doc.nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)
    db.session.delete(doc)
    db.session.commit()
    flash("Documento eliminado.", "success")
    return redirect(_destino_detalle_orden(orden))


@app.route("/ordenes/<int:orden_id>/nota", methods=["POST"])
@requiere_permiso("crear_orden", "aprobar_orden", "orden_simple")
def ordenes_nota_agregar(orden_id):
    """Nota rapida sobre una orden (ronda S, 2026-09-12) -- pensada sobre
    todo para quien tiene el permiso de Aprobación de Orden: a diferencia
    del modal "Editar" (que puede cambiar N° de PO/moneda/empresa y esta
    reservado a crear_orden), esto solo AGREGA texto al final de las notas
    existentes, con fecha y autor, sin tocar nada mas -- sirve para dejar
    constancia de por que se aprobo o anulo algo. Tambien la puede usar
    quien creo la orden con el perfil orden_simple, sobre sus propias
    ordenes."""
    orden = OrdenCompra.query.get_or_404(orden_id)
    if not _puede_gestionar_orden_simple(orden):
        flash("No tienes acceso a esta orden.", "danger")
        return redirect(url_for("dashboard"))
    texto = request.form.get("texto", "").strip()
    if not texto:
        flash("Escribe algo antes de agregar la nota.", "warning")
    else:
        sello = f"[{datetime.utcnow().strftime('%Y-%m-%d %H:%M')} · {current_user.nombre_completo}]"
        linea_nueva = f"{sello} {texto}"
        orden.notas = f"{orden.notas}\n{linea_nueva}" if orden.notas else linea_nueva
        db.session.commit()
        flash("Nota agregada.", "success")
    return redirect(_destino_detalle_orden(orden))


@app.route("/ordenes/<int:orden_id>/imprimir")
@requiere_permiso("crear_orden", "aprobar_orden")
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
@requiere_permiso("crear_orden", "aprobar_orden")
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
@requiere_permiso("crear_orden")
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
@requiere_permiso("crear_orden")
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
@requiere_permiso("actualizar_despacho", "seguimiento", "generar_costeo")
def despachos_list():
    despachos = Despacho.query.order_by(Despacho.id.desc()).all()
    hay_candidatas = bool(_ordenes_listas_para_despachar())
    return render_template("despachos/list.html", despachos=despachos, hay_candidatas=hay_candidatas)


@app.route("/despachos/nuevo", methods=["GET", "POST"])
@requiere_permiso("actualizar_despacho")
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
@requiere_permiso("actualizar_despacho", "seguimiento", "generar_costeo")
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
@requiere_permiso("actualizar_despacho")
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
@requiere_permiso("actualizar_despacho")
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
@requiere_permiso("actualizar_despacho")
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
@requiere_permiso("actualizar_despacho")
def despachos_aduana_masivo(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    marcadas = _avanzar_lineas_despacho(
        despacho, "Orden Despachada", "Internación Aduanas", "fecha_llegada_aduana"
    )
    db.session.commit()
    flash(f"{marcadas} línea(s), de todas las órdenes de este despacho, marcadas en internación de aduanas.", "success")
    return redirect(url_for("despachos_detalle", despacho_id=despacho.id))


@app.route("/despachos/<int:despacho_id>/recibido-masivo", methods=["POST"])
@requiere_permiso("actualizar_despacho")
def despachos_recibido_masivo(despacho_id):
    despacho = Despacho.query.get_or_404(despacho_id)
    marcadas = _avanzar_lineas_despacho(
        despacho, "Internación Aduanas", "Recibido", "fecha_recepcion_bodega"
    )
    db.session.commit()
    flash(f"{marcadas} línea(s), de todas las órdenes de este despacho, marcadas como recibidas en bodega.", "success")
    return redirect(url_for("despachos_detalle", despacho_id=despacho.id))


@app.route("/despachos/<int:despacho_id>/eliminar", methods=["POST"])
@requiere_permiso("actualizar_despacho")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo", "reportes")
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
@requiere_permiso("generar_costeo", "reportes")
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
@requiere_permiso("generar_costeo")
def importaciones_nueva():
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    empresas = Empresa.query.filter_by(activo=True).order_by(Empresa.nombre).all()
    if request.method == "POST":
        empresa_id = request.form.get("empresa_id") or None
        imp = Importacion(
            proveedor_id=int(request.form["proveedor_id"]),
            empresa_id=int(empresa_id) if empresa_id else None,
            numero_correlativo_inventario=parse_int(request.form.get("numero_correlativo_inventario"), default=None),
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
        "importaciones/form.html", proveedores=proveedores, empresas=empresas,
        condiciones=CONDICIONES_COMPRA, tc_mes=tc_mes,
    )


@app.route("/importaciones/<int:importacion_id>")
@requiere_permiso("generar_costeo")
def importaciones_detalle(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    resultado = costing.calcular_costeo(imp)
    tipos_cambio = TipoCambioMensual.query.order_by(
        TipoCambioMensual.anio.desc(), TipoCambioMensual.mes.desc()
    ).all()
    empresas = Empresa.query.filter_by(activo=True).order_by(Empresa.nombre).all()
    # Sugerencia de correlativo de Inventario (ronda O) para la empresa YA
    # asignada a esta importacion -- solo una sugerencia, nunca se aplica
    # sola; si la importacion ya tiene uno guardado, ese es el que manda.
    sugerido_correlativo_inventario = (
        siguiente_correlativo_inventario(imp.empresa_id) if imp.empresa_id else None
    )
    return render_template(
        "importaciones/detalle.html", imp=imp, resultado=resultado,
        regimenes=REGIMENES_PARCIAL, vias=VIAS_EMBARQUE,
        condiciones=CONDICIONES_COMPRA, conceptos_gasto=CONCEPTOS_GASTO,
        conceptos_item_factura=CONCEPTOS_ITEM_FACTURA, tipos_cambio=tipos_cambio,
        tipos_documento_gasto=TIPOS_DOCUMENTO_GASTO, empresas=empresas,
        sugerido_correlativo_inventario=sugerido_correlativo_inventario,
    )


@app.route("/importaciones/<int:importacion_id>/editar", methods=["POST"])
@requiere_permiso("generar_costeo")
def importaciones_editar(importacion_id):
    imp = Importacion.query.get_or_404(importacion_id)
    empresa_id = request.form.get("empresa_id") or None
    imp.empresa_id = int(empresa_id) if empresa_id else None
    imp.numero_correlativo_inventario = parse_int(request.form.get("numero_correlativo_inventario"), default=None)
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
def parcial_lineas_eliminar(linea_id):
    linea = ParcialLinea.query.get_or_404(linea_id)
    importacion_id = linea.parcial.importacion_id
    db.session.delete(linea)
    db.session.commit()
    flash("Línea eliminada.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


@app.route("/parcial_lineas/mover-masivo", methods=["POST"])
@requiere_permiso("generar_costeo")
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


def _normalizar_desc(texto):
    """Normaliza texto para comparar descripciones de producto sin que
    espacios de más, espacios "duros" (\\xa0, comunes en descripciones
    copiadas desde Excel) o mayúsculas/minúsculas generen un falso
    'no encontrado' al cruzar el archivo de lotes contra el catálogo."""
    if texto is None:
        return ""
    texto = str(texto).replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto).strip().lower()
    return texto


def _buscar_columna_prioridad(encabezado, *listas_claves):
    """Busca una columna probando listas de claves en orden de prioridad
    (todas las columnas contra la lista más específica antes de pasar a la
    siguiente) -- evita que, por ejemplo, una columna 'Fecha Invoice' se
    confunda con la columna de vencimiento solo porque ambas contienen la
    palabra genérica 'fecha'."""
    for claves in listas_claves:
        for i, nombre_col in enumerate(encabezado):
            for clave in claves:
                if clave in nombre_col:
                    return i
    return None


def _leer_filas_archivo_lotes(file_storage):
    """Lee un archivo .xlsx/.xls o .csv de carga masiva de lotes para TODO
    el costeo (ronda N, 2026-09-12 -- reemplaza la versión por parcial de
    la ronda J). Encabezados flexibles: identifica el producto por
    'Descripción' (formato real que usa el usuario, ver 'Carga de Lotes
    Costeo.xlsx') o por 'Código' si está presente; 'Cod. de Lote'/'Lote';
    'Vencmto'/'Fecha vencimiento'; y opcionalmente 'Cantidad' (si el
    archivo ya trae la cantidad de cada lote directamente, como en el
    formato real del usuario) -- si no hay columna de cantidad, sigue
    soportando el formato viejo de UNA FILA POR UNIDAD (se cuenta 1 por
    fila y se agrupan). Devuelve una lista de dicts
    {codigo, descripcion, lote, fecha, cantidad}, o None si no se
    reconocieron los encabezados mínimos (falta el producto o el lote)."""
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

    idx_codigo = _buscar_columna_prioridad(
        encabezado, ["codigo de producto", "código de producto"], ["codigo", "código"]
    )
    idx_descripcion = _buscar_columna_prioridad(encabezado, ["descripcion", "descripción"])
    idx_lote = _buscar_columna_prioridad(
        encabezado, ["cod.de lote", "cod. de lote", "codigo de lote", "código de lote"], ["lote"]
    )
    idx_fecha = _buscar_columna_prioridad(
        encabezado,
        ["vencmto", "fecha vencimiento", "fecha de vencimiento"],
        ["vencimiento"],
        ["fecha"],
    )
    idx_cantidad = _buscar_columna_prioridad(encabezado, ["cantidad"], ["unidades"])

    if idx_lote is None or (idx_codigo is None and idx_descripcion is None):
        return None

    def _valor_lote(crudo):
        if crudo is None:
            return ""
        if isinstance(crudo, float) and crudo.is_integer():
            return str(int(crudo))
        return str(crudo).strip()

    filas = []
    for fila in filas_crudas[1:]:
        codigo = ""
        if idx_codigo is not None and idx_codigo < len(fila):
            codigo = str(fila[idx_codigo] or "").strip()
        descripcion = ""
        if idx_descripcion is not None and idx_descripcion < len(fila):
            descripcion = str(fila[idx_descripcion] or "").strip()
        if not codigo and not descripcion:
            continue
        lote = _valor_lote(fila[idx_lote]) if idx_lote < len(fila) else ""
        fecha_cruda = fila[idx_fecha] if (idx_fecha is not None and idx_fecha < len(fila)) else None
        cantidad = None
        if idx_cantidad is not None and idx_cantidad < len(fila):
            crudo_cantidad = fila[idx_cantidad]
            if crudo_cantidad not in (None, ""):
                try:
                    cantidad = int(float(crudo_cantidad))
                except (TypeError, ValueError):
                    cantidad = None
        filas.append({
            "codigo": codigo,
            "descripcion": descripcion,
            "lote": lote,
            "fecha": _parsear_fecha_archivo_lotes(fecha_cruda),
            "cantidad": cantidad,
        })
    return filas


def _indexar_lineas_importacion(importacion):
    """Arma diccionarios (por código normalizado y por descripción
    normalizada) que apuntan a la ParcialLinea correspondiente, buscando en
    TODOS los parciales de la importación -- para la carga masiva de lotes
    a nivel de todo el costeo (un mismo producto no se repite entre
    parciales distintos del mismo costeo, confirmado por el usuario). Si
    una clave aparece en más de una línea (dato real inesperado), se marca
    como ambigua y se excluye de los diccionarios en vez de adivinar a
    cuál asignarla."""
    por_codigo, por_descripcion = {}, {}
    vistos_codigo, vistos_descripcion = set(), set()
    ambiguos_codigo, ambiguos_descripcion = set(), set()
    for parcial in importacion.parciales:
        for linea in parcial.lineas:
            clave_cod = (linea.codigo or "").strip().lower()
            if clave_cod:
                if clave_cod in vistos_codigo:
                    ambiguos_codigo.add(clave_cod)
                vistos_codigo.add(clave_cod)
                por_codigo[clave_cod] = linea
            clave_desc = _normalizar_desc(linea.descripcion)
            if clave_desc:
                if clave_desc in vistos_descripcion:
                    ambiguos_descripcion.add(clave_desc)
                vistos_descripcion.add(clave_desc)
                por_descripcion[clave_desc] = linea
    for clave in ambiguos_codigo:
        por_codigo.pop(clave, None)
    for clave in ambiguos_descripcion:
        por_descripcion.pop(clave, None)
    return por_codigo, por_descripcion, ambiguos_codigo, ambiguos_descripcion


@app.route("/importaciones/<int:importacion_id>/lotes/cargar", methods=["POST"])
@requiere_permiso("generar_costeo")
def importacion_lotes_cargar(importacion_id):
    """Carga masiva de lotes para TODO el costeo de una sola vez (ronda N,
    2026-09-12, a pedido del usuario -- reemplaza la carga por parcial de
    la ronda J). Se sube un solo archivo con los productos del embarque
    completo, sin importar el orden de las filas ni en qué parcial esté
    cada uno: el sistema busca cada producto (por Código si el archivo lo
    trae, si no por Descripción) entre TODAS las líneas de TODOS los
    parciales de esta importación y le carga los lotes al que corresponda,
    ya que un mismo producto no se repite en dos parciales del mismo
    costeo."""
    imp = Importacion.query.get_or_404(importacion_id)
    archivo = request.files.get("archivo_lotes")
    if not archivo or not archivo.filename:
        flash("Selecciona un archivo para cargar los lotes.", "warning")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))

    extension = os.path.splitext(archivo.filename)[1].lower()
    if extension not in (".xlsx", ".xls", ".csv"):
        flash("Formato no soportado. Sube un archivo .xlsx o .csv.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))

    try:
        filas = _leer_filas_archivo_lotes(archivo)
    except Exception:
        flash("No se pudo leer el archivo. Verifica que no esté dañado o abierto en otro programa.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))

    if filas is None:
        flash("No se reconocieron las columnas del archivo. Debe incluir 'Descripción' (o 'Código') y 'Lote'.", "danger")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))
    if not filas:
        flash("El archivo no tiene filas para cargar.", "warning")
        return redirect(url_for("importaciones_detalle", importacion_id=imp.id))

    por_codigo, por_descripcion, ambiguos_codigo, ambiguos_descripcion = _indexar_lineas_importacion(imp)

    conteo_por_linea = defaultdict(lambda: defaultdict(int))
    no_encontrados = set()
    ambiguos_encontrados = set()

    for fila in filas:
        codigo_key = fila["codigo"].strip().lower() if fila["codigo"] else ""
        desc_key = _normalizar_desc(fila["descripcion"]) if fila["descripcion"] else ""
        linea = None
        if codigo_key:
            linea = por_codigo.get(codigo_key)
            if linea is None and codigo_key in ambiguos_codigo:
                ambiguos_encontrados.add(fila["codigo"])
        if linea is None and desc_key:
            linea = por_descripcion.get(desc_key)
            if linea is None and desc_key in ambiguos_descripcion:
                ambiguos_encontrados.add(fila["descripcion"])
        if linea is None:
            if not (codigo_key and codigo_key in ambiguos_codigo) and not (desc_key and desc_key in ambiguos_descripcion):
                no_encontrados.add(fila["codigo"] or fila["descripcion"])
            continue
        cantidad = fila["cantidad"] if fila["cantidad"] is not None else 1
        clave_lote = (fila["lote"], fila["fecha"])
        conteo_por_linea[linea.id][clave_lote] += cantidad

    lineas_por_parcial = {}
    for linea_id, grupos in conteo_por_linea.items():
        linea = ParcialLinea.query.get(linea_id)
        nuevos = [
            ParcialLineaLote(codigo_lote=lote, fecha_vencimiento=fecha, cantidad_unidades=cantidad)
            for (lote, fecha), cantidad in grupos.items()
        ]
        linea.lotes = nuevos
        primero = min(nuevos, key=lambda lo: (lo.fecha_vencimiento is None, lo.fecha_vencimiento or date.max))
        linea.codigo_lote = ", ".join(sorted({lo.codigo_lote for lo in nuevos if lo.codigo_lote}))
        linea.fecha_vencimiento = primero.fecha_vencimiento
        nombre_parcial = linea.parcial.referencia or linea.parcial.numero_parcial or f"Parcial {linea.parcial.id}"
        lineas_por_parcial[nombre_parcial] = lineas_por_parcial.get(nombre_parcial, 0) + 1

    db.session.commit()

    if lineas_por_parcial:
        detalle = ", ".join(f"{n} en {parcial}" for parcial, n in lineas_por_parcial.items())
        mensaje = f"Lotes cargados para {sum(lineas_por_parcial.values())} producto(s) ({detalle})."
    else:
        mensaje = "No se encontró ningún producto del archivo en este costeo."
    if no_encontrados:
        mensaje += f" {len(no_encontrados)} producto(s) del archivo no se encontraron en ningún parcial y se ignoraron."
    if ambiguos_encontrados:
        mensaje += (
            f" {len(ambiguos_encontrados)} producto(s) aparecen repetidos en más de un parcial "
            "(mismo código o descripción) y no se pudieron asignar automáticamente -- revísalos a mano."
        )
    flash(mensaje, "success" if lineas_por_parcial else "warning")
    return redirect(url_for("importaciones_detalle", importacion_id=imp.id))


# --- Exportar planilla para el sistema de Inventarios (ronda N, 2026-09-12) ---

ENCABEZADOS_INVENTARIO = [
    "CODIGO", "DESCRIPCION", "DD-MM-AAAA FECHA INVOICE", "CODIGO DE BARRA ESCANEADO",
    "COD.DE LOTE", "VENCMTO AAMMDD", "UBICACIÓN", "NUL", "valor Unitario EU", "Unidades",
    "Total Ivoice EU", "TOTAL INVOICE US$", "VALOR FLETE", "US $ SEGURO", "C.I.F.",
    "TOTAL C.I.F. $", "DERECHO ADUANA $", "OTROS GASTOS $", "TOTAL COSTO $", "COSTO UNITARIO $",
]
_ANCHOS_INVENTARIO = [20, 42, 16, 16, 16, 14, 10, 6, 12, 10, 12, 14, 12, 12, 12, 16, 16, 14, 14, 16]
# La tabla de 20 columnas empieza en la columna B (no en A) -- ronda O,
# punto 2: la columna A queda siempre en blanco y angosta, igual que en la
# planilla real del sistema de Inventarios (columnas B a U).
_COL_INVENTARIO_OFFSET = 1
_RELLENO_AMARILLO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
# Formato "Contabilidad" estandar de Excel (2 decimales, separador de miles,
# negativos entre parentesis, guion para el cero) -- ronda Q, punto 2: se
# aplica a todas las columnas de valores (J a U) EXCEPTO "Unidades" (K), que
# es una cantidad, no un monto.
_FORMATO_CONTABILIDAD = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'
# Indices (1-based, dentro de ENCABEZADOS_INVENTARIO) de las columnas de
# valores a las que se les aplica el formato de contabilidad (ronda Q, punto
# 2): de "valor Unitario EU" (9) a "COSTO UNITARIO $" (20), salteando
# "Unidades" (10), que es una cantidad, no un monto.
_COLS_VALORES_INVENTARIO = [i for i in range(9, 21) if i != 10]
# Columnas que llevan fila de subtotal (ronda Q, punto 3): "desde la columna
# J en adelante" -- a diferencia del formato de contabilidad, aca SI se
# incluye "Unidades" (J es la primera de este rango, columna 10 de la hoja).
_COLS_SUBTOTAL_INVENTARIO = list(range(9, 21))
# Fila donde empieza la tabla de 20 columnas (encabezados) -- deja espacio
# arriba para el bloque de cabecera (empresa/correlativo/proveedor/etc.,
# ver _escribir_cabecera_inventario) tal como en la planilla real.
_FILA_TABLA_INVENTARIO = 13


def _fecha_larga_es(d):
    """Formatea una fecha como '14 de julio de 2026', igual que la celda
    'FECHA FACTURA' de la planilla real del sistema de Inventarios."""
    if not d:
        return None
    mes = TipoCambioMensual.MESES_NOMBRE[d.month - 1].lower()
    return f"{d.day} de {mes} de {d.year}"


def _escribir_cabecera_inventario(ws, imp, proveedor):
    """Replica el bloque de cabecera de la planilla real del sistema de
    Inventarios (ver capturas del usuario, ronda O 2026-09-12), en las
    mismas celdas que usa esa planilla: B1/B2 título y empresa, C1 el
    nombre de la empresa otra vez (ronda Q, punto 1) y su logo en el área
    E1:F3, B3:C10 los datos de la importación (uno por fila), y M1/M2:N2
    el bloque de mes/fecha de carga. Solo C3 (correlativo), C5 (código del
    proveedor en el otro sistema) y N2 (fecha de carga) quedan resaltados
    en amarillo -- son, según el usuario, los únicos campos de esta
    cabecera que ese sistema exige para poder cargar el archivo."""
    ws["B1"] = "PLANTILLA DE COSTEO"
    ws["B1"].font = Font(bold=True)
    ws["C1"] = imp.empresa.nombre if imp.empresa else None
    ws["C1"].font = Font(bold=True)
    ws["B2"] = imp.empresa.nombre if imp.empresa else None

    # Logo de la empresa compradora, en el area E1:F3 (ronda Q, punto 1) --
    # si la empresa no tiene logo cargado (o el archivo no existe en disco),
    # simplemente no se agrega ninguna imagen, sin romper la exportacion.
    if imp.empresa and imp.empresa.logo_nombre_archivo:
        ruta_logo = os.path.join(EMPRESAS_LOGOS_DIR, imp.empresa.logo_nombre_archivo)
        if os.path.isfile(ruta_logo):
            try:
                img = ExcelImage(ruta_logo)
                ancho_max, alto_max = 200, 58  # aprox. el area E1:F3
                escala = min(ancho_max / img.width, alto_max / img.height, 1)
                img.width = int(img.width * escala)
                img.height = int(img.height * escala)
                ws.merge_cells("E1:F3")
                img.anchor = "E1"
                ws.add_image(img)
            except Exception:
                pass

    ws["B3"] = "NRO. IMPORTACION"
    ws["C3"] = imp.numero_correlativo_inventario
    ws["C3"].fill = _RELLENO_AMARILLO

    ws["B4"] = "NUMERO FACTURA"
    ws["C4"] = imp.numero_factura

    ws["B5"] = "RUT"
    ws["C5"] = proveedor.codigo_sistema_inventario if proveedor else None
    ws["C5"].fill = _RELLENO_AMARILLO

    ws["B6"] = "PROVEEDOR"
    ws["C6"] = proveedor.nombre if proveedor else None

    ws["B7"] = "FECHA FACTURA"
    ws["C7"] = _fecha_larga_es(imp.fecha_factura)

    ws["B8"] = "TIPO CAMBIO"
    ws["C8"] = imp.tipo_cambio_aduanero

    ws["B9"] = "DÓLAR ADUANERO"
    ws["C9"] = imp.tipo_cambio_aduanero

    ws["B10"] = "NUMERO DE DESPACHO"
    ws["C10"] = imp.despacho.numero_tracking if (imp.despacho_id and imp.despacho) else None

    for fila in range(3, 11):
        ws.cell(row=fila, column=2).font = Font(bold=True)

    ws["M1"] = "PLANILLA DE COSTEO"
    ws["M1"].font = Font(bold=True)
    ws["M2"] = "MES/AÑO :"
    ws["M2"].font = Font(bold=True)
    ws["N2"] = date.today().strftime("%d-%m-%Y")
    ws["N2"].fill = _RELLENO_AMARILLO

    # Ancho de B/C (y del resto de la tabla) lo fija _construir_excel_inventario
    # a partir de _ANCHOS_INVENTARIO, ya que son las mismas columnas que usa
    # la tabla de 20 encabezados (CODIGO=B, DESCRIPCION=C, ...) -- acá solo
    # se deja A en blanco y angosto.
    ws.column_dimensions["A"].width = 3


def _codigo_interno_homologado(parcial_linea):
    """Ronda V (2026-09-12, punto 1): devuelve el codigo interno del sistema
    de Inventarios (Ergopyme) ya homologado para esta linea de Costeo, o
    None si todavia no se conoce (nunca rompe nada -- la planilla
    simplemente deja esa celda en blanco, igual que antes de esta ronda).
    `parcial_linea.codigo` es "codigo_mostrar" de la linea de OC original
    (ver parciales_generar): el codigo padre para un producto simple, o el
    codigo de la variante especifica elegida (ej. una dioptria) cuando el
    producto es de una familia con variantes -- por eso primero se busca en
    ProductoVariante antes de asumir que es el codigo del Producto padre."""
    producto = parcial_linea.producto
    if not producto:
        return None
    codigo_mostrado = (parcial_linea.codigo or "").strip()
    if codigo_mostrado and codigo_mostrado.upper() != (producto.codigo or "").strip().upper():
        variante = ProductoVariante.query.filter(
            ProductoVariante.producto_id == producto.id,
            db.func.upper(ProductoVariante.codigo) == codigo_mostrado.upper(),
        ).first()
        if variante and variante.codigo_interno_inventario:
            return variante.codigo_interno_inventario
    return producto.codigo_interno_inventario


def _proveedor_canonico_historico(proveedor_original):
    """Ronda AE (2026-09-14): nombre de proveedor tal como debe mostrarse en
    los reportes/Consulta de Stock, a partir del nombre CRUDO que trae la
    columna "Proveedor" del histórico de compras -- aplica el alias BVI (ver
    ALIAS_PROVEEDOR_HISTORICO) y nada más. Ver _fila_historica_dict() para
    por qué esto reemplazó, para efectos de AGRUPACIÓN, al proveedor que
    resolvía el mapeo Ergopyme."""
    nombre = (proveedor_original or "").strip()
    if not nombre:
        return nombre
    return ALIAS_PROVEEDOR_HISTORICO.get(nombre.upper(), nombre)


def _construir_homologador_ergopyme():
    """Ronda AB (2026-09-14): arma (una sola vez por request) el
    homologador de códigos internos de Ergopyme -> (proveedor real,
    proveedor_id, código de proveedor, descripción), usando el mapeo
    maestro `CodigoErgopyme` (archivo "2da Revisión códigos ergopyme" que
    mantiene el usuario, ver seed_codigos_ergopyme). Se reutiliza tanto en
    Consulta de Stock como en Reportes > Compras Proveedor -- misma regla
    en los dos lugares: "no puede haber un código interno Ergopyme sin un
    código de proveedor, o en su defecto que ya esté marcado sin marca".

    IMPORTANTE: esta es la fuente de verdad VIVA -- siempre lee la tabla
    CodigoErgopyme tal como está en este momento (que se resincroniza en
    cada arranque desde el archivo de mapeo, ver seed_codigos_ergopyme).
    No hay que confiar en los campos "congelados" que quedaron grabados en
    CompraHistorica al momento de la carga histórica (esos quedan
    desactualizados apenas el usuario corrige el archivo de mapeo)."""
    codigo_ergopyme_map = {c.codigo_interno: c for c in CodigoErgopyme.query.all()}
    proveedores_por_nombre = {p.nombre.strip().upper(): p for p in Proveedor.query.all()}

    def _homologar(codigo_interno):
        """Devuelve (proveedor_nombre, proveedor_id, codigo_proveedor,
        descripcion) si el código interno está en el mapeo Ergopyme con un
        proveedor real (no '#N/A'), o None si no se pudo resolver."""
        if not codigo_interno:
            return None
        homolog = codigo_ergopyme_map.get(codigo_interno)
        if not homolog:
            return None
        prov_nombre = (homolog.proveedor_nombre or "").strip()
        if not prov_nombre or prov_nombre.upper() == "#N/A":
            return None
        prov_real = proveedores_por_nombre.get(prov_nombre.upper())
        return (
            prov_nombre,
            prov_real.id if prov_real else None,
            (homolog.codigo_proveedor or "").strip() or None,
            (homolog.descripcion or "").strip() or None,
        )
    return _homologar


def _moneda_operacion(paridad):
    """Ronda AB (2026-09-14): en el histórico de compras (y en las
    importaciones del sistema) solo existen dos monedas de operación
    posibles -- USD (cuando la 'paridad' es 1, es decir, no hay paso de
    conversión) o EUR (cuando la paridad es distinta de 1 -- el propio
    encabezado del archivo histórico dice literalmente "Paridad EUR")."""
    try:
        p = float(paridad or 0)
    except (TypeError, ValueError):
        p = 0
    if not p or abs(p - 1) < 1e-6:
        return "USD"
    return "EUR"


def _derecho_u_otro_costo_en_moneda(valor_clp, tipo_cambio, paridad):
    """Convierte un monto en CLP (Derechos o Otros gastos de la línea) a
    USD y a la moneda de la operación, con la fórmula exacta que confirmó
    el usuario con dos ejemplos reales del archivo histórico:
      valor_USD = valor_CLP / tipo_cambio       (tipo_cambio = CLP por 1 USD)
      valor_moneda_operacion = valor_USD * paridad
    Cuando la operación es en USD, paridad = 1 y ambos valores coinciden
    (fila 6097 del histórico: 7137 / 935,57 = 7,62 USD). Cuando es en otra
    divisa (fila 6137: tipo cambio 925,25, paridad 0,8585) el segundo valor
    queda expresado en esa divisa (euros)."""
    try:
        tc = float(tipo_cambio or 0)
    except (TypeError, ValueError):
        tc = 0
    try:
        par = float(paridad or 0)
    except (TypeError, ValueError):
        par = 0
    if not par:
        par = 1
    valor_clp = float(valor_clp or 0)
    valor_usd = (valor_clp / tc) if tc else 0.0
    valor_moneda = valor_usd * par
    return valor_usd, valor_moneda


def _paridad_efectiva_historica(c):
    """Ronda AC (2026-09-14), causa raíz corregida en la ronda AD: la
    paridad tal como viene en el archivo ES la fuente de verdad (así lo
    pidió el usuario: "para definir la moneda tiene que ver ese campo de
    paridad") -- desde la ronda AD, `seed_compras_historicas` ya parsea
    correctamente el formato de texto con coma decimal que traían 54 filas
    de DORC ("0,8687"), así que en datos cargados de nuevo esta paridad
    nunca debería venir realmente en 0/vacía (ver `num()` y
    `reparar_paridad_eur_historica`, que corrige las filas que ya habían
    quedado mal cargadas por el bug de parseo). Se mantiene igual la
    reconstrucción por cociente (Total Invoice / Total USD) como red de
    seguridad puramente defensiva, por si algún dato futuro llegara
    genuinamente vacío -- no debería activarse en la práctica."""
    paridad = c.paridad_eur
    if paridad:
        return paridad
    if c.total_usd:
        try:
            return c.total_invoice / c.total_usd
        except (TypeError, ZeroDivisionError):
            return 1.0
    return 1.0


def _normalizar_nombre_proveedor_moneda(nombre):
    """Normaliza un nombre de proveedor para poder cruzarlo contra
    moneda_proveedor.xlsx tolerando las variantes que conviven en este
    sistema (ej. catálogo "BVI BEAVER" vs histórico "BEAVER", catálogo
    "CUSTOM SURGICAL" vs archivo "CUSTOMS SURGICAL", "BAUSCH" vs
    "BAUSCH + LOMB")."""
    n = (nombre or "").strip().upper()
    n = n.replace("+", " ")
    n = re.sub(r"[^A-Z0-9 ]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    if n.startswith("BVI "):
        n = n[4:]
    # Quita un plural simple palabra por palabra (CUSTOMS -> CUSTOM) para
    # tolerar esa variante puntual sin tener que hardcodear un alias.
    palabras = [p[:-1] if len(p) > 4 and p.endswith("S") else p for p in n.split(" ")]
    return " ".join(palabras)


def _moneda_recurrente_proveedores():
    """Ronda AC (2026-09-14): lee moneda_proveedor.xlsx (planilla que
    mantiene el usuario con la moneda HABITUAL de cada proveedor
    extranjero) -- se usa únicamente como desempate en _moneda_dominante,
    nunca como fuente principal."""
    resultado = {}
    if not os.path.isfile(MONEDA_PROVEEDOR_EXCEL):
        return resultado
    wb = openpyxl.load_workbook(MONEDA_PROVEEDOR_EXCEL, data_only=True)
    ws = wb.worksheets[0]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        nombre = str(row[0]).strip()
        moneda_valor = str(row[1]).strip().upper() if len(row) > 1 and row[1] is not None else ""
        moneda = "EUR" if moneda_valor.startswith("EUR") else ("USD" if moneda_valor.startswith("USD") else None)
        if nombre and moneda:
            resultado[_normalizar_nombre_proveedor_moneda(nombre)] = moneda
    return resultado


def _moneda_recurrente(nombre, mapa=None):
    """Moneda habitual del proveedor `nombre` según moneda_proveedor.xlsx,
    o None si no aparece ahí (tolera variantes de nombre, ver
    _normalizar_nombre_proveedor_moneda)."""
    if not nombre:
        return None
    mapa = mapa if mapa is not None else _moneda_recurrente_proveedores()
    clave = _normalizar_nombre_proveedor_moneda(nombre)
    if clave in mapa:
        return mapa[clave]
    for clave_archivo, moneda in mapa.items():
        if clave == clave_archivo:
            return moneda
        if clave.startswith(clave_archivo + " ") or clave_archivo.startswith(clave + " "):
            return moneda
        if clave in clave_archivo or clave_archivo in clave:
            return moneda
    return None


def _moneda_dominante(filas, proveedor_nombre=None, mapa_recurrente=None):
    """Ronda AC (2026-09-14): determina en qué moneda expresar los totales
    de un conjunto de filas de UN mismo proveedor (ya filtradas por fecha)
    -- regla exacta del usuario: se cuenta en cuántas FACTURAS distintas
    predominó cada moneda (nunca líneas sueltas -- una factura completa es
    en una sola moneda) y gana la que más se repite en el rango. Empate ->
    la moneda "habitual" de ese proveedor (moneda_proveedor.xlsx). Sin
    datos de ningún tipo -> USD."""
    facturas_por_moneda = defaultdict(set)
    for f in filas:
        factura = f.get("factura") or f"(sin-factura-{id(f)})"
        facturas_por_moneda[f.get("moneda_operacion") or "USD"].add(factura)
    usd_n = len(facturas_por_moneda.get("USD", ()))
    eur_n = len(facturas_por_moneda.get("EUR", ()))
    if usd_n > eur_n:
        return "USD"
    if eur_n > usd_n:
        return "EUR"
    if usd_n == 0 and eur_n == 0:
        return "USD"
    return _moneda_recurrente(proveedor_nombre, mapa_recurrente) or "USD"


def _paridad_representativa_eur(filas):
    """Paridad EUR "representativa" del corte actual -- promedio de las
    paridades reales de las facturas en EUR, ponderado por su propio monto
    en USD. Se usa solo para el caso raro de tener que expresar en EUR una
    factura que realmente fue en USD (esa factura, por definición, no trae
    una tasa EUR propia -- se aproxima con la tasa que sí manejó ese mismo
    proveedor en fechas cercanas dentro del mismo corte)."""
    total_peso = 0.0
    acumulado = 0.0
    for f in filas:
        if (f.get("moneda_operacion") or "USD") == "EUR":
            peso = f.get("total_usd") or 0
            acumulado += (f.get("paridad") or 1.0) * peso
            total_peso += peso
    return (acumulado / total_peso) if total_peso else 1.0


def _valor_en_moneda_dominante(valor_usd, moneda_fila, moneda_dominante, paridad_fila, paridad_representativa):
    """Convierte un monto ya expresado en USD (total_usd/derechos_usd/
    flete_usd/otros_gastos_usd -- todos disponibles siempre, sin importar
    la moneda original de la factura) a la moneda dominante del corte
    actual. Si la propia fila ya está en esa moneda, la conversión es
    EXACTA (se usa la paridad real de esa factura). Si está en la moneda
    contraria (caso raro dentro de un mismo proveedor), se aproxima con la
    paridad representativa del corte."""
    if moneda_dominante == "USD":
        return valor_usd or 0
    paridad = paridad_fila if moneda_fila == "EUR" else (paridad_representativa or paridad_fila or 1.0)
    return (valor_usd or 0) * paridad


def _totales_en_moneda(filas, moneda_dominante, paridad_representativa):
    """Arma las 5 métricas del resumen (Cantidad de productos, Total
    Invoice, Flete [+ % del Total Invoice], Derechos, Otros costos) para un
    conjunto de filas, todas expresadas en `moneda_dominante`. Incluye
    también el equivalente en USD del Total Invoice (total_invoice_usd_ref)
    -- siempre exacto sin importar la moneda mostrada -- para poder
    ordenar/calcular "% del total" entre proveedores que muestran monedas
    distintas sin mezclar peras con manzanas."""
    total = flete = derechos = otros = 0.0
    total_usd_ref = 0.0
    for f in filas:
        mf = f.get("moneda_operacion") or "USD"
        pf = f.get("paridad") or 1.0
        total_usd_ref += f.get("total_usd") or 0
        total += _valor_en_moneda_dominante(f.get("total_usd"), mf, moneda_dominante, pf, paridad_representativa)
        flete += _valor_en_moneda_dominante(f.get("flete_usd"), mf, moneda_dominante, pf, paridad_representativa)
        derechos += _valor_en_moneda_dominante(f.get("derechos_usd"), mf, moneda_dominante, pf, paridad_representativa)
        otros += _valor_en_moneda_dominante(f.get("otros_gastos_usd"), mf, moneda_dominante, pf, paridad_representativa)
    productos = {
        (f.get("codigo_producto") or f.get("descripcion") or "").strip().upper()
        for f in filas if (f.get("codigo_producto") or f.get("descripcion"))
    }
    return {
        "moneda": moneda_dominante,
        "simbolo": "€" if moneda_dominante == "EUR" else "US$",
        "cantidad_productos": len(productos),
        "total_invoice": total,
        "total_invoice_usd_ref": total_usd_ref,
        "flete": flete,
        "flete_pct": (flete / total * 100) if total else 0,
        "derechos": derechos,
        "derechos_pct": (derechos / total * 100) if total else 0,
        "otros_gastos": otros,
        "otros_gastos_pct": (otros / total * 100) if total else 0,
    }


def _construir_excel_inventario(imp, resultado, codigo_proveedor=False):
    """Arma el Excel de carga al sistema de Inventarios: mismos encabezados
    y misma cadena de conversión de moneda (factura -> USD -> CLP) que la
    planilla de referencia del usuario ('Muestra planilla de costeo
    Importacion.xlsx'), reutilizando los montos YA CALCULADOS por
    costing.py línea por línea -- para no volver a calcular nada aparte y
    evitar un descuadre entre esta planilla y lo que se ve en pantalla.
    Una fila por LOTE físico (no por producto): si un producto tiene 2
    lotes, cada lote recibe su porción proporcional de FOB/Flete/Seguro/
    CIF/Gastos según su cantidad de unidades sobre el total de la línea.
    Si un producto todavía no tiene lotes cargados, se exporta como un
    único "lote" con el total de la línea (usando codigo_lote/
    fecha_vencimiento "resumen" de siempre), para que la planilla nunca
    salga incompleta aunque falte hacer la carga masiva de lotes.
    Un sheet por parcial, en el mismo orden en que se muestran las
    pestañas en pantalla. La columna A queda en blanco (la tabla empieza
    en B, ronda O). CODIGO DE BARRA ESCANEADO queda en blanco a propósito
    -- es el código interno del sistema de Inventarios, que este sistema
    no conoce (se completa/relaciona del otro lado).

    Columnas CODIGO (B) y DESCRIPCION (C) segun el checkbox "Código
    Proveedor" (ronda T, 2026-09-12; CODIGO homologado agregado en ronda V):
    - Sin marcar (default): CODIGO lleva el código INTERNO del sistema de
      Inventarios ya homologado (ver _codigo_interno_homologado, ronda V) --
      queda en blanco solo si ese producto/variante todavía no se homologó.
      DESCRIPCION lleva el código propio de Suite Logística concatenado con
      la descripción, igual que siempre desde ronda O.
    - Marcado: CODIGO lleva el código del proveedor (el propio de Suite
      Logística) y DESCRIPCION queda solo con la descripción, sin
      concatenar."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    nombres_usados = set()
    proveedor = imp.proveedor
    for idx, pr in enumerate(resultado["parciales"], start=1):
        parcial = pr["parcial"]
        base = parcial.referencia or parcial.numero_parcial or f"Parcial {idx}"
        base = re.sub(r'[\\/*?:\[\]]', "-", str(base)).strip()[:28] or f"Parcial {idx}"
        nombre_hoja = base
        sufijo = 2
        while nombre_hoja in nombres_usados:
            nombre_hoja = f"{base} ({sufijo})"
            sufijo += 1
        nombres_usados.add(nombre_hoja)

        ws = wb.create_sheet(title=nombre_hoja)
        _escribir_cabecera_inventario(ws, imp, proveedor)

        fila_encabezado = _FILA_TABLA_INVENTARIO
        for i, titulo in enumerate(ENCABEZADOS_INVENTARIO):
            celda = ws.cell(row=fila_encabezado, column=i + 1 + _COL_INVENTARIO_OFFSET, value=titulo)
            celda.font = Font(bold=True)
        for ancho, col in zip(_ANCHOS_INVENTARIO, range(1, len(_ANCHOS_INVENTARIO) + 1)):
            ws.column_dimensions[get_column_letter(col + _COL_INVENTARIO_OFFSET)].width = ancho

        fila_actual = fila_encabezado
        for info in pr["lineas"]:
            linea = info["linea"]
            codigo_limpio = (linea.codigo or "").strip()
            descripcion_limpia = (linea.descripcion or "").strip()
            if codigo_proveedor:
                # Checkbox marcado (ronda T): CODIGO (B) = código del proveedor
                # (el propio de Suite Logística), DESCRIPCION (C) = solo la
                # descripción, sin concatenar.
                valor_codigo_col_b = codigo_limpio
                valor_descripcion_col_c = descripcion_limpia
            else:
                # Ronda V (2026-09-12, punto 1): CODIGO (B) ahora lleva el
                # código interno homologado del sistema de Inventarios,
                # cuando ya se conoce (antes quedaba siempre en blanco).
                # DESCRIPCION (C) sigue igual que siempre: "código
                # descripción" concatenados con el código propio de Suite
                # Logística (no el interno).
                valor_codigo_col_b = _codigo_interno_homologado(linea)
                valor_descripcion_col_c = f"{codigo_limpio} {descripcion_limpia}".strip()
            total_unidades = linea.cantidad_unidades or 0
            lotes = list(linea.lotes) or [None]
            for lote in lotes:
                if lote is not None:
                    cantidad_lote = lote.cantidad_unidades or 0
                    codigo_lote = lote.codigo_lote or ""
                    fecha_venc = lote.fecha_vencimiento
                else:
                    cantidad_lote = total_unidades
                    codigo_lote = linea.codigo_lote or ""
                    fecha_venc = linea.fecha_vencimiento
                share = (cantidad_lote / total_unidades) if total_unidades else 0
                costo_total_lote = info["costo_total_clp"] * share
                fila_actual += 1
                valores = [
                    valor_codigo_col_b,
                    valor_descripcion_col_c,
                    imp.fecha_factura,
                    None,
                    codigo_lote,
                    fecha_venc,
                    None,
                    None,
                    linea.valor_unitario_moneda,
                    cantidad_lote,
                    (linea.valor_unitario_moneda or 0) * cantidad_lote,
                    info["valor_usd"] * share,
                    info["flete_usd"] * share,
                    info["seguro_usd"] * share,
                    info["cif_usd"] * share,
                    info["cif_clp"] * share,
                    info["derechos_clp"] * share,
                    (info["gastos_clp"] - info["derechos_clp"]) * share,
                    costo_total_lote,
                    (costo_total_lote / cantidad_lote) if cantidad_lote else 0,
                ]
                for i, valor in enumerate(valores):
                    ws.cell(row=fila_actual, column=i + 1 + _COL_INVENTARIO_OFFSET, value=valor)
                ws.cell(row=fila_actual, column=3 + _COL_INVENTARIO_OFFSET).number_format = "DD-MM-YYYY"
                if fecha_venc is not None:
                    ws.cell(row=fila_actual, column=6 + _COL_INVENTARIO_OFFSET).number_format = "YYMMDD"
                # Formato "Contabilidad" en las columnas de valores (ronda Q,
                # punto 2) -- "Unidades" (columna K) queda fuera a propósito.
                for col_pos in _COLS_VALORES_INVENTARIO:
                    ws.cell(row=fila_actual, column=col_pos + _COL_INVENTARIO_OFFSET).number_format = _FORMATO_CONTABILIDAD

        # Fila de subtotales, filtro en el encabezado y sin cuadricula (ronda
        # Q, puntos 3, 4 y 5) -- una vez procesadas todas las lineas/lotes de
        # este parcial.
        ws.sheet_view.showGridLines = False
        if fila_actual > fila_encabezado:
            fila_subtotal = fila_actual + 1
            ws.cell(row=fila_subtotal, column=2 + _COL_INVENTARIO_OFFSET, value="SUBTOTAL").font = Font(bold=True)
            for col_pos in _COLS_SUBTOTAL_INVENTARIO:
                col_letra = get_column_letter(col_pos + _COL_INVENTARIO_OFFSET)
                celda = ws.cell(row=fila_subtotal, column=col_pos + _COL_INVENTARIO_OFFSET)
                celda.value = f"=SUM({col_letra}{fila_encabezado + 1}:{col_letra}{fila_actual})"
                celda.font = Font(bold=True)
                if col_pos != 10:  # 10 = "Unidades", sin formato de contabilidad
                    celda.number_format = _FORMATO_CONTABILIDAD
        ws.auto_filter.ref = (
            f"B{fila_encabezado}:U{fila_actual}" if fila_actual > fila_encabezado
            else f"B{fila_encabezado}:U{fila_encabezado}"
        )
    if not wb.worksheets:
        ws = wb.create_sheet(title="Sin datos")
        _escribir_cabecera_inventario(ws, imp, proveedor)
        for i, titulo in enumerate(ENCABEZADOS_INVENTARIO):
            ws.cell(row=_FILA_TABLA_INVENTARIO, column=i + 1 + _COL_INVENTARIO_OFFSET, value=titulo).font = Font(bold=True)
        ws.sheet_view.showGridLines = False
        ws.auto_filter.ref = f"B{_FILA_TABLA_INVENTARIO}:U{_FILA_TABLA_INVENTARIO}"
    return wb


@app.route("/importaciones/<int:importacion_id>/exportar-inventario")
@requiere_permiso("generar_costeo", "inventarios")
def importaciones_exportar_inventario(importacion_id):
    """Descarga la planilla de carga al sistema de Inventarios (un sheet
    por parcial, un lote físico por fila) -- ver _construir_excel_inventario.
    Ronda T (2026-09-12): checkbox opcional "Código Proveedor" -- si se
    marca, la columna CODIGO (B) lleva el código propio de Suite Logística
    (el mismo que identifica al producto en el catálogo de ese proveedor) y
    la columna DESCRIPCION (C) queda con la descripción sola, sin
    concatenar; si no se marca, se mantiene el comportamiento de siempre
    (CODIGO en blanco, DESCRIPCION con "código descripción" concatenados)."""
    imp = Importacion.query.get_or_404(importacion_id)
    codigo_proveedor = request.args.get("codigo_proveedor") == "1"
    resultado = costing.calcular_costeo(imp)
    wb = _construir_excel_inventario(imp, resultado, codigo_proveedor=codigo_proveedor)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    nombre_archivo = f"Inventario_{(imp.numero_factura or imp.id)}.xlsx".replace("/", "-").replace(" ", "_")
    return send_file(
        buffer, as_attachment=True, download_name=nombre_archivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
def gastos_documento_ver(gasto_id, doc_id):
    doc = GastoDocumento.query.get_or_404(doc_id)
    if doc.gasto_id != gasto_id:
        abort(404)
    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(doc.gasto_id))
    return send_from_directory(
        carpeta_gasto, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=False
    )


@app.route("/gastos/<int:gasto_id>/documentos/<int:doc_id>/descargar")
@requiere_permiso("generar_costeo")
def gastos_documento_descargar(gasto_id, doc_id):
    doc = GastoDocumento.query.get_or_404(doc_id)
    if doc.gasto_id != gasto_id:
        abort(404)
    carpeta_gasto = os.path.join(DOCUMENTOS_GASTOS_DIR, str(doc.gasto_id))
    return send_from_directory(
        carpeta_gasto, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=True
    )


@app.route("/gastos/<int:gasto_id>/documentos/<int:doc_id>/eliminar", methods=["POST"])
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
def importacion_legajo_ver(importacion_id, doc_id):
    doc = ImportacionDocumento.query.get_or_404(doc_id)
    if doc.importacion_id != importacion_id:
        abort(404)
    carpeta = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(doc.importacion_id))
    return send_from_directory(
        carpeta, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=False
    )


@app.route("/importaciones/<int:importacion_id>/legajo/<int:doc_id>/descargar")
@requiere_permiso("generar_costeo")
def importacion_legajo_descargar(importacion_id, doc_id):
    doc = ImportacionDocumento.query.get_or_404(doc_id)
    if doc.importacion_id != importacion_id:
        abort(404)
    carpeta = os.path.join(DOCUMENTOS_LEGAJO_DIR, str(doc.importacion_id))
    return send_from_directory(
        carpeta, doc.nombre_archivo, download_name=doc.nombre_original, as_attachment=True
    )


@app.route("/importaciones/<int:importacion_id>/legajo/<int:doc_id>/eliminar", methods=["POST"])
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
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
@requiere_permiso("generar_costeo")
def cargos_eliminar(cargo_id):
    cargo = CargoAdicionalImportacion.query.get_or_404(cargo_id)
    importacion_id = cargo.importacion_id
    db.session.delete(cargo)
    db.session.commit()
    flash("Cargo eliminado.", "success")
    return redirect(url_for("importaciones_detalle", importacion_id=importacion_id))


@app.route("/importaciones/<int:importacion_id>/cargos/eliminar-multiple", methods=["POST"])
@requiere_permiso("generar_costeo")
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
@requiere_admin
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
@requiere_admin
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
@requiere_admin
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
@requiere_admin
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
# Modulo: Configuracion -- Perfiles (Roles) y Usuarios (ronda R, 2026-09-12)
# ---------------------------------------------------------------------------

@app.route("/configuracion/perfiles")
@requiere_admin
def roles_list():
    roles = Rol.query.order_by(Rol.nombre).all()
    return render_template("configuracion/roles.html", roles=roles, permisos=PERMISOS_DISPONIBLES)


def _guardar_permisos_rol(rol, form):
    rol.es_administrador = form.get("es_administrador") == "on"
    for codigo, _nombre in PERMISOS_DISPONIBLES:
        setattr(rol, f"permiso_{codigo}", form.get(f"permiso_{codigo}") == "on")
    # Ronda W (2026-09-13, punto 5): a pedido explicito del usuario, "Inventarios"
    # y "Consultar Stock" no deben quedar como conceptos separados que se
    # puedan desalinear -- todo perfil con permiso Inventarios lleva SIEMPRE
    # tambien Consultar Stock (la consulta de saldos es un subconjunto de lo
    # que ya puede hacer alguien de Inventarios). No es reciproco: alguien
    # puede tener SOLO Consultar Stock sin Inventarios (el colaborador que
    # solo necesita ver saldos).
    if rol.permiso_inventarios:
        rol.permiso_consultar_stock = True


@app.route("/configuracion/perfiles/nuevo", methods=["POST"])
@requiere_admin
def roles_nuevo():
    nombre = request.form.get("nombre", "").strip()
    if not nombre:
        flash("El perfil necesita un nombre.", "warning")
        return redirect(url_for("roles_list"))
    if Rol.query.filter(db.func.lower(Rol.nombre) == nombre.lower()).first():
        flash(f"Ya existe un perfil llamado '{nombre}'.", "warning")
        return redirect(url_for("roles_list"))
    rol = Rol(nombre=nombre)
    _guardar_permisos_rol(rol, request.form)
    db.session.add(rol)
    db.session.commit()
    flash(f"Perfil '{rol.nombre}' creado.", "success")
    return redirect(url_for("roles_list"))


@app.route("/configuracion/perfiles/<int:rol_id>/editar", methods=["POST"])
@requiere_admin
def roles_editar(rol_id):
    rol = Rol.query.get_or_404(rol_id)
    nombre = request.form.get("nombre", "").strip()
    if not nombre:
        flash("El perfil necesita un nombre.", "warning")
        return redirect(url_for("roles_list"))
    if Rol.query.filter(db.func.lower(Rol.nombre) == nombre.lower(), Rol.id != rol.id).first():
        flash(f"Ya existe otro perfil llamado '{nombre}'.", "warning")
        return redirect(url_for("roles_list"))
    # Misma salvaguarda que en usuarios_editar: si a ESTE perfil se le quita
    # "es_administrador" y tiene usuarios activos asignados, hay que
    # asegurarse de que quede al menos otro administrador activo en algun
    # otro perfil -- si no, nadie podria volver a administrar usuarios/roles.
    dejara_de_ser_admin = rol.es_administrador and request.form.get("es_administrador") != "on"
    if dejara_de_ser_admin and rol.usuarios.filter_by(activo=True).count() > 0:
        otros_admins_activos = Usuario.query.join(Rol).filter(
            Rol.es_administrador.is_(True), Usuario.activo.is_(True), Rol.id != rol.id
        ).count()
        if otros_admins_activos == 0:
            flash("No puedes quitarle 'Administrador' a este perfil: quedaría sin ningún administrador activo en el sistema.", "danger")
            return redirect(url_for("roles_list"))
    rol.nombre = nombre
    _guardar_permisos_rol(rol, request.form)
    db.session.commit()
    flash(f"Perfil '{rol.nombre}' actualizado.", "success")
    return redirect(url_for("roles_list"))


@app.route("/configuracion/perfiles/<int:rol_id>/eliminar", methods=["POST"])
@requiere_admin
def roles_eliminar(rol_id):
    rol = Rol.query.get_or_404(rol_id)
    if rol.usuarios.count() > 0:
        flash(f"'{rol.nombre}' tiene usuarios asignados: reasígnalos a otro perfil antes de eliminarlo.", "warning")
        return redirect(url_for("roles_list"))
    db.session.delete(rol)
    db.session.commit()
    flash("Perfil eliminado.", "success")
    return redirect(url_for("roles_list"))


@app.route("/configuracion/usuarios")
@requiere_admin
def usuarios_list():
    usuarios = Usuario.query.order_by(Usuario.nombre_completo).all()
    roles = Rol.query.order_by(Rol.nombre).all()
    return render_template("configuracion/usuarios.html", usuarios=usuarios, roles=roles)


@app.route("/configuracion/usuarios/nuevo", methods=["POST"])
@requiere_admin
def usuarios_nuevo():
    email = request.form.get("email", "").strip().lower()
    nombre_completo = request.form.get("nombre_completo", "").strip()
    rol_id = request.form.get("rol_id") or None
    password = request.form.get("password", "")
    if not (email and nombre_completo and rol_id and password):
        flash("Nombre, correo, perfil y contraseña son obligatorios.", "warning")
        return redirect(url_for("usuarios_list"))
    if len(password) < 6:
        flash("La contraseña debe tener al menos 6 caracteres.", "warning")
        return redirect(url_for("usuarios_list"))
    if Usuario.query.filter(db.func.lower(Usuario.email) == email).first():
        flash(f"Ya existe un usuario con el correo '{email}'.", "warning")
        return redirect(url_for("usuarios_list"))
    usuario = Usuario(nombre_completo=nombre_completo, email=email, rol_id=int(rol_id), activo=True)
    usuario.set_password(password)
    db.session.add(usuario)
    db.session.commit()
    flash(f"Usuario '{usuario.nombre_completo}' creado.", "success")
    return redirect(url_for("usuarios_list"))


@app.route("/configuracion/usuarios/<int:usuario_id>/editar", methods=["POST"])
@requiere_admin
def usuarios_editar(usuario_id):
    usuario = Usuario.query.get_or_404(usuario_id)
    email = request.form.get("email", "").strip().lower()
    nombre_completo = request.form.get("nombre_completo", "").strip()
    rol_id = request.form.get("rol_id") or None
    if not (email and nombre_completo and rol_id):
        flash("Nombre, correo y perfil son obligatorios.", "warning")
        return redirect(url_for("usuarios_list"))
    if Usuario.query.filter(db.func.lower(Usuario.email) == email, Usuario.id != usuario.id).first():
        flash(f"Ya existe otro usuario con el correo '{email}'.", "warning")
        return redirect(url_for("usuarios_list"))
    # Salvaguardas para no dejar la app sin ningun Administrador activo
    # (ronda R, 2026-09-12) -- sin esto un error de carga (o el checkbox
    # "Usuario activo" deshabilitado en el propio formulario, que un
    # navegador simplemente no envia al hacer submit) podria desactivar o
    # quitarle el perfil de Administrador al unico usuario que puede
    # administrar usuarios, dejando la app inaccesible para siempre.
    nuevo_rol = Rol.query.get(int(rol_id))
    era_admin = bool(usuario.rol and usuario.rol.es_administrador)
    sera_admin = bool(nuevo_rol and nuevo_rol.es_administrador)
    if era_admin and not sera_admin:
        otros_admins_activos = Usuario.query.join(Rol).filter(
            Rol.es_administrador.is_(True), Usuario.activo.is_(True), Usuario.id != usuario.id
        ).count()
        if otros_admins_activos == 0:
            flash("No puedes quitarle el perfil de Administrador: es el único administrador activo.", "danger")
            return redirect(url_for("usuarios_list"))
    if usuario.id == current_user.id:
        nueva_activa = True  # nunca te desactivas a ti mismo
    else:
        nueva_activa = request.form.get("activo") == "on"
        if era_admin and sera_admin and not nueva_activa:
            otros_admins_activos = Usuario.query.join(Rol).filter(
                Rol.es_administrador.is_(True), Usuario.activo.is_(True), Usuario.id != usuario.id
            ).count()
            if otros_admins_activos == 0:
                flash("No puedes desactivar al único administrador activo.", "danger")
                return redirect(url_for("usuarios_list"))
    usuario.email = email
    usuario.nombre_completo = nombre_completo
    usuario.rol_id = int(rol_id)
    usuario.activo = nueva_activa
    nueva_password = request.form.get("password", "")
    if nueva_password:
        if len(nueva_password) < 6:
            flash("La nueva contraseña debe tener al menos 6 caracteres -- no se cambió.", "warning")
        else:
            usuario.set_password(nueva_password)
    db.session.commit()
    flash(f"Usuario '{usuario.nombre_completo}' actualizado.", "success")
    return redirect(url_for("usuarios_list"))


@app.route("/configuracion/usuarios/<int:usuario_id>/eliminar", methods=["POST"])
@requiere_admin
def usuarios_eliminar(usuario_id):
    usuario = Usuario.query.get_or_404(usuario_id)
    if usuario.id == current_user.id:
        flash("No puedes eliminar tu propio usuario mientras tienes la sesión abierta.", "warning")
        return redirect(url_for("usuarios_list"))
    nombre = usuario.nombre_completo
    db.session.delete(usuario)
    db.session.commit()
    flash(f"Usuario '{nombre}' eliminado.", "success")
    return redirect(url_for("usuarios_list"))


# ---------------------------------------------------------------------------
# Modulo: Stock (saldos de existencias) -- ronda V, 2026-09-12
# ---------------------------------------------------------------------------
# Consulta de stock por empresa (Accuvision/Accumedical juntas, se filtra en
# pantalla), pensada para colaboradores que solo necesitan ver saldos, sin
# tocar Compras/Costeo. Los datos vienen del reporte del sistema de
# Inventarios ("Ergopyme"), que no es parte de esta Suite -- alguien con
# permiso "inventarios" lo sube manualmente cuando hace falta actualizarlo
# (reemplaza TODO el stock cargado antes, ver _clasificar_y_cargar_stock).

def _transito_por_linea():
    """Ronda X (2026-09-13, punto 3B): unidades "en tránsito" por (producto,
    variante) -- lineas de ordenes YA APROBADAS que el proveedor ya
    despachó pero que todavía no llegan a bodega (etapa 'Orden Despachada'
    o 'Internación Aduanas'; una vez 'Recibido' ya deberían reflejarse en
    el Stock real y dejan de contar aquí). Devuelve un diccionario
    {(producto_id, variante_codigo_en_mayuscula_o_None): {"cantidad": int,
    "detalle": [str, ...]}} -- 'detalle' es una lista de textos para el
    tooltip (cantidad + fecha estimada de llegada, sin numero de PO) SIN
    exponer nada mas de la orden: quien solo tiene 'Consulta de Stock' (sin
    'seguimiento' ni permisos de Compras) nunca ve ni navega a la orden en
    si, solo este resumen agregado (ronda X, punto 5)."""
    ETAPAS_EN_TRANSITO = ("Orden Despachada", "Internación Aduanas")
    lineas = (
        OrdenCompraLinea.query
        .join(OrdenCompra)
        .filter(
            OrdenCompraLinea.anulada == False,  # noqa: E712
            OrdenCompraLinea.etapa.in_(ETAPAS_EN_TRANSITO),
            OrdenCompra.estado_aprobacion == "Aprobada",
            # Ronda AD (2026-09-14): bug real reportado por el usuario --
            # al cancelar una orden YA asociada a un despacho, `ordenes_
            # cancelar` la marca "Cancelada" pero (a propósito, para no
            # perder trazabilidad) no toca sus líneas ni el despacho -- así
            # que sus líneas seguían contando acá como "en tránsito" aunque
            # la orden ya estuviera cancelada/eliminada para el usuario.
            OrdenCompra.estado != "Cancelada",
        )
        .all()
    )
    resultado = {}
    for linea in lineas:
        cantidad = linea.cantidad_unidades
        if cantidad <= 0:
            continue
        clave = (linea.producto_id, (linea.variante_codigo or "").strip().upper() or None)
        info = resultado.setdefault(clave, {"cantidad": 0, "detalle": []})
        info["cantidad"] += cantidad
        despacho = linea.orden.despacho
        fecha_estimada = None
        if despacho and despacho.fecha_estimada_llegada:
            fecha_estimada = despacho.fecha_estimada_llegada
        elif linea.fecha_estimada_despacho_confirmada:
            fecha_estimada = linea.fecha_estimada_despacho_confirmada
        elif linea.fecha_estimada_despacho:
            fecha_estimada = linea.fecha_estimada_despacho
        fecha_texto = fecha_estimada.strftime("%d-%m-%Y") if fecha_estimada else "sin fecha estimada"
        # Ronda X (2026-09-13, punto 5): el tooltip NO menciona el numero de
        # PO ni ningun otro dato de la orden -- solo la cantidad y la fecha
        # estimada, para que "Consulta de Stock" (sin permiso de
        # Seguimiento) no tenga forma de identificar ni rastrear la orden.
        info["detalle"].append(f"{cantidad} un. (llegada estimada {fecha_texto})")
    return resultado


def _pedido_por_codigo():
    """Ronda X (2026-09-13, punto 3C): unidades ya comprometidas con
    clientes (columna 'Pedido'), sumadas por (empresa_id, codigo_interno) a
    partir de PedidoComprometido -- puede haber varias filas para el mismo
    código (un pedido de cliente distinto cada una).

    Ronda AA (2026-09-13): tambien devuelve la descripcion de cada codigo
    (la primera que encuentre) -- se usa en stock_list para poder armar una
    fila aunque ese codigo nunca haya aparecido en un reporte de Stock (ver
    "codigos con Pedido pero sin Stock cargado" abajo)."""
    resultado = {}
    descripciones = {}
    for p in PedidoComprometido.query.all():
        clave = (p.empresa_id, p.codigo_interno)
        resultado[clave] = resultado.get(clave, 0) + (p.cantidad or 0)
        if p.descripcion and clave not in descripciones:
            descripciones[clave] = p.descripcion
    return resultado, descripciones


def _estados_homologacion_por_codigo():
    """Ronda Z (2026-09-13): estado de HomologacionStock por codigo_interno
    -- usado en Consulta de Stock para distinguir, entre los codigos SIN
    producto/variante ligado, los que se clasificaron a proposito como
    'sin_marca' (se muestran como proveedor "Sin Marca") de los que
    todavia estan 'pendiente' de resolver (siguen mostrando
    "(sin homologar)")."""
    return {h.codigo_interno: h.estado for h in HomologacionStock.query.all()}


@app.route("/stock")
@requiere_permiso("consultar_stock", "inventarios")
def stock_list():
    empresa_id = request.args.get("empresa_id", "")
    proveedor_id = request.args.get("proveedor_id", "")
    q = request.args.get("q", "").strip()

    query = StockExistencia.query
    if empresa_id:
        query = query.filter_by(empresa_id=empresa_id)
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(
            StockExistencia.descripcion.ilike(like),
            StockExistencia.codigo_interno.ilike(like),
            StockExistencia.producto.has(Producto.codigo.ilike(like)),
            StockExistencia.variante.has(ProductoVariante.codigo.ilike(like)),
        ))
    filas = query.all()

    # Ronda Z (2026-09-13): para los codigos SIN producto/variante ligado,
    # distinguir "sin_marca" (clasificacion explicita, a pedido del
    # usuario, para que a futuro se pueda editar/asignar a un proveedor
    # real) de "pendiente" (todavia sin resolver).
    estados_homologacion = _estados_homologacion_por_codigo()

    # Ronda AB (2026-09-14): mapeo maestro Proveedor/Código Proveedor por
    # código interno Ergopyme (el mismo que usa el reporte Compras
    # Proveedor) -- se usa acá también como fuente de verdad para mostrar
    # el proveedor y código REAL de un producto que no tiene Producto/
    # Variante ligado en el catálogo, en vez de dejarlo "(sin homologar)"
    # o "(pedido sin stock cargado)" a secas. Regla del usuario: ningún
    # código interno Ergopyme debería quedar sin proveedor conocido, salvo
    # que ya esté marcado "sin marca". (helper compartido con Reportes,
    # ver _construir_homologador_ergopyme)
    _homologar_via_ergopyme = _construir_homologador_ergopyme()

    # Ronda V: stock "Desglosado por variante exacta" -- se agrupa por
    # (empresa, producto, variante) sumando todos los lotes de esa
    # combinacion, mostrando el stock total y el proximo vencimiento (el
    # detalle de cada lote individual se ve expandiendo la fila). Un
    # producto/variante sin match (sin_marca o pendiente) se agrupa por su
    # propio codigo_interno, ya que no hay Producto al que sumarle.
    grupos = {}
    for fila in filas:
        if fila.variante_id:
            clave = ("variante", fila.variante_id, fila.empresa_id)
        elif fila.producto_id:
            clave = ("producto", fila.producto_id, fila.empresa_id)
        else:
            clave = ("codigo", fila.codigo_interno, fila.empresa_id)
        grupo = grupos.get(clave)
        if grupo is None:
            proveedor_nombre = None
            proveedor_id_grupo = None
            producto_id_grupo = None
            variante_codigo_grupo = None
            via_ergopyme = False
            # Ronda X (2026-09-13, punto 2): un codigo todavia sin
            # homologar (sin marca / pendiente) ya NO muestra el codigo
            # crudo de Ergopyme en pantalla -- queda en blanco (se sigue
            # viendo la descripcion, y se lo puede seguir buscando por su
            # codigo interno desde el buscador de arriba).
            codigo_mostrar = None
            descripcion = fila.descripcion
            if fila.variante:
                codigo_mostrar = fila.variante.codigo
                descripcion = fila.variante.descripcion or fila.descripcion
                proveedor_nombre = fila.variante.producto.proveedor.nombre
                proveedor_id_grupo = fila.variante.producto.proveedor_id
                producto_id_grupo = fila.variante.producto_id
                variante_codigo_grupo = fila.variante.codigo.strip().upper()
            elif fila.producto:
                codigo_mostrar = fila.producto.codigo
                descripcion = fila.producto.descripcion
                proveedor_nombre = fila.producto.proveedor.nombre
                proveedor_id_grupo = fila.producto.proveedor_id
                producto_id_grupo = fila.producto_id
            elif estados_homologacion.get(fila.codigo_interno) == "sin_marca":
                # Ronda Z (2026-09-13): clasificado a proposito como "Sin
                # Marca" (ej. lentes genericos como ABBOTT LIO ACRILICO) --
                # se muestra en la columna/filtro de Proveedor con esa
                # etiqueta, en vez de "(sin homologar)", y queda disponible
                # en /stock/homologacion para asignarlo a un proveedor real
                # y darle mantenimiento a su codigo mas adelante.
                proveedor_nombre = "Sin Marca"
            else:
                # Ronda AB (2026-09-14): antes de rendirse a "(sin
                # homologar)", probar el mapeo Ergopyme (Proveedor +
                # Código Proveedor) -- ver _homologar_via_ergopyme arriba.
                resuelto = _homologar_via_ergopyme(fila.codigo_interno)
                if resuelto:
                    proveedor_nombre, proveedor_id_grupo, codigo_mostrar, _ = resuelto
                    via_ergopyme = True
            grupo = {
                "empresa": fila.empresa,
                "proveedor_id": proveedor_id_grupo,
                "proveedor_nombre": proveedor_nombre,
                "codigo": codigo_mostrar,
                "codigo_interno": fila.codigo_interno,
                "descripcion": descripcion,
                "stock_total": 0,
                "proximo_vencimiento": None,
                "lotes": [],
                "homologado": bool(fila.producto_id or fila.variante_id or via_ergopyme),
                "via_ergopyme": via_ergopyme,
                "_producto_id": producto_id_grupo,
                "_variante_codigo": variante_codigo_grupo,
            }
            grupos[clave] = grupo
        grupo["stock_total"] += fila.stock_fisico or 0
        if fila.fecha_vencimiento and (
            grupo["proximo_vencimiento"] is None or fila.fecha_vencimiento < grupo["proximo_vencimiento"]
        ):
            grupo["proximo_vencimiento"] = fila.fecha_vencimiento
        grupo["lotes"].append(fila)

    if proveedor_id == "sin_marca":
        grupos = {k: g for k, g in grupos.items() if g["proveedor_nombre"] == "Sin Marca"}
    elif proveedor_id:
        grupos = {k: g for k, g in grupos.items() if str(g["proveedor_id"] or "") == str(proveedor_id)}

    # Ronda X (2026-09-13, punto 3): columnas Stock actual (A) / Tránsito
    # (B) / Pedido (C) / Stock total (D = A+B-C), agregadas sobre cada
    # grupo ya armado arriba.
    transito_por_linea = _transito_por_linea()
    pedido_por_codigo, descripcion_pedido_por_codigo = _pedido_por_codigo()
    for grupo in grupos.values():
        transito_info = transito_por_linea.get((grupo["_producto_id"], grupo["_variante_codigo"]), None)
        grupo["stock_actual"] = grupo["stock_total"]
        grupo["transito"] = transito_info["cantidad"] if transito_info else 0
        grupo["transito_detalle"] = (
            "<br>".join(str(escape(linea)) for linea in transito_info["detalle"]) if transito_info else ""
        )
        grupo["pedido"] = pedido_por_codigo.get((grupo["empresa"].id if grupo["empresa"] else None, grupo["codigo_interno"]), 0)
        grupo["stock_total_proyectado"] = grupo["stock_actual"] + grupo["transito"] - grupo["pedido"]

    # Ronda AA (2026-09-13): un codigo puede tener unidades comprometidas en
    # "Notas de pedido" sin tener NINGUNA fila en el ultimo reporte de Stock
    # cargado (ej. se agoto y se cayo del reporte de Inventarios, o es un
    # codigo nuevo que todavia no aparecio en Stock) -- sin esto, esas
    # unidades quedaban invisibles en Consulta de Stock (no habia ningun
    # grupo al que sumarselas, ver validacion real hecha contra "Notas de
    # pedidos.xlsx": 26 de las 154 unidades del archivo, en 7 codigos
    # distintos, no aparecian en ningun lado de la pantalla). Se arma una
    # fila "sintetica" (Stock actual = 0, Transito = 0) solo para poder
    # mostrar el Pedido pendiente.
    # Ronda AB (2026-09-14): antes se omitia esta fila por completo si habia
    # un filtro de Proveedor activo (se asumia que estos codigos "no tienen
    # proveedor conocido") -- ahora primero se intenta homologar cada codigo
    # via el mapeo Ergopyme (mismo usado en Compras Proveedor), y solo si
    # sigue sin resolverse se sigue ocultando bajo un filtro de Proveedor
    # especifico (no tiene sentido mostrarlo bajo un proveedor que no le
    # corresponde). Con proveedor resuelto, tambien se reemplaza el codigo
    # interno de Ergopyme por el Codigo de Proveedor real en la columna
    # "Codigo" -- para reportes/estadisticas el usuario quiere ver siempre
    # el codigo del proveedor, nunca el codigo interno.
    codigos_cubiertos = {(f.empresa_id, f.codigo_interno) for f in filas}
    empresas_por_id = {e.id: e for e in Empresa.query.all()}
    q_lower = q.lower() if q else ""
    if proveedor_id != "sin_marca":
        for (emp_id, codigo), cantidad_pedido in pedido_por_codigo.items():
            if not cantidad_pedido or (emp_id, codigo) in codigos_cubiertos:
                continue
            resuelto = _homologar_via_ergopyme(codigo)
            if resuelto:
                prov_nombre_res, prov_id_res, codigo_prov_res, descripcion_ergopyme = resuelto
            else:
                prov_nombre_res = prov_id_res = codigo_prov_res = descripcion_ergopyme = None
            if proveedor_id and str(prov_id_res or "") != str(proveedor_id):
                continue
            if empresa_id and str(emp_id) != str(empresa_id):
                continue
            descripcion_pend = descripcion_pedido_por_codigo.get((emp_id, codigo), "") or descripcion_ergopyme or ""
            if (
                q_lower
                and q_lower not in (descripcion_pend or "").lower()
                and q_lower not in (codigo or "").lower()
                and q_lower not in (codigo_prov_res or "").lower()
            ):
                continue
            grupos[("pedido_sin_stock", codigo, emp_id)] = {
                "empresa": empresas_por_id.get(emp_id),
                "proveedor_id": prov_id_res,
                "proveedor_nombre": prov_nombre_res,
                "codigo": codigo_prov_res,
                "codigo_interno": codigo,
                "descripcion": descripcion_pend,
                "stock_total": 0,
                "proximo_vencimiento": None,
                "lotes": [],
                "homologado": bool(resuelto),
                "via_ergopyme": bool(resuelto),
                "sin_stock_cargado": True,
                "_producto_id": None,
                "_variante_codigo": None,
                "stock_actual": 0,
                "transito": 0,
                "transito_detalle": "",
                "pedido": cantidad_pedido,
                "stock_total_proyectado": 0 + 0 - cantidad_pedido,
            }

    # Ronda AD (2026-09-14): orden por defecto pedido por el usuario -- el
    # producto con MAYOR stock total (proyectado, columna D) primero. Como
    # este orden se recalcula en cada request (no se guarda en sesión ni en
    # el navegador), cualquier cambio de filtro vuelve a aplicar este mismo
    # orden por defecto sobre el resultado ya filtrado, sin que el usuario
    # tenga que volver a ordenar a mano -- el ordenamiento manual por
    # columna (ver JS abajo) sigue disponible y solo dura hasta el próximo
    # cambio de filtro (recarga la página).
    resultado = sorted(grupos.values(), key=lambda g: -(g["stock_total_proyectado"] or 0))
    empresas = Empresa.query.order_by(Empresa.nombre).all()
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    ultima_carga = db.session.query(db.func.max(StockExistencia.cargado_en)).scalar()
    ultima_carga_pedidos = db.session.query(db.func.max(PedidoComprometido.cargado_en)).scalar()
    return render_template(
        "stock/list.html", grupos=resultado, empresas=empresas, proveedores=proveedores,
        empresa_sel=empresa_id, proveedor_sel=proveedor_id, q=q,
        ultima_carga=ultima_carga, ultima_carga_pedidos=ultima_carga_pedidos,
    )


@app.route("/stock/cargar", methods=["POST"])
@requiere_permiso("inventarios")
def stock_cargar():
    """Sube un reporte nuevo del sistema de Inventarios ("Ergopyme", formato
    original: columnas A-H, sin cruce) y REEMPLAZA por completo el Stock
    cargado antes -- para que siempre refleje la ultima foto real, sin
    arrastrar ni duplicar datos viejos (confirmado con el usuario, ronda
    V)."""
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        flash("Selecciona el archivo de Stock para cargar.", "warning")
        return redirect(url_for("stock_list"))
    extension = os.path.splitext(archivo.filename)[1].lower()
    if extension not in (".xlsx", ".xls"):
        flash("Formato no soportado. Sube el archivo .xlsx del reporte de Stock.", "danger")
        return redirect(url_for("stock_list"))
    try:
        wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
        ws = wb.active
        filas = _leer_filas_reporte_stock(ws)
    except Exception:
        flash("No se pudo leer el archivo. Verifica que no esté dañado o abierto en otro programa.", "danger")
        return redirect(url_for("stock_list"))

    if not filas:
        flash("El archivo no tiene filas de stock reconocibles -- no se cambió nada.", "warning")
        return redirect(url_for("stock_list"))

    StockExistencia.query.delete()
    resumen = _clasificar_y_cargar_stock(filas)
    db.session.commit()

    mensaje = (
        f"Stock actualizado: {resumen['cargados']} producto(s)/lote(s) cargados "
        f"({resumen['vinculados']} ligados al catálogo, {resumen['sin_marca']} sin marca)."
    )
    if resumen["pendientes_nuevos"]:
        mensaje += (
            f" {resumen['pendientes_nuevos']} código(s) nuevo(s) todavía no están homologados -- "
            f"revísalos en 'Homologación de Stock'."
        )
    if resumen["empresas_no_encontradas"]:
        mensaje += f" No se pudo identificar la empresa para: {', '.join(sorted(resumen['empresas_no_encontradas']))}."
    flash(mensaje, "success" if not resumen["pendientes_nuevos"] and not resumen["empresas_no_encontradas"] else "warning")
    return redirect(url_for("stock_list"))


@app.route("/stock/pedidos/cargar", methods=["POST"])
@requiere_permiso("inventarios")
def stock_pedidos_cargar():
    """Ronda X (2026-09-13, punto 3C): sube el reporte 'Notas de pedido'
    (Balance de Productos) del sistema de Inventarios -- fila 1 trae el
    nombre de la empresa (igual que el reporte de Stock), encabezado en la
    fila con 'Cód.Producto' en la columna A y los datos desde la fila
    siguiente: A=código interno, C=Denominación, G=N° de pedido,
    I=fecha de pedido, K=cliente, L=cantidad comprometida (Despacho
    Comprometido). REEMPLAZA solo los registros de la empresa que trae ESE
    archivo (confirmado con el usuario) -- si mas adelante sube el
    equivalente de la otra empresa, no se pisan entre si. El archivo
    original queda guardado en data/reportes/ (uno por empresa, se
    sobrescribe) para poder volver a consultarlo despues."""
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        flash("Selecciona el archivo de Notas de pedido para cargar.", "warning")
        return redirect(url_for("stock_list"))
    extension = os.path.splitext(archivo.filename)[1].lower()
    if extension not in (".xlsx", ".xls"):
        flash("Formato no soportado. Sube el archivo .xlsx del reporte de Notas de pedido.", "danger")
        return redirect(url_for("stock_list"))

    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active
        filas_excel = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
    except Exception:
        flash("No se pudo leer el archivo. Verifica que no esté dañado o abierto en otro programa.", "danger")
        return redirect(url_for("stock_list"))

    if not filas_excel:
        flash("El archivo está vacío -- no se cambió nada.", "warning")
        return redirect(url_for("stock_list"))

    nombre_empresa_reporte = str(filas_excel[0][0]).strip() if filas_excel[0] and filas_excel[0][0] else ""
    empresa = _empresa_por_nombre_reporte(nombre_empresa_reporte, {})
    if not empresa:
        flash(
            f"No se pudo identificar a qué empresa pertenece este archivo (fila 1 dice "
            f"'{nombre_empresa_reporte or '(vacío)'}') -- no se cambió nada.", "danger",
        )
        return redirect(url_for("stock_list"))

    fila_inicio = None
    for i, row in enumerate(filas_excel[:10], start=1):
        primera = str(row[0]).strip().lower() if row and row[0] else ""
        if primera.replace("ó", "o") == "cód.producto".replace("ó", "o"):
            fila_inicio = i + 1
            break
    if fila_inicio is None:
        flash("No se encontró el encabezado 'Cód.Producto' en el archivo -- no se cambió nada.", "danger")
        return redirect(url_for("stock_list"))

    nuevas_filas = []
    for row in filas_excel[fila_inicio - 1:]:
        codigo_crudo = row[0] if len(row) > 0 else None
        if not codigo_crudo:
            continue
        codigo_interno = _normalizar_codigo_interno(codigo_crudo)
        if not codigo_interno:
            continue
        descripcion = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        nro_pedido = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
        fecha_pedido = row[8] if len(row) > 8 else None
        if isinstance(fecha_pedido, datetime):
            fecha_pedido = fecha_pedido.date()
        else:
            fecha_pedido = None
        cliente_nombre = str(row[10]).strip() if len(row) > 10 and row[10] is not None else ""
        cantidad = row[11] if len(row) > 11 else 0
        try:
            cantidad = int(cantidad) if cantidad is not None else 0
        except (TypeError, ValueError):
            cantidad = 0
        if cantidad <= 0:
            continue
        nuevas_filas.append(PedidoComprometido(
            empresa_id=empresa.id, codigo_interno=codigo_interno, descripcion=descripcion,
            nro_pedido=nro_pedido, fecha_pedido=fecha_pedido, cliente_nombre=cliente_nombre,
            cantidad=cantidad,
        ))

    if not nuevas_filas:
        flash("El archivo no tiene filas de pedidos comprometidos reconocibles -- no se cambió nada.", "warning")
        return redirect(url_for("stock_list"))

    PedidoComprometido.query.filter_by(empresa_id=empresa.id).delete()
    db.session.add_all(nuevas_filas)
    db.session.commit()

    # Guarda el archivo original para poder volver a consultarlo despues
    # (un archivo por empresa, se reemplaza cada vez).
    os.makedirs(REPORTES_DIR, exist_ok=True)
    nombre_archivo = f"notas_pedido_{re.sub(r'[^A-Za-z0-9]+', '_', empresa.nombre).strip('_').lower()}.xlsx"
    archivo.stream.seek(0)
    archivo.save(os.path.join(REPORTES_DIR, nombre_archivo))

    total_unidades = sum(f.cantidad for f in nuevas_filas)
    flash(
        f"Notas de pedido de {empresa.nombre} actualizadas: {len(nuevas_filas)} línea(s), "
        f"{total_unidades} unidad(es) comprometidas en total. Verifica que este total coincida con el "
        f"archivo que subiste antes de confiar en la columna 'Pedido'.", "success",
    )
    return redirect(url_for("stock_list"))


@app.route("/api/stock/productos")
@requiere_permiso("inventarios")
def api_stock_buscar_productos():
    """Busqueda de productos (con sus variantes, si tiene) para el
    selector de 'Homologación de Stock' -- separada de /api/proveedores/
    <id>/productos porque esa exige permiso de Creación de Orden/Costeo, y
    quien homologa Stock puede no tenerlo."""
    proveedor_id = request.args.get("proveedor_id") or None
    q = request.args.get("q", "").strip()
    query = Producto.query
    if proveedor_id:
        query = query.filter_by(proveedor_id=proveedor_id)
    if q:
        query = query.filter(db.or_(Producto.codigo.ilike(f"%{q}%"), Producto.descripcion.ilike(f"%{q}%")))
    productos = query.order_by(Producto.codigo).limit(50).all()
    return jsonify([
        {
            "id": p.id,
            "codigo": p.codigo,
            "descripcion": p.descripcion,
            "proveedor": p.proveedor.nombre,
            "variantes": [{"id": v.id, "codigo": v.codigo, "descripcion": v.descripcion} for v in p.variantes],
        }
        for p in productos
    ])


@app.route("/stock/homologacion")
@requiere_permiso("inventarios")
def stock_homologacion():
    """Códigos internos del sistema de Inventarios que todavía no están
    ligados a ningún producto de nuestro catálogo -- ver estado 'pendiente'
    en HomologacionStock -- y códigos ya clasificados como 'sin_marca' (a
    pedido explícito del usuario, ej. lentes genéricos como ABBOTT LIO
    ACRILICO). Ronda Z (2026-09-13): antes esta pantalla solo mostraba los
    'pendiente', asi que un codigo marcado 'sin_marca' quedaba sin ninguna
    forma de editarlo despues; ahora tambien aparece aca, para poder
    asignarlo mas adelante a un proveedor real y darle mantenimiento a su
    código -- desde acá se resuelven a mano, uno por uno.

    Ronda AG (2026-09-15): los códigos 'sin_marca' ya clasificados (varios
    cientos, genéricos y sin proveedor) se acumulaban SIEMPRE visibles junto
    a los 'pendiente' de verdad -- el usuario los quiere mantener así, pero
    NO viéndolos por defecto (le ensucian la pantalla de códigos realmente
    nuevos por resolver). Por defecto ahora solo se muestran los
    'pendiente' -- ?ver=sin_marca (o ?ver=todos) trae de vuelta los otros
    dos casos para cuando sí quiera revisarlos/reasignarlos."""
    ver = request.args.get("ver", "pendiente").strip().lower()
    if ver == "todos":
        estados = ["pendiente", "sin_marca"]
    elif ver == "sin_marca":
        estados = ["sin_marca"]
    else:
        ver = "pendiente"
        estados = ["pendiente"]
    pendientes = HomologacionStock.query.filter(
        HomologacionStock.estado.in_(estados)
    ).order_by(HomologacionStock.estado, HomologacionStock.codigo_interno).all()
    total_sin_marca = HomologacionStock.query.filter_by(estado="sin_marca").count()
    total_pendiente = HomologacionStock.query.filter_by(estado="pendiente").count()
    proveedores = Proveedor.query.filter_by(activo=True).order_by(Proveedor.nombre).all()
    return render_template(
        "stock/homologacion.html", pendientes=pendientes, proveedores=proveedores,
        ver=ver, total_sin_marca=total_sin_marca, total_pendiente=total_pendiente,
    )


@app.route("/stock/homologacion/asignar-codigo-interno", methods=["POST"])
@requiere_permiso("inventarios")
def stock_homologacion_asignar_codigo_interno():
    """Ronda AG (2026-09-15, punto 4): un producto nuevo agregado directo en
    una Orden de Compra (proveedor y código de producto ya conocidos) puede
    no existir todavía en el sistema Ergopyme -- nunca se vendió, o recién se
    incorporó -- así que no tiene código interno y no puede subirse en la
    planilla de Costeo. Antes la única forma de que un producto consiguiera
    su código interno era al revés (subiendo el reporte de Inventarios y
    resolviendo el 'pendiente' que aparece acá); esto permite hacerlo
    directamente: buscar el producto/variante ya existente en el catálogo y
    escribirle su código interno a mano, sin pasar por HomologacionStock (no
    aplica -- este producto nunca apareció en un reporte de Inventarios)."""
    producto_id = request.form.get("producto_id") or None
    variante_id = request.form.get("variante_id") or None
    codigo_interno = (request.form.get("codigo_interno") or "").strip()
    ver = request.form.get("ver", "pendiente")

    if not producto_id or not codigo_interno:
        flash("Selecciona un producto e indica el código interno Ergopyme a asignar.", "warning")
        return redirect(url_for("stock_homologacion", ver=ver))

    producto = Producto.query.get_or_404(producto_id)
    variante = ProductoVariante.query.get(variante_id) if variante_id else None
    objetivo = variante or producto
    etiqueta = f"{producto.codigo}" + (f" / {variante.codigo}" if variante else "")

    # Aviso (no bloqueante) si ese código interno ya está en uso en OTRO
    # producto/variante -- evita pisar sin darse cuenta una homologación
    # existente.
    ya_usado_en = None
    otro_p = Producto.query.filter(
        Producto.codigo_interno_inventario == codigo_interno, Producto.id != producto.id,
    ).first()
    otro_v = ProductoVariante.query.filter(
        ProductoVariante.codigo_interno_inventario == codigo_interno,
        ProductoVariante.id != (variante.id if variante else -1),
    ).first()
    if otro_p:
        ya_usado_en = otro_p.codigo
    elif otro_v:
        ya_usado_en = otro_v.codigo

    objetivo.codigo_interno_inventario = codigo_interno
    db.session.commit()

    if ya_usado_en:
        flash(
            f"Código interno '{codigo_interno}' asignado a {etiqueta} -- OJO: ese mismo código ya estaba "
            f"asignado también a '{ya_usado_en}', revisa que no sea un error.", "warning",
        )
    else:
        flash(f"Código interno '{codigo_interno}' asignado a {etiqueta}.", "success")
    return redirect(url_for("stock_homologacion", ver=ver))


@app.route("/stock/homologacion/<int:homolog_id>/resolver", methods=["POST"])
@requiere_permiso("inventarios")
def stock_homologacion_resolver(homolog_id):
    homolog = HomologacionStock.query.get_or_404(homolog_id)
    accion = request.form.get("accion")

    if accion == "sin_marca":
        homolog.estado = "sin_marca"
        homolog.producto_id = None
        homolog.variante_id = None
    elif accion == "excluido":
        homolog.estado = "excluido"
        homolog.producto_id = None
        homolog.variante_id = None
        StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).delete()
    elif accion == "vincular_existente":
        producto_id = request.form.get("producto_id") or None
        variante_id = request.form.get("variante_id") or None
        if not producto_id:
            flash("Selecciona un producto del catálogo para vincular.", "warning")
            return redirect(url_for("stock_homologacion"))
        producto = Producto.query.get_or_404(producto_id)
        variante = ProductoVariante.query.get(variante_id) if variante_id else None
        homolog.estado = "vinculado"
        homolog.producto_id = producto.id
        homolog.variante_id = variante.id if variante else None
        if variante:
            variante.codigo_interno_inventario = homolog.codigo_interno
        else:
            producto.codigo_interno_inventario = homolog.codigo_interno
        StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).update({
            "producto_id": producto.id, "variante_id": variante.id if variante else None,
        })
    elif accion == "crear_producto":
        proveedor_id = request.form.get("proveedor_id_nuevo")
        codigo_nuevo = (request.form.get("codigo_nuevo") or "").strip()
        if not proveedor_id or not codigo_nuevo:
            flash("Indica proveedor y código para crear el producto nuevo.", "warning")
            return redirect(url_for("stock_homologacion"))
        nuevo = Producto(
            proveedor_id=proveedor_id,
            codigo=codigo_nuevo,
            descripcion=homolog.descripcion_referencia or codigo_nuevo,
            empaque=1, precio_caja=0, precio_unitario=0, activo=True,
            codigo_interno_inventario=homolog.codigo_interno,
        )
        db.session.add(nuevo)
        db.session.flush()
        homolog.estado = "vinculado"
        homolog.producto_id = nuevo.id
        StockExistencia.query.filter_by(codigo_interno=homolog.codigo_interno).update({"producto_id": nuevo.id})
    else:
        flash("Acción no reconocida.", "danger")
        return redirect(url_for("stock_homologacion"))

    db.session.commit()
    flash(f"Código interno '{homolog.codigo_interno}' resuelto.", "success")
    return redirect(url_for("stock_homologacion"))


# ---------------------------------------------------------------------------
# Modulo: Configuracion -- Tipo de cambio mensual (2026-08-27, punto 6)
# ---------------------------------------------------------------------------

@app.route("/configuracion/tipo-cambio")
@requiere_admin
def tipo_cambio_list():
    tipos = TipoCambioMensual.query.order_by(
        TipoCambioMensual.anio.desc(), TipoCambioMensual.mes.desc()
    ).all()
    return render_template("configuracion/tipo_cambio.html", tipos=tipos, hoy_anio=date.today().year)


@app.route("/configuracion/tipo-cambio/nuevo", methods=["POST"])
@requiere_admin
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
@requiere_admin
def tipo_cambio_editar(tc_id):
    tc = TipoCambioMensual.query.get_or_404(tc_id)
    tc.tc_aduanero = float(request.form.get("tc_aduanero") or 0)
    tc.paridad_eur_usd = float(request.form.get("paridad_eur_usd") or 1)
    tc.notas = request.form.get("notas", "").strip()
    db.session.commit()
    flash(f"Tipo de cambio de {tc.nombre_mes} {tc.anio} actualizado.", "success")
    return redirect(url_for("tipo_cambio_list"))


@app.route("/configuracion/tipo-cambio/<int:tc_id>/eliminar", methods=["POST"])
@requiere_admin
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
@requiere_admin
def admin_reset():
    return render_template("configuracion/reset.html", resumen=_resumen_datos_reset())


@app.route("/configuracion/reset/ejecutar", methods=["POST"])
@requiere_admin
def admin_reset_ejecutar():
    categorias = set(request.form.getlist("categorias"))
    categorias &= {"proveedores", "ordenes", "costeos"}
    if not categorias:
        flash("No marcaste ninguna categoría para borrar.", "warning")
        return redirect(url_for("admin_reset"))
    partes = _ejecutar_reset(categorias)
    flash("Borrado: " + "; ".join(partes) + ".", "success")
    return redirect(url_for("dashboard"))


# ---------------------------------------------------------------------------
# Reportes (ronda AA, 2026-09-13 -- resumen y vista por proveedor rediseñados
# en ronda AB, 2026-09-14)
# ---------------------------------------------------------------------------

def _proveedores_reporte_compras(homologar=None):
    """Lista de proveedores para el filtro del reporte Compras Proveedor:
    los del catálogo activo + los que aparecen en el histórico de compras
    (por su proveedor CRUDO, con el alias BVI aplicado -- ver
    _proveedor_canonico_historico) -- para no dejar fuera un proveedor
    antiguo que ya no está en el catálogo (ej. una relación comercial
    discontinuada). El parámetro `homologar` se mantiene por compatibilidad
    de firma pero ya no se usa para esto (ver Ronda AE, _fila_historica_dict)."""
    nombres = {p.nombre for p in Proveedor.query.filter_by(activo=True).all()}
    for (proveedor_original,) in db.session.query(CompraHistorica.proveedor_original).distinct():
        nombre = _proveedor_canonico_historico(proveedor_original)
        if nombre and nombre.upper() != "#N/A":
            nombres.add(nombre)
    return sorted(nombres)


def _fila_historica_dict(c, homologar):
    """Ronda AB (2026-09-14): arma la fila de reporte para una línea de
    CompraHistorica homologando el código interno EN VIVO contra el mapeo
    Ergopyme actual (en vez de confiar en proveedor_homologado/
    codigo_proveedor/homologado, que quedaron "congelados" en el momento
    de la carga histórica y no se actualizan solos si el usuario corrige
    el archivo de mapeo). Regla del usuario: para reportes/estadísticas
    nunca se muestra el código interno de Ergopyme -- se usa el código y
    la descripción del PROVEEDOR."""
    # Ronda AE (2026-09-14): el PROVEEDOR de cada línea histórica se toma
    # SIEMPRE del proveedor crudo de esa factura (columna "Proveedor" del
    # archivo histórico, con el alias BVI aplicado) -- NUNCA del proveedor
    # que trae el mapeo Ergopyme. Se encontró (validado línea por línea
    # contra la planilla de comprobación que trajo el usuario, cuadrando
    # EXACTO mes a mes y proveedor por proveedor) que el mapeo Ergopyme
    # tiene varios códigos internos con el campo Proveedor mal cargado
    # (ej. repuestos de QUANTEL/BEAVER y sondas de OPTIKON marcados como
    # "DORC" en el archivo de mapeo, dos de ellos con la nota "REVISAR CON
    # LISTA PRECIO" del propio usuario) -- eso hacía que esas compras se
    # sumaran al proveedor equivocado en el reporte, descuadrando el total
    # del proveedor "de más" (y el del proveedor real "de menos"). El
    # código de PRODUCTO sí se sigue resolviendo vía el mapeo Ergopyme
    # cuando existe (es confiable y necesario para la agrupación por
    # código padre) -- solo el nombre del proveedor dejó de depender de él.
    proveedor = _proveedor_canonico_historico(c.proveedor_original)
    resuelto = homologar(c.codigo_interno)
    if resuelto:
        _prov_ergopyme, _prov_id, codigo_prov_resuelto, descripcion_resuelta = resuelto
        codigo_producto = codigo_prov_resuelto or c.codigo_proveedor or c.codigo_interno
        descripcion = descripcion_resuelta or c.descripcion
        via_ergopyme = True
        homologado = True
    else:
        codigo_producto = c.codigo_proveedor or c.codigo_interno
        descripcion = c.descripcion
        via_ergopyme = False
        homologado = False

    # Ronda AC (2026-09-14): la paridad EFECTIVA (no la cruda) es la que
    # define la moneda y las conversiones -- corrige el caso de las 54
    # líneas con "Paridad EUR" en 0 (dato faltante, no una compra en USD;
    # ver _paridad_efectiva_historica).
    paridad_efectiva = _paridad_efectiva_historica(c)

    # Derechos y Otros gastos vienen en CLP en el histórico (columnas
    # "DERECHOS $" y "OTROS GASTOS $") -- se convierten a USD y a la
    # moneda de la propia operación (USD o EUR según la paridad de esa
    # línea) con la fórmula que confirmó el usuario (ver
    # _derecho_u_otro_costo_en_moneda).
    derechos_usd, derechos_moneda = _derecho_u_otro_costo_en_moneda(c.derechos_clp, c.tipo_cambio, paridad_efectiva)
    otros_usd, otros_moneda = _derecho_u_otro_costo_en_moneda(c.otros_gastos_clp, c.tipo_cambio, paridad_efectiva)

    return {
        "origen": "Histórico",
        "fecha_factura": c.fecha_factura,
        "proveedor": proveedor,
        "factura": c.factura,
        "codigo_interno": c.codigo_interno,
        "codigo_producto": codigo_producto,
        "descripcion": descripcion,
        "unidades": c.unidades or 0,
        "total_invoice": c.total_invoice or 0,
        "total_usd": c.total_usd or 0,
        "moneda_operacion": _moneda_operacion(paridad_efectiva),
        "tipo_cambio": c.tipo_cambio,
        "paridad": paridad_efectiva,
        "flete_usd": c.flete_usd or 0,
        "derechos_usd": derechos_usd,
        "derechos_moneda": derechos_moneda,
        "otros_gastos_usd": otros_usd,
        "otros_gastos_moneda": otros_moneda,
        "empresa_compradora": c.empresa_compradora,
        "categoria": c.categoria,
        "tipo_flete": c.tipo_flete,
        "homologado": homologado,
        "via_ergopyme": via_ergopyme,
    }


def _compras_sistema(proveedor=None, fecha_desde=None, fecha_hasta=None):
    """Ronda AA (2026-09-13), extendido en ronda AB (2026-09-14) con
    Derechos/Otros gastos/Flete por línea: compras hechas DESDE la
    plataforma (Importacion/Parcial/ParcialLinea, vía costing.py),
    transformadas a filas con la MISMA forma que el histórico -- así el
    reporte "Compras Proveedor" combina ambas fuentes sin duplicar nada.
    De acá en adelante, cada Costeo que se genere en el sistema alimenta
    este mismo reporte automáticamente, sin recargar ni recapturar nada."""
    query = Importacion.query
    if proveedor:
        query = query.join(Proveedor).filter(db.func.upper(Proveedor.nombre) == proveedor.upper())
    if fecha_desde:
        query = query.filter(Importacion.fecha_factura >= fecha_desde)
    if fecha_hasta:
        query = query.filter(Importacion.fecha_factura <= fecha_hasta)

    filas = []
    for importacion in query.all():
        if importacion.parciales.count() == 0:
            continue
        try:
            costeo = costing.calcular_costeo(importacion)
        except Exception:
            continue
        proveedor_nombre = importacion.proveedor.nombre if importacion.proveedor else ""
        empresa_nombre = importacion.empresa.nombre.upper() if importacion.empresa else "SIN ASIGNAR"
        # tipo_cambio_aduanero = CLP por 1 USD (igual que la columna "USD
        # TIPO CAMBIO" del histórico); tipo_cambio_moneda_usd = unidades de
        # moneda_factura por 1 USD (igual que "Paridad EUR": queda en 1
        # cuando la factura ya es en USD) -- mismos dos numeros, mismo rol.
        tc_aduanero = importacion.tipo_cambio_aduanero or 0
        tc_moneda_usd = importacion.tipo_cambio_moneda_usd or 1
        moneda_factura = (importacion.moneda_factura or "USD").strip().upper()
        moneda_operacion = "EUR" if moneda_factura.startswith("EUR") else moneda_factura
        for info in costeo["lineas"]:
            linea = info["linea"]
            codigo_interno = _codigo_interno_homologado(linea) or ""
            derechos_usd = (info["derechos_clp"] / tc_aduanero) if tc_aduanero else 0.0
            otros_usd = (info["otros_gastos_clp"] / tc_aduanero) if tc_aduanero else 0.0
            filas.append({
                "origen": "Sistema",
                "fecha_factura": importacion.fecha_factura,
                "proveedor": proveedor_nombre,
                "factura": importacion.numero_factura,
                "codigo_interno": codigo_interno,
                # Una OC/Costeo hecho desde la plataforma ya guarda el
                # código del PROVEEDOR en la línea (no el código interno de
                # Ergopyme) -- se muestra directo, sin pasar por el mapeo.
                "codigo_producto": (linea.codigo or "").strip() or codigo_interno or "(sin código)",
                "descripcion": linea.descripcion,
                "unidades": linea.cantidad_unidades or 0,
                "total_invoice": linea.valor_total_moneda,
                "total_usd": info["valor_usd"],
                "moneda_operacion": moneda_operacion,
                "tipo_cambio": tc_aduanero,
                "paridad": tc_moneda_usd,
                "flete_usd": info.get("flete_usd", 0) or 0,
                "derechos_usd": derechos_usd,
                "derechos_moneda": derechos_usd * tc_moneda_usd,
                "otros_gastos_usd": otros_usd,
                "otros_gastos_moneda": otros_usd * tc_moneda_usd,
                "empresa_compradora": empresa_nombre,
                "categoria": "",
                "tipo_flete": "",
                "homologado": True,
                "via_ergopyme": False,
            })
    return filas


def _totales_reporte(filas, proveedor_nombre=None, mapa_recurrente=None):
    """Ronda AC (2026-09-14): arma las 5 métricas del resumen (Cantidad de
    productos, Total Invoice, Flete, Derechos, Otros costos) para un
    conjunto de filas. Ya NO fuerza todo a USD: cuando `filas` pertenece a
    UN proveedor puntual (se pasa `proveedor_nombre`), se resuelve la
    moneda DOMINANTE de ese corte (ver _moneda_dominante -- cuenta
    facturas, no líneas) y todo se expresa en esa moneda. Cuando no hay un
    proveedor puntual (la vista general, sin filtrar), se usa USD -- es la
    única forma sensata de sumar proveedores que operan en monedas
    distintas entre sí."""
    if proveedor_nombre:
        moneda = _moneda_dominante(filas, proveedor_nombre, mapa_recurrente)
        paridad_rep = _paridad_representativa_eur(filas) if moneda == "EUR" else 1.0
    else:
        moneda, paridad_rep = "USD", 1.0
    return _totales_en_moneda(filas, moneda, paridad_rep)


def _resumen_por_proveedor(filas, mapa_recurrente=None):
    """Ronda AB (2026-09-13), corregido en ronda AC (2026-09-14): agrupa
    por proveedor y arma las mismas 5 métricas que _totales_reporte, pero
    cada proveedor en SU PROPIA moneda dominante (antes se forzaba todo a
    USD parejo, lo que mostraba como "USD" o "MIXTO" compras que en
    realidad son 100% en euros -- ver _moneda_dominante). El orden y el
    "% del total" siguen usando el equivalente en USD (total_invoice_usd_ref,
    siempre exacto sin importar la moneda que se muestra) para poder
    comparar proveedores que se expresan en monedas distintas entre sí."""
    mapa_recurrente = mapa_recurrente if mapa_recurrente is not None else _moneda_recurrente_proveedores()
    agrupado = defaultdict(list)
    for f in filas:
        agrupado[f["proveedor"] or "(sin proveedor)"].append(f)

    filas_resumen = []
    for p, filas_p in agrupado.items():
        t = _totales_reporte(filas_p, proveedor_nombre=p, mapa_recurrente=mapa_recurrente)
        t["proveedor"] = p
        filas_resumen.append(t)
    return sorted(filas_resumen, key=lambda r: -r["total_invoice_usd_ref"])


def _mapa_padre_por_variante_proveedor(proveedor_nombre):
    """Ronda AD (2026-09-14): para proveedores de lentes que ya tienen
    catálogo de código padre/variante (hoy MEDICONTUR y BVI PHYSIOL, ver
    ProductoVariante y seed_variantes_lentes_medicontur/_physiol), arma
    {codigo_variante_en_mayuscula: producto_padre} -- se usa en el reporte
    Compras Proveedor para agrupar primero por código padre y recién al
    entrar a ese padre mostrar el código de producto (variante) específico,
    tal como pidió el usuario. Un proveedor SIN catálogo de variantes (la
    gran mayoría) devuelve un mapa vacío, y el reporte sigue mostrando la
    lista plana de siempre -- no hace falta ninguna lista de "proveedores
    de lentes" hardcodeada, se deriva sola de qué proveedores ya tienen
    ProductoVariante cargado."""
    prov = Proveedor.query.filter(db.func.upper(Proveedor.nombre) == (proveedor_nombre or "").upper()).first()
    if not prov:
        return {}
    variantes = ProductoVariante.query.join(Producto).filter(Producto.proveedor_id == prov.id).all()
    return {v.codigo.strip().upper(): v.producto for v in variantes if v.codigo}


def _filas_reporte_compras(proveedor=None, empresa=None, fecha_desde=None, fecha_hasta=None, homologar=None):
    """Junta histórico + sistema ya homologados y filtrados -- reutilizado
    tanto por el listado general como por la vista de un proveedor
    específico, para no repetir la lógica de armado en dos lugares."""
    homologar = homologar or _construir_homologador_ergopyme()
    query = CompraHistorica.query
    if fecha_desde:
        query = query.filter(CompraHistorica.fecha_factura >= fecha_desde)
    if fecha_hasta:
        query = query.filter(CompraHistorica.fecha_factura <= fecha_hasta)

    filas = [_fila_historica_dict(c, homologar) for c in query.all()]
    if proveedor:
        filas = [f for f in filas if (f["proveedor"] or "").upper() == proveedor.upper()]
    filas += _compras_sistema(proveedor=proveedor or None, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta)
    if empresa:
        filas = [f for f in filas if (f["empresa_compradora"] or "").upper() == empresa.upper()]
    return filas


@app.route("/reportes")
@requiere_permiso("reportes")
def reportes_index():
    return redirect(url_for("reportes_compras_proveedor"))


def _rango_fechas_reporte_compras():
    """Ronda AE (2026-09-14): al entrar SIN filtro de fecha explícito, el
    reporte de Compras Proveedor (resumen y detalle por proveedor) se
    muestra por defecto filtrado al "año comercial" en curso -- desde el
    01-01 del año actual hasta hoy. El año siguiente (ej. 2027) el default
    avanza solo, porque se calcula con date.today() en cada request, no un
    año quemado. Si la URL YA trae fecha_desde y/o fecha_hasta como
    parámetro -- aunque venga vacío, como al hacer clic en "Quitar
    filtros" -- se respeta tal cual sin forzar el default, para que
    "Quitar filtros" de verdad muestre TODO el histórico sin fecha."""
    if "fecha_desde" in request.args or "fecha_hasta" in request.args:
        return request.args.get("fecha_desde", ""), request.args.get("fecha_hasta", "")
    hoy = date.today()
    return date(hoy.year, 1, 1).isoformat(), hoy.isoformat()


@app.route("/reportes/compras-proveedor")
@requiere_permiso("reportes")
def reportes_compras_proveedor():
    """Ronda AA (2026-09-13), resumen rediseñado en ronda AB (2026-09-14):
    primer reporte del menú "Reportes" -- combina el histórico de compras a
    proveedores (cargado desde el Excel de referencia del usuario,
    homologado EN VIVO por código Ergopyme) con las compras hechas DESDE la
    plataforma (calculadas en vivo desde Costeo de Importaciones), filtrable
    por Proveedor / Empresa compradora / rango de fechas de factura. El
    resumen ya no muestra N° de facturas/líneas/costo en CLP -- ahora
    muestra Cantidad de productos, Total Invoice, Flete, Derechos y Otros
    costos (los 3 últimos en la moneda de la operación / USD agregado)."""
    proveedor = request.args.get("proveedor", "").strip()
    empresa = request.args.get("empresa", "").strip()
    fecha_desde_txt, fecha_hasta_txt = _rango_fechas_reporte_compras()
    fecha_desde = parse_date(fecha_desde_txt)
    fecha_hasta = parse_date(fecha_hasta_txt)

    homologar = _construir_homologador_ergopyme()
    mapa_recurrente = _moneda_recurrente_proveedores()
    filas = _filas_reporte_compras(
        proveedor=proveedor or None, empresa=empresa or None,
        fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, homologar=homologar,
    )

    resumen = _resumen_por_proveedor(filas, mapa_recurrente=mapa_recurrente)
    # Ronda AC (2026-09-14): si la pantalla ya está filtrada a UN proveedor
    # puntual, las tarjetas de arriba usan la moneda dominante de ESE
    # proveedor (igual que su fila en el resumen) en vez de forzar USD --
    # solo la vista sin filtrar (mezcla de proveedores con monedas
    # distintas entre sí) se muestra en USD.
    totales = _totales_reporte(filas, proveedor_nombre=proveedor or None, mapa_recurrente=mapa_recurrente)
    sin_homologar = sum(1 for f in filas if f.get("homologado") is False)

    return render_template(
        "reportes/compras_proveedor.html",
        resumen=resumen, totales=totales,
        proveedores=_proveedores_reporte_compras(homologar), empresas=Empresa.query.order_by(Empresa.nombre).all(),
        proveedor_sel=proveedor, empresa_sel=empresa,
        fecha_desde=fecha_desde_txt, fecha_hasta=fecha_hasta_txt,
        sin_homologar=sin_homologar, total_filas=len(filas),
    )


@app.route("/reportes/compras-proveedor/<path:proveedor>")
@requiere_permiso("reportes")
def reportes_compras_proveedor_detalle(proveedor):
    """Ronda AB (2026-09-14): vista "análisis por productos" al entrar a un
    proveedor específico -- tabla dinámica (como la tabla dinámica de Excel
    que trajo el usuario de referencia): filas = código de proveedor,
    columnas = meses, valores = Total Invoice en la MONEDA DE COMPRA
    (columna R del histórico, tal cual, sin convertir). Con buscador con
    texto predictivo, flechas de orden por columna, expandir una fila para
    ver las facturas, y el resumen de arriba recalculándose en vivo (JS)
    según lo que quede visible en la tabla -- ver
    reportes/compras_proveedor_detalle.html."""
    empresa = request.args.get("empresa", "").strip()
    fecha_desde_txt, fecha_hasta_txt = _rango_fechas_reporte_compras()
    fecha_desde = parse_date(fecha_desde_txt)
    fecha_hasta = parse_date(fecha_hasta_txt)

    filas = _filas_reporte_compras(
        proveedor=proveedor, empresa=empresa or None,
        fecha_desde=fecha_desde, fecha_hasta=fecha_hasta,
    )
    if not filas:
        flash(f'No hay compras registradas para el proveedor "{proveedor}" con esos filtros.', "warning")
        return redirect(url_for("reportes_compras_proveedor"))

    totales = _totales_reporte(filas, proveedor_nombre=proveedor)
    # Ronda AC (2026-09-14): TODA la página (tarjetas de arriba, el
    # resumen que se recalcula en vivo por JS, y la columna "Total" de
    # cada producto en la tabla dinámica) usa la MISMA moneda dominante --
    # la que ya resolvió _totales_reporte para este proveedor+rango de
    # fechas -- para que nunca aparezcan dos monedas distintas en una
    # misma pantalla. Las celdas mes a mes siguen mostrando el monto
    # ORIGINAL de cada factura, sin convertir (así lo pidió el usuario).
    moneda_dominante = totales["moneda"]
    paridad_representativa = _paridad_representativa_eur(filas) if moneda_dominante == "EUR" else 1.0

    # --- Arma la tabla dinámica: producto (código de proveedor) x mes ---
    # Ronda AC (2026-09-14): el "Total" de cada producto (y los montos que
    # alimentan el resumen que se recalcula en vivo por JS) se acumulan ya
    # convertidos a `moneda_dominante` -- las celdas mes a mes siguen
    # guardando el monto ORIGINAL de cada factura (columna R, sin
    # convertir), tal como pidió el usuario.
    productos = {}
    meses_set = {}
    for f in filas:
        if not f["fecha_factura"]:
            continue
        clave_mes = (f["fecha_factura"].year, f["fecha_factura"].month)
        if clave_mes not in meses_set:
            nombre_mes = TipoCambioMensual.MESES_NOMBRE[f["fecha_factura"].month - 1][:3]
            meses_set[clave_mes] = f"{nombre_mes}-{str(f['fecha_factura'].year)[2:]}"

        codigo = (f.get("codigo_producto") or "").strip() or "(sin código)"
        prod = productos.setdefault(codigo, {
            "codigo": codigo,
            "descripcion": f.get("descripcion") or "",
            "facturas_por_moneda": defaultdict(set),
            "total_dominante": 0.0,
            "flete_dominante": 0.0,
            "derechos_dominante": 0.0,
            "otros_dominante": 0.0,
            "total_unidades": 0.0,
            "meses": defaultdict(lambda: {"total_moneda": 0.0, "moneda": None, "unidades": 0.0}),
            "facturas": [],
        })
        if not prod["descripcion"] and f.get("descripcion"):
            prod["descripcion"] = f["descripcion"]
        mf = f.get("moneda_operacion") or "USD"
        pf = f.get("paridad") or 1.0
        if f.get("factura"):
            prod["facturas_por_moneda"][mf].add(f["factura"])
        prod["total_dominante"] += _valor_en_moneda_dominante(f.get("total_usd"), mf, moneda_dominante, pf, paridad_representativa)
        prod["flete_dominante"] += _valor_en_moneda_dominante(f.get("flete_usd"), mf, moneda_dominante, pf, paridad_representativa)
        prod["derechos_dominante"] += _valor_en_moneda_dominante(f.get("derechos_usd"), mf, moneda_dominante, pf, paridad_representativa)
        prod["otros_dominante"] += _valor_en_moneda_dominante(f.get("otros_gastos_usd"), mf, moneda_dominante, pf, paridad_representativa)
        prod["total_unidades"] += f.get("unidades") or 0
        celda = prod["meses"][clave_mes]
        celda["total_moneda"] += f.get("total_invoice") or 0
        celda["unidades"] += f.get("unidades") or 0
        celda["moneda"] = mf if celda["moneda"] in (None, mf) else "MIXTO"
        prod["facturas"].append({
            "factura": f.get("factura") or "-",
            "fecha_factura": f["fecha_factura"],
            "unidades": f.get("unidades") or 0,
            "total_invoice": f.get("total_invoice") or 0,
            "moneda_operacion": mf,
            "origen": f.get("origen"),
        })

    meses_ordenados = sorted(meses_set.keys())
    meses_columnas = [
        {"clave": f"{a}-{m:02d}", "etiqueta": meses_set[(a, m)]}
        for (a, m) in meses_ordenados
    ]

    productos_lista = []
    for codigo, prod in productos.items():
        # Solo se marca "mezcla" cuando el propio producto tuvo facturas
        # REALES en ambas monedas dentro del rango (no por forzar todo a
        # una moneda pareja) -- caso raro, ver _paridad_efectiva_historica.
        es_mixto = len(prod["facturas_por_moneda"]) > 1
        prod["facturas"].sort(key=lambda x: x["fecha_factura"] or date.min, reverse=True)
        productos_lista.append({
            "codigo": prod["codigo"],
            "descripcion": prod["descripcion"],
            "es_mixto": es_mixto,
            "total_dominante": prod["total_dominante"],
            "flete_dominante": prod["flete_dominante"],
            "derechos_dominante": prod["derechos_dominante"],
            "otros_dominante": prod["otros_dominante"],
            "total_unidades": prod["total_unidades"],
            "meses": {
                f"{a}-{m:02d}": prod["meses"].get((a, m), {"total_moneda": 0.0, "moneda": moneda_dominante, "unidades": 0.0})
                for (a, m) in meses_ordenados
            },
            "facturas": prod["facturas"],
            "busqueda": prod["codigo"].lower(),
        })

    # Ronda AD (2026-09-14): para proveedores de lentes con catálogo de
    # código padre/variante (MEDICONTUR, BVI PHYSIOL -- ver
    # _mapa_padre_por_variante_proveedor), se pide mostrar PRIMERO el
    # código padre (la "familia" del lente) y, al entrar a ese padre, recién
    # ahí desplegar el código de producto/variante específico (dioptría).
    # Un producto de ese mismo proveedor que NO es un lente con variante
    # conocida (ej. inyectores, viscoelástico) sigue mostrándose suelto,
    # igual que antes.
    mapa_padre = _mapa_padre_por_variante_proveedor(proveedor)
    if mapa_padre:
        agrupados = {}
        sueltos = []
        for p in productos_lista:
            producto_padre = mapa_padre.get(p["codigo"].strip().upper())
            if not producto_padre:
                sueltos.append(p)
                continue
            grupo = agrupados.get(producto_padre.id)
            if grupo is None:
                grupo = {
                    "codigo": producto_padre.codigo,
                    "descripcion": producto_padre.descripcion,
                    "es_padre": True,
                    "es_mixto": False,
                    "total_dominante": 0.0,
                    "flete_dominante": 0.0,
                    "derechos_dominante": 0.0,
                    "otros_dominante": 0.0,
                    "total_unidades": 0.0,
                    "meses": {m["clave"]: {"total_moneda": 0.0, "moneda": moneda_dominante, "unidades": 0.0} for m in meses_columnas},
                    "facturas": [],
                    "variantes": [],
                    "busqueda": producto_padre.codigo.lower(),
                }
                agrupados[producto_padre.id] = grupo
            grupo["variantes"].append(p)
            grupo["es_mixto"] = grupo["es_mixto"] or p["es_mixto"]
            grupo["total_dominante"] += p["total_dominante"]
            grupo["flete_dominante"] += p["flete_dominante"]
            grupo["derechos_dominante"] += p["derechos_dominante"]
            grupo["otros_dominante"] += p["otros_dominante"]
            grupo["total_unidades"] += p["total_unidades"]
            for clave, celda in p["meses"].items():
                grupo["meses"][clave]["total_moneda"] += celda["total_moneda"] or 0
                grupo["meses"][clave]["unidades"] += celda.get("unidades") or 0
            grupo["facturas"].extend(p["facturas"])
            grupo["busqueda"] += " " + p["busqueda"]
        for grupo in agrupados.values():
            grupo["variantes"].sort(key=lambda v: -v["total_dominante"])
            grupo["facturas"].sort(key=lambda x: x["fecha_factura"] or date.min, reverse=True)
        productos_lista = list(agrupados.values()) + sueltos

    productos_lista.sort(key=lambda p: -p["total_dominante"])

    return render_template(
        "reportes/compras_proveedor_detalle.html",
        proveedor=proveedor, totales=totales,
        productos=productos_lista, meses_columnas=meses_columnas,
        agrupado_por_padre=bool(mapa_padre),
        empresas=Empresa.query.order_by(Empresa.nombre).all(),
        empresa_sel=empresa, fecha_desde=fecha_desde_txt, fecha_hasta=fecha_hasta_txt,
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
