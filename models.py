"""
Modelos de datos para la Suite Logística.
Modulo actual: Proveedores + Compras (Ordenes de Compra), con seguimiento
de etapas por línea de producto y documentos adjuntos por orden.
"""
from datetime import datetime
from urllib.parse import quote
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

# Permisos operativos disponibles (ronda R, 2026-09-12) -- cada Rol marca
# cuales de estos tiene. "es_administrador" (ver Rol abajo) es un permiso
# aparte, mas amplio: implica TODOS estos automaticamente ademas de poder
# administrar usuarios/roles/contraseñas -- ver Rol.tiene() abajo.
PERMISOS_DISPONIBLES = [
    ("crear_orden", "Creación de Orden"),
    ("aprobar_orden", "Aprobación de Orden"),
    ("actualizar_despacho", "Actualizar despacho"),
    ("generar_costeo", "Generar Costeo"),
    ("orden_simple", "Creación de Orden Simple (sin ver precios de proveedor)"),
    ("inventarios", "Inventarios"),
    ("reportes", "Acceso a Reportes"),
    ("seguimiento", "Seguimiento de despachos (solo ver estado/tracking)"),
    # Ronda V (2026-09-12): perfil amplio, pensado para colaboradores que NO
    # necesitan crear/aprobar ordenes ni ver costeos -- solo consultar cuanto
    # stock hay de cada producto. Deliberadamente separado del permiso
    # "inventarios" de arriba (ese es para quien administra/genera el costeo
    # de importaciones); este es de solo consulta de saldos de existencias.
    ("consultar_stock", "Consulta de Stock (saldos de existencias)"),
    # Ronda AJ (2026-09-18): modulo de Pago Proveedores (facturas
    # pendientes/pagadas, plazos de credito, registro de pagos y NC) --
    # separado de "inventarios"/"reportes" porque es informacion financiera
    # sensible (montos y fechas de vencimiento de pago a proveedores) que no
    # todo el que ve Reportes deberia poder ver/operar.
    ("pagos_proveedores", "Pago Proveedores"),
]


class Rol(db.Model):
    """Perfil de acceso (ronda R, 2026-09-12): un conjunto de permisos
    operativos que se asigna a uno o mas Usuarios. "es_administrador" es un
    interruptor aparte -- un Rol administrador tiene TODOS los permisos
    automaticamente (ver tiene() abajo) ademas de poder entrar a
    Configuración > Usuarios y Perfiles para administrar cuentas, sin
    necesidad de tildar cada permiso individual por separado."""

    __tablename__ = "roles"

    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(80), unique=True, nullable=False)
    es_administrador = db.Column(db.Boolean, default=False)
    permiso_crear_orden = db.Column(db.Boolean, default=False)
    permiso_aprobar_orden = db.Column(db.Boolean, default=False)
    permiso_actualizar_despacho = db.Column(db.Boolean, default=False)
    permiso_generar_costeo = db.Column(db.Boolean, default=False)
    permiso_orden_simple = db.Column(db.Boolean, default=False)
    permiso_inventarios = db.Column(db.Boolean, default=False)
    permiso_reportes = db.Column(db.Boolean, default=False)
    permiso_seguimiento = db.Column(db.Boolean, default=False)
    # Ronda V (2026-09-12): consulta de saldos de Stock -- ver PERMISOS_
    # DISPONIBLES arriba.
    permiso_consultar_stock = db.Column(db.Boolean, default=False)
    # Ronda AJ (2026-09-18): ver PERMISOS_DISPONIBLES arriba.
    permiso_pagos_proveedores = db.Column(db.Boolean, default=False)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    usuarios = db.relationship("Usuario", backref="rol", lazy="dynamic")

    def tiene(self, permiso):
        """True si este Rol puede ejercer `permiso` (uno de los codigos de
        PERMISOS_DISPONIBLES) -- un Rol administrador siempre da True."""
        if self.es_administrador:
            return True
        return bool(getattr(self, f"permiso_{permiso}", False))

    @property
    def permisos_activos(self):
        return [nombre for codigo, nombre in PERMISOS_DISPONIBLES if self.tiene(codigo)]

    def __repr__(self):
        return f"<Rol {self.nombre}>"


class Usuario(UserMixin, db.Model):
    """Cuenta de acceso a la app (ronda R, 2026-09-12). El login se hace con
    email + contraseña (hash, nunca en texto plano); cada usuario tiene UN
    Rol que determina que puede hacer (ver Rol.tiene() arriba)."""

    __tablename__ = "usuarios"

    id = db.Column(db.Integer, primary_key=True)
    nombre_completo = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    rol_id = db.Column(db.Integer, db.ForeignKey("roles.id"), nullable=False)
    activo = db.Column(db.Boolean, default=True)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)
    ultimo_acceso = db.Column(db.DateTime, nullable=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return bool(self.password_hash) and check_password_hash(self.password_hash, password)

    def tiene_permiso(self, permiso):
        return bool(self.rol and self.rol.tiene(permiso))

    @property
    def is_active(self):
        # Sobrescribe el default de UserMixin (siempre True) -- una cuenta
        # desactivada desde Configuración > Usuarios no puede loguearse ni
        # mantener una sesion ya abierta.
        return self.activo

    def __repr__(self):
        return f"<Usuario {self.email}>"


class Empresa(db.Model):
    """Empresa compradora (ej. Accuvision, Accumedical -- ronda K,
    2026-09-07): sus datos (nombre, dirección, logo) se imprimen como
    "Comprador" en el PDF de la Orden de Compra. Cada empresa lleva su
    propio correlativo de N° de OC (ver siguiente_numero_po en app.py),
    unico por empresa sin importar a que proveedor le compre."""

    __tablename__ = "empresas"

    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(120), unique=True, nullable=False)
    direccion = db.Column(db.String(255))
    rut = db.Column(db.String(40))
    telefono = db.Column(db.String(60))
    email = db.Column(db.String(120))
    logo_nombre_archivo = db.Column(db.String(255))  # archivo en disco, ver EMPRESAS_LOGOS_DIR
    activo = db.Column(db.Boolean, default=True)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    # Plantilla de correo para el envío de la OC por Outlook (ronda L, punto 1,
    # 2026-09-09): si quedan vacíos, se usa una plantilla por defecto -- ver
    # DEFAULT_ASUNTO_CORREO / DEFAULT_CUERPO_CORREO y construir_correo_oc() en
    # app.py. Los usuarios escriben aquí con marcadores tipo [PROVEEDOR],
    # [NUMERO_OC], etc. que el programa reemplaza por los datos reales de cada
    # orden -- así no hay que editar nada a mano en Outlook cada vez.
    plantilla_asunto_correo = db.Column(db.String(200))
    plantilla_cuerpo_correo = db.Column(db.Text)

    ordenes = db.relationship("OrdenCompra", backref="empresa", lazy="dynamic")

    def __repr__(self):
        return f"<Empresa {self.nombre}>"


class Proveedor(db.Model):
    __tablename__ = "proveedores"

    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(120), unique=True, nullable=False)
    tipo = db.Column(db.String(20), nullable=False, default="Extranjero")  # Nacional / Extranjero
    pais = db.Column(db.String(80))
    contacto_nombre = db.Column(db.String(120))
    contacto_telefono = db.Column(db.String(60))
    contacto_email = db.Column(db.String(120))
    direccion = db.Column(db.String(255))
    moneda_default = db.Column(db.String(10), default="USD")
    notas = db.Column(db.Text)
    activo = db.Column(db.Boolean, default=True)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)
    # Codigo/numero que identifica a este proveedor DENTRO del sistema de
    # Inventarios del usuario (ronda O, 2026-09-12) -- no es su RUT real, es
    # un ID interno de ese otro sistema que el usuario asocia manualmente una
    # vez y que despues se reutiliza solo en la planilla de exportacion
    # "Inventario" (columna "RUT" de la cabecera, ver _construir_excel_inventario
    # en app.py). Queda en blanco hasta que el usuario lo completa.
    codigo_sistema_inventario = db.Column(db.String(50))

    # Ronda AJ (2026-09-18): plazo de credito pactado con este proveedor,
    # para el modulo de Pago Proveedores -- en dias corridos desde la fecha
    # de emision de la factura (ver FacturaProveedor.fecha_vencimiento_
    # sugerida abajo). None/0 = sin plazo definido todavia. Si
    # requiere_pago_previo esta marcado, el proveedor no despacha hasta
    # recibir el pago (factura queda "pendiente" igual, pero se muestra
    # resaltada en el listado como "pago previo").
    plazo_credito_dias = db.Column(db.Integer, nullable=True)
    requiere_pago_previo = db.Column(db.Boolean, default=False)

    productos = db.relationship(
        "Producto", backref="proveedor", cascade="all, delete-orphan", lazy="dynamic"
    )
    ordenes = db.relationship(
        "OrdenCompra", backref="proveedor", cascade="all, delete-orphan", lazy="dynamic"
    )

    def __repr__(self):
        return f"<Proveedor {self.nombre}>"


class Producto(db.Model):
    __tablename__ = "productos"

    id = db.Column(db.Integer, primary_key=True)
    proveedor_id = db.Column(db.Integer, db.ForeignKey("proveedores.id"), nullable=False)
    codigo = db.Column(db.String(120), nullable=False)
    descripcion = db.Column(db.String(500), nullable=False)
    empaque = db.Column(db.Integer, default=1)  # unidades por caja
    moneda = db.Column(db.String(10), default="USD")
    precio_caja = db.Column(db.Float, default=0)
    precio_unitario = db.Column(db.Float, default=0)
    activo = db.Column(db.Boolean, default=True)

    # Ronda V (2026-09-12): codigo interno con el que este producto se
    # identifica en el OTRO sistema de la empresa (ventas/inventario,
    # "Ergopyme") -- se homologa una vez (ver seed_homologacion_y_stock_
    # inicial en app.py, a partir de "Inventarios y codigos interno sistema
    # Inventarios.xlsx") y de ahi en adelante permite: 1) relacionar cada
    # carga nueva del reporte de stock con nuestro catalogo sin volver a
    # pedir el cruce completo, y 2) completar la columna CODIGO de la
    # planilla de descarga del Costeo (ver _construir_excel_inventario)
    # cuando el checkbox "Código Proveedor" esta SIN marcar -- antes esa
    # columna quedaba siempre en blanco porque este codigo era desconocido.
    # Nunca se muestra en pantallas de Ordenes de Compra.
    codigo_interno_inventario = db.Column(db.String(40))

    # Nota: el Excel origen trae codigos repetidos para variantes/paquetes de
    # un mismo proveedor (ej. distintas configuraciones bajo el mismo codigo
    # nominal), por lo que aqui NO se fuerza unicidad de codigo por proveedor.
    __table_args__ = (
        db.Index("ix_producto_proveedor_codigo", "proveedor_id", "codigo"),
    )

    def __repr__(self):
        return f"<Producto {self.codigo}>"


class ProductoVariante(db.Model):
    """Variante de un producto "padre" del catalogo (ronda M, 2026-09-10):
    hoy usado para los lentes intraoculares de MEDICONTUR, donde el
    catalogo guarda UN solo Producto por familia (ej. codigo padre
    "677ADY", con un solo precio) pero cada Orden de Compra real necesita
    el codigo especifico de dioptria/variante (ej. "677ADYP360"). En vez
    de cargar cada variante como un Producto separado en el catalogo (a
    pedido explicito del usuario -- inflaria la busqueda con cientos de
    codigos que no tienen precio propio), las variantes se guardan aparte
    y se eligen desde un modal al agregar la linea de la orden -- el
    precio de la linea SIEMPRE es el del Producto padre, nunca uno propio
    de la variante. Ver ProductoVariante en la ronda M para el detalle
    completo del flujo."""
    __tablename__ = "producto_variantes"

    id = db.Column(db.Integer, primary_key=True)
    producto_id = db.Column(db.Integer, db.ForeignKey("productos.id"), nullable=False)
    codigo = db.Column(db.String(120), nullable=False)
    descripcion = db.Column(db.String(500), nullable=False)

    # Ronda V (2026-09-12): mismo proposito que Producto.codigo_interno_
    # inventario, pero a nivel de variante -- en el sistema de Inventarios
    # cada dioptria/variante especifica es su propio codigo interno (el
    # producto "padre" de nuestro catalogo no es una unidad fisica real).
    codigo_interno_inventario = db.Column(db.String(40))

    producto = db.relationship("Producto", backref=db.backref("variantes", lazy="dynamic"))

    def __repr__(self):
        return f"<ProductoVariante {self.codigo}>"


class HomologacionStock(db.Model):
    """Ronda V (2026-09-12): tabla de homologacion -- una fila por cada
    codigo interno del sistema de Inventarios ('Ergopyme'), con la
    clasificacion que se le dio la primera vez que se vio (ver
    seed_homologacion_y_stock_inicial en app.py). Una vez clasificado, cada
    carga NUEVA del reporte de stock (que YA NO trae columnas de codigo de
    proveedor, solo el formato original) usa esta tabla para saber que
    hacer con cada codigo, sin tener que volver a pedir el cruce completo:
    - 'vinculado': esta ligado a un Producto (o a una ProductoVariante,
      cuando el producto es de una familia con variantes) de nuestro
      catalogo -- ver producto_id/variante_id.
    - 'sin_marca': se carga en el Stock para consulta, pero NO se crea ni
      se liga a ningun Producto del catalogo de proveedores (a pedido
      explicito del usuario -- son productos sin marca/generico).
    - 'excluido': se descarta siempre, no se carga ni al Stock (ronda U:
      equivalente a los codigos marcados "NO INCLUIR" en el cruce inicial).
    - 'pendiente': codigo nuevo que aparecio en una carga posterior y que
      todavia no se sabe a que producto/proveedor corresponde -- igual se
      carga al Stock (sin producto/variante asociado) para no perder el
      dato, y queda listado en /stock/homologacion para resolverlo a mano.
    """
    __tablename__ = "homologaciones_stock"

    id = db.Column(db.Integer, primary_key=True)
    codigo_interno = db.Column(db.String(40), unique=True, nullable=False)
    estado = db.Column(db.String(20), nullable=False, default="pendiente")
    producto_id = db.Column(db.Integer, db.ForeignKey("productos.id"), nullable=True)
    variante_id = db.Column(db.Integer, db.ForeignKey("producto_variantes.id"), nullable=True)
    # Ultima descripcion vista para este codigo en el reporte de stock -- útil
    # para mostrarla en /stock/homologacion cuando no hay Producto asociado.
    descripcion_referencia = db.Column(db.String(300))
    actualizado_en = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    producto = db.relationship("Producto")
    variante = db.relationship("ProductoVariante")

    def __repr__(self):
        return f"<HomologacionStock {self.codigo_interno} ({self.estado})>"


class StockExistencia(db.Model):
    """Ronda V (2026-09-12): saldo de existencias por empresa, tal como se
    ve en el sistema de Inventarios ('Ergopyme') -- una fila por cada
    combinacion producto/variante + lote + empresa que trae el reporte. Se
    recarga POR COMPLETO cada vez que alguien sube un reporte nuevo (ver
    stock_cargar en app.py): se borran todas las filas anteriores y se
    insertan las del archivo nuevo, para que el saldo mostrado sea siempre
    el de la ultima foto real, sin arrastrar datos viejos ni duplicarlos."""
    __tablename__ = "stock_existencias"

    id = db.Column(db.Integer, primary_key=True)
    empresa_id = db.Column(db.Integer, db.ForeignKey("empresas.id"), nullable=False)
    producto_id = db.Column(db.Integer, db.ForeignKey("productos.id"), nullable=True)
    variante_id = db.Column(db.Integer, db.ForeignKey("producto_variantes.id"), nullable=True)
    codigo_interno = db.Column(db.String(40), nullable=False)
    descripcion = db.Column(db.String(300))
    codigo_lote = db.Column(db.String(80))
    fecha_vencimiento = db.Column(db.Date, nullable=True)
    stock_fisico = db.Column(db.Integer, default=0)
    cargado_en = db.Column(db.DateTime, default=datetime.utcnow)

    empresa = db.relationship("Empresa")
    producto = db.relationship("Producto")
    variante = db.relationship("ProductoVariante")

    def __repr__(self):
        return f"<StockExistencia {self.codigo_interno} x{self.stock_fisico}>"


class StockValorizado(db.Model):
    """Ronda AM (2026-09-19): valorización de inventario (costo/valor) por
    código, tal como viene del reporte "Stock General Consolidado" del
    sistema de Inventarios (Ergopyme: Logística > Control de Existencias >
    Informes > Stock General) -- una fila por código con el detalle de
    Accuvision, Accumedical y el total combinado. A diferencia de
    StockExistencia (que no trae costos y sí trae lote/vencimiento), este
    reporte NO trae lote -- es un consolidado por código -- y SÍ trae
    costo/valor de inventario, por eso tiene su propia pantalla con su
    propio permiso ("reportes"): los usuarios que solo tienen acceso a
    Consulta de Stock (permiso "consultar_stock") no deben ver esta data.

    El PMP (precio medio ponderado) de cada bloque se calcula acá como
    Valor Final / Stock Final -- el usuario lo pidió así explícitamente en
    vez de usar la columna P.M.P. que el propio Ergopyme ya trae en el
    archivo (aunque en la práctica coinciden, salvo redondeo). Ver
    _fila_stock_valorizado_pmp en app.py.

    Se recarga POR COMPLETO cada vez que se sube un reporte nuevo (igual
    que StockExistencia): se borran todas las filas anteriores y se
    insertan las del archivo nuevo."""
    __tablename__ = "stock_valorizado"

    id = db.Column(db.Integer, primary_key=True)
    codigo_interno = db.Column(db.String(40), nullable=False, index=True)
    descripcion = db.Column(db.String(300))
    unidad_medida = db.Column(db.String(20))

    # Bloque ACCUV del reporte (columnas G=Stock Final, L=Valor Final)
    stock_accuvision = db.Column(db.Integer, default=0)
    valor_accuvision = db.Column(db.Float, default=0)
    pmp_accuvision = db.Column(db.Float, nullable=True)

    # Bloque ACCUM del reporte (columnas Q=Stock Final, V=Valor Final --
    # OJO: no son AA/AF, esas son el bloque TOT, ver nota de ronda AM)
    stock_accumedical = db.Column(db.Integer, default=0)
    valor_accumedical = db.Column(db.Float, default=0)
    pmp_accumedical = db.Column(db.Float, nullable=True)

    # Bloque TOT del reporte (columnas AA=Stock Final, AF=Valor Final)
    stock_total = db.Column(db.Integer, default=0)
    valor_total = db.Column(db.Float, default=0)
    pmp_total = db.Column(db.Float, nullable=True)

    cargado_en = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f"<StockValorizado {self.codigo_interno}>"


class PedidoComprometido(db.Model):
    """Ronda X (2026-09-13, punto 3C): unidades ya comprometidas con
    clientes (OC de clientes, no de nosotros a los proveedores), tal como
    vienen en el reporte "Notas de pedido" del sistema de Inventarios --
    una fila por cada linea de pedido de cliente todavia con saldo
    pendiente de despacho. Se recarga POR EMPRESA cada vez que se sube un
    archivo nuevo (ver stock_pedidos_cargar en app.py): se borran solo las
    filas de la empresa que trae ESE archivo (identificada por el
    encabezado del reporte, igual que el reporte de Stock) y se insertan
    las nuevas -- si mas adelante se sube el equivalente de la otra
    empresa, no se pisan entre si."""
    __tablename__ = "pedidos_comprometidos"

    id = db.Column(db.Integer, primary_key=True)
    empresa_id = db.Column(db.Integer, db.ForeignKey("empresas.id"), nullable=False)
    codigo_interno = db.Column(db.String(40), nullable=False)
    descripcion = db.Column(db.String(300))
    nro_pedido = db.Column(db.String(40))
    fecha_pedido = db.Column(db.Date, nullable=True)
    cliente_nombre = db.Column(db.String(300))
    cantidad = db.Column(db.Integer, default=0)
    cargado_en = db.Column(db.DateTime, default=datetime.utcnow)

    empresa = db.relationship("Empresa")

    def __repr__(self):
        return f"<PedidoComprometido {self.codigo_interno} x{self.cantidad}>"


class CodigoErgopyme(db.Model):
    """Ronda AA (2026-09-13): tabla maestra de homologación Proveedor +
    Código Proveedor + Descripción para cada "código interno" del sistema
    de Inventarios (Ergopyme), cargada una vez desde el archivo "2da
    Revisión códigos ergopyme" que mantiene el usuario. Es MÁS COMPLETA
    que HomologacionStock (ronda V-Z, que solo cubre los códigos que
    aparecieron alguna vez en una carga del reporte de Stock) -- se usa
    ÚNICAMENTE para homologar el histórico de compras/importaciones del
    reporte "Compras Proveedor" (ver seed_codigos_ergopyme y
    reportes_compras_proveedor en app.py). No reemplaza ni modifica
    HomologacionStock, que sigue siendo la fuente de verdad para Consulta
    de Stock."""
    __tablename__ = "codigos_ergopyme"

    id = db.Column(db.Integer, primary_key=True)
    codigo_interno = db.Column(db.String(40), unique=True, nullable=False, index=True)
    proveedor_nombre = db.Column(db.String(120))
    codigo_proveedor = db.Column(db.String(120))
    descripcion = db.Column(db.String(500))
    actualizado_en = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f"<CodigoErgopyme {self.codigo_interno} -> {self.proveedor_nombre}>"


class CompraHistorica(db.Model):
    """Ronda AA (2026-09-13): histórico "congelado" de compras/importaciones
    a proveedores anteriores a este sistema, cargado UNA SOLA VEZ desde el
    archivo "Data Histórica Compra proveedores, costo fletes y gastos
    Importación" -- una fila por cada línea de producto de cada factura
    (6.135 filas reales al 2026-09-13, de 2023-11 a 2026-08). Se homologa
    al cargarse contra CodigoErgopyme (arriba) para saber el proveedor
    "real" (mismo nombre que usa el catálogo de Proveedor de este sistema)
    y el código de ese proveedor -- homologado=False cuando el código no
    se encontró en CodigoErgopyme (queda igual con el proveedor tal como
    venía en el archivo). El reporte "Compras Proveedor" (ver app.py)
    combina estas filas con las compras hechas DESDE la plataforma
    (calculadas en vivo a partir de Importacion/Parcial/ParcialLinea), así
    que de acá en adelante cada Orden de Compra/Costeo que se haga en el
    sistema se va sumando solo al mismo reporte, sin tener que recargar
    nada."""
    __tablename__ = "compras_historicas"

    id = db.Column(db.Integer, primary_key=True)
    fecha_factura = db.Column(db.Date, nullable=True)
    mes_anio = db.Column(db.String(20))
    proveedor_original = db.Column(db.String(120))
    proveedor_homologado = db.Column(db.String(120))
    factura = db.Column(db.String(120))
    tipo_cambio = db.Column(db.Float, default=0)      # CLP por 1 USD (Dólar Aduanero)
    paridad_eur = db.Column(db.Float, default=0)      # unidades de la moneda de factura por 1 USD
    transporte = db.Column(db.String(80))
    codigo_interno = db.Column(db.String(40), index=True)
    codigo_proveedor = db.Column(db.String(120))
    descripcion = db.Column(db.String(500))
    tipo_flete = db.Column(db.String(40))
    unidades = db.Column(db.Float, default=0)
    total_invoice = db.Column(db.Float, default=0)    # en la moneda original de la factura
    total_usd = db.Column(db.Float, default=0)
    flete_usd = db.Column(db.Float, default=0)
    seguro_usd = db.Column(db.Float, default=0)
    cif_usd = db.Column(db.Float, default=0)
    cif_clp = db.Column(db.Float, default=0)
    derechos_clp = db.Column(db.Float, default=0)
    otros_gastos_clp = db.Column(db.Float, default=0)
    costo_total_clp = db.Column(db.Float, default=0)
    costo_unitario_clp = db.Column(db.Float, default=0)
    categoria = db.Column(db.String(80))              # columna "PRODUCTO" del archivo (INSTRUMENTALES/INSUMOS/...)
    otros_costos_usd = db.Column(db.Float, default=0)
    empresa_compradora = db.Column(db.String(80))     # ACCUVISION / ACCUMEDICAL, tal como viene en el archivo
    homologado = db.Column(db.Boolean, default=False)
    # Ronda AK (2026-09-19): cuando una linea "en consignacion" (factura ==
    # "CONSIGNACION", ver es_consignacion en _fila_historica_dict) se factura
    # de verdad -- ya sea a mano desde Pago Proveedores > Facturar
    # consignacion, o por la carga masiva del archivo del proveedor -- queda
    # apuntando a la FacturaProveedor real que la reemplaza, para no volver a
    # ofrecerla como "pendiente de facturar" ni facturarla dos veces.
    factura_generada_id = db.Column(db.Integer, db.ForeignKey("facturas_proveedor.id"), nullable=True)
    # Ronda AL (2026-09-19, punto 1 del pedido del usuario): True cuando esta
    # linea llego alguna vez como consignacion (factura == "CONSIGNACION") y
    # DESPUES se facturo de verdad -- ya sea por la reconciliacion masiva del
    # archivo del proveedor o por Pago Proveedores > Facturar consignacion.
    # Se usa SOLO para sombrear esas filas en el reporte Compras Proveedor
    # (ver reportes_compras_proveedor_detalle) y diferenciarlas visualmente
    # de una factura de compra directa -- no cambia ningun calculo, es
    # puramente informativo. Una linea que SIGUE en consignacion (factura ==
    # "CONSIGNACION" todavia) tiene este campo en False -- ver es_consignacion
    # en _fila_historica_dict, que se calcula de la columna factura, no de este
    # campo.
    fue_consignacion = db.Column(db.Boolean, default=False)

    def __repr__(self):
        return f"<CompraHistorica {self.codigo_interno} {self.factura}>"


# Etapas por las que avanza cada LINEA de producto de una orden, hasta la
# llegada a bodega. El estado de la orden se calcula automaticamente a
# partir de las etapas de sus lineas (ver OrdenCompra.estado_calculado).
ETAPAS_LINEA = [
    "Emisión de Orden",
    "Orden Confirmada",
    "Orden Despachada",
    "Internación Aduanas",
    "Recibido",
]

ESTADOS_OC = ETAPAS_LINEA + ["En proceso (mixto)", "Cancelada", "Anulada"]

# Puerta de aprobacion de la orden COMPLETA (ronda T, 2026-09-12), separada
# de 'estado'/'etapa' (que rastrean el avance linea por linea DESPUES de
# aprobada). Toda orden NUEVA (cualquiera sea el perfil que la cree) nace en
# "Por Aprobar" y no es todavia una Orden de Compra definitiva: no aparece en
# el listado general de Compras/Ordenes ("los registros"), solo en la vista
# de "Mis ordenes"/"Por Aprobar" de quien la creo. Alguien con permiso de
# Aprobacion la pasa a "Aprobada" (desde ahi sigue el mismo flujo de siempre
# por etapas de linea) o la "devuelve", dejandola en "Sin Emitir" -- visible
# SOLO para quien la creo, quien puede editarla y reenviarla (vuelve a "Por
# Aprobar"). Las ordenes que ya existian antes de esta ronda quedan
# "Aprobada" por defecto via la migracion (ver ensure_schema_migrations en
# app.py), para no ocultar de golpe ningun dato ya en curso.
ESTADOS_APROBACION_OC = ["Por Aprobar", "Aprobada", "Sin Emitir"]

# Una vez que una linea llega a "Orden Despachada" (o mas adelante), el
# usuario ya no quiere poder anularla, editarla ni retrocederla -- lo
# unico que se puede hacer es seguir avanzando de etapa (o subir/gestionar
# documentos del despacho). Antes de despachar, la linea sigue siendo
# libremente editable.
ETAPAS_LINEA_BLOQUEADA = {"Orden Despachada", "Internación Aduanas", "Recibido"}

TIPOS_DOCUMENTO_ORDEN = [
    "Cotización",
    "Orden de Compra (proveedor)",
    "Invoice",
    "Guía de Despacho",
    "Certificado de Origen",
    "Factura Agente de Aduanas",
    "Gastos Internación",
    "Otro",
]


class OrdenCompra(db.Model):
    __tablename__ = "ordenes_compra"

    id = db.Column(db.Integer, primary_key=True)
    # NOTA: numero_po ya NO es unico. Cuando se confirma solo una parte de
    # los productos de una orden, esta se divide automaticamente en dos
    # (una con lo confirmado, otra con lo pendiente) conservando AMBAS el
    # mismo numero de PO -- asi es como el proveedor del usuario identifica
    # los despachos parciales de una misma orden de compra.
    numero_po = db.Column(db.String(40), nullable=False)
    proveedor_id = db.Column(db.Integer, db.ForeignKey("proveedores.id"), nullable=False)
    # Empresa compradora (Accuvision / Accumedical, ronda K, 2026-09-07):
    # nullable porque las ordenes creadas ANTES de esta mejora no tienen
    # ninguna asignada todavia -- el usuario pidio dejarlas "Sin asignar" y
    # asignarlas el mismo, una por una, en vez de forzar un valor por
    # defecto. El correlativo de numero_po (ver siguiente_numero_po en
    # app.py) es UNICO POR EMPRESA (una sola secuencia por empresa, sin
    # importar a que proveedor le compre) -- confirmado explicitamente por
    # el usuario.
    empresa_id = db.Column(db.Integer, db.ForeignKey("empresas.id"), nullable=True)
    fecha_emision = db.Column(db.Date, default=datetime.utcnow)
    moneda = db.Column(db.String(10), default="USD")
    # 'estado' se mantiene sincronizado automaticamente con las etapas de las
    # lineas (ver recalcular_estado en app.py); solo se edita manualmente
    # para cancelar la orden.
    estado = db.Column(db.String(30), default="Emisión de Orden")
    notas = db.Column(db.Text)
    # Campos preparados para el modulo de despachos consolidados (futuro,
    # no se uso al final -- ver Despacho.ordenes en su lugar)
    orden_consolidada_id = db.Column(db.Integer, db.ForeignKey("ordenes_compra.id"), nullable=True)
    # Despacho fisico (courier/tracking) al que quedo asociada esta orden una
    # vez despachada. Una orden se asocia a lo sumo a UN despacho; un
    # despacho puede agrupar VARIAS ordenes (despacho consolidado).
    despacho_id = db.Column(db.Integer, db.ForeignKey("despachos.id"), nullable=True)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)
    # Autoria (ronda S, 2026-09-12): quien creo la orden -- nullable porque
    # las ordenes creadas ANTES de esta ronda no tienen usuario asociado (el
    # sistema de login no existia). Se usa para que un usuario con el perfil
    # "Creación de Orden Simple" solo pueda ver/editar/anular SUS PROPIAS
    # ordenes (ver ordenes_simple_* en app.py), nunca las de otro usuario.
    creado_por_usuario_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    creado_por = db.relationship("Usuario", foreign_keys=[creado_por_usuario_id])
    # Puerta de aprobacion de la orden completa (ronda T, 2026-09-12): ver
    # ESTADOS_APROBACION_OC mas arriba. Default "Aprobada" para que la
    # migracion automatica no oculte ninguna orden ya existente.
    estado_aprobacion = db.Column(db.String(20), default="Aprobada")
    # Consignación (ronda AH, 2026-09-16): la orden COMPLETA (no por línea)
    # se marca en consignación desde su emisión -- esos productos llegan con
    # una pro-forma invoice y el proveedor solo emite la factura real
    # semanas/meses después, cuando se le informa qué se vendió (ver
    # es_consignacion en _fila_historica_dict, que hoy solo detecta
    # consignación en el histórico de Excel por la columna Factura -- esta
    # marca es el equivalente para órdenes creadas DESDE la plataforma).
    # Default False para no afectar ninguna orden ya existente.
    es_consignacion = db.Column(db.Boolean, default=False)

    lineas = db.relationship(
        "OrdenCompraLinea", backref="orden", cascade="all, delete-orphan", lazy="dynamic"
    )
    documentos = db.relationship(
        "OrdenDocumento", backref="orden", cascade="all, delete-orphan", lazy="dynamic",
        order_by="OrdenDocumento.fecha_subida.desc()",
    )

    @property
    def total(self):
        return sum(l.subtotal for l in self.lineas)

    @property
    def total_unidades_pendientes(self):
        return sum(l.cantidad_pendiente for l in self.lineas)

    @property
    def resumen_etapas(self):
        """Cuenta cuantas lineas activas (no anuladas) hay en cada etapa.
        Util para mostrar el avance real de la orden."""
        conteo = {e: 0 for e in ETAPAS_LINEA}
        for l in self.lineas:
            if l.anulada:
                continue
            if l.etapa in conteo:
                conteo[l.etapa] += 1
        return conteo

    @property
    def total_anuladas(self):
        return sum(1 for l in self.lineas if l.anulada)

    @property
    def total_confirmadas(self):
        """Lineas activas que ya superaron 'Emision de Orden' (confirmadas por el proveedor)."""
        return sum(
            1 for l in self.lineas
            if not l.anulada and l.etapa != ETAPAS_LINEA[0]
        )

    @property
    def total_pendientes_confirmacion(self):
        return sum(
            1 for l in self.lineas
            if not l.anulada and l.etapa == ETAPAS_LINEA[0]
        )

    def __repr__(self):
        return f"<OC {self.numero_po}>"


class OrdenCompraLinea(db.Model):
    __tablename__ = "ordenes_compra_lineas"

    id = db.Column(db.Integer, primary_key=True)
    orden_id = db.Column(db.Integer, db.ForeignKey("ordenes_compra.id"), nullable=False)
    producto_id = db.Column(db.Integer, db.ForeignKey("productos.id"), nullable=False)

    cantidad_cajas = db.Column(db.Integer, default=0)
    precio_unitario_pactado = db.Column(db.Float, default=0)  # precio por caja pactado

    # Planificacion (modulo Compras)
    fecha_estimada_despacho = db.Column(db.Date, nullable=True)

    # Confirmacion del proveedor
    fecha_disponibilidad_confirmada = db.Column(db.Date, nullable=True)
    fecha_estimada_despacho_confirmada = db.Column(db.Date, nullable=True)

    # Etapa de avance de esta linea hasta la llegada a bodega
    etapa = db.Column(db.String(30), default="Emisión de Orden")
    fecha_despacho_real = db.Column(db.Date, nullable=True)
    fecha_llegada_aduana = db.Column(db.Date, nullable=True)
    fecha_recepcion_bodega = db.Column(db.Date, nullable=True)

    # Control de despacho parcial
    cantidad_despachada_cajas = db.Column(db.Integer, default=0)

    # Linea anulada: se saca de circulacion (no cuenta como pendiente, no
    # avanza mas de etapa), pero se conserva para no perder el historial.
    anulada = db.Column(db.Boolean, default=False)

    # Variante elegida para esta linea (ronda M, 2026-09-10) -- cuando el
    # producto padre tiene variantes (ver ProductoVariante), aqui se guarda
    # una FOTO del codigo/descripcion de la variante especifica elegida
    # (ej. la dioptria exacta de un lente), independiente de que esa
    # variante se edite o borre despues en el catalogo -- mismo patron
    # "snapshot" que ParcialLinea.codigo/descripcion. El precio de la linea
    # sigue siendo siempre el del Producto padre (precio_unitario_pactado),
    # nunca uno propio de la variante.
    variante_codigo = db.Column(db.String(120))
    variante_descripcion = db.Column(db.String(500))

    # Ronda S (2026-09-12): True cuando esta linea la agrego un usuario con
    # el perfil "Creación de Orden Simple" reutilizando un producto QUE YA
    # EXISTIA en el catálogo del proveedor -- ese precio es el que ya tenia
    # cargado el catálogo (no el que tecleo ese usuario, que se descarta), y
    # por eso debe seguir oculto para el aun cuando el vea el detalle de SU
    # PROPIA orden (ver ordenes/simple_detalle.html). Si en cambio el
    # producto era nuevo, el precio es el que el mismo tecleo -- no hace
    # falta ocultarselo a si mismo, asi que queda en False.
    precio_catalogo_oculto = db.Column(db.Boolean, default=False)

    # Ronda U (2026-09-12): True cuando esta linea fue la que dio de alta un
    # producto NUEVO en el catálogo del proveedor (perfil "Creación de Orden
    # Simple", código que no existía todavía) -- ese producto nace inactivo
    # (Producto.activo=False, ver ordenes_simple_nueva/ordenes_simple_linea_
    # nueva en app.py) para no ensuciar el catálogo con datos sin confirmar
    # mientras la orden esté 'Por Aprobar'/'Sin Emitir'. Al aprobar la orden
    # (ver ordenes_aprobar) se activa automáticamente. Si la línea reutilizó
    # un producto que YA existía, queda en False -- no hay nada que activar.
    producto_creado_por_esta_orden = db.Column(db.Boolean, default=False)

    producto = db.relationship("Producto")

    @property
    def codigo_mostrar(self):
        """Codigo a mostrar en pantalla/PDF/Costeo: el de la variante
        elegida si hay una, si no el del producto padre."""
        return self.variante_codigo or self.producto.codigo

    @property
    def descripcion_mostrar(self):
        """Descripcion a mostrar en pantalla/PDF/Costeo: la de la variante
        elegida si hay una, si no la del producto padre."""
        return self.variante_descripcion or self.producto.descripcion

    @property
    def cantidad_unidades(self):
        empaque = self.producto.empaque or 1
        return (self.cantidad_cajas or 0) * empaque

    @property
    def subtotal(self):
        return (self.cantidad_cajas or 0) * (self.precio_unitario_pactado or 0)

    @property
    def cantidad_pendiente(self):
        if self.anulada:
            return 0
        return max((self.cantidad_cajas or 0) - (self.cantidad_despachada_cajas or 0), 0)

    @property
    def despacho_parcial(self):
        despachado = self.cantidad_despachada_cajas or 0
        return 0 < despachado < (self.cantidad_cajas or 0)

    @property
    def siguiente_etapa(self):
        try:
            idx = ETAPAS_LINEA.index(self.etapa)
        except ValueError:
            return None
        if idx + 1 < len(ETAPAS_LINEA):
            return ETAPAS_LINEA[idx + 1]
        return None

    @property
    def etapa_anterior(self):
        try:
            idx = ETAPAS_LINEA.index(self.etapa)
        except ValueError:
            return None
        if idx - 1 >= 0:
            return ETAPAS_LINEA[idx - 1]
        return None

    @property
    def bloqueada(self):
        """True si la linea ya fue despachada (o mas adelante): ya no se
        puede anular, editar ni retroceder, solo seguir avanzando de etapa
        o gestionar sus documentos."""
        return self.etapa in ETAPAS_LINEA_BLOQUEADA


class OrdenDocumento(db.Model):
    __tablename__ = "ordenes_documentos"

    id = db.Column(db.Integer, primary_key=True)
    orden_id = db.Column(db.Integer, db.ForeignKey("ordenes_compra.id"), nullable=False)
    tipo = db.Column(db.String(60), default="Otro")
    nombre_original = db.Column(db.String(255), nullable=False)
    nombre_archivo = db.Column(db.String(255), nullable=False)  # nombre en disco
    fecha_subida = db.Column(db.DateTime, default=datetime.utcnow)


# ---------------------------------------------------------------------------
# Modulo: Despachos (control de envio bajo un numero de seguimiento/tracking)
# ---------------------------------------------------------------------------
# Una vez que una orden (o parte de una orden, tras la division automatica)
# queda con todas sus lineas activas en "Orden Despachada" o mas adelante,
# el usuario deja de operarla como "orden de compra" y pasa a operarla como
# parte de un envio fisico: un Despacho, identificado por courier/embarcador
# + numero de tracking/AWB + status. Un Despacho puede agrupar UNA o VARIAS
# ordenes (despacho consolidado). Las etapas "Internacion Aduanas" /
# "Recibido" se siguen marcando por LINEA (ver ETAPAS_LINEA) pero ahora
# tipicamente se hace en bloque desde el propio Despacho (todas las lineas
# de todas sus ordenes avanzan juntas).

ESTADOS_DESPACHO = [
    "Preparando envío",
    "En tránsito",
    "En aduana",
    "Entregado",
]


class Despacho(db.Model):
    __tablename__ = "despachos"

    id = db.Column(db.Integer, primary_key=True)
    numero_tracking = db.Column(db.String(120))
    courier = db.Column(db.String(120))  # embarcador / courier / agente de carga
    status = db.Column(db.String(40), default=ESTADOS_DESPACHO[0])
    fecha_envio = db.Column(db.Date, nullable=True)
    fecha_estimada_llegada = db.Column(db.Date, nullable=True)
    notas = db.Column(db.Text)
    # Seguimiento online del tracking (ronda L, punto 2, 2026-09-09): si el
    # usuario pega aca la URL de seguimiento que le paso el courier/agente
    # (por ejemplo un courier local que no esta en COURIER_TRACKING_URLS
    # abajo), esa URL manual tiene siempre prioridad -- ver url_seguimiento.
    url_tracking_manual = db.Column(db.String(500))
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    ordenes = db.relationship("OrdenCompra", backref="despacho", lazy="dynamic")

    # Plantillas de URL de seguimiento publico de los couriers
    # internacionales mas comunes en este tipo de operacion (envios
    # aereos/maritimos de importacion). "Mejor esfuerzo": son las URLs
    # publicas documentadas de cada courier al 2026-09; si alguna cambia o
    # el courier real no esta en esta lista, el usuario siempre puede
    # cargar su propia URL en 'url_tracking_manual', que tiene prioridad.
    COURIER_TRACKING_URLS = {
        "DHL": "https://www.dhl.com/global-en/home/tracking.html?tracking-id={tracking}",
        "FEDEX": "https://www.fedex.com/fedextrack/?trknbr={tracking}",
        "UPS": "https://www.ups.com/track?tracknum={tracking}",
        "TNT": "https://www.tnt.com/express/en_us/site/tracking.html?searchType=CON&cons={tracking}",
    }

    @property
    def url_seguimiento(self):
        """URL para hacer seguimiento online del despacho (ronda L, punto 2).
        Prioridad: 1) URL manual cargada por el usuario, 2) URL armada
        automaticamente si el campo 'courier' menciona a uno de los
        couriers conocidos (comparacion flexible, sin importar mayusculas)
        y hay numero de tracking cargado. Si nada de eso aplica, devuelve
        None (no se puede armar un link)."""
        if self.url_tracking_manual:
            return self.url_tracking_manual
        if not self.numero_tracking:
            return None
        courier_norm = (self.courier or "").strip().upper()
        for clave, plantilla in self.COURIER_TRACKING_URLS.items():
            if clave in courier_norm:
                return plantilla.format(tracking=quote(self.numero_tracking.strip()))
        return None

    @property
    def total_lineas_activas(self):
        return sum(
            1 for o in self.ordenes for l in o.lineas if not l.anulada
        )

    @property
    def proveedores(self):
        """Lista de proveedores distintos entre las ordenes asociadas (un
        despacho consolidado puede traer mas de un proveedor)."""
        vistos = {}
        for o in self.ordenes:
            vistos[o.proveedor_id] = o.proveedor
        return list(vistos.values())

    def __repr__(self):
        return f"<Despacho {self.numero_tracking or self.id}>"


# ---------------------------------------------------------------------------
# Modulo: Costeo de Importaciones
# ---------------------------------------------------------------------------
# Metodologia validada contra una planilla real del usuario (ver
# documentacion en el proyecto). Resumen:
#  - Un embarque/factura (Importacion) puede dividirse en varios "parciales"
#    (Declaraciones de Ingreso/DIN), tipicamente por regimen aduanero.
#  - Flete y Seguro del embarque se prorratean por LINEA segun el valor
#    propio de esa linea (cantidad x precio unitario en USD) sobre el FOB
#    total del embarque completo.
#  - Derechos de aduana = CIF de la linea x % advalorem del regimen de su
#    parcial (no se prorratea, se calcula directo).
#  - Honorarios del agente de aduana son propios de cada parcial: se
#    prorratean solo entre las lineas de ESE parcial (por CIF).
#  - Gastos compartidos (GastoImportacion) se prorratean entre los
#    parciales a los que aplican (por CIF de cada parcial), y dentro de
#    cada parcial, entre sus lineas (por CIF de cada linea).

REGIMENES_PARCIAL = {
    "Preferencial": 0.0,
    "General": 6.0,
    "Muestra": 0.0,
    "Consignación": 0.0,
}

VIAS_EMBARQUE = ["Marítimo", "Aéreo", "Terrestre"]

# Condicion de compra pactada con el proveedor (Incoterm simplificado).
# Renombrado de "FOB" a "EXW" el 2026-08-27 (mas ajustado al Incoterm real
# que usa el usuario) -- ver migracion de datos en ensure_schema_migrations
# (app.py) que actualiza las filas viejas con "FOB" a "EXW".
# IMPORTANTE (2026-08-27, aclarado por el usuario): esta condicion es solo
# informativa/de referencia -- NO fuerza Flete/Seguro a 0 ni oculta esos
# campos. Aunque la operacion sea CIF, el usuario igual necesita cargar
# Flete y Seguro como valores aparte (no vienen indexados en el precio del
# producto para efectos de este sistema). Lo unico que cambia segun la
# condicion es el valor por DEFECTO al crear una importacion nueva: EXW
# sugiere 0 (el proveedor tipicamente no los incluye en su factura), CIF
# no fuerza ningun valor particular -- en ambos casos el campo queda
# editable.
CONDICIONES_COMPRA = ["EXW", "CIF"]

# Conceptos fijos para los items de Flete/Seguro/Handling Fee/Otros de la
# factura (2026-09-01, punto 6, a pedido del usuario). Antes Flete y Seguro
# eran dos campos fijos separados en Importacion (flete_total_moneda /
# seguro_total_moneda, ver mas abajo, ahora DEPRECADOS); ahora son filas de
# CargoAdicionalImportacion, igual que Handling Fee y cualquier otro item
# que venga en la Invoice -- mismo mecanismo, un solo listado con selector
# de concepto. Cada item se prorratea por FOB entre las lineas y se suma
# directo al CIF -- ver costing.py.
CONCEPTOS_ITEM_FACTURA = ["Flete", "Seguro", "Handling Fee", "Otros"]

# Conceptos fijos para los Gastos compartidos de una Importacion
# (2026-08-26, a pedido del usuario -- antes era texto libre). "Flete Int."
# y "Derechos" tienen tratamiento especial en la pantalla de Costeo: se
# muestran en columnas propias (Flete entre Valor Unit. y CIF, Derechos
# entre CIF y Gastos) en vez de quedar mezclados en la columna generica de
# Gastos -- ver costing.py y templates/importaciones/detalle.html.
CONCEPTOS_GASTO = [
    "Gasto Agencia",
    "Gasto operacional",
    "Flete Int.",
    "Flete Nacional",
    "Seguro",
    "Desconsolidación",
    "Almacenaje",
    "ISP",
    "Derechos",
    "Honorarios",
    "Otros",
]

# Documentos que se pueden adjuntar a un Gasto compartido (comprobante de
# pago, factura del agente, etc.) -- 2026-09-01, ronda G, punto 7. A
# diferencia de los documentos de Orden (solo PDF/imagen), aca tambien se
# admite Word y Excel porque el usuario suele recibir el respaldo de estos
# gastos en esos formatos.
TIPOS_DOCUMENTO_GASTO = ["Comprobante de pago", "Factura", "Otro"]


class Importacion(db.Model):
    __tablename__ = "importaciones"

    id = db.Column(db.Integer, primary_key=True)
    proveedor_id = db.Column(db.Integer, db.ForeignKey("proveedores.id"), nullable=False)
    # Si esta importacion se genero automaticamente desde un Despacho (ver
    # importaciones_desde_despacho en app.py), queda la trazabilidad aqui --
    # permite mostrar en la Importacion un link de vuelta a la orden/despacho
    # de origen y a sus documentos (Invoice, Guia de despacho, etc.) ya
    # cargados, sin duplicarlos.
    despacho_id = db.Column(db.Integer, db.ForeignKey("despachos.id"), nullable=True)
    # Empresa compradora de esta importacion (Accuvision/Accumedical, ronda O
    # 2026-09-12) -- se usa para imprimir el nombre de la empresa y para
    # calcular el correlativo de importacion del sistema de Inventarios (ver
    # numero_correlativo_inventario abajo y siguiente_correlativo_inventario
    # en app.py), igual patron que OrdenCompra.empresa_id (ronda K).
    empresa_id = db.Column(db.Integer, db.ForeignKey("empresas.id"), nullable=True)
    # Correlativo de esta importacion DENTRO del sistema de Inventarios,
    # unico por empresa compradora (ronda O, 2026-09-12) -- es un numero que
    # vive en ese otro sistema, no en este, asi que no se genera solo: el
    # usuario entrega el numero de partida la primera vez y desde ahi
    # siguiente_correlativo_inventario() en app.py sugiere el siguiente
    # (maximo ya usado por esa empresa + 1), siempre editable a mano.
    numero_correlativo_inventario = db.Column(db.Integer, nullable=True)
    numero_factura = db.Column(db.String(80))
    fecha_factura = db.Column(db.Date, nullable=True)
    moneda_factura = db.Column(db.String(10), default="EURO")  # moneda en que vienen los Valor Unitario de las lineas
    # Condicion de compra con el proveedor (ver CONDICIONES_COMPRA arriba) --
    # solo informativa desde el 2026-08-27, no afecta el calculo.
    condicion_compra = db.Column(db.String(10), default="EXW")

    # Tipos de cambio (entrada manual, como en la planilla del usuario)
    tipo_cambio_aduanero = db.Column(db.Float, default=0)  # CLP por 1 USD
    tipo_cambio_moneda_usd = db.Column(db.Float, default=1)  # unidades de moneda_factura por 1 USD (ej. EUR por USD)

    # DEPRECADOS (2026-08-27, y de nuevo 2026-09-01): flete_total_usd /
    # seguro_total_usd fueron reemplazados por flete_total_moneda /
    # seguro_total_moneda (2026-08-27), y estos a su vez fueron reemplazados
    # por filas de CargoAdicionalImportacion con concepto "Flete"/"Seguro"
    # (2026-09-01, punto 6 -- ver CONCEPTOS_ITEM_FACTURA arriba). Los 4
    # campos se mantienen en el esquema solo por compatibilidad con datos
    # ya cargados; ninguno se lee ni se escribe desde el 2026-09-01. La
    # migracion en ensure_schema_migrations (app.py) copia estos valores
    # viejos a filas nuevas de CargoAdicionalImportacion la primera vez que
    # arranca la app con este cambio, para no perder datos ya cargados.
    flete_total_usd = db.Column(db.Float, default=0)
    seguro_total_usd = db.Column(db.Float, default=0)
    flete_total_moneda = db.Column(db.Float, default=0)
    seguro_total_moneda = db.Column(db.Float, default=0)

    notas = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    proveedor = db.relationship("Proveedor")
    despacho = db.relationship("Despacho", backref="importaciones_generadas")
    empresa = db.relationship("Empresa", backref="importaciones")
    parciales = db.relationship(
        "Parcial", backref="importacion", cascade="all, delete-orphan", lazy="dynamic",
        order_by="Parcial.id",
    )
    gastos = db.relationship(
        "GastoImportacion", backref="importacion", cascade="all, delete-orphan", lazy="dynamic",
        order_by="GastoImportacion.id",
    )
    cargos_adicionales = db.relationship(
        "CargoAdicionalImportacion", backref="importacion", cascade="all, delete-orphan",
        lazy="dynamic", order_by="CargoAdicionalImportacion.id",
    )
    legajos = db.relationship(
        "ImportacionDocumento", backref="importacion", cascade="all, delete-orphan",
        lazy="dynamic", order_by="ImportacionDocumento.fecha_subida.desc()",
    )

    @property
    def fob_total_usd(self):
        total = 0.0
        for parcial in self.parciales:
            for linea in parcial.lineas:
                total += linea.valor_total_usd
        return total

    @property
    def fob_total_moneda(self):
        """Ronda AJ (2026-09-18): total FOB en la MONEDA DE LA FACTURA (no en
        USD) -- es el "Valor total de la factura" que el modulo de Pago
        Proveedores usa para la cuenta por pagar (a pedido explicito del
        usuario: se paga el valor de la factura, no el costo total nacionalizado
        con derechos/gastos incluidos)."""
        total = 0.0
        for parcial in self.parciales:
            for linea in parcial.lineas:
                total += linea.valor_total_moneda
        return total

    def __repr__(self):
        return f"<Importacion {self.numero_factura}>"


class Parcial(db.Model):
    __tablename__ = "parciales"

    id = db.Column(db.Integer, primary_key=True)
    importacion_id = db.Column(db.Integer, db.ForeignKey("importaciones.id"), nullable=False)
    numero_parcial = db.Column(db.String(40))  # Nro. de Importacion / DIN
    # Etiqueta corta para identificar el parcial en su pestana (ej. "Lentes
    # intraoculares"), distinta del numero de DIN. Agregado 2026-08-26.
    referencia = db.Column(db.String(120))
    tipo_regimen = db.Column(db.String(30), default="General")  # informativo, ya no calcula derechos automaticamente
    # NOTA (2026-08-26): advalorem_pct y honorarios_* quedan en el esquema
    # por compatibilidad con datos previos, pero YA NO se usan en el calculo
    # ni se editan desde la UI -- el usuario pidio unificar Derechos de
    # Aduana y Honorarios del agente con el mismo mecanismo de "Gastos
    # compartidos" (GastoImportacion), porque los recibe como un monto ya
    # calculado (no como una tasa) y a veces cubren varios parciales a la
    # vez. Ver costing.py.
    advalorem_pct = db.Column(db.Float, default=6.0)
    via_embarque = db.Column(db.String(20), default="Marítimo")
    honorarios_agente = db.Column(db.Float, default=0)
    honorarios_referencia = db.Column(db.String(120))
    notas = db.Column(db.Text)

    lineas = db.relationship(
        "ParcialLinea", backref="parcial", cascade="all, delete-orphan", lazy="dynamic",
        order_by="ParcialLinea.id",
    )

    def __repr__(self):
        return f"<Parcial {self.numero_parcial}>"


class ParcialLinea(db.Model):
    __tablename__ = "parcial_lineas"

    id = db.Column(db.Integer, primary_key=True)
    parcial_id = db.Column(db.Integer, db.ForeignKey("parciales.id"), nullable=False)
    producto_id = db.Column(db.Integer, db.ForeignKey("productos.id"), nullable=True)
    # Si esta linea se genero automaticamente desde un Despacho, trazabilidad
    # de vuelta a la linea de la orden de compra de origen (backlog #15).
    orden_compra_linea_id = db.Column(
        db.Integer, db.ForeignKey("ordenes_compra_lineas.id"), nullable=True
    )

    codigo = db.Column(db.String(120), nullable=False)
    descripcion = db.Column(db.String(500), nullable=False)
    codigo_lote = db.Column(db.String(80))
    fecha_vencimiento = db.Column(db.Date, nullable=True)

    cantidad_unidades = db.Column(db.Integer, default=0)
    valor_unitario_moneda = db.Column(db.Float, default=0)  # en moneda_factura de la Importacion

    producto = db.relationship("Producto")
    orden_compra_linea = db.relationship("OrdenCompraLinea")
    lotes = db.relationship(
        "ParcialLineaLote",
        backref="linea",
        cascade="all, delete-orphan",
        order_by="ParcialLineaLote.fecha_vencimiento",
    )

    @property
    def valor_total_moneda(self):
        return (self.cantidad_unidades or 0) * (self.valor_unitario_moneda or 0)

    @property
    def valor_total_usd(self):
        tc = self.parcial.importacion.tipo_cambio_moneda_usd or 1
        return self.valor_total_moneda / tc

    @property
    def total_unidades_lotes(self):
        return sum((lote.cantidad_unidades or 0) for lote in self.lotes)


class ParcialLineaLote(db.Model):
    """Desglose informativo de una linea de parcial en uno o mas lotes,
    cada uno con su propia fecha de vencimiento (backlog #16). El total
    autoritativo de unidades sigue siendo ParcialLinea.cantidad_unidades;
    esta tabla es solo trazabilidad."""

    __tablename__ = "parcial_linea_lotes"

    id = db.Column(db.Integer, primary_key=True)
    linea_id = db.Column(db.Integer, db.ForeignKey("parcial_lineas.id"), nullable=False)
    codigo_lote = db.Column(db.String(80))
    fecha_vencimiento = db.Column(db.Date, nullable=True)
    cantidad_unidades = db.Column(db.Integer, default=0)


class GastoImportacion(db.Model):
    __tablename__ = "gastos_importacion"

    id = db.Column(db.Integer, primary_key=True)
    importacion_id = db.Column(db.Integer, db.ForeignKey("importaciones.id"), nullable=False)
    concepto = db.Column(db.String(120), nullable=False)
    # monto_clp es el monto YA CONVERTIDO a CLP -- el que usa costing.py
    # para el prorrateo, sin tocar. Cuando el gasto se negocia en otra
    # moneda (ej. el flete internacional facturado en USD o EUR con un
    # tipo de cambio propio, distinto del "Dolar Aduanero" de la
    # importacion), moneda/tipo_cambio/monto_original quedan solo como
    # respaldo de como se llego a ese monto_clp (2026-08-26).
    monto_clp = db.Column(db.Float, default=0)
    moneda = db.Column(db.String(10), default="CLP")
    tipo_cambio = db.Column(db.Float, default=1)  # CLP por 1 unidad de 'moneda' (1 si moneda=CLP)
    monto_original = db.Column(db.Float, default=0)  # monto en 'moneda', antes de convertir
    referencia = db.Column(db.String(150))  # ej. "FC 18720 TRANSPORTE SALAZAR"

    # Si esta vacio, el gasto se prorratea entre TODOS los parciales de la
    # importacion (por CIF). Si tiene parciales asociados, solo se
    # prorratea entre esos.
    parciales_aplicables = db.relationship(
        "Parcial", secondary="gasto_parcial_aplicable", backref="gastos_especificos"
    )
    documentos = db.relationship(
        "GastoDocumento", backref="gasto", cascade="all, delete-orphan", lazy="dynamic",
        order_by="GastoDocumento.fecha_subida.desc()",
    )


class GastoDocumento(db.Model):
    """Documento de respaldo de un Gasto compartido (comprobante de pago,
    factura del agente, etc.) -- 2026-09-01, ronda G, punto 7. Mismo patron
    que OrdenDocumento, pero admite ademas Word/Excel (ver
    TIPOS_DOCUMENTO_GASTO y EXTENSIONES_PERMITIDAS_GASTO en app.py)."""
    __tablename__ = "gastos_documentos"

    id = db.Column(db.Integer, primary_key=True)
    gasto_id = db.Column(db.Integer, db.ForeignKey("gastos_importacion.id"), nullable=False)
    tipo = db.Column(db.String(60), default="Otro")
    nombre_original = db.Column(db.String(255), nullable=False)
    nombre_archivo = db.Column(db.String(255), nullable=False)  # nombre en disco
    fecha_subida = db.Column(db.DateTime, default=datetime.utcnow)


class ImportacionDocumento(db.Model):
    """"Legajo": archivo consolidado a nivel de TODA la importacion (ej. un
    solo PDF que junta todas las facturas del embarque), distinto de los
    documentos de un Gasto puntual (GastoDocumento, que son por concepto
    especifico). Agregado 2026-09-01, ronda H, punto 3 (opcion del menu
    "Adjuntar > Legajo"). Mismo patron de archivo-en-disco + registro en BD
    que GastoDocumento/OrdenDocumento."""
    __tablename__ = "importacion_documentos"

    id = db.Column(db.Integer, primary_key=True)
    importacion_id = db.Column(db.Integer, db.ForeignKey("importaciones.id"), nullable=False)
    descripcion = db.Column(db.String(150))
    nombre_original = db.Column(db.String(255), nullable=False)
    nombre_archivo = db.Column(db.String(255), nullable=False)  # nombre en disco
    fecha_subida = db.Column(db.DateTime, default=datetime.utcnow)


gasto_parcial_aplicable = db.Table(
    "gasto_parcial_aplicable",
    db.Column("gasto_id", db.Integer, db.ForeignKey("gastos_importacion.id"), primary_key=True),
    db.Column("parcial_id", db.Integer, db.ForeignKey("parciales.id"), primary_key=True),
)


class CargoAdicionalImportacion(db.Model):
    """Item de la FACTURA/embarque que se suma al CIF: Flete, Seguro,
    Handling Fee u Otros (ver CONCEPTOS_ITEM_FACTURA arriba). Hasta el
    2026-08-27 esta tabla solo cubria cargos "extra" tipo Handling Fee,
    mientras Flete y Seguro eran dos campos fijos separados en Importacion
    -- el usuario pidio unificarlos en un solo listado con selector de
    concepto (2026-09-01, punto 6), asi que ahora Flete y Seguro tambien
    son filas aca (ver migracion de datos en ensure_schema_migrations,
    app.py, que traspasa los valores viejos la primera vez).

    A diferencia de GastoImportacion (que se prorratea por CIF entre los
    parciales que se marquen), un item de este listado se suma directo al
    CIF de CADA LINEA, prorrateado por su valor FOB -- porque
    conceptualmente es parte de la misma factura de compra, no un gasto de
    internacion que se conoce despues.

    Cada item tiene su PROPIA moneda y tipo de cambio (2026-09-01, punto 3
    -- antes todo se asumia en la misma moneda de la factura): por defecto
    se carga en la moneda de la factura, pero si el proveedor factura un
    item puntual en otra moneda (o la operacion es en USD y aparece un
    cargo en otra moneda), se puede indicar aparte sin afectar al resto.
    """
    __tablename__ = "cargos_adicionales_importacion"

    id = db.Column(db.Integer, primary_key=True)
    importacion_id = db.Column(db.Integer, db.ForeignKey("importaciones.id"), nullable=False)
    concepto = db.Column(db.String(120), nullable=False)  # uno de CONCEPTOS_ITEM_FACTURA, o texto libre (dato anterior)
    monto_moneda = db.Column(db.Float, default=0)  # en la moneda de este item (columna 'moneda' de abajo)
    # Moneda propia de este item y tipo de cambio para llevarlo a USD --
    # mismo formato que Importacion.tipo_cambio_moneda_usd: unidades de
    # 'moneda' por 1 USD (paridad). Si moneda == "USD", tipo_cambio se
    # fuerza a 1 (monto_moneda ya esta en USD, no se convierte).
    moneda = db.Column(db.String(10), default="USD")
    tipo_cambio = db.Column(db.Float, default=1)
    referencia = db.Column(db.String(150))

    @property
    def monto_usd(self):
        moneda = self.moneda or "USD"
        monto = self.monto_moneda or 0
        if moneda == "USD":
            return monto
        tc = self.tipo_cambio or 1
        return (monto / tc) if tc else 0

    def __repr__(self):
        return f"<ItemFactura {self.concepto}>"


class TipoCambioMensual(db.Model):
    """Tipo de cambio (Dolar Aduanero + paridad Euro/USD) valido para todas
    las importaciones de un mes calendario -- el usuario aclaro que ambos
    valores son los mismos para todas las operaciones de un mismo mes
    (2026-08-27, punto 6). Se define una vez por mes en Configuracion y se
    usa solo como VALOR SUGERIDO al crear una importacion nueva de ese mes
    (sigue siendo editable caso a caso, para excepciones)."""
    __tablename__ = "tipo_cambio_mensual"

    id = db.Column(db.Integer, primary_key=True)
    anio = db.Column(db.Integer, nullable=False)
    mes = db.Column(db.Integer, nullable=False)  # 1-12
    tc_aduanero = db.Column(db.Float, default=0)  # CLP por 1 USD
    paridad_eur_usd = db.Column(db.Float, default=1)  # EUR por 1 USD
    notas = db.Column(db.String(200))
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint("anio", "mes", name="uq_tipo_cambio_mensual_anio_mes"),
    )

    MESES_NOMBRE = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
        "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]

    @property
    def nombre_mes(self):
        try:
            return self.MESES_NOMBRE[self.mes - 1]
        except IndexError:
            return str(self.mes)

    def __repr__(self):
        return f"<TipoCambioMensual {self.mes}/{self.anio}>"


# ---------------------------------------------------------------------------
# Modulo: Pago Proveedores
# ---------------------------------------------------------------------------
# Ronda AJ (2026-09-18, fase 1 -- ver ronda-aj-pago-proveedores-costo-
# producto-multi-factura.md en el proyecto para el detalle completo y el
# estado de cada requerimiento). FacturaProveedor es la cuenta por pagar
# propiamente tal -- una fila por factura de compra a un proveedor
# extranjero, ya sea:
#   1) cargada UNA VEZ desde el archivo de saldos iniciales del usuario
#      ("Cuentas por pagar proveedores al 18-09-2026.xlsx", ver
#      seed_facturas_proveedor_pendientes en app.py) -- importacion_id queda
#      en None porque esas facturas son anteriores a este modulo (no hay
#      Importacion del sistema asociada), o
#   2) generada automaticamente cuando se costea una Importacion NUEVA en el
#      sistema -- importacion_id apunta a esa Importacion, y el monto es el
#      "Valor total factura" (Importacion.fob_total_moneda), NO el costo
#      nacionalizado con derechos/gastos incluidos (a pedido explicito del
#      usuario: se le paga al proveedor el valor de SU factura).
# Cada pago (parcial o total) queda como fila propia en
# PagoFacturaProveedor -- una factura puede tener varios abonos antes de
# quedar 'pagada'. El tipo de cambio de CADA pago se guarda ahi porque es
# el dato que, a pedido del usuario, se necesitara para recostear el FOB/CIF
# de la Importacion asociada una vez que la factura quede pagada por
# completo (fase 2, pendiente -- ver el documento de ronda mencionado
# arriba; por ahora el pago se registra y el estado de la factura se
# actualiza, pero el costeo de la Importacion todavia NO se recalcula solo).
ESTADOS_FACTURA_PROVEEDOR = ["pendiente", "abonada", "pagada", "anulada"]


class FacturaProveedor(db.Model):
    __tablename__ = "facturas_proveedor"

    id = db.Column(db.Integer, primary_key=True)
    proveedor_id = db.Column(db.Integer, db.ForeignKey("proveedores.id"), nullable=False)
    importacion_id = db.Column(db.Integer, db.ForeignKey("importaciones.id"), nullable=True)

    numero_factura = db.Column(db.String(80), nullable=False)
    fecha_emision = db.Column(db.Date, nullable=True)
    moneda = db.Column(db.String(10), default="USD")
    valor_factura = db.Column(db.Float, default=0)  # en 'moneda' -- el valor TOTAL de la factura
    fecha_vencimiento = db.Column(db.Date, nullable=True)
    # Valor en CLP solo de REFERENCIA (al tipo de cambio vigente cuando se
    # registro la factura) -- informativo para el listado; NO es el que se
    # usara para recostear (ese sera el tipo de cambio del PAGO real, fase 2).
    valor_clp_referencial = db.Column(db.Float, default=0)

    estado = db.Column(db.String(20), default="pendiente")
    notas = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)
    # Ronda AK (2026-09-19): de donde salio esta factura -- "costeo" (boton
    # "Generar factura" en una Importacion ya costeada), "consignacion"
    # (Pago Proveedores > Facturar consignacion) o "manual" (las que se
    # cargaron una sola vez desde el Excel de saldos iniciales, fase 1).
    origen = db.Column(db.String(20), default="manual")

    proveedor = db.relationship("Proveedor")
    importacion = db.relationship("Importacion", backref=db.backref("factura_pago", uselist=False))
    pagos = db.relationship(
        "PagoFacturaProveedor", backref="factura", cascade="all, delete-orphan",
        lazy="dynamic", order_by="PagoFacturaProveedor.fecha_pago",
    )

    @property
    def monto_pagado(self):
        return sum((p.monto or 0) for p in self.pagos)

    @property
    def total_notas_credito(self):
        """Ronda AK (2026-09-19): suma de todas las NC aplicadas a esta
        factura -- reduce el saldo real que se le debe al proveedor (ver
        NotaCreditoProveedor)."""
        return round(sum((nc.monto or 0) for nc in self.notas_credito), 2)

    @property
    def valor_neto(self):
        """Valor de la factura despues de notas de credito -- lo que
        realmente se le termina debiendo al proveedor."""
        return round((self.valor_factura or 0) - self.total_notas_credito, 2)

    @property
    def saldo_pendiente(self):
        return round((self.valor_factura or 0) - self.monto_pagado - self.total_notas_credito, 2)

    @property
    def dias_para_vencer(self):
        """Positivo = dias que faltan para vencer, negativo = dias de atraso
        (ya vencida). None si todavia no hay fecha de vencimiento cargada."""
        if not self.fecha_vencimiento:
            return None
        return (self.fecha_vencimiento - datetime.utcnow().date()).days

    def __repr__(self):
        return f"<FacturaProveedor {self.numero_factura} ({self.estado})>"


class PagoFacturaProveedor(db.Model):
    __tablename__ = "pagos_factura_proveedor"

    id = db.Column(db.Integer, primary_key=True)
    factura_id = db.Column(db.Integer, db.ForeignKey("facturas_proveedor.id"), nullable=False)
    fecha_pago = db.Column(db.Date, nullable=False)
    monto = db.Column(db.Float, default=0)  # en la MISMA moneda de la factura (FacturaProveedor.moneda)
    tipo_cambio_pago = db.Column(db.Float, default=0)  # CLP por 1 unidad de esa moneda, EL DIA DEL PAGO real
    es_abono = db.Column(db.Boolean, default=False)  # informativo -- el estado real lo decide el saldo restante
    notas = db.Column(db.String(300))
    registrado_por_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    registrado_por = db.relationship("Usuario")

    @property
    def monto_clp(self):
        return (self.monto or 0) * (self.tipo_cambio_pago or 0)

    def __repr__(self):
        return f"<PagoFacturaProveedor factura={self.factura_id} {self.monto}>"


class FacturaProveedorLinea(db.Model):
    """Ronda AK (2026-09-19): detalle por producto/lote de una
    FacturaProveedor -- hace falta para poder aplicar una Nota de Credito
    sobre productos puntuales (ver NotaCreditoProveedor). Se llena de 2
    formas: automatico al generar la factura desde un Costeo ya confirmado
    (una linea por cada ParcialLinea de la Importacion asociada), o manual
    al facturar consignacion (una linea por cada producto de
    CompraHistorica que el usuario eligio facturar)."""

    __tablename__ = "facturas_proveedor_lineas"

    id = db.Column(db.Integer, primary_key=True)
    factura_id = db.Column(db.Integer, db.ForeignKey("facturas_proveedor.id"), nullable=False)
    codigo_producto = db.Column(db.String(120))
    descripcion = db.Column(db.String(500))
    codigo_lote = db.Column(db.String(80))
    cantidad = db.Column(db.Float, default=0)
    precio_unitario = db.Column(db.Float, default=0)
    valor_total = db.Column(db.Float, default=0)
    # Trazabilidad de origen -- como mucho uno de los dos esta lleno.
    parcial_linea_id = db.Column(db.Integer, db.ForeignKey("parcial_lineas.id"), nullable=True)
    compra_historica_id = db.Column(db.Integer, db.ForeignKey("compras_historicas.id"), nullable=True)

    factura = db.relationship(
        "FacturaProveedor",
        backref=db.backref("lineas", cascade="all, delete-orphan", lazy="dynamic", order_by="FacturaProveedorLinea.id"),
    )

    @property
    def cantidad_acreditada(self):
        return sum((l.cantidad or 0) for l in self.notas_credito_lineas)

    @property
    def cantidad_disponible(self):
        return round((self.cantidad or 0) - self.cantidad_acreditada, 4)

    def __repr__(self):
        return f"<FacturaProveedorLinea {self.codigo_producto} x{self.cantidad}>"


TIPOS_NC_PROVEEDOR = ["cantidad_y_valor", "solo_valor"]


class NotaCreditoProveedor(db.Model):
    """Ronda AK (2026-09-19): nota de credito de un proveedor sobre una
    FacturaProveedor ya cargada -- por descuento comercial ("solo_valor",
    un monto que se resta sin tocar cantidades) o por devolucion de
    producto ("cantidad_y_valor", resta unidades y valor de lineas
    puntuales, ver NotaCreditoProveedorLinea). Reduce el saldo pendiente
    real de la factura (ver FacturaProveedor.saldo_pendiente) -- el motivo
    de este modulo: los reportes historicos solo contemplaban cantidades y
    valor importado, sobreestimando la compra real cuando el proveedor
    despues hace una NC por descuento o devolucion."""

    __tablename__ = "notas_credito_proveedor"

    id = db.Column(db.Integer, primary_key=True)
    factura_id = db.Column(db.Integer, db.ForeignKey("facturas_proveedor.id"), nullable=False)
    numero_nc = db.Column(db.String(80))
    fecha = db.Column(db.Date, nullable=True)
    tipo = db.Column(db.String(20), default="solo_valor")  # ver TIPOS_NC_PROVEEDOR
    monto = db.Column(db.Float, default=0)  # siempre en la moneda de la factura
    motivo = db.Column(db.Text)
    creado_por_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    creado_en = db.Column(db.DateTime, default=datetime.utcnow)

    factura = db.relationship(
        "FacturaProveedor",
        backref=db.backref("notas_credito", cascade="all, delete-orphan", lazy="dynamic", order_by="NotaCreditoProveedor.fecha.desc()"),
    )
    creado_por = db.relationship("Usuario")

    def __repr__(self):
        return f"<NotaCreditoProveedor factura={self.factura_id} {self.monto}>"


class NotaCreditoProveedorLinea(db.Model):
    """Detalle de una NC 'cantidad_y_valor': cuanta cantidad y valor se
    acredita de una linea puntual de la factura. Una NC 'solo_valor' no
    tiene lineas -- su monto ya queda directo en NotaCreditoProveedor.monto."""

    __tablename__ = "notas_credito_proveedor_lineas"

    id = db.Column(db.Integer, primary_key=True)
    nota_credito_id = db.Column(db.Integer, db.ForeignKey("notas_credito_proveedor.id"), nullable=False)
    factura_linea_id = db.Column(db.Integer, db.ForeignKey("facturas_proveedor_lineas.id"), nullable=False)
    cantidad = db.Column(db.Float, default=0)
    valor = db.Column(db.Float, default=0)

    nota_credito = db.relationship(
        "NotaCreditoProveedor",
        backref=db.backref("lineas", cascade="all, delete-orphan", lazy="dynamic"),
    )
    factura_linea = db.relationship(
        "FacturaProveedorLinea",
        backref=db.backref("notas_credito_lineas", lazy="dynamic"),
    )
