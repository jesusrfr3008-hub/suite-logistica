"""
Importa el catalogo maestro de proveedores/productos desde el Excel
'Master_Lista_precios_Proveedores.xlsx' (hoja '2026', que ya consolida
los 8 proveedores) hacia la base de datos, solo si esta vacia.
"""
import os
import openpyxl

from models import db, Proveedor, Producto

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Master_Lista_precios_Proveedores.xlsx")

# Datos de contacto conocidos (opcional). Si no se conocen, quedan vacios
# y se completan luego desde el modulo de Mantenimiento de Proveedores.
CONTACTOS_DEFAULT = {
    # "DORC": {"pais": "Holanda", "tipo": "Extranjero", "contacto_email": "ventas@dorc.nl"},
}


def _clean(value):
    if value is None:
        return ""
    return str(value).strip()


def seed_from_excel(app):
    if not os.path.exists(EXCEL_PATH):
        print(f"[seed] No se encontro el Excel en {EXCEL_PATH}, se omite la carga inicial.")
        return

    with app.app_context():
        if Proveedor.query.first() is not None:
            print("[seed] Ya existen datos en la base, se omite la importacion.")
            return

        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["2026"]

        proveedores_cache = {}
        productos_creados = 0
        filas_omitidas = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            nombre_prov, codigo, descripcion, empaque, moneda, precio_caja, precio_unt = (
                row[0], row[1], row[2], row[3], row[4], row[5], row[6]
            )
            nombre_prov = _clean(nombre_prov)
            codigo = _clean(codigo)
            if not nombre_prov or not codigo:
                filas_omitidas += 1
                continue

            if nombre_prov not in proveedores_cache:
                extra = CONTACTOS_DEFAULT.get(nombre_prov, {})
                prov = Proveedor(
                    nombre=nombre_prov,
                    tipo=extra.get("tipo", "Extranjero"),
                    pais=extra.get("pais", ""),
                    contacto_email=extra.get("contacto_email", ""),
                    moneda_default=_clean(moneda) or "USD",
                )
                db.session.add(prov)
                db.session.flush()  # obtener id
                proveedores_cache[nombre_prov] = prov
            else:
                prov = proveedores_cache[nombre_prov]

            try:
                empaque_val = int(empaque) if empaque else 1
            except (ValueError, TypeError):
                empaque_val = 1

            try:
                precio_caja_val = float(precio_caja) if precio_caja is not None else 0.0
            except (ValueError, TypeError):
                precio_caja_val = 0.0

            try:
                precio_unt_val = float(precio_unt) if precio_unt is not None else 0.0
            except (ValueError, TypeError):
                precio_unt_val = 0.0

            producto = Producto(
                proveedor_id=prov.id,
                codigo=codigo,
                descripcion=_clean(descripcion) or codigo,
                empaque=empaque_val,
                moneda=_clean(moneda) or "USD",
                precio_caja=precio_caja_val,
                precio_unitario=precio_unt_val,
            )
            db.session.add(producto)
            productos_creados += 1

        db.session.commit()
        print(
            f"[seed] Importados {len(proveedores_cache)} proveedores y "
            f"{productos_creados} productos ({filas_omitidas} filas omitidas por datos incompletos)."
        )
