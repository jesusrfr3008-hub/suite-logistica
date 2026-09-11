"""
Motor de calculo del costeo de importaciones.

Metodologia (ajustada 2026-09-01 a pedido del usuario -- ver nota abajo):
 1. Flete, Seguro, Handling Fee y cualquier otro item de la factura se
    cargan como filas de CargoAdicionalImportacion (ver CONCEPTOS_ITEM_FACTURA
    en models.py), cada una con su PROPIA moneda y tipo de cambio (unidades
    de esa moneda por 1 USD, igual formato que tipo_cambio_moneda_usd). Por
    defecto un item se carga en la moneda de la factura, pero si un item
    puntual viene facturado en otra moneda (o la operacion es en USD y
    aparece un cargo en otra moneda) se puede indicar aparte -- ver
    CargoAdicionalImportacion.monto_usd. Cuando la moneda de la factura ya
    es USD, no hay paso de paridad: el "tipo de cambio" de la operacion es
    directamente el Dolar Aduanero (tipo_cambio_moneda_usd queda en 1).
    Cada item, ya convertido a USD, se prorratea por LINEA segun el valor
    propio de esa linea (en USD) sobre el FOB total de TODO el embarque
    (todas las lineas, de todos los parciales). La condicion de compra
    (EXW/CIF) es solo informativa -- NO fuerza estos campos a 0 ni los
    oculta: el usuario aclaro que aunque la operacion sea CIF, Flete y
    Seguro igual se cargan aparte (no vienen indexados en el precio del
    producto para este sistema).
 2. CIF de la linea = valor propio + todos los items de factura asignados
    (todo en USD), convertido a CLP con el "Dolar Aduanero" de la
    importacion, y tambien expresado en la moneda de la factura
    (multiplicando por la paridad) para poder revisar el costeo sin
    convertir a mano (punto 3, 2026-09-01).
 3. TODO lo demas (Derechos de Aduana, Honorarios del agente de aduana,
    Almacenaje, Desconsolizacion, Gastos de Agencia, Transporte terrestre /
    local, Certificados, etc.) se carga como un "Gasto" (GastoImportacion):
    un monto YA CONOCIDO (del comprobante de pago o de la factura del
    agente, tipicamente propio de UN parcial/DIN en particular) que se
    prorratea entre los parciales a los que aplica (todos si no se
    especifica, por CIF de cada parcial) y, dentro de cada parcial, entre
    sus lineas (por CIF de cada linea). Este mecanismo de prorrateo por CIF
    se mantuvo SIN CAMBIOS en la ronda del 2026-09-01 -- el usuario aclaro
    que cada parcial simplemente tiene sus propios montos de
    Derechos/Honorarios, determinados independientemente, no un pool que
    deba prorratearse de otra forma.

    Para la VISUALIZACION: los gastos con concepto "Flete Int." y
    "Derechos" se muestran en columnas propias en la pantalla de Costeo en
    vez de mezclarse en la columna generica de Gastos. Esto es solo para
    mostrar el desglose: el monto total y el prorrateo NO cambian, se sigue
    sumando todo igual en costo_total_clp.
 4. Costo total de la linea = CIF + Gastos (incluye lo que antes era
    Derechos y Honorarios, ahora como conceptos de Gasto). Costo unitario
    = Costo total / cantidad de unidades de la linea.
 5. Ademas de CLP, se calculan los totales tambien en USD (asi se declara
    el CIF en la DIN, sin importar en que moneda venga la factura) y en la
    moneda de la factura (moneda_factura) cuando es distinta de USD -- para
    que una factura 100% en Euros pueda revisarse en Euros sin tener que
    convertir a mano. Esta es exactamente la cadena de conversion de la
    planilla de muestra del usuario: moneda de factura -> USD (paridad) ->
    CLP (Dolar Aduanero).
"""
from collections import defaultdict

# Conceptos de Gasto con columna propia en la pantalla de Costeo (deben
# coincidir exactamente con los valores de CONCEPTOS_GASTO en models.py).
CONCEPTO_FLETE_GASTO = "Flete Int."
CONCEPTO_DERECHOS = "Derechos"

# Conceptos de item de factura (CargoAdicionalImportacion, ver
# CONCEPTOS_ITEM_FACTURA en models.py) con columna propia en la pantalla de
# Costeo -- Flete y Seguro. Todo lo demas (Handling Fee, Otros, o un
# concepto viejo libre) se agrupa en una columna "Otros items factura".
CONCEPTO_ITEM_FLETE = "Flete"
CONCEPTO_ITEM_SEGURO = "Seguro"


def calcular_costeo(importacion):
    fob_total_usd = importacion.fob_total_usd
    tc_aduanero = importacion.tipo_cambio_aduanero or 0
    tc_moneda_usd = importacion.tipo_cambio_moneda_usd or 1

    # --- Items de la factura (Flete, Seguro, Handling Fee, Otros...),
    # cada uno ya convertido a USD con su propia moneda/tipo de cambio ---
    items_por_concepto_usd = defaultdict(float)
    for item in importacion.cargos_adicionales:
        items_por_concepto_usd[item.concepto] += item.monto_usd
    items_total_usd = sum(items_por_concepto_usd.values())

    lineas_info = []
    cif_por_parcial = defaultdict(float)
    fob_moneda_factura = 0.0

    # --- Paso 1 y 2: valor propio, items de factura prorrateados, CIF ---
    for parcial in importacion.parciales:
        for linea in parcial.lineas:
            valor_usd = linea.valor_total_usd
            valor_total_moneda = linea.valor_total_moneda
            fob_moneda_factura += valor_total_moneda
            share_fob = (valor_usd / fob_total_usd) if fob_total_usd else 0

            items_linea_usd = {c: share_fob * v for c, v in items_por_concepto_usd.items()}
            flete_usd = items_linea_usd.get(CONCEPTO_ITEM_FLETE, 0)
            seguro_usd = items_linea_usd.get(CONCEPTO_ITEM_SEGURO, 0)
            otros_items_usd = sum(
                v for c, v in items_linea_usd.items()
                if c not in (CONCEPTO_ITEM_FLETE, CONCEPTO_ITEM_SEGURO)
            )
            cif_usd = valor_usd + flete_usd + seguro_usd + otros_items_usd
            cif_clp = cif_usd * tc_aduanero
            cif_moneda = cif_usd * tc_moneda_usd

            info = {
                "linea": linea,
                "parcial": parcial,
                "valor_usd": valor_usd,
                "valor_total_moneda": valor_total_moneda,
                "flete_usd": flete_usd,
                "flete_moneda": flete_usd * tc_moneda_usd,
                "seguro_usd": seguro_usd,
                "seguro_moneda": seguro_usd * tc_moneda_usd,
                "otros_items_usd": otros_items_usd,
                "otros_items_moneda": otros_items_usd * tc_moneda_usd,
                "cif_usd": cif_usd,
                "cif_clp": cif_clp,
                "cif_moneda": cif_moneda,
            }
            lineas_info.append(info)
            cif_por_parcial[parcial.id] += cif_clp

    # --- Paso 3: pool de gastos por parcial, separado en 3 baldes para la
    # visualizacion (Flete Int., Derechos, y el resto de los gastos) ---
    flete_gasto_por_parcial = defaultdict(float)
    derechos_por_parcial = defaultdict(float)
    otros_gastos_por_parcial = defaultdict(float)
    detalle_gastos_por_parcial = defaultdict(list)

    for gasto in importacion.gastos:
        aplicables = list(gasto.parciales_aplicables) or list(importacion.parciales)
        aplicables = [p for p in aplicables if p.importacion_id == importacion.id]
        if not aplicables:
            continue
        total_cif_aplicables = sum(cif_por_parcial.get(p.id, 0) for p in aplicables)
        if gasto.concepto == CONCEPTO_FLETE_GASTO:
            balde = flete_gasto_por_parcial
        elif gasto.concepto == CONCEPTO_DERECHOS:
            balde = derechos_por_parcial
        else:
            balde = otros_gastos_por_parcial
        for p in aplicables:
            if total_cif_aplicables > 0:
                share = cif_por_parcial.get(p.id, 0) / total_cif_aplicables
            else:
                share = 1.0 / len(aplicables)
            monto_asignado = (gasto.monto_clp or 0) * share
            balde[p.id] += monto_asignado
            detalle_gastos_por_parcial[p.id].append(
                {"concepto": gasto.concepto, "monto": monto_asignado, "referencia": gasto.referencia}
            )

    def _pool_total(parcial_id):
        return (
            flete_gasto_por_parcial.get(parcial_id, 0)
            + derechos_por_parcial.get(parcial_id, 0)
            + otros_gastos_por_parcial.get(parcial_id, 0)
        )

    # --- Paso 4: distribuir cada balde de cada parcial a sus lineas por CIF, costo final ---
    for info in lineas_info:
        parcial_id = info["parcial"].id
        cif_parcial = cif_por_parcial.get(parcial_id, 0)
        share = (info["cif_clp"] / cif_parcial) if cif_parcial else 0
        info["flete_gasto_clp"] = flete_gasto_por_parcial.get(parcial_id, 0) * share
        info["derechos_clp"] = derechos_por_parcial.get(parcial_id, 0) * share
        info["otros_gastos_clp"] = otros_gastos_por_parcial.get(parcial_id, 0) * share
        info["gastos_clp"] = info["flete_gasto_clp"] + info["derechos_clp"] + info["otros_gastos_clp"]
        info["costo_total_clp"] = info["cif_clp"] + info["gastos_clp"]
        cantidad = info["linea"].cantidad_unidades or 0
        info["costo_unitario_clp"] = (info["costo_total_clp"] / cantidad) if cantidad else 0

    # --- Resumen por parcial (incluye el detalle en moneda de la factura y
    # el prorrateo de Flete/Seguro/Otros items que le corresponde SOLO a
    # ese parcial -- punto 3, 2026-09-01) ---
    resumen_parciales = []
    for parcial in importacion.parciales:
        info_parcial = [i for i in lineas_info if i["parcial"].id == parcial.id]
        fob_usd_parcial = sum(i["valor_usd"] for i in info_parcial)
        fob_moneda_parcial = sum(i["valor_total_moneda"] for i in info_parcial)
        flete_usd_parcial = sum(i["flete_usd"] for i in info_parcial)
        seguro_usd_parcial = sum(i["seguro_usd"] for i in info_parcial)
        otros_items_usd_parcial = sum(i["otros_items_usd"] for i in info_parcial)
        cif_usd_parcial = sum(i["cif_usd"] for i in info_parcial)
        resumen_parciales.append({
            "parcial": parcial,
            "fob_usd": fob_usd_parcial,
            "fob_moneda": fob_moneda_parcial,
            "flete_usd": flete_usd_parcial,
            "flete_moneda": flete_usd_parcial * tc_moneda_usd,
            "seguro_usd": seguro_usd_parcial,
            "seguro_moneda": seguro_usd_parcial * tc_moneda_usd,
            "otros_items_usd": otros_items_usd_parcial,
            "otros_items_moneda": otros_items_usd_parcial * tc_moneda_usd,
            "cif_usd": cif_usd_parcial,
            "cif_clp": cif_por_parcial.get(parcial.id, 0),
            "cif_moneda": cif_usd_parcial * tc_moneda_usd,
            "flete_gasto_clp": flete_gasto_por_parcial.get(parcial.id, 0),
            "derechos_clp": derechos_por_parcial.get(parcial.id, 0),
            "otros_gastos_clp": otros_gastos_por_parcial.get(parcial.id, 0),
            "gastos_clp": _pool_total(parcial.id),
            "detalle_gastos": detalle_gastos_por_parcial.get(parcial.id, []),
            "costo_total_clp": sum(i["costo_total_clp"] for i in info_parcial),
            "lineas": info_parcial,
            # Estas dos son POR PARCIAL (a diferencia de usa_flete_item/
            # usa_seguro_item/usa_otros_items, que son de toda la
            # importacion porque esos items siempre se reparten a TODOS los
            # parciales por FOB-share). Un Gasto "Flete Int."/"Derechos"
            # puede aplicar solo a un parcial (ver GastoImportacion.
            # parciales_aplicables) -- si el Parcial 1 no tiene Derechos
            # asignado, esa columna no debe aparecer en SU tabla aunque el
            # Parcial 2 si tenga (ronda G, punto 3).
            "usa_flete_gasto": abs(flete_gasto_por_parcial.get(parcial.id, 0)) > 0.5,
            "usa_derechos": abs(derechos_por_parcial.get(parcial.id, 0)) > 0.5,
        })

    cif_usd_total = sum(i["cif_usd"] for i in lineas_info)
    cif_clp_total = sum(i["cif_clp"] for i in lineas_info)
    costo_total_clp = sum(i["costo_total_clp"] for i in lineas_info)
    costo_total_usd = (costo_total_clp / tc_aduanero) if tc_aduanero else 0

    flete_total_usd = items_por_concepto_usd.get(CONCEPTO_ITEM_FLETE, 0)
    seguro_total_usd = items_por_concepto_usd.get(CONCEPTO_ITEM_SEGURO, 0)
    otros_items_total_usd = items_total_usd - flete_total_usd - seguro_total_usd

    # Hay columnas propias solo si algo las usa -- para no mostrar columnas
    # vacias en importaciones que no las necesitan. Se usa una tolerancia
    # (medio centavo) en vez de "!= 0" porque las restas en punto flotante
    # (items_total_usd - flete - seguro) pueden dejar un residuo minusculo
    # (ej. -0.00000000003) que con "!= 0" se leia como "si tiene datos" y
    # mostraba una columna "Otros items" vacia (punto 4, 2026-09-01 ronda F).
    TOLERANCIA_USD = 0.005
    usa_flete_gasto = any(g.concepto == CONCEPTO_FLETE_GASTO for g in importacion.gastos)
    usa_derechos = any(g.concepto == CONCEPTO_DERECHOS for g in importacion.gastos)
    usa_flete_item = abs(flete_total_usd) > TOLERANCIA_USD
    usa_seguro_item = abs(seguro_total_usd) > TOLERANCIA_USD
    usa_otros_items = abs(otros_items_total_usd) > TOLERANCIA_USD

    totales = {
        "fob_usd": fob_total_usd,
        "fob_moneda_factura": fob_moneda_factura,
        "flete_usd": flete_total_usd,
        "flete_moneda": flete_total_usd * tc_moneda_usd,
        "seguro_usd": seguro_total_usd,
        "seguro_moneda": seguro_total_usd * tc_moneda_usd,
        "otros_items_usd": otros_items_total_usd,
        "otros_items_moneda": otros_items_total_usd * tc_moneda_usd,
        "cif_usd": cif_usd_total,
        "cif_clp": cif_clp_total,
        "cif_moneda_factura": cif_usd_total * tc_moneda_usd,
        "flete_gasto_clp": sum(flete_gasto_por_parcial.values()),
        "derechos_clp": sum(derechos_por_parcial.values()),
        "gastos_clp": sum(otros_gastos_por_parcial.values()) + sum(flete_gasto_por_parcial.values()) + sum(derechos_por_parcial.values()),
        "costo_total_clp": costo_total_clp,
        "costo_total_usd": costo_total_usd,
        "costo_total_moneda_factura": costo_total_usd * tc_moneda_usd,
    }

    return {
        "lineas": lineas_info,
        "parciales": resumen_parciales,
        "totales": totales,
        "usa_flete_gasto": usa_flete_gasto,
        "usa_derechos": usa_derechos,
        "usa_flete_item": usa_flete_item,
        "usa_seguro_item": usa_seguro_item,
        "usa_otros_items": usa_otros_items,
    }
