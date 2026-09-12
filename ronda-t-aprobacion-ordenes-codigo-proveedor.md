# Ronda T (2026-09-12): Aprobación de órdenes (Por Aprobar / Aprobada / Sin Emitir) y checkbox "Código Proveedor" en Inventario

## Qué se pidió

Después de probar en producción el flujo de aprobación de órdenes, el usuario reportó 2 problemas
y pidió una mejora:

1. Un usuario con perfil aprobador devolvió una línea a la etapa anterior y la orden le seguía
   apareciendo; al entrar con el usuario que la creó (perfil "Creación de Orden Simple"), la orden
   mostraba "no tienes permiso para acceder a esta sección".
2. Pidió explícitamente un cambio de fondo: que una orden nueva NO sea una Orden de Compra
   definitiva hasta que un usuario con permiso de Aprobación la apruebe. Mientras tanto no debe
   aparecer "en los registros" (el listado general de Compras/Órdenes) y solo debe verla quien la
   emitió. Pidió probar exactamente este proceso: crear una OC, que el aprobador la devuelva, que
   quede en una etapa "Sin Emitir" visible solo para el emisor, con los estados "Por Aprobar" y
   "Aprobada" antes de que la gestión de despacho tome el control.
3. En la descarga de la planilla de Inventario, un checkbox "Código Proveedor": si se marca, la
   columna B (código interno) debe listar el código del proveedor y la columna C no debe
   concatenar los 2 campos; si no se marca, se deja como estaba.

## Diagnóstico del punto 1

El mensaje "no tienes permiso para acceder a esta sección" **no es** el aviso de "esta orden no es
tuya" (ese dice otra cosa: "Solo puedes ver las órdenes que tú mismo creaste..."). Es el mensaje
genérico de cualquier pantalla para la que el usuario no tiene permiso en absoluto -- así que no
era un problema de qué orden se veía, sino de qué acceso existía. En paralelo se confirmó un bug
real e independiente: cuando el sistema separa automáticamente una orden en partes (división
legítima por despacho parcial, ver ronda anteriores), la parte nueva que crea no copiaba
`creado_por_usuario_id` ni el nuevo `estado_aprobacion` de la orden original -- quedaba "huérfana"
de su creador. Se corrigió en `dividir_orden_si_corresponde`, sea o no la causa exacta de lo
reportado.

La causa de fondo, sin embargo, era el pedido del punto 2: no existía ningún estado que impidiera
que una orden recién creada (o devuelta) actuara como una Orden de Compra normal desde el minuto
uno -- por eso "seguía apareciendo" en todos lados.

## Diseño acordado con el usuario (4 preguntas antes de tocar código)

Como este pedido revierte una decisión explícita de la ronda R ("mismo flujo de siempre, sin
estado especial"), antes de implementar se confirmaron 4 decisiones:

1. **Dónde ve el aprobador las órdenes pendientes**: una pestaña nueva "Por Aprobar" dentro de
   Compras/Órdenes (no aparecen mezcladas en el listado general).
2. **Alcance**: el usuario pidió que aplique a **todas las órdenes nuevas**, no solo a las creadas
   con el perfil "Creación de Orden Simple" -- también las que crea alguien con el perfil completo.
3. **Qué pasa al devolver una orden**: el mismo creador la edita y la reenvía (no hay que crear una
   orden nueva).
4. **Cómo se aprueba**: un solo botón para toda la orden (no línea por línea).

## Qué se construyó

### 1. Nuevo campo `OrdenCompra.estado_aprobacion`
Valores: `"Por Aprobar"`, `"Aprobada"`, `"Sin Emitir"` (constante `ESTADOS_APROBACION_OC` en
`models.py`). Es independiente del `estado`/`etapa` de línea de siempre (que sigue rastreando el
avance físico de la compra -- Emisión de Orden, Confirmada, Despachada, etc. -- pero **solo una vez
aprobada**). Toda orden nueva (desde `/ordenes/nueva` o `/ordenes/simple/nueva`) nace `"Por
Aprobar"`. Las órdenes que ya existían antes de esta ronda quedan `"Aprobada"` automáticamente vía
la migración, para no ocultar de golpe ningún dato ya en curso.

### 2. Pestaña "Por Aprobar" (`/ordenes/por-aprobar`)
- Quien tiene permiso de **Aprobación** ve ahí TODAS las órdenes "Por Aprobar" de cualquier
  creador, más las que él mismo devolvió y siguen "Sin Emitir" (para poder aprobarlas o
  devolverlas con un motivo opcional).
- Quien tiene permiso de **Creación de Orden** (completo, sin Aprobación) ve ahí únicamente las
  suyas -- "Por Aprobar" (a la espera) y "Sin Emitir" (para corregir y reenviar) -- funciona como
  su propio "Mis Órdenes", igual que ya existía para el perfil Orden Simple.
- El perfil "Creación de Orden Simple" sigue usando su "Mis Órdenes" de siempre (ronda S), que
  ahora también muestra el estado de aprobación y, si está "Sin Emitir", un botón "Reenviar a
  aprobación".
- Un badge en el menú "Compras / Órdenes" muestra cuántas órdenes están pendientes para ese
  usuario.

### 3. El listado general ("los registros") solo muestra órdenes Aprobadas
`/ordenes` (Compras/Órdenes) filtra por `estado_aprobacion = 'Aprobada'` (o `NULL`, para las
órdenes de antes de esta ronda). Una orden "Por Aprobar" o "Sin Emitir" nunca aparece ahí ni en
Despachos/Costeo de Importaciones -- de hecho no puede llegar a esos módulos porque las acciones de
aprobación por línea (Confirmar, Cambiar a etapa anterior) quedan bloqueadas, tanto en pantalla como
en el servidor, hasta que la orden completa esté Aprobada.

### 4. Acciones nuevas
- **Aprobar orden** (permiso Aprobación): pasa a `"Aprobada"` -- desde ahí sigue el mismo flujo de
  siempre, línea por línea, exactamente como ya funcionaba.
- **Devolver** (permiso Aprobación, con motivo opcional que queda registrado en las notas de la
  orden con fecha y autor): pasa a `"Sin Emitir"`.
- **Reenviar a aprobación**: quien creó la orden (con permiso de Creación de Orden completo, sobre
  cualquier orden -- mismo criterio de acceso que ya tenía Editar/Cancelar; o con el perfil Orden
  Simple, solo sobre la suya) la corrige y la reenvía -- vuelve a `"Por Aprobar"`.

### 5. Checkbox "Código Proveedor" en la descarga de Inventario
Junto al botón "Descargar planilla Inventario" (pantalla de Importación), un checkbox opcional:
- **Sin marcar** (default): igual que siempre desde la ronda O -- columna `CODIGO` (B) en blanco,
  columna `DESCRIPCION` (C) con "código descripción" concatenados.
- **Marcado**: columna `CODIGO` (B) lleva el código del producto (el mismo que identifica ese
  producto en el catálogo de ese proveedor), columna `DESCRIPCION` (C) queda solo con la
  descripción, sin concatenar.

No se tocó ningún cálculo de costeo ni el resto del formato (logo, subtotales, filtro, etc. de la
ronda Q).

## Pruebas realizadas (sandbox aislado, nunca la base real)

- Flujo completo (perfil "Creación de Orden"): crea una orden -> nace "Por Aprobar" -> no aparece
  en `/ordenes` ni para el creador ni para el aprobador -> sí aparece en "Por Aprobar" para ambos
  -> intentar Confirmar una línea antes de aprobar queda bloqueado (server-side) -> el aprobador
  aprueba -> ahora sí aparece en `/ordenes` y Confirmar funciona con normalidad.
- Devolver y reenviar (flujo completo): el aprobador devuelve con un motivo -> queda "Sin Emitir",
  el motivo queda en las notas -> el creador la reenvía -> vuelve a "Por Aprobar". Se confirmó que
  otro usuario con Creación de Orden completa también puede reenviar la orden ajena (mismo criterio
  de acceso total que ya regía Editar/Cancelar desde antes de esta ronda) y que un usuario que NO
  es ni el creador ni tiene permiso de Aprobación no la ve en su bandeja.
- **Escenario exacto pedido por el usuario**: crear una OC con el perfil "Creación de Orden
  Simple" -> nace "Por Aprobar", visible solo en "Mis Órdenes" del creador -> el aprobador la ve en
  su bandeja "Por Aprobar" y la devuelve -> queda "Sin Emitir", visible únicamente para el usuario
  que la creó (otro usuario con el mismo perfil simple no la ve ni por URL directa; el creador la
  sigue viendo SIN el mensaje "no tienes permiso") -> el creador la reenvía -> el aprobador la
  aprueba -> recién ahí aparece en "los registros" (Compras/Órdenes).
- Fix de `dividir_orden_si_corresponde`: una orden con 2 productos se aprueba, se confirma solo 1
  -> se divide en 2 (mismo N° de PO) -> se verificó que AMBAS partes conservan
  `creado_por_usuario_id` y `estado_aprobacion` de la original, y que el creador sigue viendo las 2
  en "Mis Órdenes".
- Checkbox de Inventario: se generó la planilla con y sin el checkbox marcado y se comparó celda
  por celda -- sin marcar, columna B vacía y C concatenada (comportamiento de siempre); marcado,
  columna B con el código y columna C solo con la descripción, sin el código duplicado.
- **Migración sin pérdida de datos**: se simuló la base real tal como está hoy (con una orden ya
  cargada, sin la columna `estado_aprobacion`) y se confirmó que al arrancar la app agrega la
  columna con todas las órdenes existentes en `"Aprobada"` automáticamente, sin perder ningún dato,
  y que esa orden sigue apareciendo con normalidad en el listado general.
- Regresión: se volvió a probar la validación de N° de PO único por empresa (ronda S) y la subida
  de documentos a una orden propia (ronda S) sobre el código nuevo -- siguen funcionando. Se hizo
  además un barrido de humo por todas las pantallas principales (Dashboard, Órdenes, Por Aprobar,
  Mis Órdenes, Despachos, Costeo de Importaciones, Proveedores, Mi cuenta) con 6 perfiles distintos
  (Administrador, Creación de Orden, Aprobación, 2 usuarios de Orden Simple, otro Creación de
  Orden) sin errores.

## Despliegue

Se hizo respaldo de los 8 archivos existentes que se modificaron (`app.py`, `models.py`,
`templates/base.html`, `templates/ordenes/detalle.html`, `templates/ordenes/list.html`,
`templates/ordenes/simple_list.html`, `templates/ordenes/simple_detalle.html`,
`templates/importaciones/detalle.html`) en una carpeta `_backup_antes_ronda_t/` dentro de la misma
carpeta de la Suite, antes de sobrescribirlos. Se copió además 1 archivo nuevo
(`templates/ordenes/por_aprobar.html`). Se verificó que los 9 quedaron con exactamente el mismo
tamaño en bytes que la versión de origen. No hubo cambios de dependencias (no hace falta volver a
correr `pip install`).

## Pendiente / fuera de este alcance

- No se restringió imprimir/enviar al proveedor por correo mientras una orden está "Por Aprobar" o
  "Sin Emitir" -- sigue visible solo según el permiso de Creación de Orden, igual que siempre; si el
  usuario prefiere bloquearlo también hasta la aprobación, es un ajuste chico para la próxima ronda.
- El Dashboard (conteo de "órdenes pendientes" y "últimas órdenes") no distingue todavía el estado
  de aprobación -- sigue contando/mostrando todas las órdenes sin importar si están Por
  Aprobar/Aprobada/Sin Emitir, tal como funcionaba antes de esta ronda.
- La numeración correlativa de Accuvision para el sistema de Inventarios sigue pausada hasta que el
  usuario tenga datos reales para empezar (sin cambios este round).
