# Ronda S (2026-09-12): Documentos, edición del perfil simple, notas de aprobación y N° de PO único

## Qué se pidió

Después de desplegar el sistema de usuarios/perfiles (ronda R) y de que el usuario probara en
producción el flujo "Creación de Orden Simple", pidió 4 mejoras:

1. Que ese perfil pueda adjuntar documentos (Cotización, Orden de Compra del proveedor) al emitir
   la orden, para que quien aprueba pueda corroborar por qué se pide esa compra.
2. Que, una vez creada, ese mismo usuario pueda ver, editar y anular SU orden — pero nunca
   aprobarla ni avanzarla de etapa (eso sigue siendo trabajo exclusivo de quien tiene el permiso de
   Aprobación/Creación de Orden).
3. Reportó un aviso confuso en el detalle de una orden ("Este PO está dividido en 2 partes / Ver
   otra parte") que aparecía al aprobar una orden creada con el perfil simple.
4. Que los botones/opciones de una pantalla no aparezcan del todo cuando el perfil del usuario no
   tiene acceso a esa acción (en vez de mostrarlos y bloquear recién al hacer clic).

## Aclaración clave sobre el punto 3

Antes de tocar nada se le preguntó al usuario si el aviso debía ocultarse según el perfil de quien
lo ve. Su respuesta cambió el diagnóstico por completo: **no era una orden realmente dividida** —
en una prueba se había escrito a mano el mismo número de PO que ya tenía otra orden distinta
(sin relación alguna entre ambas). El sistema no tenía ninguna validación que impidiera esto, así
que las trató como si fueran partes de una misma división automática (una función legítima del
sistema: cuando se confirma solo una parte de los productos de una orden, esta se separa en dos
filas que comparten el mismo N° de PO a propósito, para que el proveedor identifique los envíos
parciales). La solución real no era ocultar el aviso — es una función que sigue siendo útil — sino
impedir que dos órdenes sin relación terminen compartiendo un número de PO.

**Regla implementada:** el N° de PO debe ser único por Empresa compradora (Accuvision,
Accumedical, etc.), en cualquier estado (emitida, en proceso, recibida o cancelada). Esto se valida
únicamamente al escribir un número a mano desde "Editar" (el único lugar donde se puede cambiar el
N° de PO manualmente) y solo cuando el número realmente cambia — así una orden ya dividida
legítimamente (que comparte su número con sus "hermanas" a propósito) se puede seguir guardando sin
problema. Si se intenta asignar un número que ya usa otra orden no relacionada de la misma empresa,
se bloquea con un mensaje claro indicando cuál es la orden en conflicto.

## Qué se construyó

### 1. Documentos adjuntos desde la creación (perfil simple)
El formulario de "Nueva Orden" simple ahora tiene dos campos de archivo opcionales — Cotización y
Orden de Compra del proveedor — además de un campo de notas. Se agregaron esos dos tipos a la lista
general de tipos de documento (`TIPOS_DOCUMENTO_ORDEN`), así quedan disponibles también desde la
pantalla de documentos del detalle completo.

### 2. "Mis Órdenes": ver / editar / anular las propias órdenes (perfil simple)
Se agregó un nuevo espacio exclusivo para este perfil (menú "Mis Órdenes" en la barra de
navegación, con "Nueva orden" y "Mis órdenes"):

- **Alcance:** cada usuario con este perfil solo ve las órdenes que ÉL MISMO creó (se agregó un
  campo `creado_por_usuario_id` a la orden) — nunca las de otro usuario ni el listado completo.
- **Precios:** si una línea usa un producto que YA existía en el catálogo (precio oculto desde la
  creación), ese precio sigue oculto incluso en su propia orden — se muestra un candado en vez del
  número, y el total de la orden excluye esas líneas (mostrando "Total visible" con una nota) para
  que no se pueda deducir el precio real restando el resto de los montos que sí conoce.
- **Edición:** puede agregar productos nuevos, ajustar cantidad y fecha, y agregar/editar el precio
  SOLO si es un producto que él mismo dio de alta (nunca si el precio viene del catálogo) — mismas
  reglas que al crear la orden, nunca un buscador de catálogo.
- **Anular:** puede anular una línea o la orden completa, pero solo mientras nada se haya
  confirmado todavía. En cuanto una línea pasa a "Orden Confirmada" (por alguien con permiso de
  Aprobación), este perfil ya no puede revertir esa decisión — el mensaje lo redirige a pedírselo a
  quien aprueba o crea órdenes.
- **Documentos:** puede subir, ver, descargar y eliminar documentos de sus propias órdenes desde
  esa misma pantalla.

### 3. Notas para quien aprueba
Se agregó una acción liviana "Agregar nota" (separada del modal completo de "Editar", que sigue
siendo exclusivo de quien tiene Creación de Orden): cualquiera con permiso de Aprobación — o el
perfil simple sobre su propia orden — puede dejar un comentario con fecha y su nombre, sin tocar
ningún otro dato de la orden (N° de PO, moneda, empresa). Sirve para dejar constancia de por qué se
aprobó o anuló algo, o para comentar sobre los documentos adjuntos.

### 4. Botones visibles solo según el permiso
En el detalle completo de una orden, ahora se muestran u ocultan según lo que el usuario realmente
puede hacer:

| Acción | Antes | Ahora |
|---|---|---|
| Editar cabecera (N° PO, moneda, notas) | Visible para cualquiera | Solo Creación de Orden |
| Agregar producto | Visible para cualquiera | Solo Creación de Orden |
| Cancelar / Reactivar orden | Visible para cualquiera | Solo Creación de Orden |
| Enviar al proveedor (Outlook/correo) | Visible para cualquiera | Solo Creación de Orden |
| Imprimir / Descargar PDF | Visible para cualquiera | Creación o Aprobación |
| Confirmar línea / Cambiar a etapa anterior | Visible para cualquiera | Solo Aprobación |
| Anular línea / Reactivar línea | Visible para cualquiera | Solo Creación de Orden |
| Subir/eliminar documentos | Visible para cualquiera | Solo Creación de Orden |
| Agregar nota | No existía | Aprobación (o perfil simple en su propia orden) |

Quien solo tiene Aprobación de Orden ahora ve una pantalla mucho más simple: puede revisar las
líneas, los documentos adjuntos, confirmar o devolver etapas, y dejar una nota — sin botones que de
todas formas le iban a bloquear el sistema al hacer clic.

## Pruebas realizadas (sandbox, antes de desplegar)

- Creación de una orden simple con los 2 documentos adjuntos y una nota: se guardan correctamente,
  quedan asociados al creador (`creado_por_usuario_id`).
- Línea con producto nuevo: precio visible y editable por el propio creador. Línea con producto ya
  existente en catálogo: precio oculto (candado), no editable, conserva el valor real del catálogo
  aunque se intente sobreescribirlo.
- Aislamiento: un segundo usuario con el mismo perfil no puede ver ni acceder a la orden del
  primero (ni por URL directa).
- Edición de línea propia: cantidad y fecha se actualizan siempre; el precio solo cuando no viene
  del catálogo.
- Anulación de la orden propia funciona mientras nada esté confirmado; se bloquea automáticamente
  en cuanto una línea pasa a "Orden Confirmada".
- Un usuario con Aprobación de Orden confirma una línea y agrega una nota — la nota queda con su
  nombre y fecha.
- Validación de N° de PO: bloquea asignar un número ya usado por otra orden de la misma empresa;
  permite cambiar a uno libre; re-guardar sin cambiar el número (caso de una orden ya dividida)
  sigue funcionando sin falsos bloqueos.
- Visibilidad de botones verificada en el HTML renderizado: un aprobador no ve "Editar", "Agregar
  producto" ni "Enviar al proveedor", pero sí el cuadro de notas; un creador ve todo lo suyo y no ve
  el cuadro de notas (usa el modal completo).
- Regresión completa: todos los tests de la ronda R (login, los 8 permisos, CRUD de perfiles y
  usuarios) y de la ronda O (exportación a Inventarios) se volvieron a correr sobre el código nuevo
  y siguen pasando.
- Migración de base de datos: se simuló una base de datos como la real (con las tablas de antes de
  esta ronda, sin las 2 columnas nuevas) y se confirmó que al arrancar la app agrega las columnas
  nuevas sin perder ningún dato existente.

## Despliegue

Se hizo respaldo de los 5 archivos existentes que se modificaron (`app.py`, `models.py`,
`templates/base.html`, `templates/ordenes/detalle.html`, `templates/ordenes/simple_form.html`) en
una carpeta `_backup_antes_ronda_s/` dentro de la misma carpeta de la Suite, antes de sobrescribirlos.
Se copiaron los 7 archivos nuevos/modificados y se verificó que cada uno quedó con exactamente el
mismo tamaño en bytes que la versión de origen. No hubo cambios de dependencias (no hace falta
volver a correr `pip install`).

## Pendiente / fuera de este alcance

- La numeración correlativa de Accuvision para el sistema de Inventarios sigue pausada hasta que el
  usuario tenga datos reales para empezar (sin cambios este round).
- No se construyó un reporte administrativo para detectar duplicados de N° de PO ya existentes en
  la base real — si el usuario quiere revisar si quedó algún duplicado de antes de esta ronda, puede
  ubicarlo por el aviso "está dividido en N partes" en el detalle de la orden y corregirlo a mano
  desde "Editar" (ahora protegido por la nueva validación).
