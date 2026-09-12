# Ronda R (2026-09-12): Usuarios, Perfiles y Control de Acceso

## Qué se pidió

Después de cerrar los ajustes de formato de la planilla de Inventarios (rondas O/P/Q), el usuario
pidió construir el sistema de usuarios/perfiles/permisos para controlar el acceso a la Suite:

> "Ahora podemos pasar la creacion de los perfiles, usuarios para el control de acceso. 1 usuario
> Administrador que solo pueda editar, modificar, ver las claves y usuarios. Considerar Los
> perfiles a los que cada perfil tendria acceso, ademas rol de aprobacion de las ordenes. Quiero:
> Perfiles o Permisos que pueden tener: 1- Creacion de Orden 2- Aprobacion de Orden 3- Actualizar
> despacho 4- Generar Costeo 5- Creacion de Orden Simple sin Visualizar precios, solo incorpora
> producto y propone emision de orden. 6- Inventarios 7- Acceso a Reportes."

Con dos rondas de preguntas de aclaración se definieron las reglas de negocio exactas (ver abajo).

## Decisiones de diseño (confirmadas con el usuario)

1. **Administrador = superusuario total.** No es un perfil separado con acceso limitado a
   usuarios/claves — administra usuarios/perfiles Y automáticamente tiene las 7 funciones
   operativas también.
2. **Aprobación de Orden es un permiso, no una persona.** Cualquier usuario con el permiso
   "Aprobación de Orden" puede confirmar/aprobar líneas de cualquier orden — puede haber varios
   aprobadores a la vez.
3. **Creación de Orden Simple** (perfil #5): pensado para alguien que NO debe ver los precios ya
   cargados en el catálogo de ningún proveedor.
   - Si escribe un código que ya existe en el catálogo del proveedor elegido, se usa el producto y
     precio **ya cargados** — el precio que haya escrito se descarta sin mostrárselo.
   - Si el código es nuevo, se crea un `Producto` real en el catálogo de ese proveedor con el
     precio que escribió (así queda disponible para todos, no es un dato suelto).
   - La orden resultante sigue el flujo normal de siempre ("Emisión de Orden", como cualquier
     otra) — no hay un estado especial "Propuesta". El único freno real es que sin el permiso de
     Aprobación de Orden nadie puede confirmarla.
   - Este usuario tampoco puede ver el detalle de la orden después de creada (no tiene acceso al
     módulo completo de Órdenes), así que nunca ve precios de catálogo por ningún otro camino.
4. **Permiso nuevo agregado a pedido del usuario: "Seguimiento".** Solo ver el estado/tracking de
   los despachos (sin poder editar nada) — separado de "Actualizar despacho", que sí permite
   editar.

## Los 8 permisos finales

| Código | Nombre | Qué habilita |
|---|---|---|
| `crear_orden` | Creación de Orden | Crear/editar órdenes completas, ver catálogo de precios, gestionar Proveedores/Productos |
| `aprobar_orden` | Aprobación de Orden | Confirmar/retroceder líneas de cualquier orden (cualquiera con este permiso puede aprobar) |
| `actualizar_despacho` | Actualizar despacho | Crear/editar/asociar despachos, marcar aduana/recibido |
| `generar_costeo` | Generar Costeo | Módulo completo de Importaciones/Costeo (parciales, gastos, exportar a Inventarios), también ve Proveedores y Despachos |
| `orden_simple` | Creación de Orden Simple | Solo el formulario `/ordenes/simple/nueva` — nunca ve catálogo de precios |
| `inventarios` | Inventarios | Exportar la planilla de Inventarios desde una Importación (además de `generar_costeo`) |
| `reportes` | Acceso a Reportes | Ver (solo lectura) el listado/comparativo de Importaciones |
| `seguimiento` | Seguimiento de despachos | Solo ver el estado/tracking de despachos (sin editar) |

Un **Rol** (perfil) es una combinación de estos permisos, más el flag maestro `es_administrador`
que los da todos automáticamente. Cada Usuario tiene exactamente un Rol.

## Qué se construyó

- **`models.py`**: modelos `Rol` (nombre, `es_administrador`, un booleano `permiso_<codigo>` por
  cada permiso) y `Usuario` (`UserMixin` de Flask-Login, contraseña con hash `werkzeug.security`,
  `rol_id`, `activo`, `ultimo_acceso`). Lista `PERMISOS_DISPONIBLES` como fuente única de verdad
  de los 8 permisos (se usa tanto para las pantallas de administración como para los decoradores).
- **`app.py`**:
  - Flask-Login con `before_request` global: TODA la app exige sesión iniciada excepto `/login` y
    los archivos estáticos.
  - Decoradores `@requiere_permiso("perm1", "perm2", ...)` (pasa si el usuario tiene AL MENOS UNO)
    y `@requiere_admin`, aplicados a las ~89 rutas existentes según a qué módulo pertenece cada una.
  - Rutas nuevas: `/login`, `/logout`, `/mi-cuenta` (cambio de contraseña), `/ordenes/simple/nueva`,
    y el CRUD de Configuración › Perfiles y Configuración › Usuarios.
  - `seed_administrador_inicial()`: crea el primer Rol "Administrador" + el primer Usuario
    (`jesusrfr3008@gmail.com`, contraseña temporal) la primera vez que arranca — sin usuarios
    previos en la base. No hace nada en arranques posteriores.
  - **Salvaguardas contra quedarse sin Administrador**: no se puede autoeliminar el propio usuario,
    no se puede autodesactivar, no se le puede quitar el perfil de Administrador al único
    administrador activo del sistema (ni editando el usuario ni editando el perfil).
- **Templates nuevos**: `login.html`, `mi_cuenta.html`, `configuracion/roles.html`,
  `configuracion/usuarios.html`, `ordenes/simple_form.html`.
- **`base.html`**: barra de navegación ahora muestra/oculta cada sección según los permisos del
  usuario logueado, agrega el menú de usuario (nombre, perfil, Mi cuenta, Cerrar sesión), y agrega
  Perfiles/Usuarios al menú Configuración (solo visible para Administradores).

## Pruebas realizadas (sandbox, antes de desplegar)

Se armó una copia aislada completa en un sandbox separado (nunca se tocó la base de datos real) y
se corrieron pruebas automatizadas cubriendo:

- Cualquier ruta protegida redirige a `/login` sin sesión iniciada; login con clave incorrecta no
  autentica.
- El Administrador sembrado inicialmente accede a absolutamente todo.
- Cada uno de los 8 permisos, probado de forma aislada (un usuario de prueba con solo ESE
  permiso), accede exactamente a lo que debe y es bloqueado (con mensaje + redirect al Dashboard)
  de todo lo demás — incluyendo el caso de "Creación de Orden Simple": crea un producto nuevo en
  el catálogo del proveedor con el precio tecleado, pero si el código ya existe usa el precio del
  catálogo y descarta el tecleado (se verificó explícitamente contra la base de datos).
- CRUD de Perfiles y Usuarios: alta, edición, nombres duplicados rechazados, no se puede eliminar
  un perfil con usuarios asignados, contraseñas quedan hasheadas (nunca en texto plano).
- Las salvaguardas anti-bloqueo (no autoeliminarse, no autodesactivarse, no quitarle el perfil de
  Administrador al único admin activo) se probaron y confirmaron.
- Regresión: con el Administrador logueado, se repitió el flujo completo de Importaciones/Costeo
  (incluida la exportación a Inventarios de las rondas O/P/Q), creación de una Orden completa,
  Despachos y Proveedores — todo sigue funcionando igual que antes de agregar el login.

## Despliegue

Antes de tocar la máquina real se hizo una copia de respaldo de los 4 archivos existentes
(`app.py`, `models.py`, `requirements.txt`, `templates/base.html`) en una carpeta
`_backup_antes_login/` dentro de la misma carpeta de la Suite, por si hiciera falta revertir.
Luego se copiaron los 9 archivos nuevos/modificados y se verificó que cada uno quedó con
exactamente el mismo tamaño en bytes que la versión de origen.

**Pendiente que el usuario debe hacer una sola vez:** correr `py -m pip install -r
requirements.txt` en su máquina para instalar `Flask-Login` (la nueva dependencia).

**Credenciales del Administrador inicial** (se crean solas la primera vez que arranque la app con
este código):
- Correo: `jesusrfr3008@gmail.com`
- Contraseña temporal: `CambiaEsta123!`

Se recomienda cambiarla de inmediato desde "Mi cuenta" apenas se ingresa la primera vez.

## Pendiente (explícitamente pausado por el usuario)

- El correlativo de Inventario de Accuvision: el usuario pidió esperar a tener datos reales antes
  de continuar la numeración.
- Migración a Railway/Postgres: sigue pausada, foco actual es seguir puliendo la app local.
