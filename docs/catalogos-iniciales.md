# Catálogos y parámetros iniciales

Fuente: `doc-oficial.md` v0.5 (2026-10-29), secciones 4.11 (Parámetros) y 5 (Catálogos). Este archivo fija los datos iniciales que `tools/crear_libro.py` carga en las tablas de `libro/oee_v0.5_AAAA-MM-DD.xlsm`.

Reglas aplicadas (sección 5 del Doc):

- Código estable y único; nunca se reutiliza ni se cambia.
- No se borran filas: se marcan como inactivas (`activo` = No).
- Columna `activo` (Sí/No) en todos los catálogos: validación de lista.
- Los valores nuevos se agregan a la lista sin revisión (RN-08).

Notas de implementación (no son valores nuevos, son vacíos deliberados):

- `vigente_hasta` vacío en todas las filas iniciales (siguen vigentes).
- `fecha_aprobacion` vacío en `tblParametros`: el Doc no da fecha por fila, solo "Aprobado"/"Propuesto".
- `motivo` vacío en `tblParametros` iniciales: es el motivo de cada cambio posterior; la carga inicial no tiene cambio previo.
- `cruza_medianoche`, `td_minutos` y `descanso_minutos` en `tblTurnos` vacías: columnas calculadas, no se cargan (5.2).
- Códigos `SE-*`, `MR-*`, `IN-*`, `PU-*`, `VEL-*` y de productos: el Doc no les asigna códigos; se asignan aquí con el mismo patrón `XX-nn` de TP/OR/MO. Si el ingeniero define otros, reemplazar (los códigos deben ser estables a partir de esa aprobación).

## 1. `01_Parámetros` → `tblParametros`

Sección 4.11 (8 filas). `producto_codigo` referencia `tblProductos`. Vigencia: P-03 a P-07 con vigencia hacia atrás supuesta 2026-01-01 (nota del Doc); P-01, P-02 y P-08 cargan `vigente_desde` = 2026-10-29 (fecha de la decisión del Doc; P-01 se actualizó ese día, P-02 y P-08 son parámetros globales aprobados y cargados en esta versión).

| id | nombre | parametro | ambito | producto_codigo | valor | unidad | vigente_desde | motivo | aprobado_por | estado |
|---|---|---|---|---|---|---|---|---|---|---|
| P-01 | Umbral de microparada | UMBRAL_MICROPARADA_MIN | Global | | 5 | min | 2026-10-29 | | Ingeniero + reg. 13 | Aprobado |
| P-02 | Descanso reglamentario | DESCANSO_MIN | Global | | 35 | min | 2026-10-29 | | Ingeniero + reg. 13 | Aprobado |
| P-03 | Velocidad nominal 500 sin gas | VEL_NOM_500_SG | Producto | 500SG | 90 | botellas/min | 2026-01-01 | | Ingeniero | Aprobado (vigencia supuesta) |
| P-04 | Velocidad nominal 500 con gas | VEL_NOM_500_CG | Producto | 500CG | 70 | botellas/min | 2026-01-01 | | Ingeniero | Aprobado (vigencia supuesta) |
| P-05 | Velocidad nominal 1500 sin gas | VEL_NOM_1500_SG | Producto | 1500SG | 70 | botellas/min | 2026-01-01 | | Ingeniero | Aprobado (vigencia supuesta) |
| P-06 | Velocidad nominal 2000 sin gas | VEL_NOM_2000_SG | Producto | 2000SG | 70 | botellas/min | 2026-01-01 | | Ingeniero | Aprobado (vigencia supuesta) |
| P-07 | Velocidad nominal bidón 6 L | VEL_NOM_BIDON_6L | Producto | BID6 | 10 | bidones/min | 2026-01-01 | | Ingeniero | Aprobado (vigencia supuesta) |
| P-08 | Tolerancia de comparación | TOLERANCIA_COMPARACION_PP | Global | | 0,05 | puntos porcentuales | 2026-10-29 | | Ingeniero + reg. 13 | Aprobado |

Columnas de `tblParametros` sin fila inicial: `vigente_hasta` (vacía), `motivo` (vacío), `fecha_aprobacion` (vacía; ver notas).

## 2. `02_Cat_Paradas` → `tblTiposParada`

Sección 5.1 (7 filas), columna `activo` = Sí en todas.

| codigo | tipo | clasificacion | efecto_calculo | seleccionable | activo | estado |
|---|---|---|---|---|---|---|
| TP-01 | Parada programada | PP | Descuenta TPL; no afecta OEE | Sí | Sí | Aprobado |
| TP-02 | Parada no planificada | PNP | Cabecera de clasificación; por sí sola no calcula efecto | No (cabecera) | Sí | Propuesto |
| TP-03 | Avería o falla | PNP / A/F | Resta TO; baja Disponibilidad | Sí (solo >5 min) | Sí | Aprobado |
| TP-04 | Ajuste y configuración | PNP / A y C | Resta TO; baja Disponibilidad | Sí (solo >5 min) | Sí | Aprobado |
| TP-05 | Pequeña parada / microparada | PNP / micro | No resta TO; afecta Rendimiento | Sí (automática ≤5 min) | Sí | Aprobado |
| TP-06 | Sin clasificar | PNP provisoria | Resta TO hasta clasificar | Sí (>5 min) | Sí | Aprobado |
| TP-07 | Parada externa | PE | Descuenta TPL; no afecta OEE | Sí | Sí | Aprobado |

## 3. `02_Cat_Paradas` → `tblOrigenes`

Sección 5.1: "código, nombre, clase (máquina de la línea / servicio de apoyo / otro), activo" (13 filas, OR-01 a OR-13 de 8.1).

| codigo | nombre | clase | activo |
|---|---|---|---|
| OR-01 | sopladora | máquina de la línea | Sí |
| OR-02 | llenadora | máquina de la línea | Sí |
| OR-03 | etiquetadora | máquina de la línea | Sí |
| OR-04 | forradora | máquina de la línea | Sí |
| OR-05 | refrigeración de agua | servicio de apoyo | Sí |
| OR-06 | tratamiento de agua | servicio de apoyo | Sí |
| OR-07 | suministro de agua | servicio de apoyo | Sí |
| OR-08 | descanso/personal | otro | Sí |
| OR-09 | cambio de formato/producto | otro | Sí |
| OR-10 | mantenimiento programado | otro | Sí |
| OR-11 | reunión/capacitación | otro | Sí |
| OR-12 | limpieza/sanitizado programado | otro | Sí |
| OR-13 | otro/no clasificado | otro | Sí |

## 4. `02_Cat_Paradas` → `tblMotivos`

Sección 5.1: "código, motivo, origen, tipo sugerido, activo". Los motivos son una lista abierta (MO-01 a MO-99, 8.1); estas 5 filas son los ejemplos dados en el Doc. `origen` referencia `tblOrigenes`. `tipo_sugerido` vacío: el Doc no asigna tipo a estos ejemplos (el tipo lo elige quien carga; ver 5.1).

| codigo | motivo | origen | tipo_sugerido | activo |
|---|---|---|---|---|
| MO-01 | volcador de preforma | OR-01 | | Sí |
| MO-02 | atasco a la salida de la etiquetadora | OR-03 | | Sí |
| MO-03 | limpieza de adhesivo | OR-03 | | Sí |
| MO-04 | falla de refrigeración por alta temperatura | OR-05 | | Sí |
| MO-05 | falla en horno | OR-01 | | Sí |

## 5. `03_Cat_Turnos` → `tblTurnos`

Sección 5.2: "código, nombre, equipo (SOP / RES), hora de inicio, hora de fin, hora de inicio y fin del descanso, días en que aplica, activo, vigencia desde / hasta" + columnas calculadas vacías `cruza_medianoche`, `td_minutos`, `descanso_minutos`.

RES (tabla de 5.2). SOP (5.2: 4 turnos de 6 h, 7 días; nombre = código; descanso con horario no definido, se dejan vacías las horas de descanso).

| codigo | nombre | equipo | hora_inicio | hora_fin | descanso_inicio | descanso_fin | dias_aplica | activo |
|---|---|---|---|---|---|---|---|---|
| RES-M | Mañana RES | RES | 06:00 | 14:00 | 11:55 | 12:30 | Lun–vie; sáb solo mañana | Sí |
| RES-T | Tarde RES | RES | 14:00 | 22:00 | 18:55 | 19:30 | Lun–vie | Sí |
| RES-N | Noche RES | RES | 22:00 | 06:00 | 02:55 | 03:30 | Lun–vie | Sí |
| SOP-00 | SOP-00 | SOP | 00:00 | 06:00 | | | Lun–dom | Sí |
| SOP-06 | SOP-06 | SOP | 06:00 | 12:00 | | | Lun–dom | Sí |
| SOP-12 | SOP-12 | SOP | 12:00 | 18:00 | | | Lun–dom | Sí |
| SOP-18 | SOP-18 | SOP | 18:00 | 00:00 | | | Lun–dom | Sí |

Notas:

- RES-N cruza la medianoche: pertenece a la fecha en que empieza (Q-04); `cruza_medianoche` se calcula, no se carga.
- El descanso de SOP son 35 min (P-02) pero el horario no está definido: horas de descanso vacías.
- `vigente_desde` / `vigente_hasta` vacías: el Doc no da vigencia por turno.
- Nombre de los turnos SOP = código: el Doc no les asigna nombre propio.

## 6. `04_Cat_Productos` → `tblProductos`

Sección 5.3: "código, producto, formato (botella / bidón), volumen (litros), unidades por pack, packs por pallet, activo". Códigos derivados (ver notas).

| codigo | producto | formato | volumen_litros | unidades_por_pack | packs_por_pallet | activo |
|---|---|---|---|---|---|---|
| 500SG | 500 sin gas | botella | 0,5 | 12 | 150 | Sí |
| 500CG | 500 con gas | botella | 0,5 | 12 | 132 | Sí |
| 1500SG | 1500 sin gas | botella | 1,5 | 6 | 100 | Sí |
| 2000SG | 2000 sin gas | botella | 2 | 6 | 80 | Sí |
| BID6 | Bidón | bidón | 6 | 1 | 168 | Sí |

- 1 (no se hace pack) para el bidón (tabla 5.3).
- Punto de conteo: fin de línea. Unidad de producción: botella / bidón (4.6).

## 7. `04_Cat_Productos` → `tblVelocidades`

Sección 5.3.1: "producto, velocidad (botellas/min), vigencia desde, origen del dato" + `codigo` y `activo` por reglas globales (5). `producto` referencia `tblProductos`.

| codigo | producto | velocidad | vigencia_desde | origen_dato | activo |
|---|---|---|---|---|---|
| VEL-01 | 500SG | 90 | 2026-01-01 | ingeniero | Sí |
| VEL-02 | 500CG | 70 | 2026-01-01 | ingeniero | Sí |
| VEL-03 | 1500SG | 70 | 2026-01-01 | ingeniero | Sí |
| VEL-04 | 2000SG | 70 | 2026-01-01 | ingeniero | Sí |
| VEL-05 | BID6 | 10 | 2026-01-01 | ingeniero | Sí |

- Velocidad del bidón en bidones/min (5.3.1).
- Vigencia supuesta 2026-01-01: nota de 4.11 (P-03 a P-07); origen del dato: ingeniero (Q-09).

## 8. `05_Cat_Calidad` → `tblSeccionesRechazo`

Sección 5.4: secciones de la línea + `codigo` y `activo` por reglas globales. Códigos derivados.

| codigo | seccion | activo |
|---|---|---|
| SE-01 | sopladora | Sí |
| SE-02 | llenadora | Sí |
| SE-03 | etiquetadora | Sí |
| SE-04 | forradora | Sí |

## 9. `05_Cat_Calidad` → `tblMotivosRechazo`

Sección 5.4: "motivo, sección, activo" + `codigo` por reglas globales. Los 5 motivos son los ejemplos dados en el Doc. `seccion` referencia `tblSeccionesRechazo`.

| codigo | motivo | seccion | activo |
|---|---|---|---|
| MR-01 | botella mal formada o mal tapada | SE-02 | Sí |
| MR-02 | botella aplastada por trabado de máquina | SE-02 | Sí |
| MR-03 | etiqueta mal puesta que no se puede recuperar | SE-03 | Sí |
| MR-04 | pico sucio de pegamento | SE-03 | Sí |
| MR-05 | rotura de envase | SE-01 | Sí |

- Lista abierta; valores nuevos sin revisión (RN-08).
- Retrabajo no es rechazo (RN-06): una botella con nivel bajo que se destapa y reprocesa no entra.

## 10. `06_Cat_Insumos` → `tblInsumos`

Sección 5.5: "insumo, unidad de medida (unidades o peso), motivo, activo" + `codigo` por reglas globales. Códigos derivados.

| codigo | insumo | unidad_medida | motivo | activo |
|---|---|---|---|---|
| IN-01 | preformas | unidades | | Sí |
| IN-02 | tapas | unidades | | Sí |
| IN-03 | etiquetas | peso | | Sí |
| IN-04 | termocontraíble | peso | | Sí |
| IN-05 | film stretch | peso | | Sí |

- `motivo` vacío: el Doc da motivos como ejemplos ("preformas quemadas por una parada de la sopladora", "preformas caídas al piso") pero no los lista como filas iniciales; es una lista abierta que se carga con el uso.
- Unidades: preformas y tapas en unidades; termocontraíble, etiquetas y film stretch en peso (5.5). "A confirmar por insumo" (gramos o kilos).

## 11. `07_Cat_Personal` → `tblPuestos`

Sección 5.6: "puesto, equipo (SOP / RES), activo" + `codigo` por reglas globales. Códigos derivados.

| codigo | puesto | equipo | activo |
|---|---|---|---|
| PU-01 | soplador | SOP | Sí |
| PU-02 | ayudante | SOP | Sí |
| PU-03 | líder | RES | Sí |
| PU-04 | llenador | RES | Sí |
| PU-05 | visor | RES | Sí |
| PU-06 | etiquetador | RES | Sí |
| PU-07 | forradora (final de línea) | RES | Sí |
| PU-08 | estibador | RES | Sí |
| PU-09 | autoelevadorista | RES | Sí |

- El Doc indica "confirmar nombre del séptimo puesto" de RES: PU-09 queda con el nombre propuesto (autoelevadorista), pendiente de confirmación.
- Nombres del personal (personas): Esperando definición (5.6); no hay filas de personas.
