# Estructura de datos del libro maestro

Fuente: `documentación_oficial_v0_8.md` (2026-11-04), sección 8. Este archivo fija hojas, tablas y encabezados exactos que usa `tools/crear_libro.py` para generar `libro/oee_master.xlsm`.

El VBA usa el **nombre de tabla**, no el de la hoja.

## 1. Hojas y tablas del libro

| Hoja | Tabla | Contenido |
|---|---|---|
| `00_Inicio` | — | Portada, versión, botones |
| `01_Parámetros` | `tblParametros` | Histórico con vigencia. Una fila por cambio |
| `02_Cat_Paradas` | `tblTiposParada` | TP-01 a TP-07 |
| | `tblOrigenes` | OR-01 a OR-13 |
| | `tblMotivos` | MO-01 a MO-99 |
| `03_Cat_Turnos` | `tblTurnos` | RES-M, RES-T, RES-N, SOP-00…18 |
| `04_Cat_Productos` | `tblProductos` | 500 SG/CG, 1500, 2000, bidón |
| | `tblVelocidades` | Velocidad por producto con vigencia |
| `05_Cat_Calidad` | `tblSeccionesRechazo` | Sopladora, llenadora, etiquetadora, forradora |
| | `tblMotivosRechazo` | Motivo + sección |
| `06_Cat_Insumos` | `tblInsumos` | Preforma, tapa, etiqueta, termocontraíble, film |
| `07_Cat_Personal` | `tblPuestos` | Soplador, ayudante, líder, llenador, visor, etiquetador, forradora, estibador, autoelevadorista |
| `10_Lotes` | `tblLotes` | `id_lote`, `numero_lote`, producto, inicio/fin |
| `11_Turnos` | `tblTurnosRegistro` | `id_turno`, `id_turno_lote`, `id_lote`, nro_reporte, fecha, turno_codigo, equipo, encargado, estado_revision, anulado, inicio/fin |
| `12_Paradas` | `tblParadas` | Códigos, no texto |
| `13_Produccion` | `tblProduccion` | Unidades, pallets, litros |
| `18_TramosSinEnvase` | `tblTramosSinEnvase` | Tramos sin envase de cada reporte: inicio y fin |
| `14_Rechazos` | `tblRechazos` | Por sección y motivo |
| `15_Reducciones` | `tblReducciones` | Formulario aparte |
| `16_PerdidasInsumo` | `tblPerdidasInsumo` | Seguimiento, no OEE |
| `17_PersonalTurno` | `tblPersonalTurno` | Trazabilidad |
| `20_Calculos` | — | Fórmulas de cálculo |
| `21_Dashboard` | — | Tablero |
| `90_Tests` | — | Casos de prueba |

## 2. Encabezados exactos (orden de columnas)

Nomenclatura fijada en este archivo:

- En minúscula y con `_` (snake_case), sin tildes. La forma en minúscula/tilde del Doc se conserva en los datos, no en el encabezado.
- `fecha`, `vigente_desde`, `vigente_hasta`, `fecha_aprobacion`, `vigencia_desde`: **fecha real** (formato `dd/mm/yyyy`, validación de fecha).
- `fecha_hora_inicio`, `fecha_hora_fin`: **fecha y hora reales** (formato `dd/mm/yyyy hh:mm`, validación de fecha).
- `hora_inicio`, `hora_fin`, `descanso_inicio`, `descanso_fin`: hora del día (formato `hh:mm`).
- `activo`: **lista de validación Sí/No** (regla de la sección 5 del Doc: columna `activo` (sí/no) en todos los catálogos).
- Se guardan los códigos del catálogo, no el texto.
- Sin celdas combinadas en ninguna hoja.

### 2.1 `01_Parámetros` → `tblParametros`

Sección 8.2 del Doc:

`id` · `nombre` · `parametro` · `ambito` · `producto_codigo` · `valor` · `unidad` · `vigente_desde` · `vigente_hasta` · `motivo` · `aprobado_por` · `estado` · `fecha_aprobacion`

- Fechas: `vigente_desde`, `vigente_hasta`, `fecha_aprobacion`.
- No tiene `activo` (es histórico con vigencia, no lista de opciones).

### 2.2 `02_Cat_Paradas` → `tblTiposParada`

Sección 5.1 (tabla de tipos + reglas globales con `código` y `activo`):

`codigo` · `tipo` · `clasificacion` · `efecto_calculo` · `seleccionable` · `activo` · `estado`

### 2.3 `02_Cat_Paradas` → `tblOrigenes`

Sección 5.1: "Columnas: código, nombre, clase (máquina de la línea / servicio de apoyo / otro), activo":

`codigo` · `nombre` · `clase` · `activo`

### 2.4 `02_Cat_Paradas` → `tblMotivos`

Sección 5.1: "Columnas: código, motivo, origen, tipo sugerido, activo":

`codigo` · `motivo` · `origen` · `tipo_sugerido` · `activo`

- `origen` guarda el código (`OR-xx`) del catálogo de orígenes.
- `tipo_sugerido` guarda el código del tipo (`TP-xx`); es ayuda, no obligación. Vacío si el Doc no asigna tipo.

### 2.5 `03_Cat_Turnos` → `tblTurnos`

Sección 5.2: "Columnas: código, nombre, equipo (SOP / RES), hora de inicio, hora de fin, hora de inicio y fin del descanso, días en que aplica, activo, vigencia desde / hasta" + "Calculadas (no se cargan): cruza medianoche (sí/no), duración TD en minutos, descanso en minutos (35)":

`codigo` · `nombre` · `equipo` · `hora_inicio` · `hora_fin` · `descanso_inicio` · `descanso_fin` · `dias_aplica` · `activo` · `vigente_desde` · `vigente_hasta` · `cruza_medianoche` · `td_minutos` · `descanso_minutos`

- Fechas: `vigente_desde`, `vigente_hasta`. Horas: `hora_inicio`, `hora_fin`, `descanso_inicio`, `descanso_fin`.
- `cruza_medianoche`, `td_minutos`, `descanso_minutos` son columnas calculadas: el script las crea **vacías** para que las llenen fórmulas/VBA.

### 2.6 `04_Cat_Productos` → `tblProductos`

Sección 5.3: "Columnas: código, producto, formato (botella / bidón), volumen (litros), unidades por pack, packs por pallet, activo":

`codigo` · `producto` · `formato` · `volumen_litros` · `unidades_por_pack` · `packs_por_pallet` · `activo`

### 2.7 `04_Cat_Productos` → `tblVelocidades`

Sección 5.3.1: "Columnas: producto, condición de envase (con envase / sin envase), velocidad, unidad, vigencia desde, origen del dato" + `código` y `activo` por las reglas globales de la sección 5:

`codigo` · `producto` · `condicion_envase` · `velocidad` · `unidad` · `vigencia_desde` · `origen_dato` · `activo`

- Fecha: `vigencia_desde`.
- `producto` guarda el código del producto (`500SG`, `500CG`, `1500SG`, `2000SG`, `BID6`).
- `condicion_envase`: "con envase" o "sin envase" (5.3.1).
- `unidad`: "botellas/min", "bidones/min", "botellas/hora" o "bidones/hora".

### 2.8 `05_Cat_Calidad` → `tblSeccionesRechazo`

Sección 5.4 (lista de secciones + reglas globales con `código` y `activo`):

`codigo` · `seccion` · `activo`

### 2.9 `05_Cat_Calidad` → `tblMotivosRechazo`

Sección 5.4: "Columnas: motivo, sección, activo" + `código` por las reglas globales:

`codigo` · `motivo` · `seccion` · `activo`

- `seccion` guarda el código de `tblSeccionesRechazo` (`SE-xx`).

### 2.10 `06_Cat_Insumos` → `tblInsumos`

Sección 5.5: "Columnas: insumo, unidad de medida (unidades o peso), motivo, activo" + `código` por las reglas globales:

`codigo` · `insumo` · `unidad_medida` · `motivo` · `activo`

- `motivo` es lista abierta de motivos de pérdida. El Doc no asigna motivos por insumo: se crea la columna y se deja vacía hasta que haya datos.

### 2.11 `07_Cat_Personal` → `tblPuestos`

Sección 5.6: "Columnas: puesto, equipo (SOP / RES), activo" + `código` por las reglas globales:

`codigo` · `puesto` · `equipo` · `activo`

### 2.12 `10_Lotes` → `tblLotes`

Sección 8.1: `id_lote`, `numero_lote`, producto, inicio/fin:

`id_lote` · `numero_lote` · `producto` · `fecha_hora_inicio` · `fecha_hora_fin`

### 2.13 `11_Turnos` → `tblTurnosRegistro`

Sección 8.1: reporte/segmento con `nro_reporte`, `encargado`, `estado_revision`, `anulado`, `notas_sin_horario`, `observaciones` y `velocidad_informada`, además de los campos de trazabilidad:

`id_turno` · `id_turno_lote` · `id_lote` · `nro_reporte` · `fecha` · `turno_codigo` · `equipo` · `encargado` · `estado_revision` · `anulado` · `notas_sin_horario` · `observaciones` · `velocidad_informada` · `fecha_hora_inicio` · `fecha_hora_fin` · `control_calidad` · `nro_registro_sopladora` · `registros_bolsones`

- `anulado` es una lista Sí/No. Nunca se borra un reporte: se anula.
- `estado_revision` guarda `OK`, `Pendiente de revisar` o `Anulado` según las validaciones de F2.3.
- `fecha_hora_inicio` y `fecha_hora_fin` representan el horario real anotado en el reporte y definen el TD (ya no son informativos).
- `control_calidad` y `nro_registro_sopladora`: texto libre (Q-26 y Q-23/Q-31).
- `registros_bolsones`: reservado para trazabilidad; no visible en Fase 2 (Q-31).

### 2.14 `12_Paradas` → `tblParadas`

Sección 8.2: "ID, fecha, turno, línea, `id_lote`, `id_turno_lote`, hora de inicio y fin (como fecha y hora completas), duración (en minutos), origen, motivo, tipo de parada (se asigna solo si dura 5 minutos o menos), código, interpretada (sí/no) y observaciones". `línea` se llama `linea_id` por el principio de 8.3:

`id` · `fecha` · `turno` · `linea_id` · `id_lote` · `id_turno_lote` · `fecha_hora_inicio` · `fecha_hora_fin` · `duracion_min` · `origen` · `motivo` · `tipo` · `codigo` · `interpretada` · `observaciones`

- Fechas: `fecha`. Fecha y hora: `fecha_hora_inicio`, `fecha_hora_fin`.
- `codigo` es el código de la parada (se genera solo, ver 4.9).
- `interpretada` es una lista Sí/No y vale `No` por defecto. Se marca `Sí` cuando administración interpretó el horario de una parada que el papel no traía con hora clara (Doc 4.9 y RN-11).

### 2.15 `13_Produccion` → `tblProduccion`

Sección 8.2: "fecha, turno, línea, lote, `id_turno_lote`, producto, unidades producidas (unidad base, botellas o bidones; se calculan desde los packs contados), pallets, litros (calculado y guardado) y velocidad nominal":

`fecha` · `turno` · `linea_id` · `lote` · `id_turno_lote` · `producto` · `unidades_producidas` · `pallets` · `litros` · `velocidad_nominal`

### 2.16 `14_Rechazos` → `tblRechazos`

Sección 8.2: "una fila por sección y motivo: fecha, turno, lote, `id_turno_lote`, sección, motivo, cantidad":

`fecha` · `turno` · `lote` · `id_turno_lote` · `seccion` · `motivo` · `cantidad` · `detalle`

- `detalle` (texto, opcional): por qué se rechazó (4.7).

### 2.17 `15_Reducciones` → `tblReducciones`

Sección 8.2: "fecha, turno, `id_turno_lote`, origen, motivo, hora de inicio y fin, velocidad real":

`fecha` · `turno` · `id_turno_lote` · `origen` · `motivo` · `fecha_hora_inicio` · `fecha_hora_fin` · `velocidad_real`

### 2.18 `16_PerdidasInsumo` → `tblPerdidasInsumo`

Sección 8.2: "fecha, turno, insumo, cantidad, unidad, motivo, `detalle`":

`fecha` · `turno` · `insumo` · `cantidad` · `unidad` · `motivo` · `detalle`

- `detalle` (texto, opcional).

### 2.19 `17_PersonalTurno` → `tblPersonalTurno`

Sección 8.2: "fecha, turno, puesto, persona":

`fecha` · `turno` · `puesto` · `persona`

### 2.20 `18_TramosSinEnvase` → `tblTramosSinEnvase`

Sección 8.2: "`id_turno_lote`, hora de inicio y fin del tramo sin envase (fecha y hora completas). Cero, uno o varios tramos por reporte; lo que no está en un tramo es con envase":

`id_turno_lote` · `fecha_hora_inicio` · `fecha_hora_fin`

- Fecha y hora: `fecha_hora_inicio`, `fecha_hora_fin`.
- Cero, uno o varios tramos por reporte; lo que no está en un tramo es con envase (RN-14).

## 3. Principios (sección 8.3 del Doc)

- Una fila por evento, fechas y horas reales, IDs únicos, `linea_id`, sin valores calculados en hojas de datos.
- Formularios de carga con validaciones.
- Dashboard: OEE y sus tres factores, tendencia y desglose de pérdidas.
- Sin celdas combinadas.
- **Estado:** Esperando definición (nombres finales de hoja y tabla).
