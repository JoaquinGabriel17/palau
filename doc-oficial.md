# Documentación oficial: cálculo de OEE de la línea de embotellado

Planta embotelladora de agua · Proyecto de cálculo de OEE en Excel con VBA

- **Versión del documento:** v0.5
- **Última actualización:** 2026-10-30

## 0. Reglas de este documento

- Este documento es la **única fuente de verdad** del proyecto.
- Si algo no está en este documento, no está definido. Si el código y el documento no coinciden, manda el documento.
- El dueño del documento es el ingeniero: él aprueba los cambios.
- Estados: **Esperando definición** (falta el dato), **Propuesto** (cargado a partir de lo conversado; falta aprobación) y **Aprobado** (con fecha y nombre de quien aprobó).
- Todo cambio se anota en la sección 13, con fecha y aprobación.

## 1. Objetivo y alcance

### 1.1 Objetivo

Calcular el OEE (Eficiencia General de los Equipos) de la línea de embotellado **[nombre de la línea: Línea 1 – Embotellado (supuesto)]** mediante un libro de Excel con VBA, con reportes por turno (día y semana: Esperando definición, ver 4.10).

- **Estado:** Aprobado (nombre de línea: supuesto).

### 1.2 Alcance

- **Línea medida:** Línea 1 – Embotellado (supuesto). Es la línea completa de embotellado, con 4 máquinas: sopladora, llenadora, etiquetadora y forradora.
- **Equipo o punto de medición:** Equipo de línea. Se mide la línea completa, no cada máquina por separado.
- **Producto(s) envasado(s):** Botellas (500 sin gas, 500 con gas, 1500 sin gas, 2000 sin gas) y bidones de 6 L.
- **Estado:** Aprobado (nombre de la línea: supuesto).

### 1.3 Fuera de alcance (por ahora)

- Sistema completo para otras áreas de la planta.
- Base de datos real y aplicación web.
- Otras líneas de producción.
- OEE de cada máquina por separado (requeriría una planilla de registro por máquina).
- Análisis detallado de las fallas de los servicios de apoyo (refrigeración de agua, tratamiento de agua, suministro de agua): se registran como paradas externas, sin detallar el componente que falló.
- **Estado:** Aprobado.

### 1.4 Criterio de éxito

El cálculo del libro coincide con el turno real calculado a mano (sección 7), con la tolerancia aprobada: el OEE y cada factor del libro coinciden con el cálculo manual a 1 decimal (diferencia de hasta 0,05 puntos porcentuales).

- **Estado:** Aprobado (tolerancia aprobada el 2026-10-28; aprobado por: ingeniero [confirmar nombre]).

## 2. Equipo y roles

- **Ingeniero (jefe de planta):** define las reglas y las fórmulas, es dueño de este documento y aprueba los cambios.
- **Joaquín:** desarrollo del libro de Excel con VBA.
- **Tercera persona del equipo:** rol [Esperando definición].
- **Supervisores / líderes de turno:** cargan los datos de paradas, producción, rechazos y pérdidas de insumo (ver 4.9).
- **Estado:** Propuesto.

## 3. Glosario y fórmulas

### 3.1 Fórmulas generales del OEE

**Estado:** Aprobado.

- **OEE = Disponibilidad × Rendimiento × Calidad**
- **Disponibilidad** = tiempo operativo / tiempo planificado de producción
- **Rendimiento** = (unidades producidas × tiempo de ciclo ideal) / tiempo operativo
- **Calidad** = unidades buenas / unidades totales
- **Unidades buenas** = unidades totales − unidades rechazadas
- **Forma simplificada:** OEE = (tiempo de ciclo ideal × unidades buenas) / tiempo planificado. Equivale a multiplicar los tres factores y sirve para verificar el cálculo.
- **Tiempo de ciclo ideal** = 1 / velocidad nominal (por ejemplo, 90 botellas/min equivalen a un ciclo de 1/90 min por botella).
- **Ni el OEE ni sus factores pueden superar 100%.** Si pasa, hay una variable mal definida (por ejemplo, una velocidad nominal desactualizada) y se revisa.
- Las pequeñas paradas y la reducción de velocidad no tienen un término propio en la fórmula: su efecto aparece de forma indirecta en las unidades producidas.

### 3.2 Escalera de tiempos

- Tiempo calendario.
- Tiempo planificado de producción (se descuenta lo que no se planeó producir: paradas programadas y paradas externas).
- Tiempo operativo (se descuentan las paradas no planificadas de **más de 5 minutos**: averías/fallas y ajustes y configuración; las pequeñas paradas no se descuentan).
- Tiempo plenamente productivo (se descuentan las pérdidas de rendimiento —pequeñas paradas y reducción de velocidad— y los rechazos).
- **Estado:** Aprobado.

### 3.3 Las tres familias de pérdida

- **Disponibilidad:** averías y/o fallas (A/F) y ajustes y configuración (A y C), de **más de 5 minutos**.
- **Rendimiento:** pequeñas paradas (**5 minutos o menos**) y reducción de velocidad.
- **Calidad:** productos rechazados, que fueron procesados pero no cumplen los requisitos. La pérdida de insumo no es rechazo y no entra en el cálculo, pero se registra para seguimiento (ver glosario 3.4).
- **Estado:** Aprobado.

### 3.4 Glosario

- **Tiempo disponible (TD):** tiempo total del turno.
- **Tiempo planificado de producción (TPL):** TD − PP − PE. Ver 4.2.
- **Tiempo operativo (TO):** TPL − A/F − A y C. Solo restan las paradas no planificadas de **más de 5 minutos**; las pequeñas paradas no se restan.
- **Parada planificada o programada (PP):** parada que se contempla antes del inicio del turno. Ej.: descanso por ley, limpieza y sanitizado programada, cambio de formato o producto, mantenimiento programado, reunión o capacitación. Definido en 4.3.
- **Parada externa (PE):** parada no relacionada con la línea que se quiere medir. Ej.: falla de refrigeración o del tratamiento de agua, falta de agua, falta de insumos. Definido en 4.3.
- **Parada no planificada (PNP):** parada no contemplada antes del inicio del turno, inesperada. Incluye averías/fallas, ajustes y configuración y pequeñas paradas. Definido en 4.3.
- **Avería o falla (A/F):** PNP de **más de 5 minutos** que requiere intervención de mantenimiento (ej.: mantenimiento correctivo).
- **Ajuste y configuración (A y C):** PNP de **más de 5 minutos** que resuelve el mismo operario (ej.: cambio de un acople del volcador de preformas).
- **Pequeña parada / microparada:** PNP de **5 minutos o menos** que el operario resuelve rápido (ej.: atasco de preforma o de botella, mover un sensor, regular una chapa). No se resta del tiempo operativo y afecta el Rendimiento.
- **Reducción de velocidad:** trabajar por debajo de la velocidad nominal sin parar (ej.: cuchilla de la etiquetadora sin filo; marcha a ritmo de sopladora). Afecta el Rendimiento. Se registra en un formulario aparte (ver 4.9).
- **Tiempo de ciclo ideal / velocidad nominal:** cantidad de unidades que la línea puede producir por unidad de tiempo en condiciones ideales, según el producto. Valores en 5.3, origen en 4.5.
- **Rechazo:** unidad que pasó por al menos una de las 4 máquinas de la línea y no cumple los requisitos de calidad. La unidad que se recupera (retrabajo) no es rechazo. Los insumos perdidos antes de procesar son **pérdida de insumo** y no son rechazo (no afectan el OEE).
- **Pérdida de insumo:** insumo (preforma, tapa, etiqueta, termocontraíble, film stretch) que se pierde sin haber sido procesado, o por una parada (ej.: preformas quemadas en el horno de la sopladora, preformas caídas al piso). No afecta el OEE; se registra para seguimiento.
- **Turno:** período de trabajo de un equipo de personas, con horario fijo. Ver 4.2.
- **Lote:** código de trazabilidad de una producción. Cambia cuando cambia el producto y, si no cambia, a las 00:00. Formato: 2 dígitos del mes + número secuencial de 01 a 999 (ejemplo: 09261). Ver 4.8.
- **Estado:** Aprobado.

## 4. Definiciones particulares de la planta

Cada punto se completa con el ingeniero. Formato: decisión, fecha, aprobado por, estado.

Lo cargado en esta versión (v0.5) sale de la reunión del 2026-10-21, de las respuestas posteriores del ingeniero (hasta 2026-10-28) y de las decisiones de Joaquín del 2026-10-29. Las secciones 1, 3, 4 y 5 quedaron aprobadas el 2026-10-27 (aprobado por: ingeniero [confirmar nombre]). Los puntos que dicen Esperando definición siguen abiertos.

### 4.1 Línea y equipo medido

- **Decisión:** se mide la línea completa de embotellado (sopladora, llenadora, etiquetadora y forradora), no cada máquina por separado.
- **Nombre o código de la línea:** Línea 1 – Embotellado (supuesto). Reemplazar cuando el ingeniero confirme el nombre.
- **Estado:** Aprobado (nombre: supuesto).

### 4.2 Turnos y tiempo planificado

- El OEE se calcula **por turno**.
- En la línea hay dos tipos de turnos: SOP (sopladora) y RES (resto de la línea).
- **SOP:** 4 turnos de 6 horas por día, 7 días de la semana, con 2 personas por turno. Horarios: 00:00–06:00, 06:00–12:00, 12:00–18:00 y 18:00–00:00.
- **RES:** 3 turnos de 8 horas: mañana 06:00–14:00, tarde 14:00–22:00 y noche 22:00–06:00 (cruza la medianoche). Lunes a viernes trabajan los tres turnos; el sábado, solo la mañana; el domingo no trabaja.
- **Descanso reglamentario:** 35 minutos por turno, dentro del turno y de duración fija (es parada programada). Horarios habituales en RES: mañana 11:55–12:30, tarde 18:55–19:30, noche 02:55–03:30. Descanso de SOP: 35 minutos por turno; el horario en que se toma **no está definido**.
- Idea conversada: que el libro descuente automáticamente los 35 minutos de descanso en cada turno.
- **Tiempo planificado (TPL):** TD (tiempo total del turno, Tiempo Disponible) − PP (paradas programadas o planificadas) − PE (paradas externas, no relacionadas con la línea a la cual se le quiere medir el OEE).
- Un turno en que la línea no trabaja (por ejemplo, feriado sin producción) no se mide: se omite.
- Fecha a la que pertenece un turno que cruza la medianoche: la fecha en que empieza el turno.
- Las paradas y la producción de la sopladora se asignan al turno SOP.
- **Estado:** Aprobado (horario del descanso de SOP: no definido).

### 4.3 Clasificación de paradas

Toda parada se registra (ver 4.4) y después se clasifica. Quien carga el dato elige el tipo, con ayuda del formulario. Cada tipo descuenta del tiempo planificado (PP o PE) o cuenta como pérdida (PNP):

- **Descanso reglamentario / almuerzo:** PP. Descuenta de TPL, no afecta el OEE.
- **Limpieza y sanitizado:** PP. Descuenta de TPL, no afecta el OEE, siempre que se programe antes del turno. La limpieza de entrega de turno (cerca de 10 minutos, cubierta por el líder) se intenta que no cuente como parada.
- **Cambio de formato o producto:** PP. Descuenta de TPL, no afecta el OEE.
- **Mantenimiento programado:** PP. Descuenta de TPL, no afecta el OEE.
- **Reunión o capacitación:** PP. Descuenta de TPL, no afecta el OEE.
- **Falta de insumos o materia prima:** PE. Descuenta de TPL, no afecta el OEE.
- **Falla de un servicio de apoyo** (refrigeración de agua, tratamiento de agua, suministro o falta de agua, falla del UV de la sala de tratamiento): PE. Descuenta de TPL, no afecta el OEE. No se detalla el componente que falló (por ejemplo, el compresor); solo el servicio.
- **Averías y/o fallas (A/F), de más de 5 minutos:** PNP. Resta del tiempo operativo, **baja la Disponibilidad, afecta el OEE**.
- **Ajustes y configuración (A y C), de más de 5 minutos:** PNP. Lo resuelve el mismo operario. Resta del tiempo operativo, **baja la Disponibilidad, afecta el OEE**.
- **Pequeña parada (5 minutos o menos):** PNP. No resta del tiempo operativo; **afecta solo el Rendimiento**. Se clasifica automáticamente por duración.
- **Reducción de velocidad** (no es una parada): afecta solo el Rendimiento. Se registra en un formulario aparte (ver 4.9). Cuando se acaba el envase soplado almacenado y la línea sigue al ritmo de la sopladora (por ejemplo, de 90 a 30 envases/min), se registra como reducción de velocidad con origen sopladora.
- **Parada sin tipo elegido:** se trata como PNP provisoria hasta que se la clasifique.
- **Estado:** Aprobado.

### 4.4 Umbral de registro de paradas

- Se registran **todas** las paradas, sin importar su duración.
- **5 minutos o menos:** pequeña parada (microparada). La clasificación es automática, por duración.
- **Más de 5 minutos:** avería/falla o ajuste y configuración (o parada programada o externa, según el caso). Lo elige quien carga.
- **Estado:** Aprobado (Q-07 cerrada con el criterio de 5 minutos o menos).

### 4.5 Velocidad nominal / tiempo de ciclo ideal

- **Valor por producto o formato:** valor por producto. No cambia según el equipo.
- **Unidad:** botellas por minuto (bidones: bidones por minuto). Ejemplo: 500 sin gas = 90 botellas/min.
- **Origen del dato:** lo define el ingeniero. Los valores de cada producto están cargados en 5.3.1.
- **Quién lo actualiza y cuándo:** el operario o el ingeniero.
- **Estado:** Aprobado (origen del dato: ingeniero).

### 4.6 Productos y unidad de conteo

- **Productos y formatos:** botellas (500 sin gas, 500 con gas, 1500 sin gas, 2000 sin gas) y bidones de 6 L.
- **Unidad de producción:** botella (en bidones, el bidón).
- **Punto de conteo:** fin de línea.
- **Conteo:** se registran tanto las unidades buenas como las rechazadas; las unidades totales son su suma, y ambas se agregan a los reportes.
- La venta es por fardo (pack), no por unidad. Las unidades por pack y los packs por pallet varían entre productos (valores en 5.3). Se registra también la cantidad de pallets.
- **Litros producidos:** se calculan automáticamente (unidades × volumen) y se guardan en una columna. No es un indicador principal del tablero, pero se suman por período.
- **Estado:** Aprobado.

### 4.7 Rechazos

- **¿Qué se considera rechazo?** Unidad que pasó por al menos una de las 4 máquinas de la línea y no cumple los requisitos de calidad. Los insumos perdidos antes de procesar son **pérdida de insumo** y no son rechazo (no afectan el OEE). Ver glosario 3.4.
- El retrabajo no cuenta como rechazo: si la unidad se recupera (por ejemplo, una botella con nivel bajo que se destapa y se vuelve a procesar), no se contabiliza.
- Se registra **por sección** (sopladora, llenadora, etiquetadora y forradora): cantidad de unidades y motivo.
- Unidad de conteo: **botellas** (Q-12 cerrada, también para la forradora).
- Un rechazo detectado tarde (por ejemplo, con el pallet armado) se imputa a los turnos que trabajaron ese lote. Cómo se reparte entre turnos si el lote abarca varios: **Q-06, Esperando definición**.
- El líder de turno debe conocer el criterio de qué es rechazo y qué no (capacitación).
- Los tipos y motivos de rechazo los clasifican el ingeniero, los supervisores y Joaquín.
- **Estado:** Aprobado (Q-06: Esperando definición).

### 4.8 Cambios de producto dentro de un turno

- Se realizan cambios de producto dentro de un turno. Al cambiar de producto se crea un lote nuevo, con sus nuevos reportes correspondientes.
- Si no cambia el producto, el lote cambia igual a las 00:00. Por eso un turno que cruza la medianoche (noche) puede tener dos lotes aunque produzca un solo producto.
- El lote se identifica con un número de registro (ver glosario 3.4). El lote tiene su propia fecha y hora de inicio.
- **Q-05 cerrada:** sí se registra por separado la hora de inicio y fin de cada lote dentro de un turno.
- El OEE se calcula por turno. El lote sirve para la trazabilidad (los reclamos se identifican por lote) y porque cada producto tiene su velocidad nominal.
- **Modelo de datos:** muchos turnos pertenecen a un lote. Para cambios de producto dentro de un turno, se usa una entidad `id_turno_lote` (segmento) con una fila por combinación turno–lote. Paradas, producción y rechazos cuelgan del segmento. OEE por turno: agrupar por `id_turno`. OEE por lote: agrupar por `id_lote`.
- **Estado:** Aprobado (modelo turno–lote: Propuesto; Q-06: Esperando definición).

### 4.9 Carga de datos

- **Quién carga los datos y cuándo:** el supervisor (líder de turno), durante y al final del turno.
- **Paradas:** hora de inicio, hora de fin, duración (automática), origen, motivo, tipo de parada, código (se genera solo) y observaciones.
  - Origen: dónde se origina la parada; casi siempre una máquina, pero puede ser otra cosa (por ejemplo, el tratamiento de agua).
  - Tipo: "pequeña parada" se asigna solo si dura **5 minutos o menos**; en el resto lo elige quien carga, con una ayuda en el formulario sobre cuál corresponde.
  - Observaciones: texto libre con el detalle (por ejemplo, "cambio de acople").
- **Reducción de velocidad:** formulario aparte, que se abre con un botón: origen, motivo, hora de inicio, hora de fin y velocidad real. El libro calcula el tiempo perdido por trabajar a menor velocidad (fórmula: Q-08, Esperando definición).
- **Origen de la producción:** anotación manual; el supervisor cuenta por pack y desde eso se puede saber la cantidad de botellas o bidones. Se registra también la cantidad de pallets.
- **Rechazos:** por sección, con cantidad y motivo (ver 4.7).
- **Pérdidas de insumo:** las registra el supervisor. Algunas en unidades (preformas, tapas) y otras en peso (por ejemplo, termocontraíble). Idea del ingeniero: usar una balanza y convertir el peso a unidades.
- **Personal:** quién ocupa cada puesto en cada turno, para trazabilidad (ver 5.6).
- **Valores nuevos:** todos los campos de los formularios aceptan valores personalizados, que se agregan a la lista automáticamente, sin revisión. Cualquiera puede agregarlos.
- **Calidad del dato:** hoy muchas paradas no se anotan. Se espera que los primeros valores salgan bajos; el ingeniero investigará qué falta. Para los operarios la carga es una tarea extra, así que el registro tiene que ser lo más simple posible.
- **Estado:** Aprobado.

### 4.10 Reportes y metas

- **Reportes requeridos (turno, día, semana, otros):** por turno. Día y semana: Esperando definición.
- **Tablero de control:** el OEE y los tres factores se muestran con indicadores de semáforo según el valor del OEE (por ejemplo, de 0 a 50% malo). El ingeniero pasará la tabla de rangos y un mockup visual del tablero.
- **Detalle de pérdidas que quiere ver:** ranking de motivos de parada por máquina (en horas); acumulado de pequeñas paradas (cantidad y minutos) con su ranking de motivos; reducciones de velocidad con sus motivos; resumen de paradas externas por tipo.
- **Litros producidos:** suma por período (filtrando por fecha), a partir de la columna de litros. No es un indicador principal.
- **Meta de OEE y de cada factor:** obtener un OEE correcto y detectar cuellos de botella. Valor actual estimado por el ingeniero: 70% (lo dijo como "rendimiento"; confirmar si es el factor Rendimiento o el OEE). El objetivo es mejorar, no buscar culpables.
- **Preguntas:** Semana: ¿de lunes a domingo? ¿Quién los ve y en qué formato (pantalla, impreso)? ¿Reportes por producto además de por turno? ¿Se muestra la pérdida de insumo en el tablero?
- **Estado:** Aprobado (tabla del semáforo, mockup y preguntas: Esperando definición).

### 4.11 Parámetros modificables

Los parámetros se guardan en la hoja `01_Parámetros` del libro maestro (`tblParametros`), con vigencia y motivo. No son constantes en el código. Un cambio de parámetro no altera el OEE de fechas anteriores: se cierra la fila vieja con `vigente_hasta` y se agrega una fila nueva con `vigente_desde`.

| ID | Nombre | Parámetro | Ámbito | Valor por defecto | Unidad | Quién lo cambia | Aprobación | Estado |
|---|---|---|---|---|---|---|---|---|
| P-01 | Umbral de microparada | `UMBRAL_MICROPARADA_MIN` | Global | 5 | min | Ingeniero | Ingeniero + reg. 13 | Aprobado |
| P-02 | Descanso reglamentario | `DESCANSO_MIN` | Global | 35 | min | Ingeniero | Ingeniero + reg. 13 | Aprobado |
| P-03 | Velocidad nominal 500 sin gas | `VEL_NOM_500_SG` | Producto | 90 | botellas/min | Operario/ingeniero; permanente solo ingeniero | Ingeniero | Aprobado (vigencia supuesta) |
| P-04 | Velocidad nominal 500 con gas | `VEL_NOM_500_CG` | Producto | 70 | botellas/min | Operario/ingeniero; permanente solo ingeniero | Ingeniero | Aprobado (vigencia supuesta) |
| P-05 | Velocidad nominal 1500 sin gas | `VEL_NOM_1500_SG` | Producto | 70 | botellas/min | Operario/ingeniero; permanente solo ingeniero | Ingeniero | Aprobado (vigencia supuesta) |
| P-06 | Velocidad nominal 2000 sin gas | `VEL_NOM_2000_SG` | Producto | 70 | botellas/min | Operario/ingeniero; permanente solo ingeniero | Ingeniero | Aprobado (vigencia supuesta) |
| P-07 | Velocidad nominal bidón 6 L | `VEL_NOM_BIDON_6L` | Producto | 10 | bidones/min | Operario/ingeniero; permanente solo ingeniero | Ingeniero | Aprobado (vigencia supuesta) |
| P-08 | Tolerancia de comparación | `TOLERANCIA_COMPARACION_PP` | Global | 0,05 | puntos porcentuales | Ingeniero | Ingeniero + reg. 13 | Aprobado |

Notas:
- P-01: **5 minutos o menos** = microparada; **más de 5 minutos** = A/F o A y C. Actualiza RN-03 y RN-04.
- P-02: aplica a RES y SOP. El horario del descanso SOP no está definido, pero se descuentan 35 min.
- P-03 a P-07: vigencia hacia atrás supuesta (2026-01-01) para poder calcular el caso 498. Confirmar antes de congelar F1.
- P-08: solo aceptación/control. Sale de 1.4.
- **Decisión D-01 (no es parámetro):** limpiezas de más de 5 min se cargan como PNP; subtipo A y C provisional si las resuelve el operario; si son 5 min o menos, micro automática. En el caso 498 el OEE sigue 61,4%; cambia el reparto Disponibilidad (81,5%) y Rendimiento (75,3%).
- **Estado:** Propuesto.

## 5. Catálogos

Los catálogos son las listas de opciones de los formularios. Se completan a partir de las definiciones de la sección 4.

**Reglas para todos los catálogos**

- **Código estable y único.** Nunca se reutiliza ni se cambia.
- **No se borran filas: se marcan como inactivas.** Los registros viejos siguen apuntando a ese código.
- **Columna "activo" (sí/no)** en todos.
- **Valores nuevos:** los campos de los formularios aceptan valores personalizados que se agregan a la lista automáticamente, sin revisión previa (ver 4.9).
- Los valores operativos viven en el libro maestro (`02_Cat_*`); el Doc fija códigos, nombres y reglas.

### 5.1 Causas de parada

Cada parada se registra con tres datos de catálogo: tipo, origen y motivo. El tipo lo elige quien carga la parada (no depende del motivo: el mismo motivo puede ser un ajuste o una avería según quién lo resuelva).

**Tipos de parada**

| Código | Tipo | Clasificación | Efecto en el cálculo | Seleccionable | Estado |
|---|---|---|---|---|---|
| TP-01 | Parada programada | PP | Descuenta TPL; no afecta OEE | Sí | Aprobado |
| TP-02 | Parada no planificada | PNP | Cabecera de clasificación; por sí sola no calcula efecto | No (cabecera) | Propuesto |
| TP-03 | Avería o falla | PNP / A/F | Resta TO; baja Disponibilidad | Sí (solo >5 min) | Aprobado |
| TP-04 | Ajuste y configuración | PNP / A y C | Resta TO; baja Disponibilidad | Sí (solo >5 min) | Aprobado |
| TP-05 | Pequeña parada / microparada | PNP / micro | No resta TO; afecta Rendimiento | Sí (automática ≤5 min) | Aprobado |
| TP-06 | Sin clasificar | PNP provisoria | Resta TO hasta clasificar | Sí (>5 min) | Aprobado |
| TP-07 | Parada externa | PE | Descuenta TPL; no afecta OEE | Sí | Aprobado |

**Orígenes.** Columnas: código, nombre, clase (máquina de la línea / servicio de apoyo / otro), activo.

- Máquinas de la línea: sopladora, llenadora, etiquetadora, forradora.
- Servicios de apoyo: refrigeración de agua, tratamiento de agua, suministro de agua. Se pueden seguir agregando.
- Otros: descanso/personal, cambio de formato/producto, mantenimiento programado, reunión/capacitación, limpieza/sanitizado programado, otro/no clasificado.
- Las averías se separan solo por máquina (no por tipo eléctrico o mecánico).

**Motivos.** Columnas: código, motivo, origen, tipo sugerido, activo. Lista abierta que se va armando con el uso. El "tipo sugerido" es ayuda, no obligación.

- Ejemplos: volcador de preforma (sopladora); atasco a la salida de la etiquetadora; limpieza de adhesivo (etiquetadora); falla de refrigeración por alta temperatura (refrigeración de agua); falla en horno (sopladora).
- **Observaciones:** texto libre, no es catálogo.

**Estado:** Aprobado (lista concreta de motivos: en construcción; ver `02_Cat_Paradas`).

### 5.2 Turnos

- **Columnas:** código, nombre, equipo (SOP / RES), hora de inicio, hora de fin, hora de inicio y fin del descanso, días en que aplica, activo, vigencia desde / hasta.
- **Calculadas** (no se cargan): cruza medianoche (sí/no), duración TD en minutos, descanso en minutos (35).

**RES:**

| Código | Nombre | Inicio | Fin | Descanso | Días |
|---|---|---|---|---|---|
| RES-M | Mañana RES | 06:00 | 14:00 | 11:55–12:30 | Lun–vie; sáb solo mañana |
| RES-T | Tarde RES | 14:00 | 22:00 | 18:55–19:30 | Lun–vie |
| RES-N | Noche RES | 22:00 | 06:00 | 02:55–03:30 | Lun–vie |

**SOP:** 4 turnos de 6 horas, 7 días: 00:00–06:00, 06:00–12:00, 12:00–18:00 y 18:00–00:00. Descanso de 35 minutos por turno; el horario en que se toma no está definido.

- Fecha a la que pertenece un turno que cruza la medianoche: la fecha en que empieza el turno.
- **Estado:** Aprobado (horario del descanso de SOP: no definido).

### 5.3 Productos

- **Columnas:** código, producto, formato (botella / bidón), volumen (litros), unidades por pack, packs por pallet, activo.

| Producto | Volumen | Unidades/pack | Packs/pallet |
|---|---|---|---|
| 500 sin gas | 0,5 L | 12 | 150 |
| 500 con gas | 0,5 L | 12 | 132 |
| 1500 sin gas | 1,5 L | 6 | 100 |
| 2000 sin gas | 2 L | 6 | 80 |
| Bidón | 6 L | 1 (no se hace pack) | 168 |

**5.3.1 Velocidades nominales.** Columnas: producto, velocidad (botellas/min), vigencia desde, origen del dato.

| Producto | Velocidad |
|---|---|
| 500 sin gas | 90 botellas/min |
| 500 con gas | 70 botellas/min |
| 1500 sin gas | 70 botellas/min |
| 2000 sin gas | 70 botellas/min |
| Bidón | 10 bidones/min |

- **Estado:** Aprobado (origen del dato: ingeniero; vigencia inicial supuesta).

### 5.4 Rechazos (motivos por sección)

Los rechazos se registran por sección, con cantidad y motivo (ver 4.7). No hay un catálogo detallado de tipos.

- **Secciones:** sopladora, llenadora, etiquetadora, forradora.
- **Motivos.** Columnas: motivo, sección, activo. Lista abierta.
  - Ejemplos: botella mal formada o mal tapada (llenadora); botella aplastada por trabado de máquina (llenadora); etiqueta mal puesta que no se puede recuperar o pico sucio de pegamento (etiquetadora); rotura de envase (sopladora), por ejemplo por falla de válvula o por prueba de parámetros con un insumo nuevo.
- Una botella con nivel bajo se destapa y se vuelve a procesar: es retrabajo, no rechazo.
- **Estado:** Aprobado.

### 5.5 Pérdida de insumos

- **Columnas:** insumo, unidad de medida (unidades o peso), motivo, activo.
- **Insumos principales:** preformas, tapas, etiquetas, termocontraíble (el más caro, tiene que estar) y film stretch (casi nunca se pierde; opcional). Hay más insumos; se pueden agregar.
- **Unidades:** preformas y tapas en unidades; termocontraíble, etiquetas y film stretch en peso (gramos o kilos). A confirmar por insumo.
- **Motivos (ejemplos):** preformas quemadas por una parada de la sopladora (alrededor de 60 preformas dentro del horno ya no se pueden soplar); preformas caídas al piso.
- **Quién registra:** el supervisor. Es solo seguimiento: no entra en el OEE.
- **Estado:** Aprobado.

### 5.6 Puestos de trabajo

- **Columnas:** puesto, equipo (SOP / RES), activo.
- **SOP:** soplador y ayudante.
- **RES:** líder, llenador, visor, etiquetador, forradora (final de línea), estibador y autoelevadorista (confirmar nombre del séptimo puesto).
- Se registra qué persona ocupó cada puesto en cada turno, para trazabilidad. En bidones la asignación cambia porque el proceso es manual. Los nombres del personal: Esperando definición.
- **Estado:** Aprobado.

## 6. Reglas de negocio

Cada regla tiene un identificador, un estado y una fecha de aprobación. Todas las reglas (RN-01 a RN-09) quedaron aprobadas el 2026-10-28 (aprobado por: ingeniero [confirmar nombre]).

- **RN-01:** los OEE de períodos más largos (día, semana) se calculan sumando tiempos y unidades de los turnos, nunca promediando los OEE de cada turno. Estado: Aprobado (2026-10-28).
- **RN-02:** unidades buenas = unidades totales − unidades rechazadas. Estado: Aprobado (2026-10-28).
- **RN-03:** las pequeñas paradas (**5 minutos o menos**) no se restan del tiempo operativo; solo afectan el Rendimiento. Estado: Aprobado (2026-10-28; actualizado 2026-10-29). Propuesto el 2026-10-21.
- **RN-04:** todas las paradas se registran; las de **5 minutos o menos** se clasifican automáticamente como pequeña parada. Estado: Aprobado (2026-10-28; actualizado 2026-10-29). Propuesto el 2026-10-21.
- **RN-05:** ni el OEE ni sus factores pueden superar 100%; si pasa, se revisan las variables (por ejemplo, la velocidad nominal) y los datos. Estado: Aprobado (2026-10-28).
- **RN-06:** el retrabajo no cuenta como rechazo; la pérdida de insumo no es rechazo y no entra en el OEE. Estado: Aprobado (2026-10-28).
- **RN-07:** un cambio de producto crea un lote nuevo; si no cambia el producto, el lote cambia a las 00:00. Estado: Aprobado (2026-10-28).
- **RN-08:** los valores nuevos cargados en los formularios se agregan a las listas sin revisión. Estado: Aprobado (2026-10-28).
- **RN-09:** una parada sin tipo elegido cuenta como PNP provisoria hasta que se la clasifique. Estado: Aprobado (2026-10-28).

## 7. Caso de prueba: turno real calculado a mano

Este caso es la referencia para validar el cálculo del libro.

**Caso cargado (2026-10-27):** Reporte de Producción Nº 498, fecha 28/09/26 (foto). Estado: Propuesto. Actualizado el 2026-10-28 con las respuestas del ingeniero (Q-09, Q-10, Q-18). Falta que el ingeniero lo verifique contra el papel.

### 7.1 Datos de entrada

- **Fuente:** Reporte de Producción Nº 498, fecha 28/09/26. Lote 09261.
- **Fecha y turno:** 2026-09-28, turno tarde (RES), 14:00–22:00. Producto: 1500 cc sin gas.
- **Tiempo disponible (TD):** 480 min.
- **Tiempo planificado de producción (TPL):** 405 min (= 480 − 35 de PP − 40 de PE).
- **Paradas** (inicio–fin · minutos · origen · motivo · tipo asignado). El papel no trae el tipo; las limpiezas de la etiquetadora y la baja tensión las clasificó el ingeniero (Q-18); los demás tipos son una interpretación a confirmar.

| Inicio–fin | Min | Origen | Motivo | Tipo |
|---|---|---|---|---|
| 14:10–14:50 | 40 | Compresor Sabroe | temperatura alta | PE (refrigeración) |
| 15:40–15:55 | 15 | Etiquetadora | limpieza | Micro (Q-18) |
| 16:50–17:00 | 10 | Etiquetadora | cambio de rollo y limpieza | Micro (Q-18) |
| 18:10–18:30 | 20 | Etiquetadora | limpieza | Micro (Q-18) |
| 18:55–19:30 | 35 | descanso | — | PP |
| 20:40–21:10 | 30 | baja tensión | falla de compresor | Reducción de velocidad (Q-18) |
| 21:30–22:00 | 30 | horno (sopladora) | falla en horno | A/F |

- **Nota del papel sin hora** (no entra al cálculo): "pequeñas paradas por falta de temperatura y problemas en el loose de pack".
- **Velocidad nominal:** 70 botellas/min (definida por el ingeniero, Q-18).
- **Unidades producidas OK:** 29 pallets. Con 100 packs/pallet y 6 botellas/pack: **17.400 botellas**.
- **Unidades rechazadas:** 4. El papel no indica sección ni motivo.
- **Pérdida de insumo** (no entra en el OEE): envases 2 u, tapas 52 u.

### 7.2 Cálculo manual

- **Tiempo operativo (TO):** 405 − 30 = **375 min**. Las limpiezas de la etiquetadora (microparadas) y la baja tensión (reducción de velocidad) no restan.
- **Disponibilidad:** 375 / 405 = 0,9259 → **92,6%**.
- **Unidades totales (UT):** 17.400 + 4 = 17.404.
- **Rendimiento:** (17.404 × 1/70) / 375 = 248,63 / 375 = 0,6630 → **66,3%**.
- **Calidad:** 17.400 / 17.404 = 0,99977 → **99,98%**.
- **OEE esperado:** 0,9259 × 0,6630 × 0,99977 = 0,6138 → **61,4%**.
- **Verificación con la forma simplificada:** (17.400 × 1/70) / 405 = 248,57 / 405 = 0,6138 ✓.
- **Control:** el Rendimiento no supera 100% solo si la velocidad nominal es de al menos 46,41 botellas/min (17.404 / 375). Con una velocidad menor se dispara RN-05.
- **Nota (actualizada 2026-10-29):** este cálculo sigue la clasificación del ingeniero (Q-18): las limpiezas de 10 a 20 minutos se tratan como microparada. Con la regla automática (≤5 min micro), si se trataran como A y C (PNP >5 min), el OEE sería el mismo (**61,4%**); solo cambiaría el reparto entre Disponibilidad (**81,5%**) y Rendimiento (**75,3%**). Decisión D-01: se cargan como PNP, A y C provisional. Pendiente de aclarar con el ingeniero.

### 7.3 Resultado del libro y diferencia

- Datos del libro: la baja tensión de 20:40–21:10 es una reducción de velocidad (Q-18) y no está en `tblParadas`; TO = 405 − 45 de A y C − 30 de A/F = **330 min**.
- OEE calculado por el libro: **61,38%** (Disponibilidad **81,48%**, Rendimiento **75,34%**, Calidad **99,98%**).
- Contra **D-01** (limpiezas de 15/10/20 min como A y C): TO 330 min, Disponibilidad 81,5%, Rendimiento 75,3%, Calidad 99,98%, OEE 61,4%. Diferencias: D −0,02 pp, R +0,04 pp, C 0,00 pp, OEE −0,02 pp; todas dentro de la tolerancia P-08.
- Contra **Q-18** (limpiezas como microparada): TO 375 min, Disponibilidad 92,6%, Rendimiento 66,3%, Calidad 99,98%, OEE 61,4%. Diferencias: D −11,12 pp, R +9,04 pp, C 0,00 pp, OEE −0,02 pp. La diferencia en D y R **no es un error de cálculo**: es el efecto de la decisión D-01, todavía pendiente, sobre cómo se clasifican las limpiezas. El OEE coincide en ambos criterios.
- Comparación automática: `90_Tests`, macro `ReporteCaso498`, con las columnas manuales de Q-18 y D-01 y la tolerancia P-08.
- Aprobado por / fecha: [ ]

### 7.4 Casos adicionales

Los cuatro casos usan **datos inventados**. Se calculan a mano con las fórmulas de 3.1 y las reglas de la sección 6. Estado: **Propuesto**; el ingeniero aprueba los resultados esperados. Tolerancia: la de 1.4.

**Caso 1 · Turno sin producción**

- **Datos:** 2026-12-25 (feriado), turno mañana RES 06:00–14:00. La línea no trabaja; no hay paradas ni producción cargadas.
- **Esperado:** el turno no se mide, se omite (Q-16). No se calcula OEE ni factores y no entra en reportes, sumas ni promedios. No debe aparecer como 0%.
- **Variante 1b (propuesta):** la línea estaba programada pero no produjo por una avería de todo el turno. TD 480, PP 35, PE 0, A/F 445 → TPL 445, TO 0, 0 unidades. Disponibilidad 0%; Rendimiento y Calidad sin dato (no dividir por cero); OEE 0%.

**Caso 2 · Turno con cambio de producto**

- **Datos:** 2026-10-13, turno mañana RES 06:00–14:00. Lote A: 500 sin gas (90/min). Cambio de formato de 30 min (PP). Lote B: 1500 sin gas (70/min).
- **Paradas:** descanso 35 (PP); cambio de formato 30 (PP); avería de la llenadora 25 (A/F); ajuste de la etiquetadora 8 (A y C); 4 microparadas (2, 2, 2 y 4 min = 10 min, no restan).
- **Producción:** lote A: 1.425 packs × 12 = 17.100 buenas y 40 rechazadas. Lote B: 1.600 packs × 6 = 9.600 buenas y 12 rechazadas.
- **Cálculo:** TD 480 − PP 65 − PE 0 = TPL 415. TO = 415 − 25 − 8 = 382. Disponibilidad = 382 / 415 = 0,9205 → 92,0%. Tiempo ideal = 17.140 / 90 + 9.612 / 70 = 190,44 + 137,31 = 327,76 min. Rendimiento = 327,76 / 382 = 0,8580 → 85,8%. Calidad = 26.700 / 26.752 = 0,99806 → 99,81%.
- **OEE esperado:** 0,9205 × 0,8580 × 0,99806 = 0,7882 → **78,8%**. Forma simplificada: (17.100 / 90 + 9.600 / 70) / 415 = 327,14 / 415 = 0,7883 → 78,8% ✓. La diferencia de 0,0001 sale de ponderar la Calidad por unidades y no por tiempo de ciclo; a 1 decimal coinciden.
- **Qué verifica:** el Rendimiento suma el tiempo ideal de cada lote con la velocidad de su producto. El horario de cada lote no se usa en el cálculo.

**Caso 3 · Turno con paradas que cruzan medianoche**

- **Datos:** turno noche RES que empieza el lunes 2026-10-12 a las 22:00 y termina el martes a las 06:00: pertenece al 2026-10-12 (Q-04). Producto: 1500 sin gas (70/min). Dos lotes (RN-07): A hasta las 00:00 y B desde las 00:00.
- **Paradas:** avería de la llenadora 23:40–00:25 = 45 min (A/F); falta de agua 01:15–01:50 = 35 (PE); descanso 02:55–03:30 = 35 (PP); ajuste de la etiquetadora 04:10–04:22 = 12 (A y C); 2 microparadas (2 y 4 min, no restan).
- **Producción:** lote A: 700 packs × 6 = 4.200 buenas y 6 rechazadas. Lote B: 2.900 packs × 6 = 17.400 buenas y 24 rechazadas. Total: 21.600 buenas y 30 rechazadas.
- **Cálculo:** TD 480 − PP 35 − PE 35 = TPL 410. TO = 410 − 45 − 12 = 353. Disponibilidad = 353 / 410 = 0,8610 → 86,1%. Rendimiento = (21.630 / 70) / 353 = 309,0 / 353 = 0,8754 → 87,5%. Calidad = 21.600 / 21.630 = 0,99861 → 99,86%.
- **OEE esperado:** 0,8610 × 0,8754 × 0,99861 = 0,7526 → **75,3%**. Forma simplificada: (21.600 / 70) / 410 = 308,57 / 410 = 0,7526 ✓.
- **Qué verifica:** la parada de 23:40 a 00:25 dura 45 min y queda en el turno del 2026-10-12 aunque termine el 13; el descanso de las 02:55 también pertenece a ese turno; los dos lotes suman en un solo OEE de turno.

**Caso 4 · Turno con datos incompletos**

- **Datos:** 2026-10-14, turno tarde RES 14:00–22:00. Producto: 2000 sin gas (70/min).
- **Paradas:** descanso 18:55–19:30 = 35 (PP); avería de la forradora 15:10–15:30 = 20 (A/F); parada de la llenadora 16:20–16:50 = 30 **sin tipo ni motivo** (RN-09: PNP provisoria). Nota sin hora: "pequeñas paradas por falta de temperatura".
- **Producción:** 46 pallets × 80 packs × 6 = 22.080 buenas. Rechazadas: 8, sin sección ni motivo.
- **Cálculo:** TD 480 − PP 35 − PE 0 = TPL 445. TO = 445 − 20 − 30 = 395. Disponibilidad = 395 / 445 = 0,8876 → 88,8%. Rendimiento = (22.088 / 70) / 395 = 315,54 / 395 = 0,7988 → 79,9%. Calidad = 22.080 / 22.088 = 0,99964 → 99,96%.
- **OEE esperado:** 0,8876 × 0,7988 × 0,99964 = 0,7088 → **70,9%**. Forma simplificada: (22.080 / 70) / 445 = 315,43 / 445 = 0,7088 ✓. El turno queda marcado con datos provisorios; los 8 rechazos cuentan en la Calidad con sección y motivo vacíos; la nota sin hora no cambia nada.
- **Variante 4b:** si después la parada de 30 min se clasifica como PE (por ejemplo, falta de agua): TPL = 415, TO = 395, Disponibilidad 95,2%, Rendimiento 79,9%, OEE = 315,43 / 415 = 0,7601 → **76,0%**.

## 8. Estructura de datos prevista

### 8.1 Hojas y tablas del libro maestro

El libro maestro `oee_vX.Y_AAAA-MM-DD.xlsm` tiene una Excel Table por catálogo y por conjunto de datos. El VBA usa el **nombre de tabla**, no el de la hoja.

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
| `11_Turnos` | `tblTurnosRegistro` | `id_turno`, `id_turno_lote`, `id_lote`, fecha, turno_codigo, equipo, inicio/fin |
| `12_Paradas` | `tblParadas` | Códigos, no texto |
| `13_Produccion` | `tblProduccion` | Unidades, pallets, litros |
| `14_Rechazos` | `tblRechazos` | Por sección y motivo |
| `15_Reducciones` | `tblReducciones` | Formulario aparte |
| `16_PerdidasInsumo` | `tblPerdidasInsumo` | Seguimiento, no OEE |
| `17_PersonalTurno` | `tblPersonalTurno` | Trazabilidad |
| `20_Calculos` | — | Fórmulas de cálculo |
| `21_Dashboard` | — | Tablero |
| `90_Tests` | — | Casos de prueba |

### 8.2 Estructura de las tablas de datos

- **`tblParadas`:** ID, fecha, turno, línea, `id_lote`, `id_turno_lote`, hora de inicio y fin (como fecha y hora completas), duración (en minutos), origen, motivo, tipo de parada (se asigna solo si dura 5 minutos o menos), código y observaciones. Se guardan los códigos del catálogo, no el texto.
- **`tblProduccion`:** fecha, turno, línea, lote, `id_turno_lote`, producto, unidades producidas (unidad base, botellas o bidones; se calculan desde los packs contados), pallets, litros (calculado y guardado) y velocidad nominal (se trae del catálogo con vigencia, 5.3).
- **`tblRechazos`:** una fila por sección y motivo: fecha, turno, lote, `id_turno_lote`, sección, motivo, cantidad.
- **`tblReducciones`:** fecha, turno, `id_turno_lote`, origen, motivo, hora de inicio y fin, velocidad real.
- **`tblPerdidasInsumo`:** fecha, turno, insumo, cantidad, unidad, motivo.
- **`tblPersonalTurno`:** fecha, turno, puesto, persona.
- **`tblParametros`:** id, nombre, parámetro, ámbito, producto_codigo, valor, unidad, vigente_desde, vigente_hasta, motivo, aprobado_por, estado, fecha_aprobacion.

### 8.3 Principios

- Una fila por evento, fechas y horas reales, IDs únicos, `linea_id`, sin valores calculados en hojas de datos.
- Formularios de carga con validaciones.
- Dashboard: OEE y sus tres factores, tendencia y desglose de pérdidas.
- **Estado:** Esperando definición (nombres finales de hoja y tabla).

## 9. Reportes y dashboard

- **Reporte por turno:** OEE, los tres factores y el detalle de pérdidas (ver 4.10).
- **Reporte por día:** Esperando definición.
- **Reporte por semana:** Esperando definición.
- **Período:** se elige con un selector de fecha (también se usa para sumar litros).
- **Filtros:** línea, producto, turno.
- **Contenido pedido por el ingeniero:** OEE y los tres factores con semáforo; ranking de motivos de parada por máquina (en horas); acumulado de pequeñas paradas (cantidad y minutos) con sus motivos; reducciones de velocidad con sus motivos; litros producidos por período.
- **Indicadores a mostrar:** a elección del ingeniero (lista de abajo).
- **Gráfico:** gráfico de barras.
- **Para quién y dónde:** pantalla en Excel, impreso o PDF.
- **Estado:** Esperando definición (tabla del semáforo, mockup, preguntas).

**Indicadores para proponerle (a elección del ingeniero)**

- OEE, Disponibilidad, Rendimiento y Calidad.
- Tiempo planificado, tiempo operativo y paradas (en minutos).
- Unidades producidas, buenas y rechazadas.
- Pérdidas por familia (Disponibilidad, Rendimiento, Calidad).
- Paradas por motivo y por máquina, ordenadas de mayor a menor (Pareto).
- Pequeñas paradas acumuladas (cantidad y minutos).
- Reducciones de velocidad y sus motivos.
- Rechazos por sección y motivo.
- Pérdida de insumo (solo seguimiento, no entra en el OEE).
- Litros producidos.
- OEE por producto.

## 10. Plan de entregas

**Etapa 1:** cálculo funcionando con datos cargados a mano. Terminado cuando coincide con el caso de prueba de la sección 7 dentro de la tolerancia aprobada (1.4) y pasan los casos de 7.4 (sin producción, datos incompletos, cruce de medianoche).

**Etapa 2:** formularios de carga con validaciones. Terminado cuando un operador carga un turno completo sin ayuda, las validaciones rechazan datos incorrectos (solapamientos, horas fuera del turno, causa inactiva) y el OEE coincide con el cálculo manual.

**Etapa 3:** dashboard con tendencia y desglose de pérdidas. Terminado cuando el layout fue aprobado en boceto y el ingeniero valida que los números coinciden con su criterio.

**Etapa 4:** prueba con datos reales durante una semana, correcciones, manual corto y aceptación final del ingeniero. Terminado cuando: aceptación con fecha en 13.

Fechas acordadas: [Esperando definición]. Se proponen después de cerrar las secciones 4 y 5.

## 11. Riesgos y supuestos

### Riesgos

- Definiciones incompletas.
- Datos desordenados o inconsistentes.
- Cambios de requisitos durante el desarrollo.
- Registros sin horarios de paradas.
- Paradas que hoy no se anotan: los primeros valores pueden no reflejar la realidad.
- La carga de datos es una tarea extra para los operarios; si resulta pesada, se registra menos.
- La clasificación de las paradas depende del criterio de quien carga (se prevé capacitación y una ayuda en el formulario).

### Supuestos

- Se usa Excel de escritorio con VBA en la notebook del ingeniero.
- Los datos se cargan a mano en formularios; no hay conexión con sensores ni sistemas automáticos.
- Se mide una sola línea y la velocidad nominal es una por producto.
- Los líderes de turno registran todas las paradas, sin importar su duración.
- El nombre de la línea es "Línea 1 – Embotellado (supuesto)" hasta que el ingeniero lo confirme.

## 12. Preguntas abiertas

Por cada pregunta: número, pregunta, fecha, responsable, estado.

- **Q-01** · ¿Lo que se cuenta al final de la línea son unidades buenas, o incluye las rechazadas? · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: Se cuentan ambas y se agregan a los reportes.
- **Q-02** · Las paradas y la producción de la sopladora, ¿se asignan al turno SOP o al turno RES que está corriendo? · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: se asignan al turno SOP.
- **Q-03** · Horarios de los 4 turnos SOP y su descanso. · 2026-10-26 · Ingeniero · **parcial** (horarios cerrados el 2026-10-28; falta el horario del descanso). 00–06, 06–12, 12–18, 18–00. 35 min de descanso; no está definido cuándo se toma.
- **Q-04** · ¿A qué fecha pertenece un turno que cruza la medianoche? · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: a la fecha cuando empieza el turno.
- **Q-05** · ¿Se registra hora de inicio y fin de cada lote dentro del turno? · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: sí, se registra por separado la hora de inicio y fin de cada lote dentro de un turno.
- **Q-06** · Rechazo detectado tarde en un lote que abarca varios turnos: ¿cómo se reparte entre turnos? · 2026-10-26 · Ingeniero · **abierta**.
- **Q-07** · Una parada de exactamente 5 minutos, ¿es pequeña parada o avería/ajuste? · 2026-10-26 · Ingeniero · **cerrada (2026-10-29)**. R: es pequeña parada siempre que sea 5 minutos o menos. Más de 5 minutos: A/F o A y C.
- **Q-08** · Fórmula del tiempo perdido por reducción de velocidad. · 2026-10-26 · Ingeniero · **abierta**.
- **Q-09** · Velocidades nominales de todos los productos y su origen. · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: 500 SG = 90 botellas/min; bidones = 10 bidones/min; los demás = 70 botellas/min. Origen del dato: ingeniero.
- **Q-10** · Productos: confirmar unidades por pack, packs por pallet, volumen del bidón y si el de 2000 es con o sin gas. · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: 2000 sin gas; bidón 6 L; valores confirmados en 5.3.
- **Q-11** · Lote: ¿el número secuencial se reinicia cada mes o cada año? · 2026-10-26 · Ingeniero · **abierta**.
- **Q-12** · Rechazos de la forradora: ¿se cuentan en botellas o en packs? · 2026-10-26 · Ingeniero · **cerrada (2026-10-29)**. R: en botellas.
- **Q-13** · Reportes: ¿semana de lunes a domingo? ¿Quién los ve y en qué formato? ¿Por producto? ¿Se muestra la pérdida de insumo? · 2026-10-26 · Ingeniero · **abierta**.
- **Q-14** · El 70% que mencionó, ¿es el Rendimiento o el OEE? · 2026-10-26 · Ingeniero · **abierta**.
- **Q-15** · Pasar la tabla de rangos del semáforo y el mockup del tablero. · 2026-10-26 · Ingeniero · **abierta**.
- **Q-16** · Turno en que la línea no trabaja (feriado): ¿no se mide? · 2026-10-26 · Ingeniero · **cerrada (2026-10-28)**. R: no se mide, se omite.
- **Q-17** · Fotografiar un registro viejo para el caso de prueba. · 2026-10-26 · Joaquín · **cerrada (2026-10-27)**.
- **Q-18** · Caso de prueba (7.1): ¿cómo se clasifican las limpiezas de la etiquetadora y la baja tensión? ¿Cuál es la velocidad nominal de 1500 sin gas? · 2026-10-27 · Ingeniero · **cerrada (2026-10-28)**. R: limpiezas de etiquetadora = microparada; baja tensión = reducción de velocidad; 1500 sin gas = 70 botellas/min.

## 13. Registro de cambios

- **[2026-10-21]** · Creación del documento (v0.1) · Joaquín · aprobado el 2026-10-28.
- **2026-10-26** · v0.2: se cargan como Propuesto las definiciones de la reunión del 2026-10-21 y respuestas posteriores; se rehace la sección 5; se corrige la microparada; se unifican siglas · Joaquín · aprobado el 2026-10-28.
- **2026-10-27** · v0.3: las secciones 1, 3, 4 y 5 pasan a Aprobado; se carga el caso de prueba en 7.1 y 7.2 desde el Reporte Nº 498 · Joaquín · aprobación de 1, 3, 4 y 5: ingeniero [confirmar nombre].
- **2026-10-28** · v0.4: se incorporan respuestas del ingeniero (Q-01, Q-02, Q-04, Q-10, Q-16, Q-18); se cargan horarios SOP, fecha del turno nocturno, velocidades y productos; se actualiza el caso 7.1/7.2; se cargan 4 casos de 7.4; se aprueba tolerancia 1.4 y RN-01 a RN-09 · Joaquín · aprobación: ingeniero [confirmar nombre], 2026-10-28.
- **2026-10-29** · v0.5: microparada = **5 minutos o menos**; A/F y A y C = **más de 5 minutos**; se actualizan 3.2, 3.3, 3.4, 4.3, 4.4, 4.9, 5.1, RN-03, RN-04, Q-07 y nota 7.2; se agrega 4.11 con tabla de parámetros (8 parámetros, con columna Nombre; se eliminan P-09 a P-12); se corrige catálogo de tipos de parada (TP-01 PP, TP-02 PNP cabecera, TP-03 A/F, TP-04 A y C, TP-05 micro ≤5, TP-06 sin clasificar, TP-07 PE); se cierra Q-05, Q-07, Q-09 y Q-12; se agrega modelo turno–lote con `id_turno_lote`; se define nombre de línea supuesto; se agrega estructura de hojas y tablas del xlsm; se registra D-01 y D-02 · Joaquín · estado: Propuesto (falta aprobación del ingeniero).
- **2026-10-30** · Se corrige la carga del caso 498: la baja tensión sale de `tblParadas` porque es reducción de velocidad (Q-18). Se completa 7.3 con el resultado del libro y la comparación contra Q-18 y D-01; se agrega el caso punta a punta y la comparación automática en `90_Tests` (macro `ReporteCaso498`) y la función `DetalleTurno` en `Agregacion` · Joaquín · estado: Propuesto.

## 14. Control de versiones del archivo Excel

- Existe **un único archivo maestro**, con una persona responsable de modificarlo: [nombre].
- Cada versión lleva fecha en el nombre, por ejemplo `oee_v0.5_AAAA-MM-DD.xlsm`.
- Los demás prueban siempre sobre copias, nunca sobre el maestro.
- El código (módulos y formularios) se exporta a Git.
- Los catálogos y parámetros se exportan a CSV en `data/config/` con la macro `ExportarConfigACSV`, y se versionan.
- La plantilla vacía (sin datos reales) puede publicarse como Release.