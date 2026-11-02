# Roadmap Fase 2: datos y formularios (Etapa 2)

Proyecto OEE Excel/VBA · Autor: Joaquín · Versión 1.0 · 2026-11-01

Reglas de referencia: Doc v0.5 más las decisiones del 2026-11-01, que se formalizan en el Doc v0.6 (F2.0).

## 1. Alcance y reglas de juego

- **Objetivo:** que administración cargue en Excel los reportes de producción de cada turno de forma cómoda y segura, y que el OEE salga correcto a partir de esa carga.
- **Quién usa los formularios:** un empleado de administración que transcribe en la PC el papel que llena el supervisor durante el turno. La comodidad de esa persona es el criterio de diseño.
- **El papel de hoy es el origen de los datos, no el molde del formulario.** Se respeta su orden para que transcribir sea fácil, pero el libro puede registrar más datos que el papel.
- **Dato que el papel no trae:** se acepta vacío, no bloquea la carga y no entra al cálculo hasta que se informe. Se bloquea solo lo que es imposible o rompe el cálculo.
- **Fuente de verdad:** el Doc. Si algo no está ahí, no está definido.
- **Entrega en dos tramos:** 2A núcleo (lo que alimenta el OEE) y 2B complementarios (reducciones, insumos, personal, papeles nuevos).
- **Fuera de alcance:** dashboard (Fase 3), reporte de SOP, reportes por día y semana, base de datos y web.
- **Etapa 1:** cerrada (tag etapa1). Quedan sin bloquear: la demo H3 al ingeniero y la agregación día/semana (T11).
- **Tercera persona:** sigue fuera del plan hasta que se defina su rol.

## 2. Decisiones cerradas (pasan al Doc v0.6)

1. **Carga (secciones 2 y 4.9):** el supervisor llena el papel durante el turno; administración lo transcribe a Excel en una PC.
2. **Un reporte = un segmento turno–lote (4.8):** si cambia el producto dentro del turno, se llena otro reporte. El segmento (id\_turno\_lote) es la unidad de carga.
3. **Sopladora (ajusta Q-02 y 4.2):** hoy no existe reporte de SOP. El supervisor anota las paradas de sopladora en el reporte RES y afectan el OEE del turno RES cuando corresponde. **Pendiente** hasta que se confirme la creación del reporte de SOP.
4. **Velocidad nominal:** se usa la del catálogo (70/min en el 498; el papel no la informaba). El campo del papel, si se completa, se guarda como dato informativo.
5. **Producción (4.6):** packs totales = pallets completos × packs por pallet + packs sueltos. No se cuentan pallets incompletos; sí los packs sueltos.
6. **Paradas sin horario claro (4.9):** si no se entienden la hora de inicio y la de fin, no se cargan como parada. Quedan en una nota sin horario del reporte.
7. **Parada interpretada:** si administración interpreta un horario (caso del ozonizador del 498: 21:00–21:10), la parada se carga marcada como interpretada.
8. **Datos que el papel no trae:** sección y motivo de rechazo, tipo de parada, velocidad real, etc. Se le piden al supervisor, pero se admite que falten sin bloquear la carga.
9. **Reducción de velocidad (Q-08, cerrada):** afecta al Rendimiento y al OEE por las unidades producidas, con la fórmula actual. No se resta del TO y no necesita fórmula propia. Su formulario sirve para explicar la pérdida.
10. **Formularios de reducciones:** uno en papel para imprimir (supervisor) y uno en Excel (administración).
11. **Personal (5.6):** los puestos son los del papel y se anotan como hoy (ej. Soplador / Ayudante: Carlos / Manuel), en texto libre.
12. **Lote (Q-11, cerrada):** 2 dígitos de mes + 3 de secuencial (001 a 999), con reinicio anual. Clave única: número de lote + año.
13. **Caso 498:** la fila 4 es 18:20–18:30 (10 min) y se agrega la parada del ozonizador 21:00–21:10 (A y C, interpretada).

## 3. Supuestos propios (a confirmar)

- Una parada que cruza dos reportes del mismo turno se asigna al reporte donde empieza.
- Las paradas se validan contra el turno completo y no solo contra el reporte, porque varios reportes pueden compartir turno.
- Rechazo detectado tarde (Q-06, abierta): se carga en el reporte del turno donde se informa y se anota en observaciones.
- Un turno que no corresponde al día (por ejemplo, RES un domingo) avisa y no bloquea, por posibles horas extra.
- Parada sin tipo: 5 min o menos = microparada (TP-05) automática; más de 5 min = TP-06 (sin clasificar) hasta que se clasifique (RN-09).
- Nunca se borra un reporte: se anula y queda marcado.

## 4. Principios de los formularios

**Orden y comodidad**

- Los campos siguen el orden del papel y Tab recorre en ese orden.
- Se muestran en pantalla los cálculos útiles (duración, unidades, litros) para que administración detecte errores de lectura.
- Obligatorio solo lo mínimo: fecha, turno, producto, y hora de inicio y fin de cada parada.
- Los desplegables aceptan valores nuevos (RN-08).

**Bloquea (el dato es imposible o rompe el cálculo)**

- Parada sin inicio o sin fin, o con fin menor o igual al inicio (considerando el cruce de medianoche).
- Parada fuera del horario del turno.
- Solapamiento entre paradas del mismo turno (incluye otros reportes del turno).
- Suma de paradas mayor que el tiempo del turno.
- Tipo inactivo o TP-02 (cabecera).
- Número de lote con formato inválido o número de reporte duplicado.
- Cantidades negativas o no enteras.

**Avisa y guarda como Pendiente de revisar**

- Parada sin tipo o sin motivo; rechazo sin sección o sin motivo.
- Producción en cero; turno sin paradas; turno que no corresponde al día.
- Parada interpretada; nota sin horario.
- OEE o algún factor mayor a 100% (RN-05: se marca, no se corrige).

## 5. Datos que se agregan o cambian

- **Reporte (segmento):** nro\_reporte, encargado, estado\_revision, anulado, notas\_sin\_horario, observaciones, velocidad\_informada (informativa) y horas de inicio y fin del turno según el papel (informativas).
- **Paradas:** códigos de origen, motivo y tipo (no texto), interpretada (sí/no) y observaciones.
- **Producción:** pallets y packs sueltos por separado; unidades buenas y litros calculados y guardados (Doc 8.2).
- **Rechazos:** sección y motivo opcionales; cantidad obligatoria.
- **Tablas nuevas:** tblOrigenes, tblMotivos, tblSeccionesRechazo, tblMotivosRechazo, tblInsumos, tblPuestos, tblReducciones, tblPerdidasInsumo, tblPersonalTurno.
- **Insumos (según el papel):** preformas, envases y tapas en unidades; etiquetas y termocontraíble en gramos; film stretch opcional.
- **Puestos (según el papel):** Soplador / Ayudante, Llenador / Visor, Etiquetador, Forradora / Final de línea y Autoelevador.

## 6. Fases

### F2.0 Cierre de definiciones y Doc v0.6

**Duración:** 2-3 días (parcial) / 1-2 días (completo). Depende del ingeniero.

**Objetivo:** decisiones formalizadas y caso 498 corregido.

**Tareas:**

- Reunión del 2/11 con las preguntas de la sección 9.
- Pasar las decisiones de la sección 2 a su lugar en el Doc; mover las preguntas cerradas a 12.2 (una línea cada una) y dejar solo las abiertas en 12.1.
- Corregir el caso 498 (7.1 a 7.3) con la fila 4 de 10 min y el ozonizador.
- Limpiar la sección 12 (líneas sueltas Q1 y Q2) y definir o quitar D-02.
- Registrar todo en la sección 13 con fecha y aprobación.

**Resultados esperados del 498:**

- Con D-01 (limpiezas como A y C): TO 330, Disponibilidad 81,5%, Rendimiento 75,3%, OEE 61,4%.
- Con Q-18 (limpiezas como microparada): TO 365, Disponibilidad 90,1%, Rendimiento 68,1%, OEE 61,4%.

**Impacto en pruebas:** el caso 498-D01 de 90\_Tests no cambia (−10 min de la limpieza corregida y +10 min del ozonizador se compensan). El caso 498 (Q-18) pasa de TO 375 a 365 (Disponibilidad 90,1% y Rendimiento 68,1%).

**Entregable:** Doc v0.6.

**Criterio de terminado:** el ingeniero aprobó v0.6 (con fecha en el registro de cambios) y lo no resuelto está marcado como Esperando definición.

**Riesgos:** el ingeniero tarda en responder. *Mitigación:* seguir con los supuestos de la sección 3.

### F2.1 Modelo de datos y catálogos

**Duración:** \~3 días (parcial) / 1-2 días (completo).

**Objetivo:** libro v0.2 con toda la estructura y los catálogos cargados.

**Tareas:**

- Crear la rama feature/etapa2-formularios desde etapa1 y copiar el libro a oee\_v0.2\_AAAA-MM-DD.xlsm.
- Macro CompletarEstructura (idempotente) con las tablas nuevas y los campos de la sección 5.
- Pasar tblParadas a códigos y separar pallets y packs sueltos en tblProduccion.
- Cargar catálogos: orígenes, motivos iniciales sacados de los papeles, secciones y motivos de rechazo, insumos, puestos y turnos SOP.
- Corregir el caso 498 en las tablas y en 90\_Tests.
- Verificar que los tests de Etapa 1 siguen pasando.

**Entregable:** libro v0.2 con catálogos cargados.

**Criterio de terminado:** CompletarEstructura corre dos veces sin romper nada, los tests están en verde y el 498 da los resultados de F2.0.

**Depende de:** F2.0 (puede arrancar con los supuestos).

### F2.2 Diseño con administración (en paralelo con F2.1 y F2.3)

**Duración:** \~2 días (parcial) / 1 día (completo).

**Objetivo:** que el formulario resulte cómodo para quien transcribe.

**Tareas:**

- Observar 30 minutos cómo transcribe administración hoy (con permiso del ingeniero).
- Pedir el papel en blanco y 5-10 reportes completos en foto.
- Bocetos en papel de frmReporte, frmParada y la portada, con el orden de Tab igual al del papel.
- Definir los campos mínimos obligatorios y mensajes de error claros en español.
- Acordar con administración un tiempo de carga aceptable por reporte (se mide en F2.5).

**Entregable:** bocetos aprobados.

**Criterio de terminado:** H4 (ingeniero) y conformidad de administración registradas en el Doc.

**Riesgos:** no hay acceso a administración. *Mitigación:* validar con el ingeniero y con las fotos.

### F2.3 Validaciones y acceso a datos (con IA)

**Duración:** 4-5 días (parcial) / 2 días (completo).

**Objetivo:** las reglas de la sección 4 como funciones puras, y escritura segura en las hojas.

**Tareas:**

- Módulo Validaciones: cada regla devuelve bloqueo, aviso u ok, con su mensaje.
- AccesoHojas de escritura: alta, edición y anulación de reporte, parada, producción y rechazo, con IDs únicos.
- Alta automática de valores nuevos en catálogos (RN-08).
- Parámetros (umbral de microparada, etc.) siempre vía ObtenerParametro.
- Casos en 90\_Tests y tests/casos.csv para cada regla: dato malo bloquea, dato bueno pasa, dato faltante avisa.

**Entregable:** módulos Validaciones y AccesoHojas con pruebas.

**Criterio de terminado:** pasan todas las pruebas, incluidos los bordes (cruce de medianoche, solapamiento entre reportes del mismo turno, parada sin hora).

**Depende de:** F2.1.

### F2.4 Formularios núcleo (2A)

**Duración:** \~1 semana (parcial) / 3 días (completo).

**Objetivo:** cargar un reporte completo que alimente el OEE.

**Tareas:**

- frmReporte: nº de registro, fecha, turno, encargado, lote, producto, inicio y fin del segmento; producción (pallets y packs sueltos, con unidades y litros en pantalla); no conformes (total y detalle opcional por sección y motivo); notas sin horario y observaciones.
- frmParada: lista de las paradas del reporte con alta y edición; inicio, fin, duración calculada, origen, motivo encadenado, tipo (automático si dura 5 min o menos, con ayuda en pantalla) y marca de interpretada.
- Editar y anular reportes ya cargados.
- Botón Ver OEE del turno y lista de Pendientes de revisar.
- Portada 00\_Inicio con botones.
- Los formularios solo llaman a Validaciones y AccesoHojas: no calculan ni escriben hojas directo.

**Entregable:** formularios de reporte y paradas funcionando.

**Criterio de terminado:** el 498 se carga solo con formularios y da el resultado esperado.

**Depende de:** F2.2 y F2.3.

**Riesgos:** formularios que piden más de lo que el papel trae. *Mitigación:* todo lo que el papel no trae es opcional.

### F2.5 Prueba y demo de 2A

**Duración:** \~3 días (parcial) / 1-2 días (completo).

**Tareas:**

- Recargar el 498 con los formularios (ambas clasificaciones) y los 4 casos de la sección 7.4 del Doc; comparar con 20\_Calculos.
- Intentar los errores típicos y confirmar que se bloquean o avisan.
- Cargar 3-5 reportes reales desde fotos.
- Que administración cargue un reporte completo sin ayuda; anotar dudas y medir el tiempo.
- Ajustar, exportar a Git y poner tag etapa2a.

**Entregable:** 2A probado.

**Criterio de terminado:** administración carga un turno completo sin ayuda, las validaciones bloquean los datos incorrectos y el OEE coincide con el cálculo a mano (criterio de Etapa 2, Doc sección 10). Hito H5a.

### F2.6 Complementarios (2B)

**Duración:** \~1 semana (parcial) / 3 días (completo).

**Objetivo:** cubrir el resto de lo que trae (y lo que podría traer) el papel.

**Tareas:**

- frmReducciones (origen, motivo, inicio, fin y velocidad real opcional) y su formulario en papel imprimible para el supervisor.
- Tiempo equivalente perdido por reducción como estimación informativa: duración × (1 − velocidad real / velocidad nominal), rotulada como estimado, sujeta a aprobación y sin entrar al OEE.
- frmInsumos (scrap) y personal por turno, según el papel.
- Hoja en papel para rechazos por sección y motivo (o papel v2), según decida el ingeniero.
- Lista de Pendientes de revisar con filtros.

**Entregable:** formularios 2B y papeles nuevos para imprimir.

**Criterio de terminado:** administración carga todos los campos del papel actual y de los papeles nuevos, y el ingeniero aprobó los papeles (H5b).

**Depende de:** F2.5 y de las respuestas sobre el papel y la velocidad real.

### F2.7 Protección, respaldo y cierre

**Duración:** 2-3 días (parcial) / 1-2 días (completo).

**Tareas:**

- Proteger las hojas de datos y catálogos (se editan solo por formularios).
- Respaldo automático con fecha y macro ExportarConfigACSV.
- ExportarModulos incluye .frm y .frx.
- Guía de carga de una página para administración.
- Actualizar el Doc (secciones 7.3, 10 y 13) y registrar cada pedido a la IA en docs/pedidos-ia/.
- Commit, tag etapa2 y demo al ingeniero.

**Entregable:** Etapa 2 completa.

**Criterio de terminado:** aceptación del ingeniero registrada en el Doc.

## 7. Fechas (inicio 2026-11-02, sin contar feriados)

- **15-20 h/semana:** 2A interno \~20-nov (prometer \~30-nov). Fase 2 completa interno \~4-dic (prometer \~14-dic). Incluye +35 % de margen.
- **Tiempo completo:** 2A interno \~12-nov (prometer \~17-nov). Fase 2 completa interno \~18-nov (prometer \~24-nov).
- **Impacto global:** con este alcance, la promesa global de \~1-dic ya no se sostiene. A ritmo parcial, F2 + F3 + F4 cierran internamente hacia mediados de diciembre. Recalcular y avisar al ingeniero.
- **Opción para ganar visibilidad:** después de 2A pasar a Fase 3 (dashboard) y dejar 2B para después, si el ingeniero prioriza ver el tablero.
- Recalcular según el avance real y la velocidad de respuesta del ingeniero.

## 8. Hitos de aprobación del ingeniero

- **H3 (pendiente, no bloquea):** demo de Etapa 1.
- **H4:** Doc v0.6 y diseño de formularios (bocetos), con conformidad de administración.
- **H5a:** demo de 2A con carga real.
- **H5b:** papeles nuevos (reducciones y rechazos) y formularios 2B.

Cada aprobación queda en el registro de cambios del Doc, con fecha.

## 9. Preguntas para el ingeniero (lunes 2/11)

**Prioridad alta**

1. **D-01:** las limpiezas de la etiquetadora del 498 (15, 10 y 10 min), ¿microparada (Q-18) o A y C? El OEE da 61,4% en ambos casos. Con Q-18: Disponibilidad 90,1% y Rendimiento 68,1%. Con D-01: 81,5% y 75,3%.
2. **Papel:** ¿rediseñan el reporte (papel v2) o agregamos hojas complementarias (rechazos por sección y motivo, reducciones de velocidad)? ¿Quién imprime y reparte?
3. **Reducciones de velocidad:** ¿cómo sabe el supervisor la velocidad real? Suponemos que la máquina indica la producción por minuto. ¿Se aprueba la estimación del tiempo equivalente perdido para el dashboard?
4. **Horario real del turno:** ¿puede diferir del fijo (arranque tarde, turno extendido)? ¿Qué TD se usa en ese caso?
5. **Sopladora:** confirmar la redacción: mientras no exista reporte de SOP, sus paradas van en el reporte RES y afectan el OEE de ese turno. ¿Hay fecha estimada para el reporte de SOP?
6. **Lote a las 00:00:** si el lote cambia a medianoche sin cambiar el producto, ¿también se llena otro reporte?

**Prioridad media**

7. **Q-03:** horario del descanso de SOP (mientras tanto se descuentan 35 min sin horario).
8. **Insumos:** confirmar la lista del papel: preformas, envases y tapas en unidades; etiquetas y termocontraíble en gramos; film stretch opcional.
9. **Control de Calidad y observaciones del papel:** ¿entran al libro? Propuesta: no por ahora.
10. **Administración:** ¿transcribe el mismo día o con atraso? ¿Quién corrige si algo no se entiende? ¿Puedo observar 30 minutos cómo carga hoy?
11. **Campo Velocidad nominal de la línea del papel:** si algún supervisor lo completa, ¿tiene efecto o es solo informativo? Propuesta: informativo.
12. **Material:** papel en blanco y 5-10 reportes completos en foto.

Siguen abiertas, sin bloquear la Fase 2: Q-06, Q-13, Q-14 y Q-15 (Fase 3).

## 10. Trabajo con la IA en esta fase

- Prompt base actualizado a Doc v0.6 y con estas reglas extra: los UserForms solo llaman a Validaciones y AccesoHojas; prefijos frm, cbo, txt y btn; los desplegables aceptan valores nuevos; los parámetros se leen con ObtenerParametro.
- Cada pedido incluye casos de prueba con resultado esperado, y se pide explicar los casos borde.
- Probar siempre antes de integrar y cruzar con una fórmula de hoja independiente.
- Registrar cada pedido en docs/pedidos-ia/NNN-nombre.md.

## 11. Riesgos y señales de alerta

**Riesgos:** errores de transcripción desde el papel, carga con atraso, supervisores que no completan los campos nuevos, cambios del papel que no llegan, paradas de sopladora sin reporte propio (castigan el OEE del turno RES), vencimiento de la prueba de M365 (los UserForms solo se desarrollan en Excel de escritorio), crecimiento del alcance (2B), ingeniero sin responder.

**Señales de alerta:**

- Reglas que cambian sin quedar en el Doc.
- Formularios que obligan a completar datos que el papel no trae.
- Administración tarda más en cargar un reporte que el tiempo acordado.
- Más del 30 % de las tareas en Esperando definición.
- OEE mayor a 100% o que no coincide con el cálculo a mano.
- Dos versiones del archivo maestro.
- Pegar código de IA que no puedo explicar.
- Piden sumar otras líneas o áreas antes de terminar esta.

## 12. Próximos pasos inmediatos

- Confirmar este roadmap.
- Prompt para que la IA ajuste el Doc (.md) a v0.6, tras la confirmación.
- Reunión del lunes 2/11 con las preguntas de la sección 9.
- Verificar la fecha de vencimiento de la prueba de M365 y definir un plan B en la notebook del ingeniero.
- Arrancar F2.1 con los supuestos: rama, libro v0.2 y CompletarEstructura.
