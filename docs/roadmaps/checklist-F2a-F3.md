# Checklist F2a y F3

Versión 1.0 · 2026-11-05 · Reglas: Doc v0.8 · Plan: `roadmap.md` v2

**Regla de cierre de cada tarea:** `tools/ejecutar_tests.py` en verde (hoy 100/100) + commit con prefijo (`feat:`, `test:`, `docs:`).
**No confundir:** esta F2a (puente de datos) no es el "2A" del Roadmap Fase 2 (formularios núcleo).

---

# F2a: puente de datos (2-3 días)

**Objetivo:** esquema congelado y datos suficientes para construir y probar el dashboard sin formularios.

## Bloque 0: preparación

- [ ] `ejecutar_tests.py` en 100/100 y `git status` limpio.
- [ ] Tag del estado estable (ej. `backend-estable-2026-11-05`).
- [ ] Rama `feature/etapa3-dashboard` desde `main`.
- [ ] Verificar la fecha de vencimiento de la prueba de M365 y anotar el plan B (notebook del ingeniero).
- [ ] Doc v0.9 (Propuesto): registrar en la sección 13 el reordenamiento "Etapa 3 antes de 2.4".

## Bloque 1: congelar el esquema

- [ ] Listar qué columnas necesita leer el dashboard: `tblTurnosRegistro`, `tblParadas`, `tblProduccion`, `tblRechazos`, `tblTramosSinEnvase`, `tblReducciones`, `tblProductos`, `tblVelocidades`, `tblLotes`, `tblOrigenes`, `tblMotivos`.
- [ ] Marcar cada necesidad como "ya está" o "falta" (ej. máquina de una parada: sale de `origen` → `tblOrigenes.clase`; producto de un reporte: sale de `id_lote` → `tblLotes`).
- [ ] Verificar qué lecturas existen en `AccesoHojas`. Hoy hay `ListarParadas`, `ListarCatalogo`, `ListarReportes`, `ListarTramos`. Revisar si faltan lectores de producción, rechazos y reducciones, y agregarlos ahí (no en `Dashboard.bas`).
- [ ] Si falta alguna columna: migración idempotente en `Modulo_Estructura`, `crear_libro.py`, `estructura-datos.md` y `test_compatibilidad.py`, en ese orden y con el Doc primero.
- [ ] Registrar en el Doc (sección 13) "esquema congelado para Etapa 3" con fecha.

**Terminado:** lista de lecturas cerrada y sin ítems "falta" sin resolver.

## Bloque 2: dataset de demo

- [ ] Diseñar 12-15 reportes en CSV (`data/ejemplo/`). Debe incluir:
  - [ ] El **498 real** (sin prefijo) y los 4 casos de 7.4 (inventados, con prefijo).
  - [ ] Los 5 productos (500 SG, 500 CG, 1500, 2000, bidón).
  - [ ] Turnos RES y SOP, con al menos uno que cruce medianoche.
  - [ ] Un turno con dos reportes (cambio de producto).
  - [ ] Tramos sin envase (uno y varios).
  - [ ] Microparadas (5 min o menos), paradas largas y una parada **interpretada**.
  - [ ] Rechazos con y sin sección o motivo.
  - [ ] Un turno sin paradas y un reporte **anulado**.
  - [ ] Dos semanas lunes a domingo, con un turno en el cambio de semana.
  - [ ] 3-4 reducciones de velocidad con motivos.
- [ ] Calcular a mano (o con fórmula de hoja independiente) el OEE de al menos 5 de esos turnos. Guardarlo como resultado esperado.
- [ ] Macro `CargarDatasetDemo`: lee el CSV y carga con `GuardarReporteCompleto` (pasa por `Validaciones`). Las reducciones entran por `AccesoHojas`.
- [ ] Prefijo `DEMO-` en `nro_reporte` y nota en `observaciones` de todo lo ficticio.
- [ ] Macro `LimpiarDatasetDemo` que borra solo lo `DEMO-` (patrón de `LimpiarDatosPrueba`).
- [ ] Trabajar en una **copia** del maestro: `oee_demo_AAAA-MM-DD.xlsm`. El maestro queda sin datos ficticios.
- [ ] Test nuevo: tras cargar, N reportes cargados, el 498 da 61,4% y los OEE esperados coinciden (tolerancia 0,05 pp).
- [ ] Test de limpieza: tras `LimpiarDatasetDemo`, no quedan residuos `DEMO-`.

**Terminado:** cargar y limpiar dos veces seguidas funciona sin duplicar ni dejar basura.

## Bloque 3: carga de datos reales hasta que haya formularios (opcional)

- [ ] Preguntarle al ingeniero si quiere ver el dashboard con sus turnos reales antes de F2.4.
- [ ] Si sí: macro importadora desde una plantilla simple (CSV) que llama a `GuardarReporteCompleto`. Misma ruta de validación, nada de pegar a mano en las tablas.
- [ ] Si no: saltar este bloque.

## Criterio de terminado de F2a

- [ ] Esquema congelado y registrado en el Doc.
- [ ] `oee_demo_*.xlsm` con el dataset, y el 498 da 61,4% desde las tablas.
- [ ] 100 tests previos + los nuevos en verde.
- [ ] Exportado a Git y commit.

---

# F3: dashboard (≈ 2,5 semanas, en 3 etapas)

**Objetivo:** OEE, tres factores y detalle de pérdidas, según Doc 4.10 y 9.

## D0: definir y aprobar antes de construir (H6-0)

- [ ] Doc v0.9 sección 9: listar el contenido del tablero en orden de prioridad (Propuesto).
- [ ] Boceto de **una** pantalla (papel o mockup en Excel con números ficticios): filtros arriba, cuatro tarjetas (OEE y 3 factores), cascada TD → TPL → TO, Pareto por máquina, microparadas, reducciones, litros, tendencia.
- [ ] Mandar al ingeniero, en un solo mensaje:
  - [ ] Rangos del semáforo (verde, amarillo, rojo). Si no responde, proponer valores provisorios y rotularlos así en pantalla.
  - [ ] ¿Quiere reporte por día?
  - [ ] ¿Qué quiere ver primero y con qué turnos reales probar?
  - [ ] ¿Pérdida de insumo en el tablero?
- [ ] Guardar los umbrales como filas nuevas de `tblParametros` (estado Propuesto), no en el código.
- [ ] Decidir cómo se tratan los reportes "Pendiente de revisar": incluirlos y marcarlos. Registrarlo en el Doc.
- [ ] **H6-0:** boceto aprobado (con fecha en el Doc).

## D1: datos para el dashboard, sin pantalla (F3a)

- [ ] Función de agregación por turno (fecha + turno, sumando todos sus reportes) usando `TDetalleTurno`.
- [ ] Función de **tendencia**: un punto por turno (nunca promedio de OEE).
- [ ] Función de cascada de tiempos: TD, PP, PE, TPL, averías y fallas, ajustes y configuración, TO.
- [ ] Cada función con casos en `Tests.bas` y contra los esperados del bloque 2 de F2a.
- [ ] Todas ignoran anulados. Divisiones por cero devuelven vacío.

## D2: hoja Dashboard, vista de turno (F3a)

- [ ] Hoja `21_Dashboard` (y una auxiliar `22_DashAux`, oculta, solo si hace falta; Doc primero).
- [ ] `Dashboard.bas`: `ActualizarDashboard` que lee, no calcula.
- [ ] Selectores con validación de datos: fecha, turno, producto, línea.
- [ ] Cuatro tarjetas con semáforo (formato condicional que lee `tblParametros`).
- [ ] Cascada de tiempos como gráfico.
- [ ] Señalar OEE o factor mayor a 100% (RN-05) y reportes pendientes.
- [ ] "—" si el turno no tiene datos o no existe.
- [ ] **Verificación:** el 498 muestra 61,4% (D-01); cada cifra coincide con `20_Calculos` a 1 decimal.
- [ ] Exportar, commit y copia versionada del libro.
- [ ] **H6a:** demo al ingeniero (también cumple H3). Mostrar el 498, cambiar una variable y ver el recálculo, y decir qué falta.

## D3: pérdidas (F3b)

- [ ] Función de Pareto de paradas por máquina (`origen` → `tblOrigenes.clase`) en **horas**, de mayor a menor.
- [ ] Función de Pareto por motivo dentro de cada máquina.
- [ ] Función de microparadas: cantidad, minutos y ranking de motivos.
- [ ] Función de paradas externas (PP y PE) por tipo.
- [ ] Función de reducciones de velocidad por motivo (duración). **Sin** tiempo equivalente perdido hasta resolver Q-21.
- [ ] Función de litros por período.
- [ ] Función de OEE por producto.
- [ ] Tests de cada una con sus casos borde (sin datos, un solo motivo, empate).
- [ ] Gráficos y tablas en el dashboard.

## D4: período, filtros y tendencia (F3c)

- [ ] Helper de semana lunes a domingo en `Utils` (`InicioSemana`, `FinSemana`).
- [ ] Agregación por período: **suma** tiempos y unidades y recién después calcula los factores.
- [ ] Test clave: dos turnos de OEE distinto; el período no es el promedio de ambos.
- [ ] Selector de período (turno / semana; día solo si el ingeniero lo confirma).
- [ ] Filtros de línea, producto y turno aplicados a todos los widgets.
- [ ] Gráfico de tendencia por turno.
- [ ] Pérdida de insumo: solo si la pidió.

## D5: cierre (H6b)

- [ ] Medir tiempo de refresco con un dataset grande (ej. 200-500 reportes). Si es lento, usar la hoja auxiliar.
- [ ] Probar sin datos, con un solo turno, con un reporte anulado y con un cambio de semana.
- [ ] Proteger la hoja del dashboard (sin tocar las tablas).
- [ ] Actualizar el Doc (secciones 8.1, 9, 10 y 13) y registrar cada pedido a la IA en `docs/pedidos-ia/`.
- [ ] `ExportarModulos` (incluye `Dashboard.bas`), commit y tag `etapa3`.
- [ ] **H6b:** demo completa; el ingeniero valida que los números coinciden con su criterio. Registrar con fecha.

## Criterio de terminado de F3 (Doc sección 10)

- [ ] Layout aprobado en boceto (H6-0).
- [ ] Los números del dashboard coinciden con `20_Calculos` y con el cálculo a mano.
- [ ] El ingeniero validó (H6b).
- [ ] Todos los tests en verde.

---

## Dependencias rápidas

| Tarea | Depende de | Bloquea |
|---|---|---|
| F2a | Estabilización (hecha) | D1, D2 |
| D0 | Respuesta del ingeniero (o supuestos marcados) | D2 en adelante |
| D1 | F2a (dataset) | D2, D3 |
| D2 | D1, D0 | H6a |
| D3, D4 | D2 | H6b |
| F2.2 | Q-29 | F2.4 |
