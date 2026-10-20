# Roadmap OEE: Excel/VBA, planta embotelladora

Dejaste sin completar la sección "Lo que salió de la primera reunión". Planifiqué asumiendo el peor caso: definiciones abiertas, datos parciales y paradas sin horarios. Al final van las 3 preguntas que más ajustarían las fechas.

---

## 1) Roadmap por fases

Estimación con IA generando el código. Yo conté ~70 h de trabajo total, y la mayor parte se va en definir, revisar y validar, no en tipear.

| Fase | 15-20 h/sem | Tiempo completo |
|---|---|---|
| F0 Definiciones + caso de prueba | 4-5 días | 2-3 días |
| F1 Núcleo de cálculo | ~1 sem | 2-3 días |
| F2 Datos + formularios | 1,5-2 sem | 4-5 días |
| F3 Reportes + dashboard | ~1 sem | 3-4 días |
| F4 Estabilización + entrega | 3-4 días | 2 días |

**Fechas (inicio 20-oct):**
- **Parcial:** interno ~17-nov, **prometer ~1-dic** (+35 %).
- **Completo:** interno ~3-nov, **prometer ~10-nov** (+35 %).
- **Etapa 1 (cálculo funcionando):** interno ~30-oct, prometer ~6-nov.

Las fases F0 y F1 se solapan: F0 depende del ingeniero, por eso F1 arranca con datos inventados.

---

## 2) Detalle por fase

**F0: Definiciones y caso de prueba**
- *Objetivo:* reglas cerradas y un turno real calculado a mano.
- *Tareas:* armar el Doc fuente, volcar la reunión, enviar preguntas en bloque, calcular a mano un turno real.
- *Entregable:* Doc con glosario, fórmulas, reglas y 1 turno resuelto.
- *Terminado:* el ingeniero aprobó fórmulas y el turno de prueba (con fecha en el registro de cambios).
- *Desbloquea:* Bloque A de preguntas.

**F1: Núcleo de cálculo (Etapa 1)**
- *Objetivo:* OEE correcto con datos cargados a mano.
- *Tareas:* funciones puras (Disponibilidad, Rendimiento, Calidad, OEE), agregación día/semana sumando tiempos y unidades, casos de prueba.
- *Entregable:* `oee_v0.x.xlsm` con hojas de datos y módulo de cálculo.
- *Terminado:* el turno real da exactamente lo calculado a mano y pasan 8-10 casos de prueba, incluyendo bordes (0 unidades, 0 tiempo).
- *Depende de:* F0 (Bloque A).

**F2: Datos y formularios (Etapa 2)**
- *Objetivo:* que la carga sea cómoda y no deje pasar datos malos.
- *Tareas:* tablas de Excel, catálogos, UserForms, validaciones (solapamientos, horas, causas), módulo de acceso a hojas.
- *Entregable:* formularios de Paradas y Producción funcionando.
- *Terminado:* un operador real carga un turno completo sin ayuda y el resultado coincide.
- *Depende de:* Bloque B.

**F3: Reportes y dashboard (Etapa 3)**
- *Objetivo:* ver OEE, sus 3 factores, tendencia y pérdidas.
- *Tareas:* selector turno/día/semana, Pareto de paradas, gráfico de tendencia, metas si existen.
- *Entregable:* hoja Dashboard.
- *Terminado:* el ingeniero valida que los números coinciden con su criterio.
- *Depende de:* Bloque C.

**F4: Estabilización y entrega**
- *Objetivo:* que funcione sin mí al lado.
- *Tareas:* correr con datos reales 1 semana, corregir, manual de 1-2 páginas, protección de hojas, backup.
- *Entregable:* versión 1.0 y documentación.
- *Terminado:* aceptación firmada por el ingeniero.

---

## 3) Qué pedirle al ingeniero (por prioridad)

**Bloque A: bloquea todo (F0/F1)**
1. ¿Qué es "tiempo planificado"? (turno completo vs. descontando almuerzo/descansos)
2. ¿Qué paradas descuentan del planificado (almuerzo, limpieza, mantenimiento programado) y cuáles son pérdida?
3. Velocidad nominal / ciclo ideal: valor y de dónde sale.
4. Unidad de conteo (botellas, bidones) y producto de la línea.
5. ¿Qué es un rechazo? ¿Se registra en la línea o después? ¿El retrabajo cuenta como bueno o malo?
6. **Un turno real completo con datos crudos** (paradas, producción, rechazos) y que él calcule el OEE a mano.

**Bloque B: bloquea F2**
7. Umbral de registro de paradas (¿desde cuántos minutos se anotan?).
8. ¿Se registra hora inicio/fin o solo minutos?
9. Quién carga, cuándo y con qué dispositivo.
10. Catálogos: causas de parada, tipos de rechazo, turnos (horarios, ¿alguno cruza medianoche?).
11. ¿Hay cambios de producto dentro de un turno?

**Bloque C: bloquea F3**
12. Reportes que quiere ver y metas de OEE.

**Bloque D: accesos**
13. Copias de planillas actuales (sin datos sensibles o anonimizadas).
14. Permiso para repo privado en GitHub y acceso al Doc.
15. Ver cómo se registra hoy en la línea (aunque sea 30 min).
16. Rol de la tercera persona.

---

## 4) Próximos 5 días

**Mar 20 (hoy)**
- Crear el Doc fuente con el esqueleto y volcar lo de la reunión.
- Crear el repo y el tablero.
- Dejar la lista de preguntas con fecha.

**Mié 21**
- Enviar el Bloque A al ingeniero (un solo mensaje).
- Armar `oee_v0.1` con hojas vacías y catálogos provisorios.
- Inventar 4 turnos de prueba y calcularlos a mano.

**Jue 22**
- Pedir a la IA el primer módulo `Calculos.bas`.
- Revisarlo y probarlo con los 4 turnos.
- Exportar a Git.

**Vie 23**
- Revisar respuestas del Bloque A.
- Actualizar el Doc y el registro de cambios.
- Ajustar casos de prueba con el turno real, si llegó.

**Lun 26**
- Agregación por día y semana (sumando, no promediando).
- Publicar `v0.2` y preparar el Bloque B.

---

## 5) Repositorio

```
oee-planta/
├─ README.md
├─ docs/
│  ├─ reglas.md          (espejo del Doc)
│  ├─ decisiones/
│  └─ pedidos-ia/
├─ src/vba/              (.bas .cls .frm .frx)
├─ tests/                (casos en CSV)
├─ data/ejemplo/         (datos ficticios)
├─ libro/                (el .xlsm vive acá, ignorado)
└─ tools/
```

**.gitignore**
```
*.xlsm
*.xlsx
*.xlsb
~$*
libro/
data/real/
```

- Se sube todo en `src/vba`, incluidos los `.frx` (son binarios, pero el formulario no se reconstruye sin ellos).
- Si querés compartir el libro, subí una plantilla **vacía** como Release, nunca con datos reales.
- **Ramas:** `main` (estable), `feature/etapa1-calculo`, `fix/...`.
- **Commits:** `feat(calculo): agrega Disponibilidad`, `test:`, `docs:`, `fix:`.

**Exportar VBA:** activá en Excel *Opciones > Centro de confianza > Configuración de macros > Confiar en el acceso al modelo de objetos de proyectos de VBA*. Después pegá esto en un módulo:

```vba
Option Explicit

Public Sub ExportarVBA()
    Dim comp As Object, ext As String, ruta As String
    ruta = ThisWorkbook.Path & "\..\src\vba\"
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1: ext = ".bas"
            Case 2, 100: ext = ".cls"
            Case 3: ext = ".frm"
            Case Else: ext = ""
        End Select
        If ext <> "" Then comp.Export ruta & comp.Name & ext
    Next comp
    MsgBox "Exportado en " & ruta
End Sub
```
Ejecutala al cerrar cada día y hacé `git commit`.

---

## 6) Trabajar con la IA como programadora

**Formato de pedido:**
```
CONTEXTO: [reglas del Doc, versión/fecha]
OBJETIVO: función X
ENTRADAS/SALIDAS: tipos exactos
REGLAS: división por cero, redondeo, unidades
RESTRICCIONES: Option Explicit, función pura,
  sin leer hojas, sin On Error Resume Next
CASOS DE PRUEBA: [entradas → resultado esperado]
PEDIDO EXTRA: explicá cada decisión y casos borde
```

**Revisión:**
- Cada función puede explicarse en una frase.
- No toca hojas si es de cálculo.
- Maneja divisiones por cero y tipos (`Double`).
- Los casos de prueba pasan **antes** de integrar.
- Cruzá el resultado con una fórmula de hoja independiente.
- Desconfiá de OEE > 100 % o muy parejo (~85 %): suele ser un error.

**Pruebas exigidas:** turno real, turno sin paradas, sin producción, con 100 % rechazo, turno que cruza medianoche, agregación semanal.

**Registro:** un archivo por pedido en `docs/pedidos-ia/NNN-nombre.md` con el pedido, la versión de reglas, el resultado de las pruebas, estado (aprobado/rechazado) y el commit.

---

## 7) Hitos de aprobación del ingeniero

1. Glosario y fórmulas.
2. Turno calculado a mano (con el OEE esperado).
3. Demo Etapa 1.
4. Catálogos y diseño de formularios (en papel).
5. Demo Etapa 2 con carga real.
6. Layout del dashboard.
7. Aceptación final.

Todos quedan en el registro de cambios con fecha.

---

## 8) Riesgos por fase

- **F0:** el ingeniero tarda en responder. *Mitigación:* preguntas en bloque, avanzar con supuestos marcados "Esperando definición".
- **F1:** código que parece correcto y calcula mal. *Mitigación:* casos de prueba antes de integrar.
- **F2:** datos desordenados o paradas sin horario. *Mitigación:* validaciones fuertes, aceptar "minutos" como plan B.
- **F3:** dashboard que no refleja lo que él espera. *Mitigación:* validar un mockup en papel antes de construir.
- **F4:** requisitos nuevos de último momento. *Mitigación:* lista de "versión 2" aparte.
- **Transversal:** dos versiones del .xlsm. *Mitigación:* una sola persona modifica el maestro.

---

## 9) Preparar la migración futura

- Una fila = un evento (parada o producción), nada de celdas combinadas.
- Fechas y horas reales (datetime), no texto.
- IDs únicos en cada registro y en cada catálogo (ID + nombre).
- Columna `linea_id` desde ya, aunque haya una sola línea.
- No guardar valores calculados en hojas de datos.
- Capa de acceso a hojas separada de la de cálculo.
- Casos de prueba en CSV: los vas a reutilizar con Jest/Vitest en TypeScript.
- Macro que exporte las tablas a CSV.
- Reglas del Doc con versión, para portarlas tal cual.

---

## 10) Señales de alerta

- Reglas que cambian sin quedar en el Doc.
- El ingeniero sin responder más de 3 días hábiles.
- Más del 30 % de tareas en "Esperando definición".
- OEE > 100 % o resultados que no coinciden con el cálculo a mano.
- Hacer formularios antes de tener el cálculo validado.
- Cambios de requisitos por WhatsApp.
- Dos versiones del archivo maestro.
- Pegar código de la IA que no podés explicar.
- Piden sumar otras líneas o áreas antes de terminar la primera.
- Paradas sin horarios y sin plan B definido.

---

## Mis 3 preguntas

1. ¿Qué definiciones cerraron en la reunión y qué datos reales conseguiste (formato y detalle)?
2. ¿Cómo se registran hoy paradas y producción en la línea (papel, Excel, con horarios o solo minutos)?
3. ¿Qué rol tiene la tercera persona y hay fechas acordadas con el ingeniero?

Con eso te ajusto fechas y te dejo el Bloque A redactado listo para mandar.