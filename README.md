# Palau OEE

Cálculo de **OEE (Eficiencia General de los Equipos)** para la línea de embotellado de una planta de agua, implementado como libro de Excel con VBA.

El OEE se calcula **por turno** y se agrega por día y semana. Los reportes y el tablero se generan dentro del propio libro.

> **Estado:** en desarrollo. Versión de documentación: **v0.8** (2026-11-04). Ver [Roadmap F2](<docs/roadmaps/Roadmap Fase 2 datos y formularios (Etapa 2).md>).

---

## Qué mide

| Factor | Qué representa | Pérdidas que lo afectan |
|---|---|---|
| **Disponibilidad** | Tiempo operativo / tiempo planificado | Averías, fallas y ajustes de más de 5 minutos |
| **Rendimiento** | Velocidad real frente a la nominal | Pequeñas paradas (5 minutos o menos) y reducción de velocidad |
| **Calidad** | Unidades buenas / unidades totales | Rechazos en las 4 máquinas de la línea |

```
OEE = Disponibilidad × Rendimiento × Calidad
```

La línea medida es completa (sopladora, llenadora, etiquetadora y forradora), no máquina por máquina. Los productos son botellas (500 ml, 1500 ml y 2000 ml, con y sin gas) y bidones de 6 L.

Las reglas de negocio (clasificación de paradas, turnos, lotes, criterios de rechazo) están en [`documentación_oficial_v0_8.md`](documentación_oficial_v0_8.md). Ese documento es la **única fuente de verdad**: si el código y el documento no coinciden, manda el documento.

---

## Estructura del repositorio

```
palau-OEE/
├─ documentación_oficial_v0_8.md  Reglas, fórmulas y decisiones aprobadas (fuente de verdad)
├─ docs/
│  ├─ estructura-datos.md  Hojas, tablas y encabezados del libro maestro
│  ├─ catalogos-iniciales.md  Parámetros y catálogos iniciales
│  ├─ roadmaps/             Plan de fases, dependencias y riesgos
│  └─ pedidos-ia/          Registro de pedidos de código a la IA
├─ src/vba/                Módulos VBA exportados (.bas)
│  ├─ Utils.bas            Helpers comunes (tablas, columnas, fechas, números)
│  ├─ Calculos.bas         Funciones puras de cálculo (Disponibilidad, Rendimiento, Calidad, OEE)
│  ├─ Validaciones.bas     Reglas puras de validación de carga
│  ├─ AccesoHojas.bas      Lectura y escritura de tablas (única escritura de producción)
│  ├─ Agregacion.bas       Agregación por turno y día
│  ├─ Modulo_Estructura.bas  Migración idempotente de estructura
│  ├─ DatosPrueba.bas      Fixtures de prueba
│  ├─ Tests.bas            Casos de prueba
│  └─ ExportarModulos.bas  Exporta módulos y formularios a src/vba/
├─ tools/
│  ├─ crear_libro.py       Genera el libro maestro con Excel COM
│  └─ test_compatibilidad.py  Verifica columnas del libro contra el VBA
├─ data/ejemplo/           Datos ficticios para pruebas
├─ tests/                  Casos de prueba
└─ libro/                  Libro .xlsm de trabajo (no versionado)
```

### Principios de diseño

- **Funciones de cálculo puras:** no leen ni escriben hojas. Reciben valores y devuelven resultados.
- **Capa de acceso a datos separada** de la capa de cálculo.
- **Una fila = un evento** (parada o producción). Sin celdas combinadas.
- **Fechas y horas reales** (no texto), **códigos** en los catálogos (no descripciones) e IDs únicos.
- **Sin valores calculados guardados** en las hojas de datos.
- Divisiones por cero devuelven vacío en lugar de error.

---

## Requisitos

- Windows con **Microsoft Excel** (versión de escritorio con macros habilitadas).
- Para regenerar el libro desde cero: **Python 3** y **pywin32**.

```powershell
python -m pip install pywin32
```

---

## Uso

### 1. Generar el libro maestro

```powershell
python tools/crear_libro.py
```

Crea `libro/oee_master.xlsm` con las 20 hojas, las tablas de Excel, los catálogos iniciales y las validaciones. La estructura exacta está en [`docs/estructura-datos.md`](docs/estructura-datos.md).

### 2. Importar o exportar el código VBA

Para exportar los módulos a `src/vba/` desde el libro, activar en Excel:

*Opciones > Centro de confianza > Configuración de macros > Confiar en el acceso al modelo de objetos de proyectos de VBA*

Luego ejecutar la macro `ExportarModulos` del módulo `ExportarModulos.bas`.

### 3. Ejecutar las pruebas

Los casos de prueba están en `src/vba/Tests.bas` y en `tests/`. Deben pasar **antes** de integrar cualquier cambio de cálculo.

---

## Control de cambios

- Toda regla o fórmula cambia primero en `documentación_oficial_v0_8.md`, con fecha y aprobación. Después se refleja en el código.
- Cada cambio relevante queda registrado en la sección 13 del documento.
- Ramas: `main` (estable), `feature/...`, `fix/...`.
- Commits con prefijo: `feat(...)`, `fix(...)`, `test:`, `docs:`.

### Archivos que no se versionan

Por `.gitignore`: libros de Excel (`*.xlsm`, `*.xlsx`, `*.xlsb`), archivos temporales de Office (`~$*`), `libro/`, `data/real/`, `deprecado/` y `__pycache__/`. Los datos reales de producción nunca se suben al repositorio.

---

## Roadmap

| Fase | Contenido | Estado |
|---|---|---|
| F0 | Definiciones de negocio y caso de prueba con un turno real calculado a mano | Completado |
| F1 | Núcleo de cálculo (Etapa 1) | Completado |
| F2 | Datos y formularios de carga (Etapa 2) | En curso: F2.1 y F2.3 completadas; F2.4 pausada para estabilizar |
| F3 | Reportes y dashboard (Etapa 3) | Pendiente |
| F4 | Estabilización y entrega (versión 1.0) | Pendiente |

El detalle de cada fase, sus criterios de aceptación y los riesgos asociados está en [`docs/roadmaps/Roadmap Fase 2 datos y formularios (Etapa 2).md`](<docs/roadmaps/Roadmap Fase 2 datos y formularios (Etapa 2).md>).

### Criterio de éxito

El libro debe coincidir con un turno real calculado a mano, con una tolerancia de **0,05 puntos porcentuales** para el OEE y cada factor (ver `documentación_oficial_v0_8.md`, sección 1.4).

---

## Contacto

Desarrollo: Joaquín. Las reglas de negocio las define y aprueba el ingeniero de planta, dueño de [`documentación_oficial_v0_8.md`](documentación_oficial_v0_8.md).
