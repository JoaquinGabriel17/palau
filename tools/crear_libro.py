#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera el libro maestro OEE con Excel COM (pywin32).

Uso:
    python tools/crear_libro.py

Crea libro/oee_v0.5_AAAA-MM-DD.xlsm con las 19 hojas del libro maestro,
una Excel Table por catálogo y por conjunto de datos, encabezados y datos
iniciales, validación Sí/No en las columnas `activo`, columnas de fecha en
formato fecha real (sin texto) y sin celdas combinadas. Guarda como .xlsm
(FileFormat 52) y cierra Excel.

Fuentes de verdad (los datos de este script espejan estos archivos):
    docs/estructura-datos.md    -> hojas, tablas y encabezados exactos
    docs/catalogos-iniciales.md -> parámetros y catálogos iniciales
"""

import datetime as dt
import os
import sys

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover
    sys.exit("Falta pywin32. Instalar con: python -m pip install pywin32")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VERSION = "0.5"
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR_LIBRO = os.path.join(RAIZ, "libro")

# Constantes Excel (evitan gencache / nombres localizados)
XL_SRC_RANGE = 1
XL_HAS_HEADERS = 1
XL_VALIDATE_LIST = 3
XL_VALIDATE_DATE = 4
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_OPENXML_WORKBOOK_MACROENABLED = 52  # .xlsm

HOJAS = [
    "00_Inicio",
    "01_Parámetros",
    "02_Cat_Paradas",
    "03_Cat_Turnos",
    "04_Cat_Productos",
    "05_Cat_Calidad",
    "06_Cat_Insumos",
    "07_Cat_Personal",
    "10_Lotes",
    "11_Turnos",
    "12_Paradas",
    "13_Produccion",
    "14_Rechazos",
    "15_Reducciones",
    "16_PerdidasInsumo",
    "17_PersonalTurno",
    "20_Calculos",
    "21_Dashboard",
    "90_Tests",
]

# Columnas por tipo (según docs/estructura-datos.md)
COL_SI_NO = {"activo"}
COL_FECHA = {
    "vigente_desde",
    "vigente_hasta",
    "fecha_aprobacion",
    "vigencia_desde",
    "fecha",
}
COL_FECHA_HORA = {"fecha_hora_inicio", "fecha_hora_fin"}
COL_HORA = {"hora_inicio", "hora_fin", "descanso_inicio", "descanso_fin"}

FORMATO_FECHA = "DD/MM/YYYY"
FORMATO_FECHA_HORA = "DD/MM/YYYY HH:MM"
FORMATO_HORA = "HH:MM"

BUFFER_FILAS = 50  # filas futuras con formato/validación ya aplicados


def D(anio, mes, dia):
    """Fecha real (se escribe como fecha de Excel, no como texto)."""
    return dt.datetime(anio, mes, dia)


def T(hora, minuto):
    """Hora del día como serial de Excel (fracción del día, formato HH:MM)."""
    return (hora * 60 + minuto) / 1440.0


# ---------------------------------------------------------------------------
# Parámetros (docs/catalogos-iniciales.md §1 -> tblParametros)
# ---------------------------------------------------------------------------
PARAM_HEADERS = [
    "id",
    "nombre",
    "parametro",
    "ambito",
    "producto_codigo",
    "valor",
    "unidad",
    "vigente_desde",
    "vigente_hasta",
    "motivo",
    "aprobado_por",
    "estado",
    "fecha_aprobacion",
]

PARAM_ROWS = [
    ("P-01", "Umbral de microparada", "UMBRAL_MICROPARADA_MIN", "Global", "", 5,
     "min", D(2026, 10, 29), "", "", "Ingeniero + reg. 13", "Aprobado", ""),
    ("P-02", "Descanso reglamentario", "DESCANSO_MIN", "Global", "", 35,
     "min", D(2026, 10, 29), "", "", "Ingeniero + reg. 13", "Aprobado", ""),
    ("P-03", "Velocidad nominal 500 sin gas", "VEL_NOM_500_SG", "Producto", "500SG", 90,
     "botellas/min", D(2026, 1, 1), "", "", "Ingeniero", "Aprobado (vigencia supuesta)", ""),
    ("P-04", "Velocidad nominal 500 con gas", "VEL_NOM_500_CG", "Producto", "500CG", 70,
     "botellas/min", D(2026, 1, 1), "", "", "Ingeniero", "Aprobado (vigencia supuesta)", ""),
    ("P-05", "Velocidad nominal 1500 sin gas", "VEL_NOM_1500_SG", "Producto", "1500SG", 70,
     "botellas/min", D(2026, 1, 1), "", "", "Ingeniero", "Aprobado (vigencia supuesta)", ""),
    ("P-06", "Velocidad nominal 2000 sin gas", "VEL_NOM_2000_SG", "Producto", "2000SG", 70,
     "botellas/min", D(2026, 1, 1), "", "", "Ingeniero", "Aprobado (vigencia supuesta)", ""),
    ("P-07", "Velocidad nominal bidón 6 L", "VEL_NOM_BIDON_6L", "Producto", "BID6", 10,
     "bidones/min", D(2026, 1, 1), "", "", "Ingeniero", "Aprobado (vigencia supuesta)", ""),
    ("P-08", "Tolerancia de comparación", "TOLERANCIA_COMPARACION_PP", "Global", "", 0.05,
     "puntos porcentuales", D(2026, 10, 29), "", "", "Ingeniero + reg. 13", "Aprobado", ""),
]

# ---------------------------------------------------------------------------
# Catálogo de paradas (docs/catalogos-iniciales.md §2-§4)
# ---------------------------------------------------------------------------
TIPOS_HEADERS = ["codigo", "tipo", "clasificacion", "efecto_calculo",
                 "seleccionable", "activo", "estado"]
TIPOS_ROWS = [
    ("TP-01", "Parada programada", "PP", "Descuenta TPL; no afecta OEE", "Sí", "Sí", "Aprobado"),
    ("TP-02", "Parada no planificada", "PNP",
     "Cabecera de clasificación; por sí sola no calcula efecto", "No (cabecera)", "Sí", "Propuesto"),
    ("TP-03", "Avería o falla", "PNP / A/F", "Resta TO; baja Disponibilidad",
     "Sí (solo >5 min)", "Sí", "Aprobado"),
    ("TP-04", "Ajuste y configuración", "PNP / A y C", "Resta TO; baja Disponibilidad",
     "Sí (solo >5 min)", "Sí", "Aprobado"),
    ("TP-05", "Pequeña parada / microparada", "PNP / micro", "No resta TO; afecta Rendimiento",
     "Sí (automática ≤5 min)", "Sí", "Aprobado"),
    ("TP-06", "Sin clasificar", "PNP provisoria", "Resta TO hasta clasificar",
     "Sí (>5 min)", "Sí", "Aprobado"),
    ("TP-07", "Parada externa", "PE", "Descuenta TPL; no afecta OEE", "Sí", "Sí", "Aprobado"),
]

ORIGENES_HEADERS = ["codigo", "nombre", "clase", "activo"]
ORIGENES_ROWS = [
    ("OR-01", "sopladora", "máquina de la línea", "Sí"),
    ("OR-02", "llenadora", "máquina de la línea", "Sí"),
    ("OR-03", "etiquetadora", "máquina de la línea", "Sí"),
    ("OR-04", "forradora", "máquina de la línea", "Sí"),
    ("OR-05", "refrigeración de agua", "servicio de apoyo", "Sí"),
    ("OR-06", "tratamiento de agua", "servicio de apoyo", "Sí"),
    ("OR-07", "suministro de agua", "servicio de apoyo", "Sí"),
    ("OR-08", "descanso/personal", "otro", "Sí"),
    ("OR-09", "cambio de formato/producto", "otro", "Sí"),
    ("OR-10", "mantenimiento programado", "otro", "Sí"),
    ("OR-11", "reunión/capacitación", "otro", "Sí"),
    ("OR-12", "limpieza/sanitizado programado", "otro", "Sí"),
    ("OR-13", "otro/no clasificado", "otro", "Sí"),
]

MOTIVOS_HEADERS = ["codigo", "motivo", "origen", "tipo_sugerido", "activo"]
MOTIVOS_ROWS = [
    ("MO-01", "volcador de preforma", "OR-01", "", "Sí"),
    ("MO-02", "atasco a la salida de la etiquetadora", "OR-03", "", "Sí"),
    ("MO-03", "limpieza de adhesivo", "OR-03", "", "Sí"),
    ("MO-04", "falla de refrigeración por alta temperatura", "OR-05", "", "Sí"),
    ("MO-05", "falla en horno", "OR-01", "", "Sí"),
]

# ---------------------------------------------------------------------------
# Turnos (docs/catalogos-iniciales.md §5 -> tblTurnos)
# ---------------------------------------------------------------------------
TURNOS_HEADERS = [
    "codigo",
    "nombre",
    "equipo",
    "hora_inicio",
    "hora_fin",
    "descanso_inicio",
    "descanso_fin",
    "dias_aplica",
    "activo",
    "vigente_desde",
    "vigente_hasta",
    "cruza_medianoche",
    "td_minutos",
    "descanso_minutos",
]
TURNOS_ROWS = [
    ("RES-M", "Mañana RES", "RES", T(6, 0), T(14, 0), T(11, 55), T(12, 30),
     "Lun–vie; sáb solo mañana", "Sí", "", "", "", "", ""),
    ("RES-T", "Tarde RES", "RES", T(14, 0), T(22, 0), T(18, 55), T(19, 30),
     "Lun–vie", "Sí", "", "", "", "", ""),
    ("RES-N", "Noche RES", "RES", T(22, 0), T(6, 0), T(2, 55), T(3, 30),
     "Lun–vie", "Sí", "", "", "", "", ""),
    ("SOP-00", "SOP-00", "SOP", T(0, 0), T(6, 0), "", "",
     "Lun–dom", "Sí", "", "", "", "", ""),
    ("SOP-06", "SOP-06", "SOP", T(6, 0), T(12, 0), "", "",
     "Lun–dom", "Sí", "", "", "", "", ""),
    ("SOP-12", "SOP-12", "SOP", T(12, 0), T(18, 0), "", "",
     "Lun–dom", "Sí", "", "", "", "", ""),
    ("SOP-18", "SOP-18", "SOP", T(18, 0), T(0, 0), "", "",
     "Lun–dom", "Sí", "", "", "", "", ""),
]

# ---------------------------------------------------------------------------
# Productos (docs/catalogos-iniciales.md §6-§7)
# ---------------------------------------------------------------------------
PRODUCTOS_HEADERS = ["codigo", "producto", "formato", "volumen_litros",
                     "unidades_por_pack", "packs_por_pallet", "activo"]
PRODUCTOS_ROWS = [
    ("500SG", "500 sin gas", "botella", 0.5, 12, 150, "Sí"),
    ("500CG", "500 con gas", "botella", 0.5, 12, 132, "Sí"),
    ("1500SG", "1500 sin gas", "botella", 1.5, 6, 100, "Sí"),
    ("2000SG", "2000 sin gas", "botella", 2.0, 6, 80, "Sí"),
    ("BID6", "Bidón", "bidón", 6.0, 1, 168, "Sí"),
]

VELOCIDADES_HEADERS = ["codigo", "producto", "velocidad", "vigencia_desde",
                       "origen_dato", "activo"]
VELOCIDADES_ROWS = [
    ("VEL-01", "500SG", 90, D(2026, 1, 1), "ingeniero", "Sí"),
    ("VEL-02", "500CG", 70, D(2026, 1, 1), "ingeniero", "Sí"),
    ("VEL-03", "1500SG", 70, D(2026, 1, 1), "ingeniero", "Sí"),
    ("VEL-04", "2000SG", 70, D(2026, 1, 1), "ingeniero", "Sí"),
    ("VEL-05", "BID6", 10, D(2026, 1, 1), "ingeniero", "Sí"),
]

# ---------------------------------------------------------------------------
# Calidad (docs/catalogos-iniciales.md §8-§9)
# ---------------------------------------------------------------------------
SECCIONES_HEADERS = ["codigo", "seccion", "activo"]
SECCIONES_ROWS = [
    ("SE-01", "sopladora", "Sí"),
    ("SE-02", "llenadora", "Sí"),
    ("SE-03", "etiquetadora", "Sí"),
    ("SE-04", "forradora", "Sí"),
]

MOT_RECH_HEADERS = ["codigo", "motivo", "seccion", "activo"]
MOT_RECH_ROWS = [
    ("MR-01", "botella mal formada o mal tapada", "SE-02", "Sí"),
    ("MR-02", "botella aplastada por trabado de máquina", "SE-02", "Sí"),
    ("MR-03", "etiqueta mal puesta que no se puede recuperar", "SE-03", "Sí"),
    ("MR-04", "pico sucio de pegamento", "SE-03", "Sí"),
    ("MR-05", "rotura de envase", "SE-01", "Sí"),
]

# ---------------------------------------------------------------------------
# Insumos y personal (docs/catalogos-iniciales.md §10-§11)
# ---------------------------------------------------------------------------
INSUMOS_HEADERS = ["codigo", "insumo", "unidad_medida", "motivo", "activo"]
INSUMOS_ROWS = [
    ("IN-01", "preformas", "unidades", "", "Sí"),
    ("IN-02", "tapas", "unidades", "", "Sí"),
    ("IN-03", "etiquetas", "peso", "", "Sí"),
    ("IN-04", "termocontraíble", "peso", "", "Sí"),
    ("IN-05", "film stretch", "peso", "", "Sí"),
]

PUESTOS_HEADERS = ["codigo", "puesto", "equipo", "activo"]
PUESTOS_ROWS = [
    ("PU-01", "soplador", "SOP", "Sí"),
    ("PU-02", "ayudante", "SOP", "Sí"),
    ("PU-03", "líder", "RES", "Sí"),
    ("PU-04", "llenador", "RES", "Sí"),
    ("PU-05", "visor", "RES", "Sí"),
    ("PU-06", "etiquetador", "RES", "Sí"),
    ("PU-07", "forradora (final de línea)", "RES", "Sí"),
    ("PU-08", "estibador", "RES", "Sí"),
    ("PU-09", "autoelevadorista", "RES", "Sí"),
]

# ---------------------------------------------------------------------------
# Hojas de datos (docs/estructura-datos.md §2.12-§2.19): encabezados, sin datos
# ---------------------------------------------------------------------------
TABLAS = [
    ("01_Parámetros", "tblParametros", PARAM_HEADERS, PARAM_ROWS),
    ("02_Cat_Paradas", "tblTiposParada", TIPOS_HEADERS, TIPOS_ROWS),
    ("02_Cat_Paradas", "tblOrigenes", ORIGENES_HEADERS, ORIGENES_ROWS),
    ("02_Cat_Paradas", "tblMotivos", MOTIVOS_HEADERS, MOTIVOS_ROWS),
    ("03_Cat_Turnos", "tblTurnos", TURNOS_HEADERS, TURNOS_ROWS),
    ("04_Cat_Productos", "tblProductos", PRODUCTOS_HEADERS, PRODUCTOS_ROWS),
    ("04_Cat_Productos", "tblVelocidades", VELOCIDADES_HEADERS, VELOCIDADES_ROWS),
    ("05_Cat_Calidad", "tblSeccionesRechazo", SECCIONES_HEADERS, SECCIONES_ROWS),
    ("05_Cat_Calidad", "tblMotivosRechazo", MOT_RECH_HEADERS, MOT_RECH_ROWS),
    ("06_Cat_Insumos", "tblInsumos", INSUMOS_HEADERS, INSUMOS_ROWS),
    ("07_Cat_Personal", "tblPuestos", PUESTOS_HEADERS, PUESTOS_ROWS),
    ("10_Lotes", "tblLotes",
     ["id_lote", "numero_lote", "producto", "fecha_hora_inicio", "fecha_hora_fin"], []),
    ("11_Turnos", "tblTurnosRegistro",
     ["id_turno", "id_turno_lote", "id_lote", "fecha", "turno_codigo", "equipo",
      "fecha_hora_inicio", "fecha_hora_fin"], []),
    ("12_Paradas", "tblParadas",
     ["id", "fecha", "turno", "linea_id", "id_lote", "id_turno_lote",
      "fecha_hora_inicio", "fecha_hora_fin", "duracion_min", "origen", "motivo",
      "tipo", "codigo", "observaciones"], []),
    ("13_Produccion", "tblProduccion",
     ["fecha", "turno", "linea_id", "lote", "id_turno_lote", "producto",
      "unidades_producidas", "pallets", "litros", "velocidad_nominal"], []),
    ("14_Rechazos", "tblRechazos",
     ["fecha", "turno", "lote", "id_turno_lote", "seccion", "motivo", "cantidad"], []),
    ("15_Reducciones", "tblReducciones",
     ["fecha", "turno", "id_turno_lote", "origen", "motivo",
      "fecha_hora_inicio", "fecha_hora_fin", "velocidad_real"], []),
    ("16_PerdidasInsumo", "tblPerdidasInsumo",
     ["fecha", "turno", "insumo", "cantidad", "unidad", "motivo"], []),
    ("17_PersonalTurno", "tblPersonalTurno",
     ["fecha", "turno", "puesto", "persona"], []),
]


def normalizar(valor):
    """Convierte un valor de datos a un valor escribible en una celda."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, dt.datetime):
        return valor
    if isinstance(valor, dt.date):
        return dt.datetime(valor.year, valor.month, valor.day)
    return valor


def agregar_validacion_sino(rango, separador):
    # Posicional: con DispatchEx (late binding) los kwargs de parámetros no se resuelven
    rango.Validation.Add(
        XL_VALIDATE_LIST,
        XL_VALID_ALERT_STOP,
        XL_BETWEEN,
        '"Sí{0}No"'.format(separador),
    )
    rango.Validation.IgnoreBlank = True
    rango.Validation.InCellDropdown = True


def agregar_validacion_fecha(rango):
    # Excel acepta límites de fecha solo en formato local (d/m/aaaa en es-AR)
    for fecha_ini, fecha_fin in (("1/1/1900", "31/12/2199"),
                                 ("1/1/1900", "12/31/2199")):
        try:
            rango.Validation.Add(
                XL_VALIDATE_DATE,
                XL_VALID_ALERT_STOP,
                XL_BETWEEN,
                fecha_ini,
                fecha_fin,
            )
            break
        except Exception:
            continue
    rango.Validation.IgnoreBlank = True


def escribir_tabla(ws, col_inicial, nombre, headers, rows, separador):
    """Escribe encabezados+filas, crea la Table y aplica formatos/validaciones.

    Devuelve la columna desde donde empezar la próxima tabla de la hoja.
    """
    ncols = len(headers)
    nrows = max(len(rows), 1)  # tablas vacías llevan una fila en blanco

    for j, encabezado in enumerate(headers):
        ws.Cells(1, col_inicial + j).Value = encabezado

    for i, fila in enumerate(rows):
        if len(fila) != ncols:
            raise ValueError(
                "{0}: fila {1} tiene {2} valores, se esperaban {3}".format(
                    nombre, i + 1, len(fila), ncols))
        for j, valor in enumerate(fila):
            valor = normalizar(valor)
            if valor is not None:
                ws.Cells(2 + i, col_inicial + j).Value = valor

    rango_tabla = ws.Range(
        ws.Cells(1, col_inicial),
        ws.Cells(1 + nrows, col_inicial + ncols - 1),
    )
    tabla = ws.ListObjects.Add(XL_SRC_RANGE, rango_tabla, 0, XL_HAS_HEADERS)
    tabla.Name = nombre

    ultima_fila = 1 + nrows + BUFFER_FILAS
    for j, encabezado in enumerate(headers):
        col = col_inicial + j
        rango_col = ws.Range(ws.Cells(2, col), ws.Cells(ultima_fila, col))
        if encabezado in COL_SI_NO:
            agregar_validacion_sino(rango_col, separador)
        if encabezado in COL_FECHA:
            rango_col.NumberFormat = FORMATO_FECHA
            agregar_validacion_fecha(rango_col)
        elif encabezado in COL_FECHA_HORA:
            rango_col.NumberFormat = FORMATO_FECHA_HORA
            agregar_validacion_fecha(rango_col)
        elif encabezado in COL_HORA:
            rango_col.NumberFormat = FORMATO_HORA

    return col_inicial + ncols + 2  # deja una columna en blanco de separación


def escribir_portada(ws, hoy):
    ws.Cells(1, 1).Value = "Libro maestro OEE - Línea 1 - Embotellado (supuesto)"
    ws.Cells(2, 1).Value = "Versión"
    ws.Cells(2, 2).Value = "v{0}".format(VERSION)
    ws.Cells(3, 1).Value = "Generado"
    ws.Cells(3, 2).Value = D(hoy.year, hoy.month, hoy.day)
    ws.Cells(3, 2).NumberFormat = FORMATO_FECHA
    ws.Cells(4, 1).Value = "Fuente"
    ws.Cells(4, 2).Value = "doc-oficial.md v0.5 (2026-10-29)"
    ws.Cells(5, 1).Value = "Generador"
    ws.Cells(5, 2).Value = "tools/crear_libro.py"


def main():
    hoy = dt.date.today()
    nombre_archivo = "oee_v0.5_{0}.xlsm".format(hoy.isoformat())
    ruta = os.path.join(DIR_LIBRO, nombre_archivo)
    os.makedirs(DIR_LIBRO, exist_ok=True)

    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            sys.exit("No se pudo iniciar Excel COM: {0}".format(exc))

        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False

        # Separador de lista regional: con DispatchEx International devuelve la
        # tupla completa del enum (xlListSeparator=5 -> índice 4).
        try:
            intl = app.International
            separador = intl[4] if isinstance(intl, tuple) else app.International(5)
        except Exception:
            separador = ","
        if separador not in (",", ";"):
            separador = ","

        wb = app.Workbooks.Add()
        while wb.Sheets.Count > 1:
            wb.Sheets(wb.Sheets.Count).Delete()
        wb.Sheets(1).Name = HOJAS[0]
        for nombre_hoja in HOJAS[1:]:
            # Add(Before, After): posicional; kwargs no se resuelven con DispatchEx
            ws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
            ws.Name = nombre_hoja

        escribir_portada(wb.Sheets("00_Inicio"), hoy)

        proxima_col = {}
        for hoja, nombre, headers, rows in TABLAS:
            ws = wb.Sheets(hoja)
            col = proxima_col.get(hoja, 1)
            proxima_col[hoja] = escribir_tabla(
                ws, col, nombre, headers, rows, separador)

        # Ancho de columnas según contenido (sin celdas combinadas)
        for nombre_hoja in HOJAS:
            usado = wb.Sheets(nombre_hoja).UsedRange
            if usado is not None and usado.Value is not None:
                usado.EntireColumn.AutoFit()

        wb.SaveAs(ruta, FileFormat=XL_OPENXML_WORKBOOK_MACROENABLED)

        print("Libro creado: {0}".format(ruta))
        print("Hojas: {0} | Tablas: {1}".format(len(HOJAS), len(TABLAS)))
        print("Separador de lista usado: '{0}'".format(separador))
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        wb = None
        app = None
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
