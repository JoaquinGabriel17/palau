#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Verifica la compatibilidad de columnas entre el libro .xlsm y el VBA.

Abre `libro/oee_master.xlsm` (o el libro legado v0.2 si el maestro no
existe) con openpyxl y comprueba que las columnas de `tblVelocidades`,
`tblTurnosRegistro`, `tblParadas` y `tblTramosSinEnvase` coincidan
exactamente (orden incluido) con las que espera el VBA segun
`docs/estructura-datos.md`.

Uso:
    python tools/test_compatibilidad.py [ruta/al/libro.xlsm]
"""

import os
import sys
import warnings

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("Falta openpyxl. Instalar con: python -m pip install openpyxl")

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO_POR_DEFECTO = os.path.join(RAIZ, "libro", "oee_master.xlsm")
ARCHIVO_LEGADO = os.path.join(RAIZ, "libro", "oee_v0.2_2026-11-04.xlsm")

# Encabezados exactos de docs/estructura-datos.md (secciones 2.7 y 2.13).
COLUMNAS_ESPERADAS = {
    "tblVelocidades": [
        "codigo",
        "producto",
        "condicion_envase",
        "velocidad",
        "unidad",
        "vigencia_desde",
        "origen_dato",
        "activo",
    ],
    "tblTurnosRegistro": [
        "id_turno",
        "id_turno_lote",
        "id_lote",
        "nro_reporte",
        "fecha",
        "turno_codigo",
        "equipo",
        "encargado",
        "estado_revision",
        "anulado",
        "notas_sin_horario",
        "observaciones",
        "velocidad_informada",
        "fecha_hora_inicio",
        "fecha_hora_fin",
        "control_calidad",
        "nro_registro_sopladora",
        "registros_bolsones",
    ],
    # Seccion 2.14 del Doc (la migracion F2.1 agrega `interpretada`).
    "tblParadas": [
        "id",
        "fecha",
        "turno",
        "linea_id",
        "id_lote",
        "id_turno_lote",
        "fecha_hora_inicio",
        "fecha_hora_fin",
        "duracion_min",
        "origen",
        "motivo",
        "tipo",
        "codigo",
        "interpretada",
        "observaciones",
    ],
    # Seccion 2.20 del Doc.
    "tblTramosSinEnvase": [
        "id_turno_lote",
        "fecha_hora_inicio",
        "fecha_hora_fin",
    ],
    # Seccion 2.17 del Doc (la migracion F2a agrega `id`).
    "tblReducciones": [
        "id",
        "fecha",
        "turno",
        "id_turno_lote",
        "origen",
        "motivo",
        "fecha_hora_inicio",
        "fecha_hora_fin",
        "velocidad_real",
    ],
}


def resolver_archivo(argumento):
    if argumento:
        return os.path.abspath(argumento)
    if os.path.exists(ARCHIVO_POR_DEFECTO):
        return ARCHIVO_POR_DEFECTO
    if os.path.exists(ARCHIVO_LEGADO):
        return ARCHIVO_LEGADO
    candidato = os.path.join(os.getcwd(), "oee_master.xlsm")
    if os.path.exists(candidato):
        return candidato
    return ARCHIVO_POR_DEFECTO


def buscar_tabla(workbook, nombre):
    for hoja in workbook.worksheets:
        if nombre in hoja.tables:
            return hoja.tables[nombre]
    return None


def verificar_tabla(workbook, nombre, esperadas):
    tabla = buscar_tabla(workbook, nombre)
    if tabla is None:
        print("ERROR: Tabla no encontrada {}".format(nombre))
        return False

    actuales = [columna.name for columna in tabla.tableColumns]
    ok = True

    for columna in esperadas:
        if columna not in actuales:
            print("ERROR: Columna faltante {}".format(columna))
            ok = False

    for columna in actuales:
        if columna not in esperadas:
            print("ERROR: Columna inesperada {}".format(columna))
            ok = False

    esperadas_presentes = [c for c in esperadas if c in actuales]
    actuales_presentes = [c for c in actuales if c in esperadas]
    if ok and esperadas_presentes != actuales_presentes:
        print("ERROR: Orden de columnas distinto en {}".format(nombre))
        print("  Esperado:  {}".format(", ".join(esperadas)))
        print("  Encontrado: {}".format(", ".join(actuales)))
        ok = False

    return ok


def main():
    ruta = resolver_archivo(sys.argv[1] if len(sys.argv) > 1 else None)
    if not os.path.exists(ruta):
        print("ERROR: No se encontro el libro {}".format(ruta))
        return 1

    workbook = openpyxl.load_workbook(
        ruta, keep_vba=True, read_only=False, data_only=True)

    ok = True
    for nombre, esperadas in COLUMNAS_ESPERADAS.items():
        if not verificar_tabla(workbook, nombre, esperadas):
            ok = False

    if ok:
        print("COMPATIBILIDAD OK")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
