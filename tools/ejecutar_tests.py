#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ejecuta la suite VBA (Tests.ReportePruebas) y cuenta los resultados.

Uso:
    python tools/ejecutar_tests.py [ruta/al/libro.xlsm]

Abre el libro con Excel COM (pywin32), corre Tests.ReportePruebas (sin
cuadros de dialogo), lee el reporte de la hoja 90_Tests y cuenta cuantos
casos PASA/FALLA. Sale con exit code 1 si hay FALLA.

Guarda el libro al final para que el reporte de 90_Tests quede persistido
en el estado devuelto.
"""

import os
import sys

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover
    sys.exit("Falta pywin32. Instalar con: python -m pip install pywin32")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO = os.path.join(RAIZ, "libro", "oee_v0.2_2026-11-04.xlsm")

FILA_INICIO = 6
MAX_FILAS = 500

DIR_BAS = os.path.join(RAIZ, "src", "vba")
MOD_BASE = ["Calculos", "Modulo_Estructura", "Utils"]


def reimportar_modulos(excel, wb):
    """Reimporta los .bas de src/vba en la sesion activa.

    Tras guardar/reabrir, el proyecto VBA puede quedar en un estado donde
    excel.Run devuelve ' macros deshabilitadas'. Borrar los modulos estandar
    e importarlos de nuevo en la misma sesion (base primero, con un Run de
    CompletarEstructura en el medio) restablece el estado.
    """
    nombres = sorted(f[:-4] for f in os.listdir(DIR_BAS)
                     if f.lower().endswith(".bas"))
    base = [n for n in MOD_BASE if n in nombres]
    resto = [n for n in nombres if n not in base]
    proyecto = wb.VBProject
    for componente in list(proyecto.VBComponents):
        if componente.Type == 1:
            proyecto.VBComponents.Remove(componente)
    for nombre in base:
        proyecto.VBComponents.Import(os.path.join(DIR_BAS, nombre + ".bas"))
    excel.Run("Modulo_Estructura.CompletarEstructura", True)
    for nombre in resto:
        proyecto.VBComponents.Import(os.path.join(DIR_BAS, nombre + ".bas"))
    excel.Run("Modulo_Estructura.CompletarEstructura", True)


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else ARCHIVO
    ruta = os.path.abspath(ruta)
    if not os.path.exists(ruta):
        sys.exit("ERROR: No existe el libro {}".format(ruta))

    bloqueo = os.path.join(os.path.dirname(ruta), "~$" + os.path.basename(ruta))
    if os.path.exists(bloqueo):
        sys.exit("ERROR: El libro parece estar abierto en Excel ({}). "
                 "Cerrarlo y volver a intentar.".format(bloqueo))

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    wb = None
    casos = []
    global_ok = False
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityForceEnable

        wb = excel.Workbooks.Open(ruta, 0, False)

        ws = wb.Worksheets("90_Tests")

        # Sonda: si un Run simple falla, reimportar modulos en esta sesion.
        try:
            excel.Run("Utils.Tabla", "tblParametros")
        except Exception:
            print("AVISO: proyecto VBA inactivo; reimportando modulos...")
            reimportar_modulos(excel, wb)
            excel.Run("Utils.Tabla", "tblParametros")

        # ReportePruebas reescribe la hoja; limpiar para no arrastrar restos.
        ws.Range("A1:J200").Clear()
        excel.Run("Tests.ReportePruebas")

        # A3 = "Resultado global: PASA|FALLA"
        global_ok = "PASA" in str(ws.Cells(3, 1).Value or "")

        fila = FILA_INICIO
        while fila < MAX_FILAS + FILA_INICIO:
            caso = ws.Cells(fila, 1).Value
            metrica = ws.Cells(fila, 2).Value
            resultado = ws.Cells(fila, 7).Value
            if not caso:
                break
            casos.append((str(caso), str(metrica or ""), str(resultado or "")))
            fila += 1

        wb.Save()
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        pythoncom.CoUninitialize()

    pasan = [c for c in casos if c[2].upper() == "PASA"]
    fallan = [c for c in casos if c[2].upper() != "PASA"]

    print("Total casos reportados: {}".format(len(casos)))
    print("PASA: {}".format(len(pasan)))
    print("FALLA: {}".format(len(fallan)))
    print("Resultado global: {}".format("PASA" if global_ok else "FALLA"))

    for caso, metrica, _ in fallan:
        print("  FALLA: {} / {}".format(caso, metrica))

    if fallan:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
