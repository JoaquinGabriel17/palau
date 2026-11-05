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

        wb = excel.Workbooks.Open(ruta, 0, False)

        ws = wb.Worksheets("90_Tests")

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
