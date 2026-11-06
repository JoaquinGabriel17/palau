#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sincroniza el libro .xlsm con los modulos de src/vba.

Uso:
    python tools/sincronizar_libro.py [ruta/al/libro.xlsm]

Abre el libro con Excel COM (pywin32) y:
    1. Elimina los modulos estandar (.bas) presentes y re-importa los de
       src/vba (fuente de verdad), hasta la Version 0.2 del libro.
    2. Ejecuta Modulo_Estructura.CompletarEstructura (silencioso) para
       aplicar la migracion de estructura (ej. columna `interpretada` en
       tblParadas).
    3. Guarda y cierra el libro.

Requiere "Confiar en el acceso al modelo de objetos de proyectos de VBA";
el script habilita la clave AccessVBOM del registro por el usuario.
"""

import os
import re
import sys
import winreg

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover
    sys.exit("Falta pywin32. Instalar con: python -m pip install pywin32")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR_BAS = os.path.join(RAIZ, "src", "vba")
ARCHIVO = os.path.join(RAIZ, "libro", "oee_v0.2_2026-11-04.xlsm")


def habilitar_vbom():
    """Habilita AccessVBOM (y macros) para todas las versiones de Office."""
    base = r"Software\Microsoft\Office"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base) as k:
            versiones = [winreg.EnumKey(k, i)
                         for i in range(winreg.QueryInfoKey(k)[0])]
    except OSError:
        return
    for version in versiones:
        if not re.match(r"^\d+\.\d$", version):
            continue
        if int(version.split(".")[0]) < 12:
            continue
        clave = r"{}\Excel\Security".format(version)
        try:
            with winreg.CreateKeyEx(
                    winreg.HKEY_CURRENT_USER, clave, 0,
                    winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as sk:
                winreg.SetValueEx(sk, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(sk, "VBAWarnings", 0, winreg.REG_DWORD, 1)
        except OSError as exc:
            print("AVISO: no se pudo escribir el registro ({}): {}".format(
                clave, exc))


def modulos_bas():
    return sorted(nombres[:-4] for nombres in os.listdir(DIR_BAS)
                  if nombres.lower().endswith(".bas"))


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else ARCHIVO
    ruta = os.path.abspath(ruta)
    if not os.path.exists(ruta):
        sys.exit("ERROR: No existe el libro {}".format(ruta))

    bloqueo = os.path.join(os.path.dirname(ruta), "~$" + os.path.basename(ruta))
    if os.path.exists(bloqueo):
        sys.exit("ERROR: El libro parece estar abierto en Excel ({}). "
                 "Cerrarlo y volver a intentar.".format(bloqueo))

    nombres = modulos_bas()
    if not nombres:
        sys.exit("ERROR: No hay modulos .bas en {}".format(DIR_BAS))

    habilitar_vbom()

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    wb = None
    ok = False
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityForceEnable

        wb = excel.Workbooks.Open(ruta, 0, False)

        proyecto = wb.VBProject
        componentes = list(proyecto.VBComponents)
        for componente in componentes:
            if componente.Type == 1 and componente.Name in nombres:
                proyecto.VBComponents.Remove(componente)

        # Nota (hallazgo del 2026-11-07): quitar modulos y/o importar todos
        # en un solo lote deja el proyecto VBA en un estado donde excel.Run
        # falla con " macros deshabilitadas" (sin dialogo visible). El flujo
        # estable es: quitar -> guardar/cerrar/reabrir -> importar la base
        # (Calculos, Modulo_Estructura, Utils) -> correr CompletarEstructura
        # -> importar el resto -> correrla de nuevo -> guardar.
        for componente in list(proyecto.VBComponents):
            if componente.Type == 1 and componente.Name in nombres:
                proyecto.VBComponents.Remove(componente)
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = excel.Workbooks.Open(ruta, 0, False)

        proyecto = wb.VBProject
        base = [n for n in nombres if n in ("Calculos", "Modulo_Estructura", "Utils")]
        resto = [n for n in nombres if n not in base]
        for nombre in base:
            proyecto.VBComponents.Import(os.path.join(DIR_BAS, nombre + ".bas"))
        excel.Run("Modulo_Estructura.CompletarEstructura", True)
        for nombre in resto:
            proyecto.VBComponents.Import(os.path.join(DIR_BAS, nombre + ".bas"))
        excel.Run("Modulo_Estructura.CompletarEstructura", True)
        print("Modulos importados: {}".format(", ".join(nombres)))

        wb.Save()
        ok = True
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        pythoncom.CoUninitialize()

    if not ok:
        sys.exit(1)
    print("SINCRONIZACION OK: {}".format(ruta))
    return 0


if __name__ == "__main__":
    sys.exit(main())
