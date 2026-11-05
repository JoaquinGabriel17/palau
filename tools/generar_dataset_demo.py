#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera el dataset de demo de F2a (Bloque 2) en data/ejemplo/.

Cada CSV es un TURNO completo (uno o mas reportes). Esquema fijo de columnas;
la primera columna (`seccion`) clasifica cada fila:
    REPORTE | PARADA | PRODUCCION | RECHAZO | TRAMO | REDUCCION.

Reglas del dataset (checklist F2a Bloque 2):
- El caso 498 real (sin prefijo), datos exactos del Doc 7.1/7.2: OEE 61,4%.
- Casos 1 a 4 de la seccion 7.4 replicados con fechas de noviembre 2026.
- Los 5 productos, turnos RES y SOP, un cruce de medianoche y un reporte en
  el cambio de semana (domingo 08/11 SOP-18 hasta el lunes 09/11).
- Un turno con dos reportes (cambio de producto), tramos sin envase, microparadas,
  paradas largas, una parada interpretada, rechazos con y sin seccion/motivo,
  un turno sin paradas, un reporte anulado y 4 reducciones de velocidad.
- Todo lo ficticio: prefijo DEMO- en nro_reporte y "DATO DE DEMO" en observaciones.
"""

import csv
import os

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DESTINO = os.path.join(RAIZ, "data", "ejemplo")

ENCABEZADO = [
    "seccion", "nro_reporte", "fecha_reporte", "turno", "equipo", "encargado",
    "lote", "producto", "horario_inicio", "horario_fin", "anulado",
    "notas_sin_horario", "observaciones", "velocidad_informada",
    "inicio", "fin", "origen", "motivo", "tipo_parada", "interpretada",
    "rechazo_seccion", "rechazo_motivo", "cantidad", "detalle",
    "unidades", "pallets", "velocidad",
]


def f(anio, mes, dia):
    return "{0:02d}/{1:02d}/{2}".format(dia, mes, anio)


def fh(anio, mes, dia, hora, minuto):
    return "{0:02d}/{1:02d}/{2} {3:02d}:{4:02d}".format(dia, mes, anio, hora, minuto)


def fila(**campos):
    """Devuelve un dict con solo las claves del esquema que vengan en campos."""
    return {clave: campos.get(clave, "") for clave in ENCABEZADO}


def seccion(nombre, **campos):
    return fila(seccion=nombre, **campos)


# ---------------------------------------------------------------------------
# Dataset
# ---------------------------------------------------------------------------

DATASET = {}

# 1) Caso real 498 (Doc 7.1/7.2, D-01: OEE 61,4%). Sin prefijo ni marcas DEMO.
DATASET["498.csv"] = [
    seccion("REPORTE", nro_reporte="498", fecha_reporte=f(2026, 9, 28),
            turno="RES-T", equipo="RES", lote="09261", producto="1500SG",
            horario_inicio=fh(2026, 9, 28, 14, 0), horario_fin=fh(2026, 9, 28, 22, 0),
            notas_sin_horario="pequenas paradas por falta de temperatura y problemas en el loose de pack",
            velocidad_informada=70),
    # Paradas de la seccion 7.1 (tipos segun D-01 / Q-19):
    seccion("PARADA", inicio=fh(2026, 9, 28, 14, 10), fin=fh(2026, 9, 28, 14, 50),
            origen="OR-05", motivo="temperatura alta", tipo_parada="TP-07"),
    seccion("PARADA", inicio=fh(2026, 9, 28, 15, 40), fin=fh(2026, 9, 28, 15, 55),
            origen="OR-03", motivo="limpieza", tipo_parada="TP-04"),
    seccion("PARADA", inicio=fh(2026, 9, 28, 16, 50), fin=fh(2026, 9, 28, 17, 0),
            origen="OR-03", motivo="cambio de rollo y limpieza", tipo_parada="TP-04"),
    seccion("PARADA", inicio=fh(2026, 9, 28, 18, 20), fin=fh(2026, 9, 28, 18, 30),
            origen="OR-03", motivo="limpieza", tipo_parada="TP-04"),
    seccion("PARADA", inicio=fh(2026, 9, 28, 18, 55), fin=fh(2026, 9, 28, 19, 30),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    # Ozonizador: parada interpretada (marcada como interpretada = Si).
    seccion("PARADA", inicio=fh(2026, 9, 28, 21, 0), fin=fh(2026, 9, 28, 21, 10),
            origen="OR-06", motivo="se espera 10 minutos", tipo_parada="TP-04",
            interpretada="Si"),
    seccion("PARADA", inicio=fh(2026, 9, 28, 21, 30), fin=fh(2026, 9, 28, 22, 0),
            origen="OR-01", motivo="falla en horno", tipo_parada="TP-03"),
    # Produccion: 29 pallets = 17.400 buenas + 4 rechazadas (papel no indica
    # seccion ni motivo: entra por AVISA quedando pendiente de revisar).
    seccion("PRODUCCION", unidades=17400, pallets=29),
    seccion("RECHAZO", cantidad=4),
    # Baja tension 20:40-21:10: reduccion de velocidad (Q-18), no parada.
    seccion("REDUCCION", inicio=fh(2026, 9, 28, 20, 40), fin=fh(2026, 9, 28, 21, 10),
            origen="OR-13", motivo="falla de compresor", velocidad=45),
]

# 2) Caso 1 de 7.4: turno sin produccion (DEMO-500).
DATASET["DEMO-500.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-500", fecha_reporte=f(2026, 11, 11),
            turno="RES-M", equipo="RES", lote="11012", producto="500SG",
            horario_inicio=fh(2026, 11, 11, 6, 0), horario_fin=fh(2026, 11, 11, 14, 0),
            observaciones="DATO DE DEMO - caso 7.4 turno sin produccion"),
]

# 3) Caso 2 de 7.4: cambio de producto en el mismo turno (DEMO-501 y DEMO-502,
#    mismo id_turno por fecha + RES-M + equipo RES).
DATASET["DEMO-501-502.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-501", fecha_reporte=f(2026, 11, 10),
            turno="RES-M", equipo="RES", lote="11001", producto="500SG",
            horario_inicio=fh(2026, 11, 10, 6, 0), horario_fin=fh(2026, 11, 10, 10, 0),
            observaciones="DATO DE DEMO - caso 7.4 cambio de producto lote A"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 9, 0), fin=fh(2026, 11, 10, 9, 25),
            origen="OR-02", motivo="averia llenadora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 8, 0), fin=fh(2026, 11, 10, 8, 2),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 8, 30), fin=fh(2026, 11, 10, 8, 32),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 9, 30), fin=fh(2026, 11, 10, 9, 32),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PRODUCCION", unidades=17100, pallets=0),
    seccion("RECHAZO", rechazo_seccion="SE-02", rechazo_motivo="MR-01", cantidad=40),
    seccion("REPORTE", nro_reporte="DEMO-502", fecha_reporte=f(2026, 11, 10),
            turno="RES-M", equipo="RES", lote="11002", producto="1500SG",
            horario_inicio=fh(2026, 11, 10, 10, 0), horario_fin=fh(2026, 11, 10, 14, 0),
            observaciones="DATO DE DEMO - caso 7.4 cambio de producto lote B"),
    # Cambio de formato = A y C (Q-30).
    seccion("PARADA", inicio=fh(2026, 11, 10, 10, 0), fin=fh(2026, 11, 10, 10, 30),
            origen="OR-09", motivo="cambio de formato", tipo_parada="TP-04"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 11, 55), fin=fh(2026, 11, 10, 12, 30),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 12, 45), fin=fh(2026, 11, 10, 12, 53),
            origen="OR-03", motivo="ajuste etiquetadora", tipo_parada="TP-04"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 13, 0), fin=fh(2026, 11, 10, 13, 4),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PRODUCCION", unidades=9600, pallets=0),
    # Rechazo con motivo pero sin seccion (AVISA).
    seccion("RECHAZO", rechazo_motivo="MR-03", cantidad=12),
]

# 4) Caso 3 de 7.4: turno nocturno con paradas que cruzan medianoche
#    (DEMO-503 y DEMO-504, lunes 09/11).
DATASET["DEMO-503-504.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-503", fecha_reporte=f(2026, 11, 9),
            turno="RES-N", equipo="RES", lote="11003", producto="1500SG",
            horario_inicio=fh(2026, 11, 9, 22, 0), horario_fin=fh(2026, 11, 10, 0, 0),
            observaciones="DATO DE DEMO - caso 7.4 cruce de medianoche lote A"),
    seccion("PARADA", inicio=fh(2026, 11, 9, 23, 0), fin=fh(2026, 11, 9, 23, 2),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PRODUCCION", unidades=4200, pallets=0),
    seccion("RECHAZO", rechazo_seccion="SE-02", rechazo_motivo="MR-01", cantidad=6),
    seccion("REPORTE", nro_reporte="DEMO-504", fecha_reporte=f(2026, 11, 9),
            turno="RES-N", equipo="RES", lote="11004", producto="1500SG",
            horario_inicio=fh(2026, 11, 10, 0, 0), horario_fin=fh(2026, 11, 10, 6, 0),
            observaciones="DATO DE DEMO - caso 7.4 cruce de medianoche lote B"),
    # Averia 23:40-00:25: queda en el turno del 09/11 (Q-04).
    seccion("PARADA", inicio=fh(2026, 11, 9, 23, 40), fin=fh(2026, 11, 10, 0, 25),
            origen="OR-02", motivo="averia llenadora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 1, 0), fin=fh(2026, 11, 10, 1, 4),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 1, 15), fin=fh(2026, 11, 10, 1, 50),
            origen="OR-07", motivo="falta de agua", tipo_parada="TP-07"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 2, 55), fin=fh(2026, 11, 10, 3, 30),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PARADA", inicio=fh(2026, 11, 10, 4, 10), fin=fh(2026, 11, 10, 4, 22),
            origen="OR-03", motivo="ajuste etiquetadora", tipo_parada="TP-04"),
    seccion("PRODUCCION", unidades=17400, pallets=0),
    seccion("RECHAZO", rechazo_seccion="SE-03", rechazo_motivo="MR-03", cantidad=24),
]

# 5) Caso 4 de 7.4: datos incompletos (DEMO-505). Parada 16:20-16:50 sin tipo
#    ni motivo (RN-09: PNP provisoria).
DATASET["DEMO-505.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-505", fecha_reporte=f(2026, 11, 12),
            turno="RES-T", equipo="RES", lote="11005", producto="2000SG",
            horario_inicio=fh(2026, 11, 12, 14, 0), horario_fin=fh(2026, 11, 12, 22, 0),
            notas_sin_horario="pequenas paradas por falta de temperatura",
            observaciones="DATO DE DEMO - caso 7.4 datos incompletos"),
    seccion("PARADA", inicio=fh(2026, 11, 12, 15, 10), fin=fh(2026, 11, 12, 15, 30),
            origen="OR-04", motivo="averia forradora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 12, 16, 20), fin=fh(2026, 11, 12, 16, 50)),
    seccion("PARADA", inicio=fh(2026, 11, 12, 18, 55), fin=fh(2026, 11, 12, 19, 30),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PRODUCCION", unidades=22080, pallets=46),
    seccion("RECHAZO", cantidad=8),
]

# 6) Caso 5 de 7.4: tramos con y sin envase (DEMO-506, un tramo).
DATASET["DEMO-506.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-506", fecha_reporte=f(2026, 11, 13),
            turno="RES-M", equipo="RES", lote="11006", producto="500SG",
            horario_inicio=fh(2026, 11, 13, 6, 0), horario_fin=fh(2026, 11, 13, 14, 0),
            observaciones="DATO DE DEMO - caso 7.4 tramos con y sin envase"),
    seccion("TRAMO", inicio=fh(2026, 11, 13, 10, 0), fin=fh(2026, 11, 13, 14, 0)),
    seccion("PARADA", inicio=fh(2026, 11, 13, 7, 0), fin=fh(2026, 11, 13, 7, 20),
            origen="OR-02", motivo="averia llenadora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 13, 8, 0), fin=fh(2026, 11, 13, 8, 3),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 13, 11, 0), fin=fh(2026, 11, 13, 11, 10),
            origen="OR-03", motivo="ajuste etiquetadora", tipo_parada="TP-04"),
    seccion("PARADA", inicio=fh(2026, 11, 13, 11, 55), fin=fh(2026, 11, 13, 12, 30),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    # Micro del tramo sin envase, corrida para no solapar el descanso.
    seccion("PARADA", inicio=fh(2026, 11, 13, 12, 45), fin=fh(2026, 11, 13, 12, 47),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PRODUCCION", unidades=22800, pallets=12),
    seccion("RECHAZO", cantidad=60),
]

# 7) Reporte con DOS tramos sin envase (DEMO-507).
DATASET["DEMO-507.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-507", fecha_reporte=f(2026, 11, 5),
            turno="RES-M", equipo="RES", lote="11007", producto="500SG",
            horario_inicio=fh(2026, 11, 5, 6, 0), horario_fin=fh(2026, 11, 5, 14, 0),
            observaciones="DATO DE DEMO - reporte con dos tramos sin envase"),
    seccion("TRAMO", inicio=fh(2026, 11, 5, 9, 0), fin=fh(2026, 11, 5, 10, 0)),
    seccion("TRAMO", inicio=fh(2026, 11, 5, 12, 0), fin=fh(2026, 11, 5, 13, 30)),
    seccion("PARADA", inicio=fh(2026, 11, 5, 7, 0), fin=fh(2026, 11, 5, 7, 15),
            origen="OR-02", motivo="averia llenadora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 5, 8, 0), fin=fh(2026, 11, 5, 8, 35),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PARADA", inicio=fh(2026, 11, 5, 10, 30), fin=fh(2026, 11, 5, 10, 33),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 5, 12, 30), fin=fh(2026, 11, 5, 12, 32),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 5, 13, 45), fin=fh(2026, 11, 5, 13, 50),
            origen="OR-03", motivo="ajuste etiquetadora", tipo_parada="TP-04"),
    seccion("PRODUCCION", unidades=16520, pallets=9),
    seccion("RECHAZO", rechazo_seccion="SE-01", rechazo_motivo="MR-05", cantidad=20),
]

# 8) Turno SOP con bidon 6 L (DEMO-508; producto BID6).
DATASET["DEMO-508.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-508", fecha_reporte=f(2026, 11, 4),
            turno="SOP-12", equipo="SOP", lote="11008", producto="BID6",
            horario_inicio=fh(2026, 11, 4, 12, 0), horario_fin=fh(2026, 11, 4, 18, 0),
            observaciones="DATO DE DEMO - turno SOP con bidon 6 litros"),
    seccion("PARADA", inicio=fh(2026, 11, 4, 13, 0), fin=fh(2026, 11, 4, 13, 40),
            origen="OR-02", motivo="averia llenadora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 4, 14, 30), fin=fh(2026, 11, 4, 14, 34),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 4, 16, 0), fin=fh(2026, 11, 4, 16, 35),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PRODUCCION", unidades=1690, pallets=10),
    seccion("RECHAZO", rechazo_seccion="SE-02", rechazo_motivo="MR-01", cantidad=10),
    seccion("REDUCCION", inicio=fh(2026, 11, 4, 15, 0), fin=fh(2026, 11, 4, 15, 20),
            origen="OR-06", motivo="presion de aire baja", velocidad=8),
]

# 9) Turno SOP con 500 con gas (DEMO-509).
DATASET["DEMO-509.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-509", fecha_reporte=f(2026, 11, 6),
            turno="SOP-06", equipo="SOP", lote="11009", producto="500CG",
            horario_inicio=fh(2026, 11, 6, 6, 0), horario_fin=fh(2026, 11, 6, 12, 0),
            observaciones="DATO DE DEMO - turno SOP con 500 con gas"),
    seccion("PARADA", inicio=fh(2026, 11, 6, 8, 10), fin=fh(2026, 11, 6, 8, 30),
            origen="OR-01", motivo="averia sopladora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 6, 9, 0), fin=fh(2026, 11, 6, 9, 2),
            origen="OR-03", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 6, 10, 0), fin=fh(2026, 11, 6, 10, 35),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PRODUCCION", unidades=19500, pallets=12),
    seccion("RECHAZO", rechazo_seccion="SE-02", rechazo_motivo="MR-01", cantidad=50),
    seccion("REDUCCION", inicio=fh(2026, 11, 6, 7, 0), fin=fh(2026, 11, 6, 7, 15),
            origen="OR-13", motivo="falta de personal", velocidad=55),
]

# 10) Domingo noche: cruce de medianoche y cambio de semana (DEMO-510,
#     domingo 08/11 hasta el lunes 09/11).
DATASET["DEMO-510.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-510", fecha_reporte=f(2026, 11, 8),
            turno="SOP-18", equipo="SOP", lote="11010", producto="500SG",
            horario_inicio=fh(2026, 11, 8, 18, 0), horario_fin=fh(2026, 11, 9, 0, 0),
            observaciones="DATO DE DEMO - domingo noche cambio de semana"),
    seccion("PARADA", inicio=fh(2026, 11, 8, 20, 0), fin=fh(2026, 11, 8, 20, 15),
            origen="OR-03", motivo="averia etiquetadora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 8, 21, 0), fin=fh(2026, 11, 8, 21, 2),
            origen="OR-01", motivo="micro", tipo_parada="TP-05"),
    seccion("PARADA", inicio=fh(2026, 11, 8, 22, 0), fin=fh(2026, 11, 8, 22, 35),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PRODUCCION", unidades=15000, pallets=8),
    seccion("RECHAZO", rechazo_seccion="SE-01", rechazo_motivo="MR-05", cantidad=30),
    seccion("REDUCCION", inicio=fh(2026, 11, 8, 21, 30), fin=fh(2026, 11, 8, 22, 0),
            origen="OR-13", motivo="espera de envase", velocidad=70),
]

# 11) Reporte anulado (DEMO-511): se carga completo y el loader lo anula.
DATASET["DEMO-511.csv"] = [
    seccion("REPORTE", nro_reporte="DEMO-511", fecha_reporte=f(2026, 11, 6),
            turno="RES-T", equipo="RES", lote="11011", producto="500SG",
            horario_inicio=fh(2026, 11, 6, 14, 0), horario_fin=fh(2026, 11, 6, 22, 0),
            anulado="Si",
            observaciones="DATO DE DEMO - reporte anulado de prueba"),
    seccion("PARADA", inicio=fh(2026, 11, 6, 15, 0), fin=fh(2026, 11, 6, 15, 20),
            origen="OR-01", motivo="averia sopladora", tipo_parada="TP-03"),
    seccion("PARADA", inicio=fh(2026, 11, 6, 18, 55), fin=fh(2026, 11, 6, 19, 30),
            origen="OR-08", motivo="descanso", tipo_parada="TP-01"),
    seccion("PRODUCCION", unidades=6000, pallets=3),
    seccion("RECHAZO", rechazo_seccion="SE-02", rechazo_motivo="MR-01", cantidad=10),
]


def main():
    if not os.path.isdir(DESTINO):
        os.makedirs(DESTINO)
    for nombre in os.listdir(DESTINO):
        if nombre.lower().endswith(".csv"):
            os.remove(os.path.join(DESTINO, nombre))

    total_reportes = 0
    total_reducciones = 0
    for nombre, filas in sorted(DATASET.items()):
        ruta = os.path.join(DESTINO, nombre)
        with open(ruta, "w", newline="") as archivo:
            escritor = csv.writer(archivo, lineterminator="\r\n")
            escritor.writerow(ENCABEZADO)
            for fila_actual in filas:
                escritor.writerow([fila_actual[clave] for clave in ENCABEZADO])
        n_rep = sum(1 for fila_actual in filas if fila_actual.get("seccion") == "REPORTE")
        n_red = sum(1 for fila_actual in filas if fila_actual.get("seccion") == "REDUCCION")
        total_reportes += n_rep
        total_reducciones += n_red
        print("Generado: {0} ({1} reportes)".format(nombre, n_rep))

    print("Total: {0} archivos, {1} reportes, {2} reducciones".format(
        len(DATASET), total_reportes, total_reducciones))


if __name__ == "__main__":
    main()
