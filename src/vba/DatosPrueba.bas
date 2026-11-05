Attribute VB_Name = "DatosPrueba"
Option Explicit

' Datos de los cinco casos oficiales de la Doc v0.8 (seccion 7).
' Se cargan con IDs que empiezan con TST- para poder borrarlos sin tocar
' los datos reales del libro.

Private Const PREFIJO As String = "TST-"

Public Sub PrepararDatosPrueba()
    LimpiarDatosPrueba
    PrepararVelocidades
    PrepararCaso498
    PrepararCaso2
    PrepararCaso3
    PrepararCaso4
    PrepararCaso5
    PrepararCasoRN15
End Sub

Public Sub LimpiarDatosPrueba()
    BorrarPorPrefijo "tblLotes", "id_lote", PREFIJO
    BorrarPorPrefijo "tblTurnosRegistro", "id_turno", PREFIJO
    BorrarPorPrefijo "tblParadas", "id", PREFIJO
    BorrarPorPrefijo "tblProduccion", "id_turno_lote", PREFIJO
    BorrarPorPrefijo "tblRechazos", "id_turno_lote", PREFIJO
    BorrarPorPrefijo "tblTramosSinEnvase", "id_turno_lote", PREFIJO
    BorrarPorPrefijo "tblVelocidades", "codigo", PREFIJO
    BorrarPorPrefijo "tblReducciones", "id_turno_lote", PREFIJO
    BorrarPorPrefijo "tblPerdidasInsumo", "turno", PREFIJO
    BorrarPorPrefijo "tblPersonalTurno", "turno", PREFIJO
    BorrarPorPrefijo "tblParametros", "id", PREFIJO
End Sub

Private Sub PrepararVelocidades()
    AgregarVelocidad "TST-VEL-500-CON", "500SG", "con envase", 90, "botellas/min", DateSerial(2026, 1, 1)
    AgregarVelocidad "TST-VEL-500-SIN", "500SG", "sin envase", 2500, "botellas/hora", DateSerial(2026, 1, 1)
    AgregarVelocidad "TST-VEL-1500-CON", "1500SG", "con envase", 70, "botellas/min", DateSerial(2026, 1, 1)
    AgregarVelocidad "TST-VEL-2000-CON", "2000SG", "con envase", 60, "botellas/min", DateSerial(2026, 1, 1)
End Sub

Private Sub PrepararCaso498()
    Dim f As Date
    f = DateSerial(2026, 9, 28)

    AgregarLote "TST-498-L", "TST9261", "1500SG", f + TimeSerial(14, 0, 0), f + TimeSerial(22, 0, 0)
    AgregarTurno "TST-498", "TST-498-TL", "TST-498-L", f, "RES-T", "RES", _
        f + TimeSerial(14, 0, 0), f + TimeSerial(22, 0, 0)

    AgregarParada "TST-498-P1", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(14, 10, 0), f + TimeSerial(14, 50, 0), 40, "OR-05", "temperatura alta", "TP-07"
    AgregarParada "TST-498-P2", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(15, 40, 0), f + TimeSerial(15, 55, 0), 15, "OR-03", "limpieza", "TP-04"
    AgregarParada "TST-498-P3", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(16, 50, 0), f + TimeSerial(17, 0, 0), 10, "OR-03", "cambio de rollo y limpieza", "TP-04"
    AgregarParada "TST-498-P4", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(18, 20, 0), f + TimeSerial(18, 30, 0), 10, "OR-03", "limpieza", "TP-04"
    AgregarParada "TST-498-P5", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(18, 55, 0), f + TimeSerial(19, 30, 0), 35, "OR-08", "descanso", "TP-01"
    AgregarParada "TST-498-P6", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(21, 0, 0), f + TimeSerial(21, 10, 0), 10, "OR-01", "se espera 10 minutos", "TP-04"
    AgregarParada "TST-498-P7", f, "TST-498", 1, "TST-498-L", "TST-498-TL", _
        f + TimeSerial(21, 30, 0), f + TimeSerial(22, 0, 0), 30, "OR-01", "falla en horno", "TP-03"

    AgregarProduccion f, "TST-498", 1, "TST-498-L", "TST-498-TL", "1500SG", 17400, 29
    AgregarRechazo f, "TST-498", "TST-498-L", "TST-498-TL", "SE-01", "MR-05", 4
End Sub

Private Sub PrepararCaso2()
    Dim f As Date
    f = DateSerial(2026, 10, 13)

    AgregarLote "TST-C2-LA", "TSTC2A", "500SG", f + TimeSerial(6, 0, 0), f + TimeSerial(10, 0, 0)
    AgregarLote "TST-C2-LB", "TSTC2B", "1500SG", f + TimeSerial(10, 0, 0), f + TimeSerial(14, 0, 0)
    AgregarTurno "TST-C2", "TST-C2-A", "TST-C2-LA", f, "RES-M", "RES", _
        f + TimeSerial(6, 0, 0), f + TimeSerial(10, 0, 0)
    AgregarTurno "TST-C2", "TST-C2-B", "TST-C2-LB", f, "RES-M", "RES", _
        f + TimeSerial(10, 0, 0), f + TimeSerial(14, 0, 0)

    AgregarParada "TST-C2-P1", f, "TST-C2", 1, "TST-C2-LA", "TST-C2-A", _
        f + TimeSerial(9, 0, 0), f + TimeSerial(9, 25, 0), 25, "OR-02", "averia llenadora", "TP-03"
    AgregarParada "TST-C2-P2", f, "TST-C2", 1, "TST-C2-LB", "TST-C2-B", _
        f + TimeSerial(10, 0, 0), f + TimeSerial(10, 30, 0), 30, "OR-09", "cambio de formato", "TP-04"
    AgregarParada "TST-C2-P3", f, "TST-C2", 1, "TST-C2-LB", "TST-C2-B", _
        f + TimeSerial(11, 55, 0), f + TimeSerial(12, 30, 0), 35, "OR-08", "descanso", "TP-01"
    AgregarParada "TST-C2-P4", f, "TST-C2", 1, "TST-C2-LB", "TST-C2-B", _
        f + TimeSerial(12, 45, 0), f + TimeSerial(12, 53, 0), 8, "OR-03", "ajuste etiquetadora", "TP-04"

    AgregarParada "TST-C2-M1", f, "TST-C2", 1, "TST-C2-LA", "TST-C2-A", _
        f + TimeSerial(8, 0, 0), f + TimeSerial(8, 2, 0), 2, "OR-01", "micro", "TP-05"
    AgregarParada "TST-C2-M2", f, "TST-C2", 1, "TST-C2-LA", "TST-C2-A", _
        f + TimeSerial(8, 30, 0), f + TimeSerial(8, 32, 0), 2, "OR-01", "micro", "TP-05"
    AgregarParada "TST-C2-M3", f, "TST-C2", 1, "TST-C2-LA", "TST-C2-A", _
        f + TimeSerial(9, 30, 0), f + TimeSerial(9, 32, 0), 2, "OR-01", "micro", "TP-05"
    AgregarParada "TST-C2-M4", f, "TST-C2", 1, "TST-C2-LB", "TST-C2-B", _
        f + TimeSerial(13, 0, 0), f + TimeSerial(13, 4, 0), 4, "OR-01", "micro", "TP-05"

    AgregarProduccion f, "TST-C2", 1, "TST-C2-LA", "TST-C2-A", "500SG", 17100, 0
    AgregarProduccion f, "TST-C2", 1, "TST-C2-LB", "TST-C2-B", "1500SG", 9600, 0
    AgregarRechazo f, "TST-C2", "TST-C2-LA", "TST-C2-A", "SE-02", "MR-01", 40
    AgregarRechazo f, "TST-C2", "TST-C2-LB", "TST-C2-B", "SE-03", "MR-03", 12
End Sub

Private Sub PrepararCaso3()
    Dim f1 As Date
    Dim f2 As Date
    f1 = DateSerial(2026, 10, 12)
    f2 = DateSerial(2026, 10, 13)

    AgregarLote "TST-C3-LA", "TSTC3A", "1500SG", f1 + TimeSerial(22, 0, 0), f2 + TimeSerial(0, 0, 0)
    AgregarLote "TST-C3-LB", "TSTC3B", "1500SG", f2 + TimeSerial(0, 0, 0), f2 + TimeSerial(6, 0, 0)
    AgregarTurno "TST-C3", "TST-C3-A", "TST-C3-LA", f1, "RES-N", "RES", _
        f1 + TimeSerial(22, 0, 0), f2 + TimeSerial(0, 0, 0)
    AgregarTurno "TST-C3", "TST-C3-B", "TST-C3-LB", f1, "RES-N", "RES", _
        f2 + TimeSerial(0, 0, 0), f2 + TimeSerial(6, 0, 0)

    AgregarParada "TST-C3-P1", f1, "TST-C3", 1, "TST-C3-LA", "TST-C3-A", _
        f1 + TimeSerial(23, 40, 0), f2 + TimeSerial(0, 25, 0), 45, "OR-02", "averia llenadora", "TP-03"
    AgregarParada "TST-C3-P2", f1, "TST-C3", 1, "TST-C3-LB", "TST-C3-B", _
        f2 + TimeSerial(1, 15, 0), f2 + TimeSerial(1, 50, 0), 35, "OR-07", "falta de agua", "TP-07"
    AgregarParada "TST-C3-P3", f1, "TST-C3", 1, "TST-C3-LB", "TST-C3-B", _
        f2 + TimeSerial(2, 55, 0), f2 + TimeSerial(3, 30, 0), 35, "OR-08", "descanso", "TP-01"
    AgregarParada "TST-C3-P4", f1, "TST-C3", 1, "TST-C3-LB", "TST-C3-B", _
        f2 + TimeSerial(4, 10, 0), f2 + TimeSerial(4, 22, 0), 12, "OR-03", "ajuste etiquetadora", "TP-04"

    AgregarParada "TST-C3-M1", f1, "TST-C3", 1, "TST-C3-LA", "TST-C3-A", _
        f1 + TimeSerial(23, 0, 0), f1 + TimeSerial(23, 2, 0), 2, "OR-01", "micro", "TP-05"
    AgregarParada "TST-C3-M2", f1, "TST-C3", 1, "TST-C3-LB", "TST-C3-B", _
        f2 + TimeSerial(1, 0, 0), f2 + TimeSerial(1, 4, 0), 4, "OR-01", "micro", "TP-05"

    AgregarProduccion f1, "TST-C3", 1, "TST-C3-LA", "TST-C3-A", "1500SG", 4200, 0
    AgregarProduccion f1, "TST-C3", 1, "TST-C3-LB", "TST-C3-B", "1500SG", 17400, 0
    AgregarRechazo f1, "TST-C3", "TST-C3-LA", "TST-C3-A", "SE-02", "MR-01", 6
    AgregarRechazo f1, "TST-C3", "TST-C3-LB", "TST-C3-B", "SE-03", "MR-03", 24
End Sub

Private Sub PrepararCaso4()
    Dim f As Date
    f = DateSerial(2026, 10, 14)

    AgregarLote "TST-C4-L", "TSTC4", "2000SG", f + TimeSerial(14, 0, 0), f + TimeSerial(22, 0, 0)
    AgregarTurno "TST-C4", "TST-C4-TL", "TST-C4-L", f, "RES-T", "RES", _
        f + TimeSerial(14, 0, 0), f + TimeSerial(22, 0, 0)

    AgregarParada "TST-C4-P1", f, "TST-C4", 1, "TST-C4-L", "TST-C4-TL", _
        f + TimeSerial(15, 10, 0), f + TimeSerial(15, 30, 0), 20, "OR-04", "averia forradora", "TP-03"
    AgregarParada "TST-C4-P2", f, "TST-C4", 1, "TST-C4-L", "TST-C4-TL", _
        f + TimeSerial(16, 20, 0), f + TimeSerial(16, 50, 0), 30, "", "", ""
    AgregarParada "TST-C4-P3", f, "TST-C4", 1, "TST-C4-L", "TST-C4-TL", _
        f + TimeSerial(18, 55, 0), f + TimeSerial(19, 30, 0), 35, "OR-08", "descanso", "TP-01"

    AgregarProduccion f, "TST-C4", 1, "TST-C4-L", "TST-C4-TL", "2000SG", 22080, 46
    AgregarRechazo f, "TST-C4", "TST-C4-L", "TST-C4-TL", "", "", 8
End Sub

Private Sub PrepararCaso5()
    Dim f As Date
    f = DateSerial(2026, 10, 15)

    AgregarLote "TST-C5-L", "TSTC5", "500SG", f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0)
    AgregarTurno "TST-C5", "TST-C5-TL", "TST-C5-L", f, "RES-M", "RES", _
        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0)
    AgregarTramo "TST-C5-TL", f + TimeSerial(10, 0, 0), f + TimeSerial(14, 0, 0)

    AgregarParada "TST-C5-P1", f, "TST-C5", 1, "TST-C5-L", "TST-C5-TL", _
        f + TimeSerial(7, 0, 0), f + TimeSerial(7, 20, 0), 20, "OR-02", "averia llenadora", "TP-03"
    AgregarParada "TST-C5-P2", f, "TST-C5", 1, "TST-C5-L", "TST-C5-TL", _
        f + TimeSerial(11, 0, 0), f + TimeSerial(11, 10, 0), 10, "OR-03", "ajuste etiquetadora", "TP-04"
    AgregarParada "TST-C5-P3", f, "TST-C5", 1, "TST-C5-L", "TST-C5-TL", _
        f + TimeSerial(11, 55, 0), f + TimeSerial(12, 30, 0), 35, "OR-08", "descanso", "TP-01"

    AgregarParada "TST-C5-M1", f, "TST-C5", 1, "TST-C5-L", "TST-C5-TL", _
        f + TimeSerial(8, 0, 0), f + TimeSerial(8, 3, 0), 3, "OR-01", "micro", "TP-05"
    AgregarParada "TST-C5-M2", f, "TST-C5", 1, "TST-C5-L", "TST-C5-TL", _
        f + TimeSerial(12, 0, 0), f + TimeSerial(12, 2, 0), 2, "OR-01", "micro", "TP-05"

    AgregarProduccion f, "TST-C5", 1, "TST-C5-L", "TST-C5-TL", "500SG", 22800, 0
    AgregarRechazo f, "TST-C5", "TST-C5-L", "TST-C5-TL", "", "", 60
End Sub

Private Sub PrepararCasoRN15()
    Dim f As Date
    f = DateSerial(2026, 10, 16)

    AgregarLote "TST-RN15-L", "TSTRN15", "500SG", f + TimeSerial(6, 0, 0), f + TimeSerial(10, 0, 0)
    AgregarTurno "TST-RN15", "TST-RN15-TL", "TST-RN15-L", f, "RES-M", "RES", _
        f + TimeSerial(6, 0, 0), f + TimeSerial(10, 0, 0)
    AgregarTramo "TST-RN15-TL", f + TimeSerial(8, 0, 0), f + TimeSerial(10, 0, 0)

    ' Parada de 60 min que cruza el limite: 30 min con envase y 30 min sin envase.
    AgregarParada "TST-RN15-P1", f, "TST-RN15", 1, "TST-RN15-L", "TST-RN15-TL", _
        f + TimeSerial(7, 30, 0), f + TimeSerial(8, 30, 0), 60, "OR-02", "averia llenadora", "TP-03"

    AgregarProduccion f, "TST-RN15", 1, "TST-RN15-L", "TST-RN15-TL", "500SG", 11850, 0
End Sub

Private Sub AgregarLote(ByVal idLote As String, ByVal numeroLote As String, ByVal producto As String, _
                        ByVal ini As Date, ByVal fin As Date)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "id_lote", idLote
    SetValor lo, f, "numero_lote", numeroLote
    SetValor lo, f, "producto", producto
    SetValor lo, f, "fecha_hora_inicio", ini
    SetValor lo, f, "fecha_hora_fin", fin
End Sub

Private Sub AgregarTurno(ByVal idTurno As String, ByVal idTurnoLote As String, ByVal idLote As String, _
                         ByVal fecha As Date, ByVal turnoCodigo As String, ByVal equipo As String, _
                         ByVal ini As Date, ByVal fin As Date)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "id_turno", idTurno
    SetValor lo, f, "id_turno_lote", idTurnoLote
    SetValor lo, f, "id_lote", idLote
    SetValor lo, f, "fecha", fecha
    SetValor lo, f, "turno_codigo", turnoCodigo
    SetValor lo, f, "equipo", equipo
    SetValor lo, f, "fecha_hora_inicio", ini
    SetValor lo, f, "fecha_hora_fin", fin
End Sub

Private Sub AgregarParada(ByVal id As String, ByVal fecha As Date, ByVal turno As String, _
                          ByVal linea As Long, ByVal idLote As String, ByVal idTurnoLote As String, _
                          ByVal ini As Date, ByVal fin As Date, ByVal duracionMin As Double, _
                          ByVal origen As String, ByVal motivo As String, ByVal tipo As String)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "id", id
    SetValor lo, f, "fecha", fecha
    SetValor lo, f, "turno", turno
    SetValor lo, f, "linea_id", linea
    SetValor lo, f, "id_lote", idLote
    SetValor lo, f, "id_turno_lote", idTurnoLote
    SetValor lo, f, "fecha_hora_inicio", ini
    SetValor lo, f, "fecha_hora_fin", fin
    SetValor lo, f, "duracion_min", duracionMin
    SetValor lo, f, "origen", origen
    SetValor lo, f, "motivo", motivo
    SetValor lo, f, "tipo", tipo
    SetValor lo, f, "interpretada", "No"
End Sub

Private Sub AgregarProduccion(ByVal fecha As Date, ByVal turno As String, ByVal linea As Long, _
                              ByVal lote As String, ByVal idTurnoLote As String, ByVal producto As String, _
                              ByVal unidades As Double, ByVal pallets As Double)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblProduccion")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "fecha", fecha
    SetValor lo, f, "turno", turno
    SetValor lo, f, "linea_id", linea
    SetValor lo, f, "lote", lote
    SetValor lo, f, "id_turno_lote", idTurnoLote
    SetValor lo, f, "producto", producto
    SetValor lo, f, "unidades_producidas", unidades
    SetValor lo, f, "pallets", pallets
    SetValor lo, f, "litros", 0
    SetValor lo, f, "velocidad_nominal", 0
End Sub

Private Sub AgregarRechazo(ByVal fecha As Date, ByVal turno As String, ByVal lote As String, _
                           ByVal idTurnoLote As String, ByVal seccion As String, _
                           ByVal motivo As String, ByVal cantidad As Double)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblRechazos")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "fecha", fecha
    SetValor lo, f, "turno", turno
    SetValor lo, f, "lote", lote
    SetValor lo, f, "id_turno_lote", idTurnoLote
    SetValor lo, f, "seccion", seccion
    SetValor lo, f, "motivo", motivo
    SetValor lo, f, "cantidad", cantidad
End Sub

Private Sub AgregarTramo(ByVal idTurnoLote As String, ByVal ini As Date, ByVal fin As Date)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblTramosSinEnvase")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "id_turno_lote", idTurnoLote
    SetValor lo, f, "fecha_hora_inicio", ini
    SetValor lo, f, "fecha_hora_fin", fin
End Sub

Private Sub AgregarVelocidad(ByVal codigo As String, ByVal producto As String, ByVal condicion As String, _
                             ByVal velocidad As Double, ByVal unidad As String, ByVal vigencia As Date)
    Dim lo As ListObject
    Dim f As ListRow

    Set lo = Utils.Tabla("tblVelocidades")
    If lo Is Nothing Then Exit Sub
    Set f = lo.ListRows.Add

    SetValor lo, f, "codigo", codigo
    SetValor lo, f, "producto", producto
    SetValor lo, f, "condicion_envase", condicion
    SetValor lo, f, "velocidad", velocidad
    SetValor lo, f, "unidad", unidad
    SetValor lo, f, "vigencia_desde", vigencia
    SetValor lo, f, "origen_dato", "prueba"
    SetValor lo, f, "activo", "Si"
End Sub

Private Sub SetValor(ByVal lo As ListObject, ByVal fila As ListRow, ByVal columna As String, ByVal valor As Variant)
    Dim c As Long

    c = Utils.IndiceColumna(lo, columna)
    If c > 0 Then fila.Range.Cells(1, c).Value = valor
End Sub

Private Sub BorrarPorPrefijo(ByVal nombreTabla As String, ByVal columna As String, ByVal prefijo As String)
    Dim lo As ListObject
    Dim rng As Range
    Dim c As Long
    Dim i As Long
    Dim v As Variant

    Set lo = Utils.Tabla(nombreTabla)
    If lo Is Nothing Then Exit Sub

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = rng.Rows.Count To 1 Step -1
        v = rng.Cells(i, c).Value
        If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
            If InStr(1, CStr(v), prefijo, vbTextCompare) = 1 Then
                lo.ListRows(i).Delete
            End If
        End If
    Next i
End Sub
