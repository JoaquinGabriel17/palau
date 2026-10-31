Attribute VB_Name = "Agregacion"
Option Explicit

Private Const UMBRAL_MICRO_DEFECTO As Double = 5

Private Type TResumenTurno
    TD As Double
    PP As Double
    PE As Double
    AF As Double
    AC As Double
    Micro As Double
    PNP As Double
    UnidadesBuenas As Double
    UnidadesRechazadas As Double
    TiempoIdeal As Double
    Encontrado As Boolean
End Type

Public Type TDetalleTurno
    TD As Double
    PP As Double
    PE As Double
    AF As Double
    AC As Double
    Micro As Double
    PNP As Double
    TPL As Double
    TO As Double
    TiempoIdeal As Double
    UnidadesBuenas As Double
    UnidadesRechazadas As Double
    UnidadesTotales As Double
    Disponibilidad As Variant
    Rendimiento As Variant
    Calidad As Variant
    OEE As Double
    Encontrado As Boolean
End Type

Private Function NumeroSeguro(ByVal v As Variant) As Double
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then NumeroSeguro = CDbl(v)
End Function

Public Function CalcularOEE_Turno(ByVal id_turno As String) As Double
    Dim detalle As TDetalleTurno

    detalle = DetalleTurnoPorId(id_turno)
    CalcularOEE_Turno = detalle.OEE
End Function

Public Function CalcularTurno(ByVal fecha As Date, ByVal turno_codigo As String) As Double
    Dim detalle As TDetalleTurno

    detalle = DetalleTurno(fecha, turno_codigo)
    CalcularTurno = detalle.OEE
End Function

Public Function DetalleTurno(ByVal fecha As Date, ByVal turno_codigo As String) As TDetalleTurno
    Dim lo As ListObject
    Dim rng As Range
    Dim ids As Collection
    Dim resumen As TResumenTurno
    Dim total As TResumenTurno
    Dim cIdTurno As Long
    Dim cFecha As Long
    Dim cTurno As Long
    Dim i As Long
    Dim k As Variant
    Dim dia As Date
    Dim f As Date

    Set lo = Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cIdTurno = IndiceColumna(lo, "id_turno")
    cFecha = IndiceColumna(lo, "fecha")
    cTurno = IndiceColumna(lo, "turno_codigo")
    If cIdTurno = 0 Or cFecha = 0 Or cTurno = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    Set ids = New Collection
    dia = DateSerial(Year(fecha), Month(fecha), Day(fecha))

    For i = 1 To rng.Rows.Count
        f = AFecha(rng.Cells(i, cFecha).Value)
        If f > 0 Then
            If DateSerial(Year(f), Month(f), Day(f)) = dia Then
                If StrComp(Trim$(CStr(rng.Cells(i, cTurno).Value)), turno_codigo, vbTextCompare) = 0 Then
                    AgregarUnico ids, CStr(rng.Cells(i, cIdTurno).Value)
                End If
            End If
        End If
    Next i

    If ids.Count = 0 Then Exit Function

    For Each k In ids
        If ResumenTurno(CStr(k), resumen) Then
            total.TD = total.TD + resumen.TD
            total.PP = total.PP + resumen.PP
            total.PE = total.PE + resumen.PE
            total.AF = total.AF + resumen.AF
            total.AC = total.AC + resumen.AC
            total.Micro = total.Micro + resumen.Micro
            total.PNP = total.PNP + resumen.PNP
            total.UnidadesBuenas = total.UnidadesBuenas + resumen.UnidadesBuenas
            total.UnidadesRechazadas = total.UnidadesRechazadas + resumen.UnidadesRechazadas
            total.TiempoIdeal = total.TiempoIdeal + resumen.TiempoIdeal
            total.Encontrado = True
        End If
    Next k

    If Not total.Encontrado Then Exit Function

    DetalleTurno = DetalleDesdeResumen(total)
End Function

Private Function DetalleTurnoPorId(ByVal id_turno As String) As TDetalleTurno
    Dim resumen As TResumenTurno

    If ResumenTurno(id_turno, resumen) Then
        DetalleTurnoPorId = DetalleDesdeResumen(resumen)
    End If
End Function

Private Function DetalleDesdeResumen(ByRef resumen As TResumenTurno) As TDetalleTurno
    Dim detalle As TDetalleTurno

    detalle = DetalleDesdeTotales( _
        resumen.TD, _
        resumen.PP, _
        resumen.PE, _
        resumen.AF, _
        resumen.AC, _
        resumen.Micro, _
        resumen.PNP, _
        resumen.UnidadesBuenas, _
        resumen.UnidadesRechazadas, _
        resumen.TiempoIdeal)
    detalle.Encontrado = resumen.Encontrado
    DetalleDesdeResumen = detalle
End Function

Private Function DetalleDesdeTotales(ByVal TD As Double, ByVal PP As Double, ByVal PE As Double, ByVal AF As Double, ByVal AC As Double, ByVal Micro As Double, ByVal PNP As Double, ByVal UnidadesBuenas As Double, ByVal UnidadesRechazadas As Double, ByVal TiempoIdealMin As Double) As TDetalleTurno
    Dim detalle As TDetalleTurno
    Dim velEfectiva As Double

    detalle.TD = TD
    detalle.PP = PP
    detalle.PE = PE
    detalle.AF = AF
    detalle.AC = AC
    detalle.Micro = Micro
    detalle.PNP = PNP

    detalle.TPL = TD - PP - PE
    If detalle.TPL < 0 Then detalle.TPL = 0

    detalle.TO = CalcularTO(detalle.TPL, PNP)
    If detalle.TO < 0 Then detalle.TO = 0

    detalle.TiempoIdeal = TiempoIdealMin
    detalle.UnidadesBuenas = UnidadesBuenas
    detalle.UnidadesRechazadas = UnidadesRechazadas
    detalle.UnidadesTotales = UnidadesBuenas + UnidadesRechazadas

    If TiempoIdealMin > 0 Then velEfectiva = detalle.UnidadesTotales / TiempoIdealMin

    detalle.Disponibilidad = CalcularDisponibilidad(detalle.TO, detalle.TPL)
    detalle.Rendimiento = CalcularRendimientoOEE(detalle.UnidadesTotales, velEfectiva, detalle.TO)
    detalle.Calidad = CalcularCalidad(detalle.UnidadesBuenas, detalle.UnidadesTotales)
    detalle.OEE = CalcularOEE( _
        NumeroSeguro(detalle.Disponibilidad), _
        NumeroSeguro(detalle.Rendimiento), _
        NumeroSeguro(detalle.Calidad))

    DetalleDesdeTotales = detalle
End Function

Private Function ResumenTurno(ByVal id_turno As String, ByRef resumen As TResumenTurno) As Boolean
    Dim loTurnos As ListObject
    Dim loParadas As ListObject
    Dim loProduccion As ListObject
    Dim loRechazos As ListObject
    Dim loVelocidades As ListObject
    Dim rng As Range
    Dim segs As Collection
    Dim lotes As Collection
    Dim rechSeg As Object
    Dim velSeg As Object
    Dim k As Variant
    Dim i As Long
    Dim cIdTurno As Long
    Dim cIdTurnoLote As Long
    Dim cIdLote As Long
    Dim cFecha As Long
    Dim cTurnoCodigo As Long
    Dim cInicio As Long
    Dim cFin As Long
    Dim cSegParada As Long
    Dim cTurnoParada As Long
    Dim cDurParada As Long
    Dim cIniParada As Long
    Dim cFinParada As Long
    Dim cTipoParada As Long
    Dim cSegProduccion As Long
    Dim cLoteProduccion As Long
    Dim cUnidades As Long
    Dim cProducto As Long
    Dim cVelocidad As Long
    Dim cFechaProduccion As Long
    Dim cSegRechazo As Long
    Dim cLoteRechazo As Long
    Dim cCantidad As Long
    Dim seg As String
    Dim lote As String
    Dim producto As String
    Dim tipo As String
    Dim turnoCodigo As String
    Dim dtIni As Date
    Dim dtFin As Date
    Dim dtTurno As Date
    Dim hayTurno As Boolean
    Dim dur As Double
    Dim umbral As Double
    Dim td As Double
    Dim pp As Double
    Dim pe As Double
    Dim af As Double
    Dim ac As Double
    Dim micro As Double
    Dim pnp As Double
    Dim prod As Double
    Dim rech As Double
    Dim buenas As Double
    Dim tiempoIdeal As Double
    Dim vel As Double
    Dim incluir As Boolean

    resumen.Encontrado = False
    Set loTurnos = Tabla("tblTurnosRegistro")
    Set loParadas = Tabla("tblParadas")
    Set loProduccion = Tabla("tblProduccion")
    Set loRechazos = Tabla("tblRechazos")
    Set loVelocidades = Tabla("tblVelocidades")

    If loTurnos Is Nothing Then Exit Function

    cIdTurno = IndiceColumna(loTurnos, "id_turno")
    If cIdTurno = 0 Then Exit Function
    cIdTurnoLote = IndiceColumna(loTurnos, "id_turno_lote")
    cIdLote = IndiceColumna(loTurnos, "id_lote")
    cFecha = IndiceColumna(loTurnos, "fecha")
    cTurnoCodigo = IndiceColumna(loTurnos, "turno_codigo")
    cInicio = IndiceColumna(loTurnos, "fecha_hora_inicio")
    cFin = IndiceColumna(loTurnos, "fecha_hora_fin")

    Set segs = New Collection
    Set lotes = New Collection
    Set rng = loTurnos.DataBodyRange

    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            If StrComp(CStr(rng.Cells(i, cIdTurno).Value), id_turno, vbTextCompare) = 0 Then
                If cIdTurnoLote > 0 Then
                    seg = CStr(rng.Cells(i, cIdTurnoLote).Value)
                    AgregarUnico segs, seg
                End If
                If cIdLote > 0 Then
                    lote = CStr(rng.Cells(i, cIdLote).Value)
                    AgregarUnico lotes, lote
                End If
                If cFecha > 0 And dtTurno = 0 Then dtTurno = AFecha(rng.Cells(i, cFecha).Value)
                If cTurnoCodigo > 0 And Len(turnoCodigo) = 0 Then turnoCodigo = CStr(rng.Cells(i, cTurnoCodigo).Value)
                If cInicio > 0 And cFin > 0 Then
                    If AFecha(rng.Cells(i, cInicio).Value) > 0 And AFecha(rng.Cells(i, cFin).Value) > 0 Then
                        If Not hayTurno Then
                            dtIni = AFecha(rng.Cells(i, cInicio).Value)
                            dtFin = AFecha(rng.Cells(i, cFin).Value)
                        Else
                            If AFecha(rng.Cells(i, cInicio).Value) < dtIni Then dtIni = AFecha(rng.Cells(i, cInicio).Value)
                            If AFecha(rng.Cells(i, cFin).Value) > dtFin Then dtFin = AFecha(rng.Cells(i, cFin).Value)
                        End If
                        hayTurno = True
                    End If
                End If
            End If
        Next i
    End If

    If Not hayTurno And segs.Count = 0 Then Exit Function

    If dtIni > 0 And dtFin > 0 Then
        td = DateDiff("n", dtIni, dtFin)
        If td < 0 Then td = td + 1440
    End If
    If td <= 0 Then td = TDDesdeCatalogo(turnoCodigo)

    umbral = UmbralMicro()

    If Not loParadas Is Nothing Then
        cSegParada = IndiceColumna(loParadas, "id_turno_lote")
        cTurnoParada = IndiceColumna(loParadas, "turno")
        cDurParada = IndiceColumna(loParadas, "duracion_min")
        cIniParada = IndiceColumna(loParadas, "fecha_hora_inicio")
        cFinParada = IndiceColumna(loParadas, "fecha_hora_fin")
        cTipoParada = IndiceColumna(loParadas, "tipo")
        Set rng = loParadas.DataBodyRange
        If Not rng Is Nothing Then
            For i = 1 To rng.Rows.Count
                seg = ""
                If cSegParada > 0 Then seg = CStr(rng.Cells(i, cSegParada).Value)
                incluir = Pertenece(segs, seg)
                If Not incluir And Len(seg) = 0 And cTurnoParada > 0 Then
                    incluir = (StrComp(CStr(rng.Cells(i, cTurnoParada).Value), id_turno, vbTextCompare) = 0)
                End If
                If incluir Then
                    dur = 0
                    If cDurParada > 0 Then dur = NumeroSeguro(rng.Cells(i, cDurParada).Value)
                    If dur <= 0 And cIniParada > 0 And cFinParada > 0 Then
                        dur = DuracionMin(0, rng.Cells(i, cIniParada).Value, rng.Cells(i, cFinParada).Value)
                    End If
                    tipo = ""
                    If cTipoParada > 0 Then tipo = UCase$(Trim$(CStr(rng.Cells(i, cTipoParada).Value)))
                    If dur <= umbral Or tipo = "TP-05" Then
                        micro = micro + dur
                        dur = 0
                    ElseIf tipo = "TP-01" Then
                        pp = pp + dur
                    ElseIf tipo = "TP-07" Then
                        pe = pe + dur
                    ElseIf tipo = "TP-03" Then
                        af = af + dur
                        pnp = pnp + dur
                    ElseIf tipo = "TP-04" Then
                        ac = ac + dur
                        pnp = pnp + dur
                    Else
                        pnp = pnp + dur
                    End If
                End If
            Next i
        End If
    End If

    Set rechSeg = CreateObject("Scripting.Dictionary")
    rechSeg.CompareMode = 1

    If Not loRechazos Is Nothing Then
        cSegRechazo = IndiceColumna(loRechazos, "id_turno_lote")
        cLoteRechazo = IndiceColumna(loRechazos, "lote")
        cCantidad = IndiceColumna(loRechazos, "cantidad")
        Set rng = loRechazos.DataBodyRange
        If Not rng Is Nothing Then
            For i = 1 To rng.Rows.Count
                seg = ""
                lote = ""
                If cSegRechazo > 0 Then seg = CStr(rng.Cells(i, cSegRechazo).Value)
                If cLoteRechazo > 0 Then lote = CStr(rng.Cells(i, cLoteRechazo).Value)
                incluir = Pertenece(segs, seg)
                If Not incluir And Len(seg) = 0 Then incluir = Pertenece(lotes, lote)
                If incluir And cCantidad > 0 Then
                    dur = NumeroSeguro(rng.Cells(i, cCantidad).Value)
                    rech = rech + dur
                    If Len(seg) > 0 Then
                        If rechSeg.Exists(seg) Then
                            rechSeg(seg) = rechSeg(seg) + dur
                        Else
                            rechSeg.Add seg, dur
                        End If
                    End If
                End If
            Next i
        End If
    End If

    Set velSeg = CreateObject("Scripting.Dictionary")
    velSeg.CompareMode = 1

    If Not loProduccion Is Nothing Then
        cSegProduccion = IndiceColumna(loProduccion, "id_turno_lote")
        cLoteProduccion = IndiceColumna(loProduccion, "lote")
        cUnidades = IndiceColumna(loProduccion, "unidades_producidas")
        cProducto = IndiceColumna(loProduccion, "producto")
        cVelocidad = IndiceColumna(loProduccion, "velocidad_nominal")
        cFechaProduccion = IndiceColumna(loProduccion, "fecha")
        Set rng = loProduccion.DataBodyRange
        If Not rng Is Nothing Then
            For i = 1 To rng.Rows.Count
                seg = ""
                lote = ""
                If cSegProduccion > 0 Then seg = CStr(rng.Cells(i, cSegProduccion).Value)
                If cLoteProduccion > 0 Then lote = CStr(rng.Cells(i, cLoteProduccion).Value)
                incluir = Pertenece(segs, seg)
                If Not incluir And Len(seg) = 0 Then incluir = Pertenece(lotes, lote)
                If incluir And cUnidades > 0 Then
                    dur = NumeroSeguro(rng.Cells(i, cUnidades).Value)
                    prod = prod + dur
                    vel = 0
                    If cVelocidad > 0 Then vel = NumeroSeguro(rng.Cells(i, cVelocidad).Value)
                    If vel <= 0 And cProducto > 0 Then
                        producto = CStr(rng.Cells(i, cProducto).Value)
                        If cFechaProduccion > 0 Then
                            vel = BuscarVelocidad(loVelocidades, producto, AFecha(rng.Cells(i, cFechaProduccion).Value))
                        Else
                            vel = BuscarVelocidad(loVelocidades, producto, 0)
                        End If
                    End If
                    If vel > 0 Then
                        tiempoIdeal = tiempoIdeal + dur / vel
                        If Len(seg) > 0 Then
                            If Not velSeg.Exists(seg) Then velSeg.Add seg, vel
                        End If
                    End If
                End If
            Next i
        End If
    End If

    For Each k In rechSeg.Keys
        If velSeg.Exists(CStr(k)) Then
            tiempoIdeal = tiempoIdeal + CDbl(rechSeg(k)) / CDbl(velSeg(CStr(k)))
        End If
    Next k

    buenas = prod

    resumen.TD = td
    resumen.PP = pp
    resumen.PE = pe
    resumen.AF = af
    resumen.AC = ac
    resumen.Micro = micro
    resumen.PNP = pnp
    resumen.UnidadesBuenas = buenas
    resumen.UnidadesRechazadas = rech
    resumen.TiempoIdeal = tiempoIdeal
    resumen.Encontrado = True
    ResumenTurno = True
End Function

Public Function CalcularOEE_DesdeTotales(ByVal TD As Double, ByVal PP As Double, ByVal PE As Double, ByVal PNP As Double, ByVal UnidadesBuenas As Double, ByVal UnidadesRechazadas As Double, ByVal TiempoIdealMin As Double) As Double
    Dim detalle As TDetalleTurno

    detalle = DetalleDesdeTotales(TD, PP, PE, 0, 0, 0, PNP, UnidadesBuenas, UnidadesRechazadas, TiempoIdealMin)
    CalcularOEE_DesdeTotales = detalle.OEE
End Function

Public Function CalcularOEE_Dia(ByVal fecha As Date) As Double
    Dim lo As ListObject
    Dim rng As Range
    Dim ids As Collection
    Dim resumen As TResumenTurno
    Dim cIdTurno As Long
    Dim cFecha As Long
    Dim i As Long
    Dim k As Variant
    Dim dia As Date
    Dim td As Double
    Dim pp As Double
    Dim pe As Double
    Dim pnp As Double
    Dim buenas As Double
    Dim rech As Double
    Dim tiempoIdeal As Double
    Dim n As Long
    Dim f As Date

    Set lo = Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cIdTurno = IndiceColumna(lo, "id_turno")
    cFecha = IndiceColumna(lo, "fecha")
    If cIdTurno = 0 Or cFecha = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    Set ids = New Collection
    dia = DateSerial(Year(fecha), Month(fecha), Day(fecha))

    For i = 1 To rng.Rows.Count
        f = AFecha(rng.Cells(i, cFecha).Value)
        If f > 0 Then
            If DateSerial(Year(f), Month(f), Day(f)) = dia Then
                AgregarUnico ids, CStr(rng.Cells(i, cIdTurno).Value)
            End If
        End If
    Next i

    If ids.Count = 0 Then Exit Function

    For Each k In ids
        If ResumenTurno(CStr(k), resumen) Then
            If resumen.UnidadesBuenas + resumen.UnidadesRechazadas > 0 Or resumen.PP + resumen.PE + resumen.PNP > 0 Then
                td = td + resumen.TD
                pp = pp + resumen.PP
                pe = pe + resumen.PE
                pnp = pnp + resumen.PNP
                buenas = buenas + resumen.UnidadesBuenas
                rech = rech + resumen.UnidadesRechazadas
                tiempoIdeal = tiempoIdeal + resumen.TiempoIdeal
                n = n + 1
            End If
        End If
    Next k

    If n = 0 Then Exit Function

    CalcularOEE_Dia = CalcularOEE_DesdeTotales(td, pp, pe, pnp, buenas, rech, tiempoIdeal)
End Function

Public Function ExisteTurno(ByVal id_turno As String) As Boolean
    Dim lo As ListObject
    Dim rng As Range
    Dim cIdTurno As Long
    Dim i As Long

    Set lo = Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function
    cIdTurno = IndiceColumna(lo, "id_turno")
    If cIdTurno = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cIdTurno).Value), id_turno, vbTextCompare) = 0 Then
            ExisteTurno = True
            Exit Function
        End If
    Next i
End Function

Public Function VelocidadVigente(ByVal producto As String, ByVal fecha As Date) As Double
    VelocidadVigente = BuscarVelocidad(Tabla("tblVelocidades"), producto, fecha)
End Function

Public Function Tabla(ByVal nombre As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nombre, vbTextCompare) = 0 Then
                Set Tabla = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function IndiceColumna(ByVal lo As ListObject, ByVal nombre As String) As Long
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, nombre, vbTextCompare) = 0 Then
            IndiceColumna = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Sub AgregarUnico(ByVal col As Collection, ByVal valor As String)
    Dim v As Variant

    If Len(valor) = 0 Then Exit Sub
    For Each v In col
        If StrComp(CStr(v), valor, vbTextCompare) = 0 Then Exit Sub
    Next v
    col.Add valor
End Sub

Private Function Pertenece(ByVal col As Collection, ByVal valor As String) As Boolean
    Dim v As Variant

    If Len(valor) = 0 Then Exit Function
    For Each v In col
        If StrComp(CStr(v), valor, vbTextCompare) = 0 Then
            Pertenece = True
            Exit Function
        End If
    Next v
End Function

Private Function AFecha(ByVal v As Variant) As Date
    If IsDate(v) Then AFecha = CDate(v)
End Function

Private Function DuracionMin(ByVal dur As Variant, ByVal ini As Variant, ByVal fin As Variant) As Double
    Dim d As Double
    Dim f1 As Date
    Dim f2 As Date

    d = NumeroSeguro(dur)
    If d > 0 Then
        DuracionMin = d
        Exit Function
    End If

    If IsDate(ini) And IsDate(fin) Then
        f1 = CDate(ini)
        f2 = CDate(fin)
        d = DateDiff("n", f1, f2)
        If d < 0 Then d = d + 1440
        DuracionMin = d
    End If
End Function

Private Function UmbralMicro() As Double
    Dim lo As ListObject
    Dim rng As Range
    Dim cParametro As Long
    Dim cValor As Long
    Dim i As Long
    Dim v As Double

    Set lo = Tabla("tblParametros")
    If lo Is Nothing Then
        UmbralMicro = UMBRAL_MICRO_DEFECTO
        Exit Function
    End If

    cParametro = IndiceColumna(lo, "parametro")
    cValor = IndiceColumna(lo, "valor")
    If cParametro = 0 Or cValor = 0 Then
        UmbralMicro = UMBRAL_MICRO_DEFECTO
        Exit Function
    End If

    Set rng = lo.DataBodyRange
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            If StrComp(CStr(rng.Cells(i, cParametro).Value), "UMBRAL_MICROPARADA_MIN", vbTextCompare) = 0 Then
                v = NumeroSeguro(rng.Cells(i, cValor).Value)
                If v > 0 Then
                    UmbralMicro = v
                    Exit Function
                End If
            End If
        Next i
    End If

    UmbralMicro = UMBRAL_MICRO_DEFECTO
End Function

Private Function TDDesdeCatalogo(ByVal codigo As String) As Double
    Dim lo As ListObject
    Dim rng As Range
    Dim cCodigo As Long
    Dim cTD As Long
    Dim cInicio As Long
    Dim cFin As Long
    Dim i As Long
    Dim hi As Double
    Dim hf As Double

    If Len(codigo) = 0 Then Exit Function

    Set lo = Tabla("tblTurnos")
    If lo Is Nothing Then Exit Function

    cCodigo = IndiceColumna(lo, "codigo")
    If cCodigo = 0 Then Exit Function
    cTD = IndiceColumna(lo, "td_minutos")
    cInicio = IndiceColumna(lo, "hora_inicio")
    cFin = IndiceColumna(lo, "hora_fin")

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cCodigo).Value), codigo, vbTextCompare) = 0 Then
            If cTD > 0 Then TDDesdeCatalogo = NumeroSeguro(rng.Cells(i, cTD).Value)
            If TDDesdeCatalogo <= 0 And cInicio > 0 And cFin > 0 Then
                hi = NumeroSeguro(rng.Cells(i, cInicio).Value) * 1440
                hf = NumeroSeguro(rng.Cells(i, cFin).Value) * 1440
                If hf <= hi Then hf = hf + 1440
                TDDesdeCatalogo = hf - hi
            End If
            Exit Function
        End If
    Next i
End Function

Private Function BuscarVelocidad(ByVal lo As ListObject, ByVal producto As String, ByVal fecha As Date) As Double
    Dim rng As Range
    Dim cProducto As Long
    Dim cVelocidad As Long
    Dim cDesde As Long
    Dim cHasta As Long
    Dim cActivo As Long
    Dim i As Long
    Dim v As Double
    Dim desde As Date
    Dim hasta As Date
    Dim activo As String
    Dim mejorFecha As Date
    Dim mejorVel As Double

    If lo Is Nothing Then Exit Function

    cProducto = IndiceColumna(lo, "producto")
    cVelocidad = IndiceColumna(lo, "velocidad")
    cDesde = IndiceColumna(lo, "vigencia_desde")
    cHasta = IndiceColumna(lo, "vigencia_hasta")
    cActivo = IndiceColumna(lo, "activo")
    If cProducto = 0 Or cVelocidad = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cProducto).Value), producto, vbTextCompare) = 0 Then
            v = NumeroSeguro(rng.Cells(i, cVelocidad).Value)
            If v > 0 Then
                activo = ""
                If cActivo > 0 Then activo = CStr(rng.Cells(i, cActivo).Value)
                If StrComp(activo, "No", vbTextCompare) <> 0 Then
                    desde = 0
                    hasta = 0
                    If cDesde > 0 Then desde = AFecha(rng.Cells(i, cDesde).Value)
                    If cHasta > 0 Then hasta = AFecha(rng.Cells(i, cHasta).Value)
                    If (desde = 0 Or fecha = 0 Or fecha >= desde) And (hasta = 0 Or fecha = 0 Or fecha <= hasta) Then
                        If mejorVel = 0 Or desde >= mejorFecha Then
                            mejorVel = v
                            mejorFecha = desde
                        End If
                    End If
                End If
            End If
        End If
    Next i

    BuscarVelocidad = mejorVel
End Function
