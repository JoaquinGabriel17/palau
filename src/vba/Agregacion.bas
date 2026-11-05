Attribute VB_Name = "Agregacion"
Option Explicit

' Respaldo solo si falta tblParametros o el parametro P-01.
' El valor normal se lee siempre de tblParametros.
Private Const UMBRAL_MICRO_DEFECTO As Double = 5

Private Type TResumenTurno
    td As Double
    pp As Double
    pe As Double
    af As Double
    ac As Double
    micro As Double
    pnp As Double
    UnidadesBuenas As Double
    UnidadesRechazadas As Double
    tiempoIdeal As Double
    produccionTeorica As Double
    Encontrado As Boolean
End Type

Public Type TDetalleTurno
    td As Double
    pp As Double
    pe As Double
    af As Double
    ac As Double
    micro As Double
    pnp As Double
    tpl As Double
    TO As Double
    tiempoIdeal As Double
    produccionTeorica As Double
    UnidadesBuenas As Double
    UnidadesRechazadas As Double
    UnidadesTotales As Double
    disponibilidad As Variant
    rendimiento As Variant
    calidad As Variant
    oee As Double
    Encontrado As Boolean
End Type

Public Function CalcularOEE_Turno(ByVal id_turno As String) As Double
    Dim detalle As TDetalleTurno

    detalle = DetalleTurnoPorId(id_turno)
    CalcularOEE_Turno = detalle.oee
End Function

Public Function CalcularTurno(ByVal fecha As Date, ByVal turno_codigo As String) As Double
    Dim detalle As TDetalleTurno

    detalle = DetalleTurno(fecha, turno_codigo)
    CalcularTurno = detalle.oee
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
    Dim cAnulado As Long
    Dim i As Long
    Dim k As Variant
    Dim dia As Date
    Dim f As Date

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cIdTurno = Utils.IndiceColumna(lo, "id_turno")
    cFecha = Utils.IndiceColumna(lo, "fecha")
    cTurno = Utils.IndiceColumna(lo, "turno_codigo")
    cAnulado = Utils.IndiceColumna(lo, "anulado")
    If cIdTurno = 0 Or cFecha = 0 Or cTurno = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    Set ids = New Collection
    dia = DateSerial(Year(fecha), Month(fecha), Day(fecha))

    For i = 1 To rng.Rows.Count
        If Not FilaAnulada(rng.Rows(i), cAnulado) Then
            f = Utils.AFecha(rng.Cells(i, cFecha).Value)
            If f > 0 Then
                If DateSerial(Year(f), Month(f), Day(f)) = dia Then
                    If StrComp(Trim$(CStr(rng.Cells(i, cTurno).Value)), turno_codigo, vbTextCompare) = 0 Then
                        AgregarUnico ids, CStr(rng.Cells(i, cIdTurno).Value)
                    End If
                End If
            End If
        End If
    Next i

    If ids.Count = 0 Then Exit Function

    For Each k In ids
        If ResumenTurno(CStr(k), resumen) Then
            total.td = total.td + resumen.td
            total.pp = total.pp + resumen.pp
            total.pe = total.pe + resumen.pe
            total.af = total.af + resumen.af
            total.ac = total.ac + resumen.ac
            total.micro = total.micro + resumen.micro
            total.pnp = total.pnp + resumen.pnp
            total.UnidadesBuenas = total.UnidadesBuenas + resumen.UnidadesBuenas
            total.UnidadesRechazadas = total.UnidadesRechazadas + resumen.UnidadesRechazadas
            total.tiempoIdeal = total.tiempoIdeal + resumen.tiempoIdeal
            total.Encontrado = True
        End If
    Next k

    If Not total.Encontrado Then Exit Function

    DetalleTurno = DetalleDesdeResumen(total)
End Function

Public Function DetalleTurnoPorId(ByVal id_turno As String) As TDetalleTurno
    Dim resumen As TResumenTurno

    If ResumenTurno(id_turno, resumen) Then
        DetalleTurnoPorId = DetalleDesdeResumen(resumen)
    End If
End Function

Private Function DetalleDesdeResumen(ByRef resumen As TResumenTurno) As TDetalleTurno
    Dim detalle As TDetalleTurno

    detalle = DetalleDesdeTotales( _
        resumen.td, _
        resumen.pp, _
        resumen.pe, _
        resumen.af, _
        resumen.ac, _
        resumen.micro, _
        resumen.pnp, _
        resumen.UnidadesBuenas, _
        resumen.UnidadesRechazadas, _
        resumen.tiempoIdeal, _
        resumen.produccionTeorica)
    detalle.Encontrado = resumen.Encontrado
    DetalleDesdeResumen = detalle
End Function

Private Function DetalleDesdeTotales(ByVal td As Double, ByVal pp As Double, ByVal pe As Double, ByVal af As Double, ByVal ac As Double, ByVal micro As Double, ByVal pnp As Double, ByVal UnidadesBuenas As Double, ByVal UnidadesRechazadas As Double, ByVal TiempoIdealMin As Double, Optional ByVal ProduccionTeorica As Double = 0) As TDetalleTurno
    Dim detalle As TDetalleTurno
    Dim velEfectiva As Double

    detalle.td = td
    detalle.pp = pp
    detalle.pe = pe
    detalle.af = af
    detalle.ac = ac
    detalle.micro = micro
    detalle.pnp = pnp

    detalle.tpl = td - pp - pe
    If detalle.tpl < 0 Then detalle.tpl = 0

    detalle.TO = CalcularTO(detalle.tpl, pnp)
    If detalle.TO < 0 Then detalle.TO = 0

    detalle.tiempoIdeal = TiempoIdealMin
    detalle.produccionTeorica = ProduccionTeorica
    detalle.UnidadesBuenas = UnidadesBuenas
    detalle.UnidadesRechazadas = UnidadesRechazadas
    detalle.UnidadesTotales = UnidadesBuenas + UnidadesRechazadas

    detalle.disponibilidad = CalcularDisponibilidad(detalle.TO, detalle.tpl)
    If ProduccionTeorica > 0 Then
        detalle.rendimiento = CalcularRendimientoTeorico(detalle.UnidadesTotales, ProduccionTeorica)
    Else
        If TiempoIdealMin > 0 Then velEfectiva = detalle.UnidadesTotales / TiempoIdealMin
        detalle.rendimiento = CalcularRendimientoOEE(detalle.UnidadesTotales, velEfectiva, detalle.TO)
    End If
    detalle.calidad = CalcularCalidad(detalle.UnidadesBuenas, detalle.UnidadesTotales)
    detalle.oee = CalcularOEE( _
        Utils.NumeroSeguro(detalle.disponibilidad), _
        Utils.NumeroSeguro(detalle.rendimiento), _
        Utils.NumeroSeguro(detalle.calidad))

    DetalleDesdeTotales = detalle
End Function

Private Function ResumenTurno(ByVal id_turno As String, ByRef resumen As TResumenTurno) As Boolean
    Dim loTurnos As ListObject
    Dim loParadas As ListObject
    Dim loProduccion As ListObject
    Dim loRechazos As ListObject
    Dim loVelocidades As ListObject
    Dim loTramos As ListObject
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
    Dim cTurnoCodigo As Long
    Dim cInicio As Long
    Dim cFin As Long
    Dim cAnulado As Long
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
    Dim fIni As Date
    Dim fFin As Date
    Dim durSeg As Double
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
    Dim produccionTeorica As Double
    Dim toTotal As Double
    Dim hayTramos As Boolean
    Dim vel As Double
    Dim incluir As Boolean

    resumen.Encontrado = False
    Set loTurnos = Utils.Tabla("tblTurnosRegistro")
    Set loParadas = Utils.Tabla("tblParadas")
    Set loProduccion = Utils.Tabla("tblProduccion")
    Set loRechazos = Utils.Tabla("tblRechazos")
    Set loVelocidades = Utils.Tabla("tblVelocidades")
    Set loTramos = Utils.Tabla("tblTramosSinEnvase")

    If loTurnos Is Nothing Then Exit Function

    cIdTurno = Utils.IndiceColumna(loTurnos, "id_turno")
    If cIdTurno = 0 Then Exit Function
    cIdTurnoLote = Utils.IndiceColumna(loTurnos, "id_turno_lote")
    cIdLote = Utils.IndiceColumna(loTurnos, "id_lote")
    cTurnoCodigo = Utils.IndiceColumna(loTurnos, "turno_codigo")
    cInicio = Utils.IndiceColumna(loTurnos, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(loTurnos, "fecha_hora_fin")
    cAnulado = Utils.IndiceColumna(loTurnos, "anulado")

    Set segs = New Collection
    Set lotes = New Collection
    Set rng = loTurnos.DataBodyRange

    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            If Not FilaAnulada(rng.Rows(i), cAnulado) Then
                If StrComp(CStr(rng.Cells(i, cIdTurno).Value), id_turno, vbTextCompare) = 0 Then
                    If cIdTurnoLote > 0 Then
                        seg = CStr(rng.Cells(i, cIdTurnoLote).Value)
                        AgregarUnico segs, seg
                    End If
                    If cIdLote > 0 Then
                        lote = CStr(rng.Cells(i, cIdLote).Value)
                        AgregarUnico lotes, lote
                    End If
                    If cTurnoCodigo > 0 And Len(turnoCodigo) = 0 Then turnoCodigo = CStr(rng.Cells(i, cTurnoCodigo).Value)
                    If cInicio > 0 And cFin > 0 Then
                        fIni = Utils.AFecha(rng.Cells(i, cInicio).Value)
                        fFin = Utils.AFecha(rng.Cells(i, cFin).Value)
                        If fIni > 0 And fFin > 0 Then
                            durSeg = DateDiff("n", fIni, fFin)
                            If durSeg < 0 Then durSeg = durSeg + 1440
                            td = td + durSeg
                            hayTurno = True
                        End If
                    End If
                End If
            End If
        Next i
    End If

    If Not hayTurno And segs.Count = 0 Then Exit Function

    If td <= 0 Then td = TDDesdeCatalogo(turnoCodigo)

    umbral = UmbralMicro()

    ' La produccion teorica por tramos solo se usa cuando el turno tiene
    ' tramos sin envase. Si no hay tramos, se mantiene el calculo por
    ' tiempo ideal (unidades / velocidad), equivalente a todo con envase.
    hayTramos = HayTramosTurno(loTramos, segs)
    If hayTramos Then
        produccionTeorica = ProduccionTeoricaPorTramos( _
            loTurnos, loParadas, loProduccion, loTramos, loVelocidades, id_turno, umbral)
    End If

    If Not loParadas Is Nothing Then
        cSegParada = Utils.IndiceColumna(loParadas, "id_turno_lote")
        cTurnoParada = Utils.IndiceColumna(loParadas, "turno")
        cDurParada = Utils.IndiceColumna(loParadas, "duracion_min")
        cIniParada = Utils.IndiceColumna(loParadas, "fecha_hora_inicio")
        cFinParada = Utils.IndiceColumna(loParadas, "fecha_hora_fin")
        cTipoParada = Utils.IndiceColumna(loParadas, "tipo")
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
                    If cDurParada > 0 Then dur = Utils.NumeroSeguro(rng.Cells(i, cDurParada).Value)
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
        cSegRechazo = Utils.IndiceColumna(loRechazos, "id_turno_lote")
        cLoteRechazo = Utils.IndiceColumna(loRechazos, "lote")
        cCantidad = Utils.IndiceColumna(loRechazos, "cantidad")
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
                    dur = Utils.NumeroSeguro(rng.Cells(i, cCantidad).Value)
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
        cSegProduccion = Utils.IndiceColumna(loProduccion, "id_turno_lote")
        cLoteProduccion = Utils.IndiceColumna(loProduccion, "lote")
        cUnidades = Utils.IndiceColumna(loProduccion, "unidades_producidas")
        cProducto = Utils.IndiceColumna(loProduccion, "producto")
        cVelocidad = Utils.IndiceColumna(loProduccion, "velocidad_nominal")
        cFechaProduccion = Utils.IndiceColumna(loProduccion, "fecha")
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
                    dur = Utils.NumeroSeguro(rng.Cells(i, cUnidades).Value)
                    prod = prod + dur
                    vel = 0
                    If cVelocidad > 0 Then vel = Utils.NumeroSeguro(rng.Cells(i, cVelocidad).Value)
                    If vel <= 0 And cProducto > 0 Then
                        producto = CStr(rng.Cells(i, cProducto).Value)
                        If cFechaProduccion > 0 Then
                            vel = BuscarVelocidad(loVelocidades, producto, Utils.AFecha(rng.Cells(i, cFechaProduccion).Value))
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

    ' Sin tramos, se guarda la produccion teorica equivalente al calculo por
    ' tiempo ideal (UT * TO / tiempo ideal), para poder sumar varios turnos
    ' en el OEE del dia con la misma formula.
    If Not hayTramos Then
        If tiempoIdeal > 0 And (buenas + rech) > 0 Then
            toTotal = td - pp - pe - pnp
            If toTotal < 0 Then toTotal = 0
            produccionTeorica = (buenas + rech) * toTotal / tiempoIdeal
        End If
    End If

    resumen.td = td
    resumen.pp = pp
    resumen.pe = pe
    resumen.af = af
    resumen.ac = ac
    resumen.micro = micro
    resumen.pnp = pnp
    resumen.UnidadesBuenas = buenas
    resumen.UnidadesRechazadas = rech
    resumen.tiempoIdeal = tiempoIdeal
    resumen.produccionTeorica = produccionTeorica
    resumen.Encontrado = True
    ResumenTurno = True
End Function

Public Function CalcularOEE_DesdeTotales(ByVal td As Double, ByVal pp As Double, ByVal pe As Double, ByVal pnp As Double, ByVal UnidadesBuenas As Double, ByVal UnidadesRechazadas As Double, ByVal TiempoIdealMin As Double, Optional ByVal ProduccionTeorica As Double = 0) As Double
    Dim detalle As TDetalleTurno

    detalle = DetalleDesdeTotales(td, pp, pe, 0, 0, 0, pnp, UnidadesBuenas, UnidadesRechazadas, TiempoIdealMin, ProduccionTeorica)
    CalcularOEE_DesdeTotales = detalle.oee
End Function

Public Function CalcularOEE_Dia(ByVal fecha As Date) As Double
    Dim lo As ListObject
    Dim rng As Range
    Dim ids As Collection
    Dim resumen As TResumenTurno
    Dim cIdTurno As Long
    Dim cFecha As Long
    Dim cAnulado As Long
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
    Dim produccionTeorica As Double
    Dim n As Long
    Dim f As Date

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cIdTurno = Utils.IndiceColumna(lo, "id_turno")
    cFecha = Utils.IndiceColumna(lo, "fecha")
    cAnulado = Utils.IndiceColumna(lo, "anulado")
    If cIdTurno = 0 Or cFecha = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    Set ids = New Collection
    dia = DateSerial(Year(fecha), Month(fecha), Day(fecha))

    For i = 1 To rng.Rows.Count
        If Not FilaAnulada(rng.Rows(i), cAnulado) Then
            f = Utils.AFecha(rng.Cells(i, cFecha).Value)
            If f > 0 Then
                If DateSerial(Year(f), Month(f), Day(f)) = dia Then
                    AgregarUnico ids, CStr(rng.Cells(i, cIdTurno).Value)
                End If
            End If
        End If
    Next i

    If ids.Count = 0 Then Exit Function

    For Each k In ids
        If ResumenTurno(CStr(k), resumen) Then
            If resumen.UnidadesBuenas + resumen.UnidadesRechazadas > 0 Or resumen.pp + resumen.pe + resumen.pnp > 0 Then
                td = td + resumen.td
                pp = pp + resumen.pp
                pe = pe + resumen.pe
                pnp = pnp + resumen.pnp
                buenas = buenas + resumen.UnidadesBuenas
                rech = rech + resumen.UnidadesRechazadas
                tiempoIdeal = tiempoIdeal + resumen.tiempoIdeal
                produccionTeorica = produccionTeorica + resumen.produccionTeorica
                n = n + 1
            End If
        End If
    Next k

    If n = 0 Then Exit Function

    CalcularOEE_Dia = CalcularOEE_DesdeTotales(td, pp, pe, pnp, buenas, rech, tiempoIdeal, produccionTeorica)
End Function

Public Function ExisteTurno(ByVal id_turno As String) As Boolean
    Dim lo As ListObject
    Dim rng As Range
    Dim cIdTurno As Long
    Dim i As Long

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function
    cIdTurno = Utils.IndiceColumna(lo, "id_turno")
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

Public Function VelocidadVigente(ByVal producto As String, ByVal fecha As Date, Optional ByVal condicion As String = "con envase") As Double
    VelocidadVigente = BuscarVelocidad(Utils.Tabla("tblVelocidades"), producto, fecha, condicion)
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

Private Function FilaAnulada(ByVal fila As Range, ByVal colAnulado As Long) As Boolean
    Dim v As Variant

    If colAnulado = 0 Then Exit Function
    v = fila.Cells(1, colAnulado).Value
    FilaAnulada = Utils.EsAnulado(v)
End Function

Private Function DuracionMin(ByVal dur As Variant, ByVal ini As Variant, ByVal fin As Variant) As Double
    Dim d As Double
    Dim f1 As Date
    Dim f2 As Date

    d = Utils.NumeroSeguro(dur)
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
    Dim v As Variant

    v = ObtenerParametro("UMBRAL_MICROPARADA_MIN")
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            UmbralMicro = CDbl(v)
            Exit Function
        End If
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

    Set lo = Utils.Tabla("tblTurnos")
    If lo Is Nothing Then Exit Function

    cCodigo = Utils.IndiceColumna(lo, "codigo")
    If cCodigo = 0 Then Exit Function
    cTD = Utils.IndiceColumna(lo, "td_minutos")
    cInicio = Utils.IndiceColumna(lo, "hora_inicio")
    cFin = Utils.IndiceColumna(lo, "hora_fin")

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cCodigo).Value), codigo, vbTextCompare) = 0 Then
            If cTD > 0 Then TDDesdeCatalogo = Utils.NumeroSeguro(rng.Cells(i, cTD).Value)
            If TDDesdeCatalogo <= 0 And cInicio > 0 And cFin > 0 Then
                hi = Utils.NumeroSeguro(rng.Cells(i, cInicio).Value) * 1440
                hf = Utils.NumeroSeguro(rng.Cells(i, cFin).Value) * 1440
                If hf <= hi Then hf = hf + 1440
                TDDesdeCatalogo = hf - hi
            End If
            Exit Function
        End If
    Next i
End Function

Private Function BuscarVelocidad(ByVal lo As ListObject, ByVal producto As String, ByVal fecha As Date, _
                                  Optional ByVal condicion As String = "con envase") As Double
    Dim rng As Range
    Dim cProducto As Long
    Dim cVelocidad As Long
    Dim cDesde As Long
    Dim cHasta As Long
    Dim cActivo As Long
    Dim cCondicion As Long
    Dim cUnidad As Long
    Dim i As Long
    Dim v As Double
    Dim desde As Date
    Dim hasta As Date
    Dim activo As String
    Dim cond As String
    Dim unidad As String
    Dim buscada As String
    Dim mejorFecha As Date
    Dim mejorVel As Double

    If lo Is Nothing Then Exit Function

    cProducto = Utils.IndiceColumna(lo, "producto")
    cVelocidad = Utils.IndiceColumna(lo, "velocidad")
    cDesde = Utils.IndiceColumna(lo, "vigencia_desde")
    cHasta = Utils.IndiceColumna(lo, "vigencia_hasta")
    cActivo = Utils.IndiceColumna(lo, "activo")
    cCondicion = Utils.IndiceColumna(lo, "condicion_envase")
    cUnidad = Utils.IndiceColumna(lo, "unidad")
    If cProducto = 0 Or cVelocidad = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    buscada = LCase$(Trim$(condicion))
    If Len(buscada) = 0 Then buscada = "con envase"

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cProducto).Value), producto, vbTextCompare) = 0 Then
            v = Utils.NumeroSeguro(rng.Cells(i, cVelocidad).Value)
            If v > 0 Then
                activo = ""
                If cActivo > 0 Then activo = CStr(rng.Cells(i, cActivo).Value)
                If StrComp(activo, "No", vbTextCompare) <> 0 Then
                    cond = "con envase"
                    If cCondicion > 0 Then
                        cond = LCase$(Trim$(CStr(rng.Cells(i, cCondicion).Value)))
                        If Len(cond) = 0 Then cond = "con envase"
                    End If
                    If cond = buscada Then
                        If cUnidad > 0 Then
                            unidad = LCase$(Trim$(CStr(rng.Cells(i, cUnidad).Value)))
                            If InStr(1, unidad, "hora", vbTextCompare) > 0 Then v = v / 60
                        End If
                        desde = 0
                        hasta = 0
                        If cDesde > 0 Then desde = Utils.AFecha(rng.Cells(i, cDesde).Value)
                        If cHasta > 0 Then hasta = Utils.AFecha(rng.Cells(i, cHasta).Value)
                        If (desde = 0 Or fecha = 0 Or Int(fecha) >= Int(desde)) And (hasta = 0 Or fecha = 0 Or Int(fecha) <= Int(hasta)) Then
                            If mejorVel = 0 Or Int(desde) >= Int(mejorFecha) Then
                                mejorVel = v
                                mejorFecha = desde
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If mejorVel = 0 And buscada <> "con envase" Then
        BuscarVelocidad = BuscarVelocidad(lo, producto, fecha, "con envase")
    Else
        BuscarVelocidad = mejorVel
    End If
End Function

Private Function HayTramosTurno(ByVal loTramos As ListObject, ByVal segs As Collection) As Boolean
    Dim rng As Range
    Dim cSeg As Long
    Dim i As Long
    Dim seg As String

    If loTramos Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(loTramos, "id_turno_lote")
    If cSeg = 0 Then Exit Function

    Set rng = loTramos.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        seg = CStr(rng.Cells(i, cSeg).Value)
        If Pertenece(segs, seg) Then
            HayTramosTurno = True
            Exit Function
        End If
    Next i
End Function

Private Function ProduccionTeoricaPorTramos(ByVal loTurnos As ListObject, ByVal loParadas As ListObject, _
        ByVal loProduccion As ListObject, ByVal loTramos As ListObject, ByVal loVelocidades As ListObject, _
        ByVal id_turno As String, ByVal umbral As Double) As Double
    Dim rng As Range
    Dim cIdTurno As Long
    Dim cSeg As Long
    Dim cLote As Long
    Dim cFecha As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim i As Long
    Dim seg As String
    Dim lote As String
    Dim producto As String
    Dim ini As Date
    Dim fin As Date
    Dim fecha As Date
    Dim total As Double

    If loTurnos Is Nothing Then Exit Function

    cIdTurno = Utils.IndiceColumna(loTurnos, "id_turno")
    cSeg = Utils.IndiceColumna(loTurnos, "id_turno_lote")
    cLote = Utils.IndiceColumna(loTurnos, "id_lote")
    cFecha = Utils.IndiceColumna(loTurnos, "fecha")
    cIni = Utils.IndiceColumna(loTurnos, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(loTurnos, "fecha_hora_fin")
    If cIdTurno = 0 Then Exit Function

    Set rng = loTurnos.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cIdTurno).Value), id_turno, vbTextCompare) = 0 Then
            seg = ""
            lote = ""
            If cSeg > 0 Then seg = CStr(rng.Cells(i, cSeg).Value)
            If cLote > 0 Then lote = CStr(rng.Cells(i, cLote).Value)

            ini = 0
            fin = 0
            fecha = 0
            If cIni > 0 Then ini = Utils.AFecha(rng.Cells(i, cIni).Value)
            If cFin > 0 Then fin = Utils.AFecha(rng.Cells(i, cFin).Value)
            If cFecha > 0 Then fecha = Utils.AFecha(rng.Cells(i, cFecha).Value)

            If ini > 0 And fin > 0 Then
                If fin <= ini Then fin = fin + 1
                If fecha = 0 Then fecha = Int(ini)

                producto = ProductoDeSegmento(loProduccion, seg, lote, id_turno)
                If Len(producto) > 0 Then
                    total = total + TeoricoSegmento( _
                        loParadas, loTramos, loVelocidades, id_turno, seg, lote, _
                        producto, fecha, ini, fin, umbral)
                End If
            End If
        End If
    Next i

    ProduccionTeoricaPorTramos = total
End Function

Private Function ProductoDeSegmento(ByVal loProduccion As ListObject, ByVal seg As String, _
        ByVal lote As String, ByVal id_turno As String) As String
    Dim rng As Range
    Dim cSeg As Long
    Dim cLote As Long
    Dim cTurno As Long
    Dim cProducto As Long
    Dim i As Long
    Dim sSeg As String
    Dim sLote As String
    Dim sTurno As String
    Dim producto As String

    If loProduccion Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(loProduccion, "id_turno_lote")
    cLote = Utils.IndiceColumna(loProduccion, "lote")
    cTurno = Utils.IndiceColumna(loProduccion, "turno")
    cProducto = Utils.IndiceColumna(loProduccion, "producto")
    If cProducto = 0 Then Exit Function

    Set rng = loProduccion.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ""
        sLote = ""
        sTurno = ""
        If cSeg > 0 Then sSeg = CStr(rng.Cells(i, cSeg).Value)
        If cLote > 0 Then sLote = CStr(rng.Cells(i, cLote).Value)
        If cTurno > 0 Then sTurno = CStr(rng.Cells(i, cTurno).Value)

        If (Len(seg) > 0 And StrComp(sSeg, seg, vbTextCompare) = 0) _
           Or (Len(seg) = 0 And Len(lote) > 0 And StrComp(sLote, lote, vbTextCompare) = 0) _
           Or (Len(seg) = 0 And Len(lote) = 0 And StrComp(sTurno, id_turno, vbTextCompare) = 0) Then
            producto = CStr(rng.Cells(i, cProducto).Value)
            If Len(producto) > 0 Then
                ProductoDeSegmento = producto
                Exit Function
            End If
        End If
    Next i
End Function

Private Function TeoricoSegmento(ByVal loParadas As ListObject, ByVal loTramos As ListObject, _
        ByVal loVelocidades As ListObject, ByVal id_turno As String, ByVal seg As String, _
        ByVal lote As String, ByVal producto As String, ByVal fecha As Date, _
        ByVal ini As Date, ByVal fin As Date, ByVal umbral As Double) As Double
    Dim limites() As Double
    Dim n As Long
    Dim sinIni As Collection
    Dim sinFin As Collection
    Dim i As Long
    Dim intIni As Date
    Dim intFin As Date
    Dim medio As Date
    Dim cond As String
    Dim vel As Double
    Dim toTramo As Double
    Dim perdido As Double
    Dim total As Double
    Dim primero As Boolean

    On Error GoTo Fallo

    Set sinIni = New Collection
    Set sinFin = New Collection

    ReDim limites(0 To 1)
    limites(0) = CDbl(ini)
    limites(1) = CDbl(fin)
    n = 1

    CollectarTramos loTramos, seg, ini, fin, sinIni, sinFin, limites, n
    OrdenarArray limites, n

    primero = True
    For i = 0 To n - 1
        If limites(i + 1) > limites(i) Then
            intIni = CDate(limites(i))
            intFin = CDate(limites(i + 1))
            medio = CDate((limites(i) + limites(i + 1)) / 2)

            If PuntoEnTramos(medio, sinIni, sinFin) Then
                cond = "sin envase"
            Else
                cond = "con envase"
            End If

            vel = BuscarVelocidad(loVelocidades, producto, fecha, cond)
            If vel > 0 Then
                perdido = MinutosPerdidos(loParadas, id_turno, seg, lote, intIni, intFin, umbral, primero)
                toTramo = MinutosEntre(intIni, intFin) - perdido
                If toTramo < 0 Then toTramo = 0
                total = total + toTramo * vel
            End If
            primero = False
        End If
    Next i

    TeoricoSegmento = total
    Exit Function

Fallo:
    Debug.Print "Agregacion.TeoricoSegmento: " & Err.Description
    TeoricoSegmento = 0
End Function

Private Sub CollectarTramos(ByVal loTramos As ListObject, ByVal seg As String, _
        ByVal ini As Date, ByVal fin As Date, ByRef sinIni As Collection, _
        ByRef sinFin As Collection, ByRef limites() As Double, ByRef n As Long)
    Dim rng As Range
    Dim cSeg As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim i As Long
    Dim tIni As Date
    Dim tFin As Date
    Dim a As Date
    Dim z As Date

    If loTramos Is Nothing Then Exit Sub
    If Len(seg) = 0 Then Exit Sub

    cSeg = Utils.IndiceColumna(loTramos, "id_turno_lote")
    cIni = Utils.IndiceColumna(loTramos, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(loTramos, "fecha_hora_fin")
    If cSeg = 0 Or cIni = 0 Or cFin = 0 Then Exit Sub

    Set rng = loTramos.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cSeg).Value), seg, vbTextCompare) = 0 Then
            tIni = Utils.AFecha(rng.Cells(i, cIni).Value)
            tFin = Utils.AFecha(rng.Cells(i, cFin).Value)
            If tIni > 0 And tFin > 0 Then
                If tFin <= tIni Then tFin = tFin + 1
                a = tIni
                z = tFin
                If a < ini Then a = ini
                If z > fin Then z = fin
                If z > a Then
                    sinIni.Add CDbl(a)
                    sinFin.Add CDbl(z)
                    AgregarLimite limites, n, CDbl(a)
                    AgregarLimite limites, n, CDbl(z)
                End If
            End If
        End If
    Next i
End Sub

Private Sub AgregarLimite(ByRef limites() As Double, ByRef n As Long, ByVal valor As Double)
    n = n + 1
    ReDim Preserve limites(0 To n)
    limites(n) = valor
End Sub

Private Sub OrdenarArray(ByRef limites() As Double, ByVal n As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Double

    For i = 0 To n - 1
        For j = i + 1 To n
            If limites(j) < limites(i) Then
                tmp = limites(i)
                limites(i) = limites(j)
                limites(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function PuntoEnTramos(ByVal p As Date, ByVal sinIni As Collection, _
        ByVal sinFin As Collection) As Boolean
    Dim i As Long

    For i = 1 To sinIni.Count
        If CDbl(p) >= CDbl(sinIni(i)) And CDbl(p) < CDbl(sinFin(i)) Then
            PuntoEnTramos = True
            Exit Function
        End If
    Next i
End Function

Private Function MinutosPerdidos(ByVal loParadas As ListObject, ByVal id_turno As String, _
        ByVal seg As String, ByVal lote As String, ByVal intIni As Date, ByVal intFin As Date, _
        ByVal umbral As Double, ByVal esPrimerTramo As Boolean) As Double
    Dim rng As Range
    Dim cSeg As Long
    Dim cTurno As Long
    Dim cLote As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim cDur As Long
    Dim cTipo As Long
    Dim i As Long
    Dim pSeg As String
    Dim pTurno As String
    Dim pLote As String
    Dim tipo As String
    Dim pIni As Date
    Dim pFin As Date
    Dim dur As Double
    Dim overlap As Double
    Dim incluir As Boolean
    Dim a As Date
    Dim z As Date

    If loParadas Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(loParadas, "id_turno_lote")
    cTurno = Utils.IndiceColumna(loParadas, "turno")
    cLote = Utils.IndiceColumna(loParadas, "lote")
    cIni = Utils.IndiceColumna(loParadas, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(loParadas, "fecha_hora_fin")
    cDur = Utils.IndiceColumna(loParadas, "duracion_min")
    cTipo = Utils.IndiceColumna(loParadas, "tipo")

    Set rng = loParadas.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        pSeg = ""
        pTurno = ""
        pLote = ""
        If cSeg > 0 Then pSeg = CStr(rng.Cells(i, cSeg).Value)
        If cTurno > 0 Then pTurno = CStr(rng.Cells(i, cTurno).Value)
        If cLote > 0 Then pLote = CStr(rng.Cells(i, cLote).Value)

        incluir = False
        If Len(pSeg) > 0 Then
            incluir = (StrComp(pSeg, seg, vbTextCompare) = 0)
        ElseIf StrComp(pTurno, id_turno, vbTextCompare) = 0 Then
            incluir = True
        ElseIf Len(lote) > 0 And Len(pLote) > 0 Then
            incluir = (StrComp(pLote, lote, vbTextCompare) = 0)
        End If

        If incluir Then
            pIni = 0
            pFin = 0
            If cIni > 0 Then pIni = Utils.AFecha(rng.Cells(i, cIni).Value)
            If cFin > 0 Then pFin = Utils.AFecha(rng.Cells(i, cFin).Value)
            If pIni > 0 And pFin > 0 And pFin <= pIni Then pFin = pFin + 1

            dur = 0
            If cDur > 0 Then dur = Utils.NumeroSeguro(rng.Cells(i, cDur).Value)
            If dur <= 0 And pIni > 0 And pFin > 0 Then dur = MinutosEntre(pIni, pFin)

            tipo = ""
            If cTipo > 0 Then tipo = UCase$(Trim$(CStr(rng.Cells(i, cTipo).Value)))

            If Not (dur <= umbral Or tipo = "TP-05") Then
                If pIni > 0 And pFin > 0 Then
                    a = MaxDate(pIni, intIni)
                    z = MinDate(pFin, intFin)
                    If z > a Then
                        overlap = MinutosEntre(a, z)
                        If overlap > 0 Then MinutosPerdidos = MinutosPerdidos + overlap
                    End If
                ElseIf esPrimerTramo And dur > 0 Then
                    MinutosPerdidos = MinutosPerdidos + dur
                End If
            End If
        End If
    Next i
End Function

Private Function MinutosEntre(ByVal ini As Date, ByVal fin As Date) As Double
    Dim d As Double

    d = DateDiff("n", ini, fin)
    If d < 0 Then d = d + 1440
    MinutosEntre = d
End Function

Private Function MinDate(ByVal a As Date, ByVal b As Date) As Date
    If a < b Then
        MinDate = a
    Else
        MinDate = b
    End If
End Function

Private Function MaxDate(ByVal a As Date, ByVal b As Date) As Date
    If a > b Then
        MaxDate = a
    Else
        MaxDate = b
    End If
End Function
