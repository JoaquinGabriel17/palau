Attribute VB_Name = "DashboardDatos"
Option Explicit

' F3 D1 - Agregaciones por periodo para el dashboard (sin pantalla).
' Todas las funciones aceptan filtros opcionales: fechaDesde/fechaHasta
' (0 = todo el historial), turno_codigo, producto_codigo y linea_id.
'
' Regla RN-01: se suman tiempos y unidades de todos los reportes del
' periodo y recien despues se calculan los factores. Nunca se promedia
' el OEE ni los porcentajes entre turnos, reportes o productos.
'
' Exclusion de anulados (anulado = Si o estado_revision = Anulado):
' se apoya en los helpers Listar...PorPeriodo de AccesoHojas, que ya
' usan AnuladosDelPeriodo. tblPerdidasInsumo no tiene id_turno_lote,
' asi que la exclusion se hace combinando (fecha, turno) con los
' reportes anulados registrados.
'
' Nota (F3b/D2): el punto calcula Rendimiento = tiempo_ideal / TO
' directamente (suma de tiempo ideal sobre suma de TO). El ajuste por
' tramos sin envase (produccion teorica, RN-14/15) sigue aplicandose
' por turno en Agregacion; si el periodo incluyera tramos, ese ajuste
' se incorpora en D2.

Private Const UMBRAL_MICRO_DEFECTO As Double = 5
Private Const ESTADO_ANULADO As String = "Anulado"

Public Type TCascadaTiempos
    td As Double
    pp As Double
    pe As Double
    tpl As Double
    af As Double
    ac As Double
    to As Double
End Type

Public Type TDetallePeriodo
    td As Double
    tpl As Double
    to As Double
    tiempo_ideal As Double
    unidades_totales As Double
    unidades_buenas As Double
    unidades_rechazadas As Double
    disponibilidad As Variant
    rendimiento As Variant
    calidad As Variant
    oee As Double
    Encontrado As Boolean
End Type

Private Type TAcumPeriodo
    td As Double
    pp As Double
    pe As Double
    af As Double
    ac As Double
    micro As Double
    pnp As Double
    tiempoIdeal As Double
    unidadesBuenas As Double
    unidadesRechazadas As Double
    hayDatos As Boolean
End Type

' ==========================================================================
' 3. Cascada de tiempos del periodo
' ==========================================================================

Public Function ObtenerCascadaTiempos(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As TCascadaTiempos

    Dim acc As TAcumPeriodo

    AcumularPeriodo fechaDesde, fechaHasta, turno, producto, linea, acc

    With ObtenerCascadaTiempos
        .td = acc.td
        .pp = acc.pp
        .pe = acc.pe
        .tpl = acc.td - acc.pp - acc.pe
        If .tpl < 0 Then .tpl = 0
        .af = acc.af
        .ac = acc.ac
        .to = Calculos.CalcularTO(.tpl, acc.pnp)
        If .to < 0 Then .to = 0
    End With
End Function

' ==========================================================================
' 2. Detalle de factores del periodo (semilla de las cuatro tarjetas)
' ==========================================================================

Public Function ObtenerDetallePeriodo(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As TDetallePeriodo

    Dim acc As TAcumPeriodo

    AcumularPeriodo fechaDesde, fechaHasta, turno, producto, linea, acc

    With ObtenerDetallePeriodo
        .td = acc.td
        .tpl = acc.td - acc.pp - acc.pe
        If .tpl < 0 Then .tpl = 0
        .to = Calculos.CalcularTO(.tpl, acc.pnp)
        If .to < 0 Then .to = 0
        .tiempo_ideal = acc.tiempoIdeal
        .unidades_buenas = acc.unidadesBuenas
        .unidades_rechazadas = acc.unidadesRechazadas
        .unidades_totales = acc.unidadesBuenas + acc.unidadesRechazadas

        .disponibilidad = Calculos.CalcularDisponibilidad(.to, .tpl)
        If .to > 0 Then
            .rendimiento = .tiempo_ideal / .to
        Else
            .rendimiento = Empty
        End If
        .calidad = Calculos.CalcularCalidad(.unidades_buenas, .unidades_totales)
        .oee = Calculos.CalcularOEE( _
            Utils.NumeroSeguro(.disponibilidad), _
            Utils.NumeroSeguro(.rendimiento), _
            Utils.NumeroSeguro(.calidad))
        .Encontrado = (.td > 0 Or .unidades_totales > 0 Or acc.hayDatos)
    End With
End Function

' ==========================================================================
' 8. Tendencia de OEE por turno (un punto por turno, nunca promedio)
' ==========================================================================

Public Function ObtenerTendenciaTurnos(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As Collection

    Dim col As Collection
    Dim rep As Collection
    Dim item As Object
    Dim d As Object
    Dim pares As Object
    Dim v As Variant
    Dim clave As String
    Dim det As TDetallePeriodo
    Dim fecha As Date

    Set col = New Collection
    Set ObtenerTendenciaTurnos = col

    ' 1. Pares unicos (fecha, turno) con al menos un reporte no anulado.
    Set pares = NewDictionary()
    Set rep = AccesoHojas.ListarReportesPorPeriodo(fechaDesde, fechaHasta)
    For Each item In rep
        If CumpleFiltros(CStr(item("turno_codigo")), Empty, turno, linea) Then
            If Len(producto) = 0 Or StrComp(CStr(item("producto")), producto, vbTextCompare) = 0 Then
                fecha = AFechaSoloDia(item("fecha_hora_inicio"))
                If fecha > 0 Then
                    clave = ClaveFechaTurno(fecha, CStr(item("turno_codigo")))
                    If Not pares.Exists(clave) Then
                        Set d = NewDictionary()
                        d("fecha") = fecha
                        d("turno") = CStr(item("turno_codigo"))
                        pares.Add clave, d
                    End If
                End If
            End If
        End If
    Next item

    ' 2. Un punto por par: el detalle del propio turno, sumado recien ahi.
    For Each v In pares.Keys
        Set d = pares(v)
        det = ObtenerDetallePeriodo(CDate(d("fecha")), CDate(d("fecha")), _
            CStr(d("turno")), producto, linea)
        If det.td > 0 Then
            Set item = NewDictionary()
            item("fecha") = CDate(d("fecha"))
            item("fechaOrd") = CDbl(CDate(d("fecha")))
            item("turno_codigo") = CStr(d("turno"))
            item("oee") = det.oee
            item("disponibilidad") = det.disponibilidad
            item("rendimiento") = det.rendimiento
            item("calidad") = det.calidad
            col.Add item
        End If
    Next v

    OrdenarColeccion col, "fechaOrd", False
End Function

' ==========================================================================
' 4. Pareto de paradas del periodo (por clase, origen o motivo)
' ==========================================================================

Public Function ObtenerParetoParadas(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal agruparPor As String = "clase", _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As Collection

    Dim col As Collection
    Dim pars As Collection
    Dim grupos As Object
    Dim selSegs As Object
    Dim clases As Object
    Dim item As Object
    Dim d As Object
    Dim v As Variant
    Dim clave As String
    Dim origen As String
    Dim minutos As Double
    Dim total As Double

    Set col = New Collection
    Set ObtenerParetoParadas = col

    If StrComp(agruparPor, "motivo", vbTextCompare) <> 0 _
       And StrComp(agruparPor, "origen", vbTextCompare) <> 0 Then
        agruparPor = "clase"
    End If

    Set grupos = NewDictionary()
    Set clases = MapaClasesOrigen()
    Set pars = AccesoHojas.ListarParadasPorPeriodo(fechaDesde, fechaHasta)
    Set selSegs = SegsFiltrados(fechaDesde, fechaHasta, turno, producto, linea)

    For Each item In pars
        If CumpleFiltros("", item("linea_id"), "", linea) Then
            If EnSegSel(selSegs, CStr(item("id_turno_lote")), turno, producto) Then
                minutos = MinutosDeParada(item)
                If minutos > 0 Then
                    If StrComp(agruparPor, "motivo", vbTextCompare) = 0 Then
                        clave = TextoNoVacio(CStr(item("motivo")))
                    ElseIf StrComp(agruparPor, "origen", vbTextCompare) = 0 Then
                        clave = TextoNoVacio(CStr(item("origen")))
                    Else
                        origen = CStr(item("origen"))
                        If Len(Trim$(origen)) > 0 And clases.Exists(origen) Then
                            clave = Trim$(CStr(clases(origen)))
                        End If
                        If Len(clave) = 0 Then clave = TextoNoVacio(origen)
                    End If
                    If Not grupos.Exists(clave) Then grupos.Add clave, 0
                    grupos(clave) = CDbl(grupos(clave)) + minutos
                    total = total + minutos
                End If
            End If
        End If
    Next item

    For Each v In grupos.Keys
        Set d = NewDictionary()
        d("clave") = CStr(v)
        d("minutos") = CDbl(grupos(v))
        If total > 0 Then
            d("porcentaje") = CDbl(grupos(v)) / total
        Else
            d("porcentaje") = Empty
        End If
        col.Add d
    Next v

    OrdenarColeccion col, "minutos", True
End Function

' ==========================================================================
' 5. Acumulado de pequenas paradas (5 minutos o menos) por motivo
' ==========================================================================

Public Function ObtenerMicroparadas(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As Collection

    Dim col As Collection
    Dim pars As Collection
    Dim grupos As Object
    Dim selSegs As Object
    Dim item As Object
    Dim d As Object
    Dim v As Variant
    Dim clave As String
    Dim minutos As Double
    Dim umbral As Double

    Set col = New Collection
    Set ObtenerMicroparadas = col

    Set grupos = NewDictionary()
    Set pars = AccesoHojas.ListarParadasPorPeriodo(fechaDesde, fechaHasta)
    Set selSegs = SegsFiltrados(fechaDesde, fechaHasta, turno, producto, linea)
    umbral = UmbralMicro()

    For Each item In pars
        If CumpleFiltros("", item("linea_id"), "", linea) Then
            If EnSegSel(selSegs, CStr(item("id_turno_lote")), turno, producto) Then
                minutos = MinutosDeParada(item)
                If minutos > 0 Then
                    If minutos <= umbral Or StrComp(UCase$(Trim$(CStr(item("tipo")))), "TP-05", vbTextCompare) = 0 Then
                        clave = TextoNoVacio(CStr(item("motivo")))
                        If Not grupos.Exists(clave) Then
                            Set d = NewDictionary()
                            d("motivo") = clave
                            d("cantidad") = 0
                            d("minutos") = 0
                            grupos.Add clave, d
                        End If
                        Set d = grupos(clave)
                        d("cantidad") = CDbl(d("cantidad")) + 1
                        d("minutos") = CDbl(d("minutos")) + minutos
                    End If
                End If
            End If
        End If
    Next item

    For Each v In grupos.Keys
        col.Add grupos(v)
    Next v

    OrdenarColeccion col, "minutos", True
End Function

' ==========================================================================
' 6. Reducciones de velocidad del periodo
' ==========================================================================

Public Function ObtenerReducciones(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As Collection

    Dim col As Collection
    Dim red As Collection
    Dim item As Object
    Dim d As Object
    Dim selSegs As Object

    Set col = New Collection
    Set ObtenerReducciones = col

    ' tblReducciones no tiene producto ni linea_id: esos filtros no aplican
    ' aca; solo el turno y la exclusion de anulados (ya la hace el helper).
    ' El filtro por turno va por set de segmentos (ver nota en Acumular).
    Set red = AccesoHojas.ListarReduccionesPorPeriodo(fechaDesde, fechaHasta)
    Set selSegs = SegsFiltrados(fechaDesde, fechaHasta, turno, producto, linea)

    For Each item In red
        If EnSegSel(selSegs, CStr(item("id_turno_lote")), turno, producto) Then
            Set d = NewDictionary()
            d("origen") = TextoNoVacio(CStr(item("origen")))
            d("motivo") = TextoNoVacio(CStr(item("motivo")))
            d("minutos") = DuracionEntre(item("fecha_hora_inicio"), item("fecha_hora_fin"))
            col.Add d
        End If
    Next item

    OrdenarColeccion col, "minutos", True
End Function

' ==========================================================================
' 7. Litros producidos por periodo (suma, no promedio)
' ==========================================================================

Public Function ObtenerLitros(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As Double

    Dim prod As Collection
    Dim item As Object
    Dim selSegs As Object
    Dim total As Double

    Set prod = AccesoHojas.ListarProduccionPorPeriodo(fechaDesde, fechaHasta)
    Set selSegs = SegsFiltrados(fechaDesde, fechaHasta, turno, producto, linea)

    For Each item In prod
        If CumpleFiltros("", item("linea_id"), "", linea) Then
            If EnSegSel(selSegs, CStr(item("id_turno_lote")), turno, producto) _
               Or (Len(producto) > 0 And Len(turno) = 0 And StrComp(CStr(item("producto")), producto, vbTextCompare) = 0) Then
                total = total + Utils.NumeroSeguro(item("litros"))
            End If
        End If
    Next item

    ObtenerLitros = total
End Function

' ==========================================================================
' 9. Perdida de insumo por periodo
' ==========================================================================

Public Function ObtenerPerdidaInsumo(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        Optional ByVal turno As String = "", Optional ByVal producto As String = "", _
        Optional ByVal linea As String = "") As Collection

    Dim col As Collection
    Dim grupos As Object
    Dim d As Object
    Dim fila As Range
    Dim clave As String
    Dim v As Variant
    Dim lo As ListObject
    Dim rng As Range
    Dim cFecha As Long
    Dim cTurno As Long
    Dim cInsumo As Long
    Dim cCantidad As Long
    Dim cUnidad As Long
    Dim anulados As Object
    Dim sTurno As String
    Dim f As Date
    Dim i As Long

    Set col = New Collection
    Set ObtenerPerdidaInsumo = col

    ' La tabla no trae producto ni linea_id: el filtro de producto no aplica.
    Set grupos = NewDictionary()

    Set lo = Utils.Tabla("tblPerdidasInsumo")
    If lo Is Nothing Then Exit Function

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cTurno = Utils.IndiceColumna(lo, "turno")
    cInsumo = Utils.IndiceColumna(lo, "insumo")
    cCantidad = Utils.IndiceColumna(lo, "cantidad")
    cUnidad = Utils.IndiceColumna(lo, "unidad")
    If cInsumo = 0 Or cCantidad = 0 Then Exit Function

    Set anulados = AnuladosPorFechaTurno(fechaDesde, fechaHasta)

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        Set fila = rng.Cells(i, 1)
        f = Utils.AFecha(fila.Cells(1, cFecha).Value)
        If Not EnPeriodo(f, fechaDesde, fechaHasta) Then GoTo SeguirFila
        sTurno = TextoNoVacio(CStr(fila.Cells(1, cTurno).Value))
        If anulados.Exists(ClaveFechaTurno(f, sTurno)) Then GoTo SeguirFila
        If Len(turno) > 0 And StrComp(sTurno, turno, vbTextCompare) <> 0 Then GoTo SeguirFila

        clave = TextoNoVacio(CStr(fila.Cells(1, cInsumo).Value)) & ChrW(166) & _
                TextoNoVacio(CStr(fila.Cells(1, cUnidad).Value))
        If Not grupos.Exists(clave) Then
            Set d = NewDictionary()
            d("insumo") = TextoNoVacio(CStr(fila.Cells(1, cInsumo).Value))
            d("unidad") = TextoNoVacio(CStr(fila.Cells(1, cUnidad).Value))
            d("cantidad") = 0
            grupos.Add clave, d
        End If
        Set d = grupos(clave)
        d("cantidad") = CDbl(d("cantidad")) + Utils.NumeroSeguro(fila.Cells(1, cCantidad).Value)
SeguirFila:
    Next i

    For Each v In grupos.Keys
        col.Add grupos(v)
    Next v
End Function

' ==========================================================================
' Nucleo: acumulacion por periodo (sumas crudas, sin factores)
' ==========================================================================

Private Sub AcumularPeriodo(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        ByVal turno As String, ByVal producto As String, ByVal linea As String, _
        ByRef acc As TAcumPeriodo)

    Dim rep As Collection
    Dim prod As Collection
    Dim rech As Collection
    Dim item As Object
    Dim prodLote As Object
    Dim selSegs As Object
    Dim velSeg As Object
    Dim velLote As Object
    Dim umbral As Double
    Dim dur As Double
    Dim vel As Double
    Dim seg As String
    Dim lote As String

    umbral = UmbralMicro()

    ' 1. Reportes: TD por registro (horario anotado, Q-22/RN-17).
    '    Arma dos mapas: lote -> producto y el set de segmentos
    '    (id_turno_lote) que cumplen los filtros, para filtrar las tablas
    '    hijas. Ojo: tblParadas/tblProduccion/tblRechazos guardan en la
    '    columna "turno" el ID del turno (ej. 498), no el codigo (RES-T),
    '    asi que el filtro por turno se resuelve con el set.
    Set prodLote = NewDictionary()
    Set selSegs = NewDictionary()
    Set rep = AccesoHojas.ListarReportesPorPeriodo(fechaDesde, fechaHasta)
    For Each item In rep
        If Len(CStr(item("id_lote"))) > 0 Then
            If Not prodLote.Exists(CStr(item("id_lote"))) Then
                prodLote.Add CStr(item("id_lote")), CStr(item("producto"))
            End If
        End If
        If CumpleFiltros(CStr(item("turno_codigo")), Empty, turno, linea) Then
            If Len(producto) = 0 Or StrComp(CStr(item("producto")), producto, vbTextCompare) = 0 Then
                dur = DuracionEntre(item("fecha_hora_inicio"), item("fecha_hora_fin"))
                If dur > 0 Then
                    acc.td = acc.td + dur
                    acc.hayDatos = True
                End If
                If Len(CStr(item("id_turno_lote"))) > 0 Then selSegs(CStr(item("id_turno_lote"))) = True
            End If
        End If
    Next item

    ' 2. Paradas: misma clasificacion que Agregacion (pp/pe/af/ac/micro/pnp).
    '    micro queda fuera de pnp: no se resta del TO.
    For Each item In AccesoHojas.ListarParadasPorPeriodo(fechaDesde, fechaHasta)
        If CumpleFiltros("", item("linea_id"), "", linea) Then
            If EnSegSel(selSegs, CStr(item("id_turno_lote")), turno, producto) _
               Or CumpleProducto(prodLote, CStr(item("id_lote")), producto) And Len(turno) = 0 Then
                dur = MinutosDeParada(item)
                If dur > 0 Then
                    acc.hayDatos = True
                    Select Case UCase$(Trim$(CStr(item("tipo"))))
                    Case "TP-01"
                        If dur <= umbral Then
                            acc.micro = acc.micro + dur
                        Else
                            acc.pp = acc.pp + dur
                        End If
                    Case "TP-07"
                        If dur <= umbral Then
                            acc.micro = acc.micro + dur
                        Else
                            acc.pe = acc.pe + dur
                        End If
                    Case "TP-03"
                        If dur <= umbral Then
                            acc.micro = acc.micro + dur
                        Else
                            acc.af = acc.af + dur
                            acc.pnp = acc.pnp + dur
                        End If
                    Case "TP-04"
                        If dur <= umbral Then
                            acc.micro = acc.micro + dur
                        Else
                            acc.ac = acc.ac + dur
                            acc.pnp = acc.pnp + dur
                        End If
                    Case "TP-05"
                        acc.micro = acc.micro + dur
                    Case Else
                        If dur <= umbral Then
                            acc.micro = acc.micro + dur
                        Else
                            acc.pnp = acc.pnp + dur
                        End If
                    End Select
                End If
            End If
        End If
    Next item

    ' 3. Produccion: unidades y tiempo ideal (unidades / velocidad nominal).
    Set velSeg = NewDictionary()
    Set velLote = NewDictionary()
    Set prod = AccesoHojas.ListarProduccionPorPeriodo(fechaDesde, fechaHasta)
    For Each item In prod
        seg = CStr(item("id_turno_lote"))
        lote = CStr(item("lote"))
        If CumpleFiltros("", item("linea_id"), "", linea) Then
            If EnSegSel(selSegs, seg, turno, producto) Or (CumpleProducto(prodLote, lote, producto) And Len(turno) = 0) Then
                vel = Utils.NumeroSeguro(item("velocidad_nominal"))
                If vel <= 0 Then
                    vel = Agregacion.VelocidadVigente(CStr(item("producto")), AFecha(item("fecha")))
                End If
                If vel > 0 Then
                    If Len(seg) > 0 Then
                        If Not velSeg.Exists(seg) Then velSeg.Add seg, vel
                    End If
                    If Len(lote) > 0 Then
                        If Not velLote.Exists(lote) Then velLote.Add lote, vel
                    End If
                End If

                If Utils.NumeroSeguro(item("unidades_producidas")) <> 0 Then
                    acc.unidadesBuenas = acc.unidadesBuenas + Utils.NumeroSeguro(item("unidades_producidas"))
                    acc.hayDatos = True
                End If
                If vel > 0 And Utils.NumeroSeguro(item("unidades_producidas")) > 0 Then
                    acc.tiempoIdeal = acc.tiempoIdeal + Utils.NumeroSeguro(item("unidades_producidas")) / vel
                End If
            End If
        End If
    Next item

    ' 4. Rechazos: unidades rechazadas y su tiempo ideal (rechazo / velocidad).
    For Each item In AccesoHojas.ListarRechazosPorPeriodo(fechaDesde, fechaHasta)
        If EnSegSel(selSegs, CStr(item("id_turno_lote")), turno, producto) _
           Or (CumpleProducto(prodLote, CStr(item("lote")), producto) And Len(turno) = 0) Then
            seg = CStr(item("id_turno_lote"))
            lote = CStr(item("lote"))
            If velSeg.Exists(seg) Then
                vel = CDbl(velSeg(seg))
            ElseIf velLote.Exists(lote) Then
                vel = CDbl(velLote(lote))
            Else
                vel = 0
            End If

            If Utils.NumeroSeguro(item("cantidad")) <> 0 Then
                acc.unidadesRechazadas = acc.unidadesRechazadas + Utils.NumeroSeguro(item("cantidad"))
                acc.hayDatos = True
            End If
            If vel > 0 And Utils.NumeroSeguro(item("cantidad")) > 0 Then
                acc.tiempoIdeal = acc.tiempoIdeal + Utils.NumeroSeguro(item("cantidad")) / vel
            End If
        End If
    Next item
End Sub

' ==========================================================================
' Helpers internos
' ==========================================================================

Private Function NewDictionary() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NewDictionary = d
End Function

' Mapa codigo de origen -> clase legible (tblOrigenes.clase), armado una vez
' por consulta para no golpear la tabla en cada parada.
Private Function MapaClasesOrigen() As Object
    Dim d As Object
    Dim item As Object

    Set d = NewDictionary()
    Set MapaClasesOrigen = d

    For Each item In AccesoHojas.ListarOrigenes()
        If Len(Trim$(CStr(item("codigo")))) > 0 Then
            d(CStr(item("codigo"))) = CStr(item("clase"))
        End If
    Next item
End Function

Private Function ProductoDeClave(ByVal mapa As Object, ByVal clave As String) As String
    If mapa.Exists(clave) Then ProductoDeClave = CStr(mapa(clave))
End Function

Private Function CumpleProducto(ByVal mapa As Object, ByVal clave As String, ByVal producto As String) As Boolean
    Dim p As String

    If Len(producto) = 0 Then
        CumpleProducto = True
        Exit Function
    End If
    p = ProductoDeClave(mapa, clave)
    If Len(p) > 0 Then CumpleProducto = (StrComp(p, producto, vbTextCompare) = 0)
End Function

Private Function CumpleFiltros(ByVal turnoval As String, ByVal lineaval As Variant, _
        ByVal turno As String, ByVal linea As String) As Boolean

    If Len(turno) > 0 Then
        If StrComp(turnoval, turno, vbTextCompare) <> 0 Then Exit Function
    End If
    If Len(linea) > 0 Then
        If StrComp(CStr(lineaval), linea, vbTextCompare) <> 0 Then Exit Function
    End If
    CumpleFiltros = True
End Function

' Set de segmentos (id_turno_lote) de los reportes que cumplen los filtros
' de turno, producto y linea. Con filtros vacios todo entra.
Private Function SegsFiltrados(ByVal fechaDesde As Date, ByVal fechaHasta As Date, _
        ByVal turno As String, ByVal producto As String, ByVal linea As String) As Object

    Dim selSegs As Object
    Set selSegs = NewDictionary()
    Set SegsFiltrados = selSegs

    Dim rep As Collection
    Dim item As Object

    Set rep = AccesoHojas.ListarReportesPorPeriodo(fechaDesde, fechaHasta)
    For Each item In rep
        If Len(CStr(item("id_turno_lote"))) > 0 Then
            If CumpleFiltros(CStr(item("turno_codigo")), Empty, turno, linea) Then
                If Len(producto) = 0 Or StrComp(CStr(item("producto")), producto, vbTextCompare) = 0 Then
                    selSegs(CStr(item("id_turno_lote"))) = True
                End If
            End If
        End If
    Next item
End Function

' Sin filtros de turno/producto, todo segmento pasa. Con filtros, solo
' los segmentos del set (las tablas hijas guardan en "turno" el ID,
' no el codigo de turno).
Private Function EnSegSel(ByVal selSegs As Object, ByVal seg As String, _
        ByVal turno As String, ByVal producto As String) As Boolean

    If Len(turno) = 0 And Len(producto) = 0 Then
        EnSegSel = True
        Exit Function
    End If
    If Len(seg) = 0 Then Exit Function
    EnSegSel = selSegs.Exists(seg)
End Function

Private Function MinutosDeParada(ByVal item As Object) As Double
    Dim dur As Double

    dur = Utils.NumeroSeguro(item("duracion_min"))
    If dur <= 0 Then
        dur = DuracionEntre(item("fecha_hora_inicio"), item("fecha_hora_fin"))
    End If
    MinutosDeParada = dur
End Function

Private Function DuracionEntre(ByVal ini As Variant, ByVal fin As Variant) As Double
    Dim fIni As Date
    Dim fFin As Date
    Dim d As Double

    If IsDate(ini) Then fIni = CDate(ini)
    If IsDate(fin) Then fFin = CDate(fin)
    If fIni <= 0 Or fFin <= 0 Then Exit Function

    d = DateDiff("n", fIni, fFin)
    If d < 0 Then d = d + 1440
    DuracionEntre = d
End Function

Private Function AFechaSoloDia(ByVal v As Variant) As Date
    Dim f As Date

    f = Utils.AFecha(v)
    If f > 0 Then AFechaSoloDia = Int(f)
End Function

Private Function UmbralMicro() As Double
    Dim v As Variant

    UmbralMicro = UMBRAL_MICRO_DEFECTO
    v = AccesoHojas.ObtenerParametro("UMBRAL_MICROPARADA_MIN", 0)
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then UmbralMicro = CDbl(v)
    End If
End Function

Private Function TextoNoVacio(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) = 0 Then s = "(sin dato)"
    TextoNoVacio = s
End Function

Private Function EnPeriodo(ByVal fecha As Date, ByVal fechaDesde As Date, ByVal fechaHasta As Date) As Boolean
    If fechaDesde = 0 And fechaHasta = 0 Then
        EnPeriodo = True
        Exit Function
    End If
    If fecha <= 0 Then Exit Function
    If fechaDesde > 0 And Int(fecha) < Int(fechaDesde) Then Exit Function
    If fechaHasta > 0 And Int(fecha) > Int(fechaHasta) Then Exit Function
    EnPeriodo = True
End Function

Private Function ClaveFechaTurno(ByVal fecha As Date, ByVal turno As String) As String
    ClaveFechaTurno = Format$(Int(fecha), "yyyy-mm-dd") & "|" & turno
End Function

' Reportes anulados (anulado=Si o estado_revision=Anulado) como combinaciones
' fecha+turno, para excluir filas de tblPerdidasInsumo (no tiene
' id_turno_lote). Mismo criterio que AnuladosDelPeriodo en AccesoHojas.
Private Function AnuladosPorFechaTurno(ByVal fechaDesde As Date, ByVal fechaHasta As Date) As Object
    Dim d As Object
    Dim lo As ListObject
    Dim rng As Range
    Dim cFecha As Long
    Dim cTurno As Long
    Dim cAnulado As Long
    Dim cEstado As Long
    Dim i As Long
    Dim f As Date
    Dim sTurno As String
    Dim enRango As Boolean
    Dim esAnulado As Boolean

    Set d = NewDictionary()
    Set AnuladosPorFechaTurno = d

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cTurno = Utils.IndiceColumna(lo, "turno_codigo")
    cAnulado = Utils.IndiceColumna(lo, "anulado")
    cEstado = Utils.IndiceColumna(lo, "estado_revision")
    If cAnulado = 0 And cEstado = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        esAnulado = (cAnulado > 0 And Utils.EsAnulado(rng.Cells(i, cAnulado).Value)) _
                    Or (cEstado > 0 And StrComp(Trim$(CStr(rng.Cells(i, cEstado).Value)), ESTADO_ANULADO, vbTextCompare) = 0)
        If esAnulado Then
            If fechaDesde = 0 And fechaHasta = 0 Then
                enRango = True
            Else
                f = 0
                If cFecha > 0 Then f = Utils.AFecha(rng.Cells(i, cFecha).Value)
                enRango = EnPeriodo(f, fechaDesde, fechaHasta)
            End If
            If enRango And cFecha > 0 Then
                f = Utils.AFecha(rng.Cells(i, cFecha).Value)
                sTurno = ""
                If cTurno > 0 Then sTurno = TextoNoVacio(CStr(rng.Cells(i, cTurno).Value))
                If f > 0 Then d(ClaveFechaTurno(f, sTurno)) = True
            End If
        End If
    Next i
End Function

' ==========================================================================
' Orden de los resultados
' ==========================================================================

' Orden estable: seleccion del extremo (o maximo o minimo segun
' descendente) en pasadas sucesivas, copiando a una coleccion nueva.
Private Sub OrdenarColeccion(ByVal col As Collection, ByVal clave As String, _
        Optional ByVal descendente As Boolean = True)

    Dim orden As Collection
    Dim i As Long
    Dim j As Long
    Dim mejor As Long
    Dim vi As Double
    Dim vm As Double

    Set orden = New Collection
    Do While col.Count > 0
        mejor = 1
        For j = 2 To col.Count
            vi = Utils.NumeroSeguro(col(j)(clave))
            vm = Utils.NumeroSeguro(col(mejor)(clave))
            If descendente Then
                If vi > vm Then mejor = j
            Else
                If vi < vm Then mejor = j
            End If
        Next j
        orden.Add col(mejor)
        col.Remove mejor
    Loop

    For i = 1 To orden.Count
        col.Add orden(i)
    Next i
End Sub
