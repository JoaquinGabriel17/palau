Attribute VB_Name = "Dashboard"
Option Explicit

' F3 D2 - Vista de turno del dashboard.
' Este modulo solo dibuja: lee los filtros de la hoja 21_Dashboard, pide los
' numeros a DashboardDatos y escribe en la hoja. Nunca calcula OEE ni factores
' (RN-01: no se promedia el OEE; la agregacion vive en DashboardDatos).
'
' Los umbrales del semaforo se leen de tblParametros (P-20 a P-27) y se
' expresan en porcentaje. Las tarjetas trabajan con fracciones (0,614 = 61,4%).

Private Const HOJA_DASHBOARD As String = "21_Dashboard"
Private Const ESTADO_PENDIENTE As String = "Pendiente de revisar"

Public Const COLOR_VERDE As Long = 13561798      ' RGB(198, 239, 206)
Public Const COLOR_AMARILLO As Long = 10284031   ' RGB(255, 235, 156)
Public Const COLOR_ROJO As Long = 13551615       ' RGB(255, 199, 206)
Public Const COLOR_SIN_DATO As Long = 15921906   ' RGB(242, 242, 242)

Public Type TDatosFiltros
    fechaDesde As Date
    fechaHasta As Date
    turno_codigo As String
    producto_codigo As String
    linea_id As String
End Type

' ==========================================================================
' 1. Macro principal: lee, pide los numeros y dibuja
' ==========================================================================

Public Sub ActualizarDashboard()
    Dim ws As Worksheet
    Dim filtros As TDatosFiltros
    Dim detalle As TDetallePeriodo
    Dim cascada As TCascadaTiempos
    Dim pareto As Collection
    Dim micro As Collection
    Dim reducciones As Collection
    Dim litros As Double
    Dim protegida As Boolean
    Dim pendientes As Boolean
    Dim lineaFiltro As String
    Dim sErr As String

    On Error GoTo Fallo

    Set ws = ThisWorkbook.Worksheets(HOJA_DASHBOARD)
    protegida = ws.ProtectContents
    If protegida Then ws.Unprotect

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    filtros = LeerSelectores()

    ' Hoy existe una sola linea (Linea 1). La tabla de reportes no guarda
    ' linea_id, asi que pasar "1" dejaria TD en cero; con una sola linea el
    ' filtro no aporta nada y se omite. Con mas lineas, se pasa el numero.
    lineaFiltro = filtros.linea_id
    If lineaFiltro = "1" Then lineaFiltro = ""

    detalle = DashboardDatos.ObtenerDetallePeriodo( _
        filtros.fechaDesde, filtros.fechaHasta, _
        filtros.turno_codigo, filtros.producto_codigo, lineaFiltro)

    cascada = DashboardDatos.ObtenerCascadaTiempos( _
        filtros.fechaDesde, filtros.fechaHasta, _
        filtros.turno_codigo, filtros.producto_codigo, lineaFiltro)

    Set pareto = DashboardDatos.ObtenerParetoParadas( _
        filtros.fechaDesde, filtros.fechaHasta, "origen", _
        filtros.turno_codigo, filtros.producto_codigo, lineaFiltro)
    Set micro = DashboardDatos.ObtenerMicroparadas( _
        filtros.fechaDesde, filtros.fechaHasta, _
        filtros.turno_codigo, filtros.producto_codigo, lineaFiltro)
    Set reducciones = DashboardDatos.ObtenerReducciones( _
        filtros.fechaDesde, filtros.fechaHasta, _
        filtros.turno_codigo, filtros.producto_codigo, lineaFiltro)
    litros = DashboardDatos.ObtenerLitros( _
        filtros.fechaDesde, filtros.fechaHasta, _
        filtros.turno_codigo, filtros.producto_codigo, lineaFiltro)

    If detalle.Encontrado Then
        EscribirTarjeta ws.Range("B7"), detalle.oee, "0.0%"
        EscribirTarjeta ws.Range("D7"), detalle.disponibilidad, "0.0%"
        EscribirTarjeta ws.Range("F7"), detalle.rendimiento, "0.0%"
        EscribirTarjeta ws.Range("H7"), detalle.calidad, "0.00%"
    Else
        EscribirTarjeta ws.Range("B7"), Empty, "0.0%"
        EscribirTarjeta ws.Range("D7"), Empty, "0.0%"
        EscribirTarjeta ws.Range("F7"), Empty, "0.0%"
        EscribirTarjeta ws.Range("H7"), Empty, "0.0%"
    End If

    EscribirCascada ws.Range("B12"), cascada
    EscribirPareto ws.Range("A24"), pareto
    EscribirMicroparadas ws.Range("A34"), micro
    EscribirReducciones ws.Range("A42"), reducciones
    EscribirLitros ws.Range("B49"), litros
    MarcarExcesos ws, detalle, detalle.Encontrado

    pendientes = MarcarPendientes(filtros.fechaDesde, filtros.fechaHasta)
    If pendientes Then
        ws.Range("B3").Value = IconoAlerta() & " Hay reportes pendientes de revisar"
    Else
        ws.Range("B3").Value = ""
    End If

    ws.Range("A2").Value = "Actualizado: " & Format$(Now, "dd/mm/yyyy hh:mm")

Cierre:
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If protegida Then ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    On Error GoTo 0
    Exit Sub

Fallo:
    sErr = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Range("B7").Value = Guion()
        ws.Range("D7").Value = Guion()
        ws.Range("F7").Value = Guion()
        ws.Range("H7").Value = Guion()
        ws.Range("A24").Value = Guion()
        ws.Range("A34").Value = Guion()
        ws.Range("A42").Value = Guion()
        ws.Range("B49").Value = Guion()
        ws.Range("A2").Value = "Error al actualizar el dashboard: " & sErr
    End If
    Resume Cierre
End Sub

' ==========================================================================
' 2. Selectores de la hoja
' ==========================================================================

Public Function LeerSelectores() As TDatosFiltros
    Dim ws As Worksheet
    Dim f As TDatosFiltros

    Set ws = ThisWorkbook.Worksheets(HOJA_DASHBOARD)

    f.fechaDesde = FechaCelda(ws.Range("B4"))
    f.fechaHasta = FechaCelda(ws.Range("D4"))
    If f.fechaDesde = 0 Then f.fechaDesde = f.fechaHasta
    If f.fechaHasta = 0 Then f.fechaHasta = f.fechaDesde
    If f.fechaDesde > 0 And f.fechaHasta > 0 Then
        If f.fechaDesde > f.fechaHasta Then
            f.fechaDesde = f.fechaHasta
            f.fechaHasta = FechaCelda(ws.Range("B4"))
        End If
    End If

    f.turno_codigo = TextoCelda(ws.Range("F4"))
    f.producto_codigo = NormalizarProducto(TextoCelda(ws.Range("H4")))
    f.linea_id = NormalizarLinea(TextoCelda(ws.Range("J4")))

    LeerSelectores = f
End Function

' ==========================================================================
' 3. Semaforo segun tblParametros (P-20 a P-27)
' ==========================================================================

Public Function AplicarSemaforo(ByVal valor As Double, ByVal tipo As String) As Long
    Dim v As Double
    Dim verde As Double
    Dim amarillo As Double

    v = valor
    If v > 1 Then v = v / 100

    Select Case UCase$(Trim$(tipo))
    Case "OEE"
        verde = Umbral("SEMAFORO_OEE_VERDE", 0.75)
        amarillo = Umbral("SEMAFORO_OEE_AMARILLO", 0.6)
    Case "DISPONIBILIDAD", "DISP"
        verde = Umbral("SEMAFORO_DISP_VERDE", 0.85)
        amarillo = Umbral("SEMAFORO_DISP_AMARILLO", 0.7)
    Case "RENDIMIENTO", "REND"
        verde = Umbral("SEMAFORO_REND_VERDE", 0.85)
        amarillo = Umbral("SEMAFORO_REND_AMARILLO", 0.7)
    Case "CALIDAD", "CAL"
        verde = Umbral("SEMAFORO_CAL_VERDE", 0.98)
        amarillo = Umbral("SEMAFORO_CAL_AMARILLO", 0.95)
    Case Else
        verde = 0.75
        amarillo = 0.6
    End Select

    If v >= verde Then
        AplicarSemaforo = COLOR_VERDE
    ElseIf v >= amarillo Then
        AplicarSemaforo = COLOR_AMARILLO
    Else
        AplicarSemaforo = COLOR_ROJO
    End If
End Function

' ==========================================================================
' 4. Una tarjeta
' ==========================================================================

Public Sub EscribirTarjeta(ByVal celda As Range, ByVal valor As Variant, _
        ByVal formato As String)

    Dim tipo As String
    Dim color As Long
    Dim rango As Range
    Dim etiqueta As Range

    On Error GoTo Fallo

    tipo = TipoTarjeta(celda)
    Set rango = celda.MergeArea
    Set etiqueta = celda.Offset(-1, 0).MergeArea

    If SinValor(valor) Then
        rango.Value = Guion()
        rango.NumberFormat = "@"
        rango.Font.Color = RGB(89, 89, 89)
        rango.Interior.Color = COLOR_SIN_DATO
        etiqueta.Interior.Color = COLOR_SIN_DATO
        Exit Sub
    End If

    color = AplicarSemaforo(CDbl(valor), tipo)
    rango.Value = CDbl(valor)
    rango.NumberFormat = formato
    rango.Font.Color = RGB(0, 0, 0)
    rango.Interior.Color = color
    etiqueta.Interior.Color = color
    Exit Sub

Fallo:
    On Error Resume Next
    celda.MergeArea.Value = Guion()
    celda.MergeArea.NumberFormat = "@"
    celda.MergeArea.Interior.Color = COLOR_SIN_DATO
    celda.Offset(-1, 0).MergeArea.Interior.Color = COLOR_SIN_DATO
End Sub

' ==========================================================================
' 5. Cascada de tiempos
' ==========================================================================

Public Sub EscribirCascada(ByVal celdaInicio As Range, casc As TCascadaTiempos)
    Dim valores(1 To 7) As Double
    Dim i As Long
    Dim sinDatos As Boolean

    valores(1) = casc.td
    valores(2) = casc.pp
    valores(3) = casc.pe
    valores(4) = casc.tpl
    valores(5) = casc.af
    valores(6) = casc.ac
    valores(7) = casc.to

    sinDatos = (casc.td = 0 And casc.pp = 0 And casc.pe = 0 And _
                casc.tpl = 0 And casc.af = 0 And casc.ac = 0 And casc.to = 0)

    On Error GoTo Fallo

    For i = 1 To 7
        If sinDatos Then
            celdaInicio.Offset(i - 1, 0).Value = Guion()
            celdaInicio.Offset(i - 1, 0).NumberFormat = "@"
        Else
            celdaInicio.Offset(i - 1, 0).Value = valores(i)
            celdaInicio.Offset(i - 1, 0).NumberFormat = "0.0"
        End If
    Next i
    Exit Sub

Fallo:
    On Error Resume Next
    For i = 1 To 7
        celdaInicio.Offset(i - 1, 0).Value = Guion()
    Next i
End Sub

' ==========================================================================
' 5b. Widgets de perdidas: Pareto, microparadas, reducciones y litros
' ==========================================================================

Public Sub EscribirPareto(ByVal celdaInicio As Range, pareto As Collection)
    Dim clases As Object
    Dim origen As Object
    Dim item As Object
    Dim i As Long
    Dim acum As Double
    Dim clase As String

    On Error GoTo Fallo

    Set clases = NuevoDiccionario()
    For Each origen In AccesoHojas.ListarOrigenes()
        clases(CStr(origen("codigo"))) = CStr(origen("clase"))
    Next origen

    celdaInicio.Resize(7, 3).ClearContents

    If pareto Is Nothing Then GoTo SinDatos
    If pareto.Count = 0 Then GoTo SinDatos

    acum = 0
    For i = 1 To pareto.Count
        Set item = pareto(i)

        clase = CStr(item("clave"))
        If clases.Exists(clase) Then
            If Len(Trim$(CStr(clases(clase)))) > 0 Then clase = CStr(clases(clase))
        End If

        celdaInicio.Offset(i - 1, 0).Value = clase
        celdaInicio.Offset(i - 1, 1).Value = Utils.NumeroSeguro(item("minutos")) / 60
        celdaInicio.Offset(i - 1, 1).NumberFormat = "0.0"
        acum = acum + Utils.NumeroSeguro(item("porcentaje"))
        celdaInicio.Offset(i - 1, 2).Value = acum
        celdaInicio.Offset(i - 1, 2).NumberFormat = "0.0%"
    Next i
    Exit Sub

SinDatos:
    celdaInicio.Value = Guion()
    celdaInicio.NumberFormat = "@"
    Exit Sub

Fallo:
    On Error Resume Next
    celdaInicio.Resize(7, 3).ClearContents
    celdaInicio.Value = Guion()
    celdaInicio.NumberFormat = "@"
End Sub

Public Sub EscribirMicroparadas(ByVal celdaInicio As Range, micro As Collection)
    Dim item As Object
    Dim i As Long

    On Error GoTo Fallo

    celdaInicio.Resize(6, 3).ClearContents

    If micro Is Nothing Then GoTo SinDatos
    If micro.Count = 0 Then GoTo SinDatos

    For i = 1 To micro.Count
        Set item = micro(i)
        celdaInicio.Offset(i - 1, 0).Value = CStr(item("motivo"))
        celdaInicio.Offset(i - 1, 1).Value = Utils.NumeroSeguro(item("cantidad"))
        celdaInicio.Offset(i - 1, 1).NumberFormat = "0"
        celdaInicio.Offset(i - 1, 2).Value = Utils.NumeroSeguro(item("minutos"))
        celdaInicio.Offset(i - 1, 2).NumberFormat = "0.0"
    Next i
    Exit Sub

SinDatos:
    celdaInicio.Value = Guion()
    celdaInicio.NumberFormat = "@"
    Exit Sub

Fallo:
    On Error Resume Next
    celdaInicio.Resize(6, 3).ClearContents
    celdaInicio.Value = Guion()
    celdaInicio.NumberFormat = "@"
End Sub

Public Sub EscribirReducciones(ByVal celdaInicio As Range, red As Collection)
    Dim item As Object
    Dim i As Long

    On Error GoTo Fallo

    celdaInicio.Resize(6, 3).ClearContents

    If red Is Nothing Then GoTo SinDatos
    If red.Count = 0 Then GoTo SinDatos

    For i = 1 To red.Count
        Set item = red(i)
        celdaInicio.Offset(i - 1, 0).Value = CStr(item("origen"))
        celdaInicio.Offset(i - 1, 1).Value = CStr(item("motivo"))
        celdaInicio.Offset(i - 1, 2).Value = Utils.NumeroSeguro(item("minutos"))
        celdaInicio.Offset(i - 1, 2).NumberFormat = "0.0"
    Next i
    Exit Sub

SinDatos:
    celdaInicio.Value = Guion()
    celdaInicio.NumberFormat = "@"
    Exit Sub

Fallo:
    On Error Resume Next
    celdaInicio.Resize(6, 3).ClearContents
    celdaInicio.Value = Guion()
    celdaInicio.NumberFormat = "@"
End Sub

Public Sub EscribirLitros(ByVal celda As Range, ByVal litros As Double)
    On Error GoTo Fallo

    If litros <= 0 Then
        celda.Value = Guion()
        celda.NumberFormat = "@"
        celda.Font.Color = RGB(89, 89, 89)
        Exit Sub
    End If

    celda.Value = litros
    celda.NumberFormat = "#,##0 ""L"""
    celda.Font.Color = RGB(0, 0, 0)
    Exit Sub

Fallo:
    On Error Resume Next
    celda.Value = Guion()
    celda.NumberFormat = "@"
End Sub

' ==========================================================================
' 6. Reportes pendientes de revisar en el periodo
' ==========================================================================

Public Function MarcarPendientes(ByVal fechaDesde As Date, ByVal fechaHasta As Date) As Boolean
    Dim reportes As Collection
    Dim item As Object

    On Error GoTo Fallo

    Set reportes = AccesoHojas.ListarReportesPorPeriodo(fechaDesde, fechaHasta)

    For Each item In reportes
        If StrComp(Trim$(CStr(item("estado_revision"))), ESTADO_PENDIENTE, vbTextCompare) = 0 Then
            MarcarPendientes = True
            Exit Function
        End If
    Next item

    Exit Function

Fallo:
    MarcarPendientes = False
End Function

' ==========================================================================
' 7. Handler de cambio de selectores (desde Worksheet_Change)
' ==========================================================================

Public Sub OnSelectorChange()
    On Error GoTo Fallo

    ActualizarDashboard
    Exit Sub

Fallo:
    MsgBox "No se pudo actualizar el dashboard: " & Err.Description, _
        vbExclamation, "Dashboard"
End Sub

' ==========================================================================
' Helpers internos
' ==========================================================================

Private Sub MarcarExcesos(ByVal ws As Worksheet, ByRef detalle As TDetallePeriodo, _
        ByVal hayDatos As Boolean)

    ws.Range("B9").Value = TextoExceso(detalle.oee, hayDatos)
    ws.Range("D9").Value = TextoExceso(detalle.disponibilidad, hayDatos)
    ws.Range("F9").Value = TextoExceso(detalle.rendimiento, hayDatos)
    ws.Range("H9").Value = TextoExceso(detalle.calidad, hayDatos)
End Sub

Private Function TextoExceso(ByVal valor As Variant, ByVal hayDatos As Boolean) As String
    If Not hayDatos Then Exit Function
    If Not IsNumeric(valor) Then Exit Function
    If CDbl(valor) > 1 Then
        TextoExceso = IconoAlerta() & " > 100% (RN-05)"
    End If
End Function

Private Function TipoTarjeta(ByVal celda As Range) As String
    Dim etiqueta As String

    etiqueta = UCase$(Trim$(CStr(celda.Offset(-1, 0).Value)))

    If InStr(etiqueta, "OEE") > 0 Then
        TipoTarjeta = "OEE"
    ElseIf InStr(etiqueta, "DISPON") > 0 Then
        TipoTarjeta = "DISPONIBILIDAD"
    ElseIf InStr(etiqueta, "REND") > 0 Then
        TipoTarjeta = "RENDIMIENTO"
    ElseIf InStr(etiqueta, "CALIDAD") > 0 Then
        TipoTarjeta = "CALIDAD"
    Else
        Select Case celda.Column
        Case 2
            TipoTarjeta = "OEE"
        Case 4
            TipoTarjeta = "DISPONIBILIDAD"
        Case 6
            TipoTarjeta = "RENDIMIENTO"
        Case 8
            TipoTarjeta = "CALIDAD"
        Case Else
            TipoTarjeta = "OEE"
        End Select
    End If
End Function

Private Function Umbral(ByVal nombre As String, ByVal defecto As Double) As Double
    Dim v As Variant

    v = AccesoHojas.ObtenerParametro(nombre, 0)
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        Umbral = defecto
    ElseIf IsNumeric(v) Then
        Umbral = CDbl(v) / 100
    Else
        Umbral = defecto
    End If
End Function

Private Function FechaCelda(ByVal celda As Range) As Date
    FechaCelda = Utils.AFecha(celda.Value)
End Function

Private Function TextoCelda(ByVal celda As Range) As String
    Dim v As Variant

    v = celda.Value
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    TextoCelda = Trim$(CStr(v))
End Function

Private Function SinValor(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        SinValor = True
    ElseIf Not IsNumeric(v) Then
        SinValor = True
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        SinValor = True
    End If
End Function

Private Function NormalizarProducto(ByVal s As String) As String
    Dim t As String

    t = UCase$(Trim$(s))
    If Len(t) = 0 Then Exit Function

    If InStr(t, "500") = 1 Then
        If InStr(t, "CG") > 0 Or InStr(t, "CON") > 0 Then
            NormalizarProducto = "500CG"
        Else
            NormalizarProducto = "500SG"
        End If
    ElseIf InStr(t, "1500") = 1 Then
        NormalizarProducto = "1500SG"
    ElseIf InStr(t, "2000") = 1 Then
        NormalizarProducto = "2000SG"
    ElseIf InStr(t, "BID") = 1 Then
        NormalizarProducto = "BID6"
    Else
        NormalizarProducto = s
    End If
End Function

Private Function NormalizarLinea(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim n As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then n = n & ch
    Next i

    NormalizarLinea = n
End Function

Private Function NuevoDiccionario() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NuevoDiccionario = d
End Function

Private Function Guion() As String
    Guion = ChrW(8212)
End Function

Private Function IconoAlerta() As String
    IconoAlerta = ChrW(9888)
End Function
