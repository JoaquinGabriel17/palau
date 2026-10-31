Attribute VB_Name = "Tests"
Option Explicit

Private Const tolerancia As Double = 0.0005

Public Sub EjecutarPruebas()
    Call RunAllTests
End Sub

Public Function RunAllTests() As Boolean
    Dim aciertos As Long
    Dim total As Long
    Dim oee As Double

    Debug.Print "=== Pruebas OEE ==="

    total = total + 1
    If Comparar("Caso 498", CalcularOEE_Turno("T-498"), 0.614) Then aciertos = aciertos + 1

    total = total + 1
    If Test_Caso498_PuntaAPunta() Then aciertos = aciertos + 1

    total = total + 1
    If ExisteTurno("T-74-2") Then
        oee = CalcularOEE_Turno("T-74-2")
    Else
        oee = CalcularOEE_DesdeTotales(480, 65, 0, 33, 26700, 52, 17140 / 90 + 9612 / 70)
    End If
    If Comparar("Caso 7.4-2", oee, 0.788) Then aciertos = aciertos + 1

    total = total + 1
    If ExisteTurno("T-74-3") Then
        oee = CalcularOEE_Turno("T-74-3")
    Else
        oee = CalcularOEE_DesdeTotales(480, 35, 35, 57, 21600, 30, 21630 / 70)
    End If
    If Comparar("Caso 7.4-3", oee, 0.753) Then aciertos = aciertos + 1

    total = total + 1
    If ExisteTurno("T-74-4") Then
        oee = CalcularOEE_Turno("T-74-4")
    Else
        oee = CalcularOEE_DesdeTotales(480, 35, 0, 50, 22080, 8, 22088 / 70)
    End If
    If Comparar("Caso 7.4-4", oee, 0.709) Then aciertos = aciertos + 1

    total = total + 1
    If Test_CeroUnidades() Then aciertos = aciertos + 1

    total = total + 1
    If Test_CienRechazo() Then aciertos = aciertos + 1

    total = total + 1
    If Test_TurnoSinParadas() Then aciertos = aciertos + 1

    total = total + 1
    If Test_VigenciaParametros() Then aciertos = aciertos + 1

    Debug.Print "Resultado: " & aciertos & "/" & total
    RunAllTests = (aciertos = total)
End Function

Public Function Test_Caso498_PuntaAPunta() As Boolean
    Dim detalle As TDetalleTurno
    Dim oee As Double
    Dim tolerancia As Double
    Dim ok As Boolean

    tolerancia = ToleranciaComparacionPP()
    detalle = DetalleTurno(DateSerial(2026, 9, 28), "RES-T")
    oee = CalcularTurno(DateSerial(2026, 9, 28), "RES-T")

    ok = detalle.Encontrado _
        And CercaMinutos(detalle.TO, 330, tolerancia) _
        And CercaPP(CDbl(detalle.disponibilidad), 0.8148, tolerancia) _
        And CercaPP(CDbl(detalle.rendimiento), 0.7534, tolerancia) _
        And CercaPP(CDbl(detalle.calidad), 0.9998, tolerancia) _
        And CercaPP(oee, 0.6138, tolerancia)

    Debug.Print "Caso 498 punta a punta: " & IIf(ok, "OK", "ERROR") & _
        " (TO=" & Format$(detalle.TO, "0.00") & _
        ", D=" & Format$(CDbl(detalle.disponibilidad), "0.00%") & _
        ", R=" & Format$(CDbl(detalle.rendimiento), "0.00%") & _
        ", C=" & Format$(CDbl(detalle.calidad), "0.00%") & _
        ", OEE=" & Format$(oee, "0.00%") & ")"
    Test_Caso498_PuntaAPunta = ok
End Function

Public Sub ReporteCaso498()
    Dim ws As Worksheet
    Dim detalle As TDetalleTurno
    Dim tolerancia As Double
    Dim fila As Long

    Set ws = ThisWorkbook.Worksheets("90_Tests")
    ws.Range("A1:J40").ClearContents
    ws.Range("A1:J40").ClearFormats

    tolerancia = ToleranciaComparacionPP()
    detalle = DetalleTurno(DateSerial(2026, 9, 28), "RES-T")

    ws.Cells(1, 1).Value = "Caso 498 - Reporte de Produccion Nro. 498 - 2026-09-28 - turno RES-T"
    ws.Cells(2, 1).Value = "Comparacion contra Q-18 (limpiezas micro) y D-01 (limpiezas A y C) - tolerancia P-08: " & Format$(tolerancia, "0.00") & " pp"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12

    If Not detalle.Encontrado Then
        ws.Cells(4, 1).Value = "No se encontro el turno 2026-09-28 / RES-T."
        Exit Sub
    End If

    ws.Cells(4, 1).Value = "Valores intermedios calculados por el libro"
    ws.Cells(4, 1).Font.Bold = True
    ws.Cells(5, 1).Value = "Indicador"
    ws.Cells(5, 2).Value = "Calculado"
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 2)).Font.Bold = True
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 2)).Interior.Color = RGB(221, 235, 247)

    fila = 6
    EscribirValor ws, fila, "TD (min)", detalle.td, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "PP (min)", detalle.pp, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "PE (min)", detalle.pe, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "A/F (min)", detalle.af, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "A y C (min)", detalle.ac, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "Microparadas (min)", detalle.micro, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "TPL (min)", detalle.tpl, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "TO (min)", detalle.TO, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "Tiempo ideal (min)", detalle.tiempoIdeal, "0.00"
    fila = fila + 1
    EscribirValor ws, fila, "Unidades buenas", detalle.UnidadesBuenas, "#,##0"
    fila = fila + 1
    EscribirValor ws, fila, "Unidades rechazadas", detalle.UnidadesRechazadas, "#,##0"
    fila = fila + 1
    EscribirValor ws, fila, "Unidades totales (UT)", detalle.UnidadesTotales, "#,##0"

    fila = fila + 2
    ws.Cells(fila, 1).Value = "Comparacion contra valores manuales"
    ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 1
    ws.Cells(fila, 1).Value = "Indicador"
    ws.Cells(fila, 2).Value = "Calculado por el libro"
    ws.Cells(fila, 3).Value = "Manual Q-18"
    ws.Cells(fila, 4).Value = "Dif. Q-18"
    ws.Cells(fila, 5).Value = "Resultado Q-18"
    ws.Cells(fila, 6).Value = "Manual D-01"
    ws.Cells(fila, 7).Value = "Dif. D-01"
    ws.Cells(fila, 8).Value = "Resultado D-01"
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 8)).Font.Bold = True
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 8)).Interior.Color = RGB(221, 235, 247)

    fila = fila + 1
    EscribirComparacionDoble ws, fila, "TO (min)", detalle.TO, 375, 330, tolerancia, False, "0.00"
    fila = fila + 1
    EscribirComparacionDoble ws, fila, "Disponibilidad", CDbl(detalle.disponibilidad), 0.926, 0.815, tolerancia, True, "0.00%"
    fila = fila + 1
    EscribirComparacionDoble ws, fila, "UT (unidades)", detalle.UnidadesTotales, 17404, 17404, tolerancia, False, "#,##0"
    fila = fila + 1
    EscribirComparacionDoble ws, fila, "Rendimiento", CDbl(detalle.rendimiento), 0.663, 0.753, tolerancia, True, "0.00%"
    fila = fila + 1
    EscribirComparacionDoble ws, fila, "Calidad", CDbl(detalle.calidad), 0.9998, 0.9998, tolerancia, True, "0.00%"
    fila = fila + 1
    EscribirComparacionDoble ws, fila, "OEE", detalle.oee, 0.614, 0.614, tolerancia, True, "0.00%"

    fila = fila + 2
    ws.Cells(fila, 1).Value = "Diferencias: pp para D, R, C y OEE; min para TO; unidades para UT. PASA/FALLA con tolerancia P-08 (Doc 1.4)."
    ws.Cells(fila, 1).Font.Italic = True
    fila = fila + 1
    ws.Cells(fila, 1).Value = "Q-18: limpiezas de la etiquetadora como microparada. D-01: limpiezas como A y C (decision pendiente)."
    ws.Cells(fila, 1).Font.Italic = True

    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B:H").ColumnWidth = 20
    ws.Columns("A:H").AutoFit

    Debug.Print "ReporteCaso498: TO=" & Format$(detalle.TO, "0.00") & _
        ", Disponibilidad=" & Format$(CDbl(detalle.disponibilidad), "0.00%") & _
        ", UT=" & Format$(detalle.UnidadesTotales, "#,##0") & _
        ", Rendimiento=" & Format$(CDbl(detalle.rendimiento), "0.00%") & _
        ", Calidad=" & Format$(CDbl(detalle.calidad), "0.00%") & _
        ", OEE=" & Format$(detalle.oee, "0.00%")
End Sub

Private Sub EscribirValor(ByVal ws As Worksheet, ByVal fila As Long, ByVal nombre As String, ByVal valor As Double, ByVal formato As String)
    ws.Cells(fila, 1).Value = nombre
    ws.Cells(fila, 2).Value = valor
    ws.Cells(fila, 2).NumberFormat = formato
End Sub

Private Sub EscribirComparacionDoble(ByVal ws As Worksheet, ByVal fila As Long, ByVal nombre As String, ByVal calculado As Double, ByVal manualQ18 As Double, ByVal manualD01 As Double, ByVal tolerancia As Double, ByVal esPorcentaje As Boolean, ByVal formatoValor As String)
    Dim difQ18 As Double
    Dim difD01 As Double
    Dim resQ18 As String
    Dim resD01 As String
    Dim formatoDif As String

    If esPorcentaje Then
        difQ18 = (calculado - manualQ18) * 100
        difD01 = (calculado - manualD01) * 100
        formatoDif = "0.00"
    Else
        difQ18 = calculado - manualQ18
        difD01 = calculado - manualD01
        formatoDif = formatoValor
    End If

    resQ18 = IIf(Abs(difQ18) <= tolerancia, "PASA", "FALLA")
    resD01 = IIf(Abs(difD01) <= tolerancia, "PASA", "FALLA")

    ws.Cells(fila, 1).Value = nombre
    ws.Cells(fila, 2).Value = calculado
    ws.Cells(fila, 3).Value = manualQ18
    ws.Cells(fila, 4).Value = difQ18
    ws.Cells(fila, 5).Value = resQ18
    ws.Cells(fila, 6).Value = manualD01
    ws.Cells(fila, 7).Value = difD01
    ws.Cells(fila, 8).Value = resD01

    ws.Cells(fila, 2).NumberFormat = formatoValor
    ws.Cells(fila, 3).NumberFormat = formatoValor
    ws.Cells(fila, 4).NumberFormat = formatoDif
    ws.Cells(fila, 6).NumberFormat = formatoValor
    ws.Cells(fila, 7).NumberFormat = formatoDif

    PintarResultado ws, fila, 5, resQ18
    PintarResultado ws, fila, 8, resD01
End Sub

Private Sub PintarResultado(ByVal ws As Worksheet, ByVal fila As Long, ByVal columna As Long, ByVal resultado As String)
    If resultado = "PASA" Then
        ws.Cells(fila, columna).Interior.Color = RGB(198, 239, 206)
    Else
        ws.Cells(fila, columna).Interior.Color = RGB(255, 199, 206)
    End If
End Sub

Private Function ToleranciaComparacionPP() As Double
    Dim lo As ListObject
    Dim rng As Range
    Dim cParametro As Long
    Dim cValor As Long
    Dim i As Long

    ToleranciaComparacionPP = 0.05

    Set lo = Tabla("tblParametros")
    If lo Is Nothing Then Exit Function

    cParametro = IndiceColumnaTest(lo, "parametro")
    cValor = IndiceColumnaTest(lo, "valor")
    If cParametro = 0 Or cValor = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(CStr(rng.Cells(i, cParametro).Value), "TOLERANCIA_COMPARACION_PP", vbTextCompare) = 0 Then
            If IsNumeric(rng.Cells(i, cValor).Value) Then
                If CDbl(rng.Cells(i, cValor).Value) > 0 Then ToleranciaComparacionPP = CDbl(rng.Cells(i, cValor).Value)
            End If
            Exit Function
        End If
    Next i
End Function

Public Function Test_CeroUnidades() As Boolean
    Dim oee As Double
    Dim calidad As Variant
    Dim rendimiento As Variant
    Dim ok As Boolean

    oee = CalcularOEE_DesdeTotales(480, 0, 0, 0, 0, 0, 0)
    calidad = CalcularCalidad(0, 0)
    rendimiento = CalcularRendimientoOEE(0, 70, 480)

    ok = Cerca(oee, 0) And IsEmpty(calidad) And Cerca(CDbl(rendimiento), 0)
    Debug.Print "Test_CeroUnidades: " & IIf(ok, "OK", "ERROR") & " (OEE=" & Format$(oee, "0.0000") & ")"
    Test_CeroUnidades = ok
End Function

Public Function Test_CienRechazo() As Boolean
    Dim oee As Double
    Dim calidad As Variant
    Dim ok As Boolean

    oee = CalcularOEE_DesdeTotales(480, 0, 0, 0, 0, 100, 100 / 70)
    calidad = CalcularCalidad(0, 100)

    ok = Cerca(CDbl(calidad), 0) And Cerca(oee, 0)
    Debug.Print "Test_CienRechazo: " & IIf(ok, "OK", "ERROR") & " (Calidad=" & Format$(CDbl(calidad), "0.00%") & ", OEE=" & Format$(oee, "0.0000") & ")"
    Test_CienRechazo = ok
End Function

Public Function Test_TurnoSinParadas() As Boolean
    Dim tpl As Double
    Dim tOper As Double
    Dim disponibilidad As Variant
    Dim rendimiento As Variant
    Dim calidad As Variant
    Dim oee As Double
    Dim ok As Boolean

    tpl = 480
    tOper = CalcularTO(tpl, 0)
    disponibilidad = CalcularDisponibilidad(tOper, tpl)
    rendimiento = CalcularRendimientoOEE(1000, 70, tOper)
    calidad = CalcularCalidad(1000, 1000)
    oee = CalcularOEE(CDbl(disponibilidad), CDbl(rendimiento), CDbl(calidad))

    ok = Cerca(CDbl(disponibilidad), 1) _
        And Cerca(CDbl(calidad), 1) _
        And Cerca(CDbl(rendimiento), (1000 / 70) / tOper) _
        And Cerca(oee, CDbl(disponibilidad) * CDbl(rendimiento) * CDbl(calidad))

    Debug.Print "Test_TurnoSinParadas: " & IIf(ok, "OK", "ERROR") & " (Disponibilidad=" & Format$(CDbl(disponibilidad), "0.00%") & ", Rendimiento=" & Format$(CDbl(rendimiento), "0.00%") & ", Calidad=" & Format$(CDbl(calidad), "0.00%") & ")"
    Test_TurnoSinParadas = ok
End Function

Public Function Test_VigenciaParametros() As Boolean
    Dim lo As ListObject
    Dim v2025 As Double
    Dim v2026 As Double
    Dim oee2025 As Double
    Dim oee2026 As Double
    Dim ok As Boolean

    On Error GoTo Fallo

    Set lo = Tabla("tblVelocidades")
    If lo Is Nothing Then
        Debug.Print "Test_VigenciaParametros: ERROR (falta tblVelocidades)"
        Exit Function
    End If

    LimpiarVelocidadTest lo, "TST-VIG"
    AgregarVelocidadTest lo, "TST-VIG-A", 10, DateSerial(2025, 1, 1), DateSerial(2026, 1, 1)
    AgregarVelocidadTest lo, "TST-VIG-B", 20, DateSerial(2026, 1, 2), 0

    v2025 = VelocidadVigente("TST-VIG", DateSerial(2025, 12, 31))
    v2026 = VelocidadVigente("TST-VIG", DateSerial(2026, 1, 2))

    If v2025 > 0 And v2026 > 0 Then
        oee2025 = CalcularOEE_DesdeTotales(480, 0, 0, 0, 1000, 0, 1000 / v2025)
        oee2026 = CalcularOEE_DesdeTotales(480, 0, 0, 0, 1000, 0, 1000 / v2026)
        ok = Cerca(v2025, 10) And Cerca(v2026, 20) And (oee2025 > oee2026)
    Else
        ok = False
    End If

    Debug.Print "Test_VigenciaParametros: " & IIf(ok, "OK", "ERROR") & " (2025: vel=" & Format$(v2025, "0.00") & ", OEE=" & Format$(oee2025, "0.0000") & "; 2026: vel=" & Format$(v2026, "0.00") & ", OEE=" & Format$(oee2026, "0.0000") & ")"
    Test_VigenciaParametros = ok

    LimpiarVelocidadTest lo, "TST-VIG"
    Exit Function

Fallo:
    Debug.Print "Test_VigenciaParametros: ERROR (" & Err.Description & ")"
    On Error Resume Next
    LimpiarVelocidadTest lo, "TST-VIG"
    Test_VigenciaParametros = False
End Function

Private Function Comparar(ByVal nombre As String, ByVal obtenido As Double, ByVal esperado As Double) As Boolean
    If Abs(obtenido - esperado) <= tolerancia Then
        Debug.Print nombre & ": OK (" & Format$(obtenido, "0.0000") & ")"
        Comparar = True
    Else
        Debug.Print nombre & ": ERROR (obtenido " & Format$(obtenido, "0.0000") & ", esperado " & Format$(esperado, "0.0000") & ")"
    End If
End Function

Private Function Cerca(ByVal a As Double, ByVal b As Double) As Boolean
    Cerca = (Abs(a - b) <= tolerancia)
End Function

Private Function CercaPP(ByVal calculado As Double, ByVal esperado As Double, ByVal tolerancia As Double) As Boolean
    CercaPP = (Abs(calculado - esperado) * 100 <= tolerancia)
End Function

Private Function CercaMinutos(ByVal calculado As Double, ByVal esperado As Double, ByVal tolerancia As Double) As Boolean
    CercaMinutos = (Abs(calculado - esperado) <= tolerancia)
End Function

Private Sub AgregarVelocidadTest(ByVal lo As ListObject, ByVal codigo As String, ByVal velocidad As Double, ByVal desde As Date, ByVal hasta As Date)
    Dim fila As ListRow
    Dim rng As Range

    Set fila = lo.ListRows.Add
    Set rng = fila.Range

    SetValorTest rng, lo, "codigo", codigo
    SetValorTest rng, lo, "producto", "TST-VIG"
    SetValorTest rng, lo, "velocidad", velocidad
    SetValorTest rng, lo, "vigencia_desde", desde
    If hasta > 0 Then
        SetValorTest rng, lo, "vigencia_hasta", hasta
    Else
        SetValorTest rng, lo, "vigencia_hasta", Empty
    End If
    SetValorTest rng, lo, "origen_dato", "prueba"
    SetValorTest rng, lo, "activo", "Si"
End Sub

Private Sub SetValorTest(ByVal rng As Range, ByVal lo As ListObject, ByVal columna As String, ByVal valor As Variant)
    Dim c As Long

    c = IndiceColumnaTest(lo, columna)
    If c > 0 Then rng.Cells(1, c).Value = valor
End Sub

Private Function IndiceColumnaTest(ByVal lo As ListObject, ByVal nombre As String) As Long
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, nombre, vbTextCompare) = 0 Then
            IndiceColumnaTest = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Sub LimpiarVelocidadTest(ByVal lo As ListObject, ByVal producto As String)
    Dim rng As Range
    Dim cProducto As Long
    Dim i As Long

    If lo Is Nothing Then Exit Sub
    cProducto = IndiceColumnaTest(lo, "producto")
    If cProducto = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = rng.Rows.Count To 1 Step -1
        If StrComp(CStr(rng.Cells(i, cProducto).Value), producto, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
        End If
    Next i
End Sub
