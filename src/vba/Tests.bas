Attribute VB_Name = "Tests"
Option Explicit

Private Const TOLERANCIA As Double = 0.0005

Public Sub EjecutarPruebas()
    Dim aciertos As Long
    Dim total As Long
    Dim oee As Double

    Debug.Print "=== Pruebas OEE ==="

    total = total + 1
    If Comparar("Caso 498", CalcularOEE_Turno("T-498"), 0.614) Then aciertos = aciertos + 1

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
End Sub

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
    If Abs(obtenido - esperado) <= TOLERANCIA Then
        Debug.Print nombre & ": OK (" & Format$(obtenido, "0.0000") & ")"
        Comparar = True
    Else
        Debug.Print nombre & ": ERROR (obtenido " & Format$(obtenido, "0.0000") & ", esperado " & Format$(esperado, "0.0000") & ")"
    End If
End Function

Private Function Cerca(ByVal a As Double, ByVal b As Double) As Boolean
    Cerca = (Abs(a - b) <= TOLERANCIA)
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
