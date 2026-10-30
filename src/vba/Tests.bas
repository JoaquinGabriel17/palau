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

    Debug.Print "Resultado: " & aciertos & "/" & total
End Sub

Private Function Comparar(ByVal nombre As String, ByVal obtenido As Double, ByVal esperado As Double) As Boolean
    If Abs(obtenido - esperado) <= TOLERANCIA Then
        Debug.Print nombre & ": OK (" & Format$(obtenido, "0.0000") & ")"
        Comparar = True
    Else
        Debug.Print nombre & ": ERROR (obtenido " & Format$(obtenido, "0.0000") & ", esperado " & Format$(esperado, "0.0000") & ")"
    End If
End Function
