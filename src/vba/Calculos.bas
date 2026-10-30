Attribute VB_Name = "Calculos"
Option Explicit

Public Function CalcularTO(ByVal TPL As Double, ByVal PNP As Double) As Double
    CalcularTO = TPL - PNP
End Function

Public Function CalcularDisponibilidad(ByVal TO As Double, ByVal TPL As Double) As Variant
    If TPL = 0 Then
        CalcularDisponibilidad = Empty
    Else
        CalcularDisponibilidad = TO / TPL
    End If
End Function

Public Function CalcularRendimiento(ByVal UnidadesTotales As Double, ByVal VelocidadNominal As Double, ByVal TO As Double) As Variant
    If VelocidadNominal = 0 Or TO = 0 Then
        CalcularRendimiento = Empty
    Else
        CalcularRendimiento = (UnidadesTotales / VelocidadNominal) / TO
    End If
End Function

Public Function CalcularCalidad(ByVal UnidadesBuenas As Double, ByVal UnidadesTotales As Double) As Variant
    If UnidadesTotales = 0 Then
        CalcularCalidad = Empty
    Else
        CalcularCalidad = UnidadesBuenas / UnidadesTotales
    End If
End Function

Public Function CalcularOEE(ByVal Disponibilidad As Double, ByVal Rendimiento As Double, ByVal Calidad As Double) As Double
    CalcularOEE = Disponibilidad * Rendimiento * Calidad
End Function
