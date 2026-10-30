Attribute VB_Name = "Calculos"
Option Explicit

Public Function CalcularTO(ByVal TPL As Double, ByVal PNP As Double) As Double
    CalcularTO = TPL - PNP
End Function

Public Function CalcularDisponibilidad(ByVal TiempoOperativo As Double, ByVal TPL As Double) As Variant
    If TPL = 0 Then
        CalcularDisponibilidad = Empty
    Else
        CalcularDisponibilidad = TiempoOperativo / TPL
    End If
End Function

Public Function CalcularRendimientoOEE(ByVal UnidadesTotales As Double, ByVal VelocidadNominal As Double, ByVal TiempoOperativo As Double) As Variant
    If VelocidadNominal = 0 Or TiempoOperativo = 0 Then
        CalcularRendimientoOEE = Empty
    Else
        CalcularRendimientoOEE = (UnidadesTotales / VelocidadNominal) / TiempoOperativo
    End If
End Function

Public Function CalcularRendimiento(ByVal UnidadesTotales As Double, ByVal VelocidadNominal As Double, ByVal TiempoOperativo As Double) As Variant
    CalcularRendimiento = CalcularRendimientoOEE(UnidadesTotales, VelocidadNominal, TiempoOperativo)
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
