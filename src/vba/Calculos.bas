Attribute VB_Name = "Calculos"
Option Explicit

Public Function CalcularTO(ByVal tpl As Double, ByVal pnp As Double) As Double
    CalcularTO = tpl - pnp
End Function

Public Function CalcularDisponibilidad(ByVal TiempoOperativo As Double, ByVal tpl As Double) As Variant
    If tpl = 0 Then
        CalcularDisponibilidad = Empty
    Else
        CalcularDisponibilidad = TiempoOperativo / tpl
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

Public Function CalcularOEE(ByVal disponibilidad As Double, ByVal rendimiento As Double, ByVal calidad As Double) As Double
    CalcularOEE = disponibilidad * rendimiento * calidad
End Function
