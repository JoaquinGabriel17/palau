Attribute VB_Name = "Utils"
Option Explicit

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

Public Function IndiceColumna(ByVal lo As ListObject, ByVal nombre As String) As Long
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, nombre, vbTextCompare) = 0 Then
            IndiceColumna = lc.Index
            Exit Function
        End If
    Next lc
End Function

Public Function AFecha(ByVal v As Variant) As Date
    If IsDate(v) Then AFecha = CDate(v)
End Function

Public Function NumeroSeguro(ByVal v As Variant) As Double
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then NumeroSeguro = CDbl(v)
End Function

Public Function EsAnulado(ByVal valor As Variant) As Boolean
    Dim s As String

    If IsEmpty(valor) Or IsNull(valor) Then Exit Function
    If IsError(valor) Then Exit Function

    s = Trim$(CStr(valor))
    EsAnulado = (StrComp(s, "S" & ChrW(237), vbTextCompare) = 0) _
                Or (StrComp(s, "Si", vbTextCompare) = 0)
End Function
