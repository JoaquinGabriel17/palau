Attribute VB_Name = "Validaciones"
Option Explicit

Public Enum ResultadoValidacion
    OK = 0
    AVISA = 1
    BLOQUEA = 2
End Enum

Public Type TResultadoValidacion
    Resultado As ResultadoValidacion
    Mensaje As String
    Bloquea As Boolean
    Avisa As Boolean
End Type

Public Function ValidarParada( _
        ByVal fechaHoraInicio As Variant, _
        ByVal fechaHoraFin As Variant, _
        ByVal turnoInicio As Variant, _
        ByVal turnoFin As Variant, _
        Optional ByVal tipo As String = "", _
        Optional ByVal motivo As String = "", _
        Optional ByVal duracionMin As Double = -1, _
        Optional ByVal idParada As String = "", _
        Optional ByVal otrasParadas As Variant, _
        Optional ByVal tdMinutos As Double = 0, _
        Optional ByVal sumaParadasMin As Double = -1, _
        Optional ByVal tipoActivo As Boolean = True) As TResultadoValidacion

    Dim ini As Date
    Dim fin As Date
    Dim tIni As Date
    Dim tFin As Date
    Dim dur As Double
    Dim td As Double
    Dim sumaOtras As Double
    Dim solapa As Boolean
    Dim total As Double

    If Not EsFechaValida(fechaHoraInicio) Or Not EsFechaValida(fechaHoraFin) Then
        ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "La parada debe tener hora de inicio y fin.")
        Exit Function
    End If

    ini = CDate(fechaHoraInicio)
    fin = CDate(fechaHoraFin)

    If fin = ini Then
        ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "La hora de fin debe ser posterior a la de inicio.")
        Exit Function
    End If
    If fin < ini Then fin = DateAdd("d", 1, fin)

    If duracionMin < 0 Then
        dur = DateDiff("n", ini, fin)
    Else
        dur = duracionMin
    End If
    If dur <= 0 Then
        ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "La duracion de la parada debe ser mayor que cero.")
        Exit Function
    End If

    If EsFechaValida(turnoInicio) And EsFechaValida(turnoFin) Then
        tIni = CDate(turnoInicio)
        tFin = CDate(turnoFin)
        If tFin <= tIni Then tFin = DateAdd("d", 1, tFin)
        If ini < tIni Or fin > tFin Then
            ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "La parada cae fuera del horario del turno.")
            Exit Function
        End If
        td = DateDiff("n", tIni, tFin)
    End If
    If tdMinutos > 0 Then td = tdMinutos

    If Not IsMissing(otrasParadas) Then
        AnalizarOtrasParadas ini, fin, otrasParadas, idParada, solapa, sumaOtras
        If solapa Then
            ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "La parada se superpone con otra parada del mismo turno.")
            Exit Function
        End If
    End If

    If sumaParadasMin >= 0 Then
        total = sumaParadasMin + dur
    Else
        total = sumaOtras + dur
    End If

    If td > 0 And total > td Then
        ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "La suma de paradas supera el tiempo del turno.")
        Exit Function
    End If

    If Len(Trim$(tipo)) = 0 Then
        ValidarParada = Resultado(ResultadoValidacion.AVISA, "La parada no tiene tipo; se guardara como TP-06.")
    ElseIf UCase$(Trim$(tipo)) = "TP-02" Then
        ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "El tipo TP-02 es una cabecera y no se puede usar.")
    ElseIf Not tipoActivo Then
        ValidarParada = Resultado(ResultadoValidacion.BLOQUEA, "El tipo de parada esta inactivo.")
    ElseIf Len(Trim$(motivo)) = 0 Then
        ValidarParada = Resultado(ResultadoValidacion.AVISA, "La parada no tiene motivo.")
    Else
        ValidarParada = Resultado(ResultadoValidacion.OK, "Parada valida.")
    End If
End Function

Public Function ValidarProduccion(ByVal cantidad As Variant, _
        Optional ByVal nombreCampo As String = "La produccion") As TResultadoValidacion

    If IsEmpty(cantidad) Or IsNull(cantidad) Then
        ValidarProduccion = Resultado(ResultadoValidacion.BLOQUEA, nombreCampo & " debe tener una cantidad.")
        Exit Function
    End If
    If Not IsNumeric(cantidad) Then
        ValidarProduccion = Resultado(ResultadoValidacion.BLOQUEA, nombreCampo & " debe ser numerica.")
        Exit Function
    End If
    If CDbl(cantidad) < 0 Then
        ValidarProduccion = Resultado(ResultadoValidacion.BLOQUEA, nombreCampo & " no puede ser negativa.")
        Exit Function
    End If
    If CDbl(cantidad) <> Int(CDbl(cantidad)) Then
        ValidarProduccion = Resultado(ResultadoValidacion.BLOQUEA, nombreCampo & " debe ser un numero entero.")
        Exit Function
    End If
    If CDbl(cantidad) = 0 Then
        ValidarProduccion = Resultado(ResultadoValidacion.AVISA, nombreCampo & " es cero; se guardara como pendiente de revisar.")
        Exit Function
    End If

    ValidarProduccion = Resultado(ResultadoValidacion.OK, "Produccion valida.")
End Function

Public Function ValidarRechazo(ByVal seccion As String, ByVal motivo As String, _
        Optional ByVal cantidad As Variant = Empty) As TResultadoValidacion

    If IsEmpty(cantidad) Or IsNull(cantidad) Then
        ValidarRechazo = Resultado(ResultadoValidacion.BLOQUEA, "La cantidad del rechazo es obligatoria.")
        Exit Function
    End If
    If Not IsNumeric(cantidad) Then
        ValidarRechazo = Resultado(ResultadoValidacion.BLOQUEA, "La cantidad del rechazo debe ser numerica.")
        Exit Function
    End If
    If CDbl(cantidad) < 0 Then
        ValidarRechazo = Resultado(ResultadoValidacion.BLOQUEA, "La cantidad del rechazo no puede ser negativa.")
        Exit Function
    End If
    If CDbl(cantidad) <> Int(CDbl(cantidad)) Then
        ValidarRechazo = Resultado(ResultadoValidacion.BLOQUEA, "La cantidad del rechazo debe ser un numero entero.")
        Exit Function
    End If

    If Len(Trim$(seccion)) = 0 And Len(Trim$(motivo)) = 0 Then
        ValidarRechazo = Resultado(ResultadoValidacion.AVISA, "El rechazo no tiene seccion ni motivo; se guardara como pendiente de revisar.")
    ElseIf Len(Trim$(seccion)) = 0 Then
        ValidarRechazo = Resultado(ResultadoValidacion.AVISA, "El rechazo no tiene seccion; se guardara como pendiente de revisar.")
    ElseIf Len(Trim$(motivo)) = 0 Then
        ValidarRechazo = Resultado(ResultadoValidacion.AVISA, "El rechazo no tiene motivo; se guardara como pendiente de revisar.")
    Else
        ValidarRechazo = Resultado(ResultadoValidacion.OK, "Rechazo valido.")
    End If
End Function

Public Function ValidarLote(ByVal numeroLote As Variant) As TResultadoValidacion

    Dim s As String
    Dim mes As Long
    Dim secuencial As Long

    If IsEmpty(numeroLote) Or IsNull(numeroLote) Then
        ValidarLote = Resultado(ResultadoValidacion.BLOQUEA, "El numero de lote es obligatorio.")
        Exit Function
    End If

    s = Trim$(CStr(numeroLote))
    If Len(s) <> 5 Then
        ValidarLote = Resultado(ResultadoValidacion.BLOQUEA, "El numero de lote debe tener 2 digitos de mes y 3 secuenciales (ej. 11261).")
        Exit Function
    End If
    If Not SoloDigitos(s) Then
        ValidarLote = Resultado(ResultadoValidacion.BLOQUEA, "El numero de lote solo puede tener digitos.")
        Exit Function
    End If

    mes = CLng(Left$(s, 2))
    secuencial = CLng(Right$(s, 3))

    If mes < 1 Or mes > 12 Then
        ValidarLote = Resultado(ResultadoValidacion.BLOQUEA, "El mes del lote debe estar entre 01 y 12.")
        Exit Function
    End If
    If secuencial < 1 Or secuencial > 999 Then
        ValidarLote = Resultado(ResultadoValidacion.BLOQUEA, "El secuencial del lote debe estar entre 001 y 999.")
        Exit Function
    End If

    ValidarLote = Resultado(ResultadoValidacion.OK, "Numero de lote valido.")
End Function

Public Function ValidarReporte(ByVal nroReporte As String, _
        Optional ByVal nrosExistentes As Variant, _
        Optional ByVal tipoParada As String = "", _
        Optional ByVal tipoActivo As Boolean = True) As TResultadoValidacion

    Dim s As String

    s = Trim$(nroReporte)
    If Len(s) > 0 Then
        If Not IsMissing(nrosExistentes) Then
            If ValorEnLista(s, nrosExistentes) Then
                ValidarReporte = Resultado(ResultadoValidacion.BLOQUEA, "El numero de reporte esta duplicado.")
                Exit Function
            End If
        End If
    End If

    If Len(Trim$(tipoParada)) > 0 Then
        If UCase$(Trim$(tipoParada)) = "TP-02" Then
            ValidarReporte = Resultado(ResultadoValidacion.BLOQUEA, "El tipo TP-02 es una cabecera y no se puede usar.")
            Exit Function
        End If
        If Not tipoActivo Then
            ValidarReporte = Resultado(ResultadoValidacion.BLOQUEA, "El tipo de parada esta inactivo.")
            Exit Function
        End If
    End If

    ValidarReporte = Resultado(ResultadoValidacion.OK, "Reporte valido.")
End Function

Private Function Resultado(ByVal estado As ResultadoValidacion, ByVal mensaje As String) As TResultadoValidacion
    Resultado.Resultado = estado
    Resultado.Mensaje = mensaje
    Resultado.Bloquea = (estado = ResultadoValidacion.BLOQUEA)
    Resultado.Avisa = (estado = ResultadoValidacion.AVISA)
End Function

Private Function EsFechaValida(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    If CDbl(v) = 0 Then Exit Function
    EsFechaValida = True
End Function

Private Function SoloDigitos(ByVal s As String) As Boolean
    Dim i As Long

    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    SoloDigitos = True
End Function

Private Sub AnalizarOtrasParadas(ByVal ini As Date, ByVal fin As Date, ByVal otras As Variant, _
        ByVal idExcluir As String, ByRef solapa As Boolean, ByRef suma As Double)

    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim lb2 As Long
    Dim nCols As Long
    Dim item As Variant
    Dim oIni As Variant
    Dim oFin As Variant
    Dim oId As Variant

    solapa = False
    suma = 0

    If IsEmpty(otras) Then Exit Sub

    If IsArray(otras) Then
        On Error GoTo Fin
        lb = LBound(otras, 1)
        ub = UBound(otras, 1)
        lb2 = LBound(otras, 2)
        nCols = UBound(otras, 2) - lb2 + 1
        For i = lb To ub
            oIni = otras(i, lb2)
            oFin = otras(i, lb2 + 1)
            oId = Empty
            If nCols >= 3 Then oId = otras(i, lb2 + 2)
            EvaluarOtraParada oIni, oFin, oId, idExcluir, ini, fin, solapa, suma
        Next i
    ElseIf TypeName(otras) = "Collection" Then
        For Each item In otras
            If IsArray(item) Then
                On Error Resume Next
                oIni = item(LBound(item))
                oFin = item(LBound(item) + 1)
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                Else
                    oId = Empty
                    oId = item(LBound(item) + 2)
                    Err.Clear
                    On Error GoTo 0
                    EvaluarOtraParada oIni, oFin, oId, idExcluir, ini, fin, solapa, suma
                End If
            End If
        Next item
    End If
    Exit Sub

Fin:
    On Error GoTo 0
End Sub

Private Sub EvaluarOtraParada(ByVal oIni As Variant, ByVal oFin As Variant, ByVal oId As Variant, _
        ByVal idExcluir As String, ByVal ini As Date, ByVal fin As Date, _
        ByRef solapa As Boolean, ByRef suma As Double)

    Dim dIni As Date
    Dim dFin As Date

    If Not EsFechaValida(oIni) Or Not EsFechaValida(oFin) Then Exit Sub
    If CDbl(oIni) <= 0 Or CDbl(oFin) <= 0 Then Exit Sub

    If Len(idExcluir) > 0 Then
        If StrComp(Trim$(CStr(oId)), idExcluir, vbTextCompare) = 0 Then Exit Sub
    End If

    dIni = CDate(oIni)
    dFin = CDate(oFin)
    If dFin = dIni Then Exit Sub
    If dFin < dIni Then dFin = DateAdd("d", 1, dFin)

    suma = suma + DateDiff("n", dIni, dFin)
    If ini < dFin And fin > dIni Then solapa = True
End Sub

Private Function ValorEnLista(ByVal valor As String, ByVal lista As Variant) As Boolean
    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim item As Variant

    If IsEmpty(lista) Or IsNull(lista) Then Exit Function

    If IsArray(lista) Then
        On Error GoTo Fallo
        lb = LBound(lista)
        ub = UBound(lista)
        For i = lb To ub
            If StrComp(Trim$(CStr(lista(i))), valor, vbTextCompare) = 0 Then
                ValorEnLista = True
                Exit Function
            End If
        Next i
    ElseIf TypeName(lista) = "Collection" Then
        For Each item In lista
            If StrComp(Trim$(CStr(item)), valor, vbTextCompare) = 0 Then
                ValorEnLista = True
                Exit Function
            End If
        Next item
    Else
        If StrComp(Trim$(CStr(lista)), valor, vbTextCompare) = 0 Then ValorEnLista = True
    End If
    Exit Function

Fallo:
    ValorEnLista = False
End Function
