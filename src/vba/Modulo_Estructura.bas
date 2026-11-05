Attribute VB_Name = "Modulo_Estructura"
Option Explicit

' F2.1 - CompletarEstructura
' Lleva el libro maestro a la estructura de docs/estructura-datos.md (Doc v0.8).
' Es idempotente: verifica hoja, tabla y columna antes de crear.
' Los nombres de hoja, tabla y encabezados van en snake_case, sin tildes.
' Las listas con tildes se arman con ChrW para mantener el .bas en ASCII.

Private Const BUFFER_FILAS As Long = 50
Private Const HOJA_TRAMOS As String = "18_TramosSinEnvase"
Private Const TABLA_TRAMOS As String = "tblTramosSinEnvase"

Private mTablasCreadas As Long
Private mColumnasAgregadas As Long
Private mValidacionesAplicadas As Long
Private mAdvertencias As String

Public Sub CompletarEstructura(Optional ByVal silencioso As Boolean = False)
    Dim sep As String
    Dim lo As ListObject

    On Error GoTo Fallo

    mTablasCreadas = 0
    mColumnasAgregadas = 0
    mValidacionesAplicadas = 0
    mAdvertencias = ""

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    sep = SeparadorLista()

    ' 1. Tabla nueva: tramos sin envase de cada reporte.
    Set lo = AsegurarTabla(HOJA_TRAMOS, TABLA_TRAMOS, HeadersTramos())
    AsegurarColumna lo, "id_turno_lote", 1
    AsegurarColumna lo, "fecha_hora_inicio", 2
    AsegurarColumna lo, "fecha_hora_fin", 3
    FormatearFechas lo

    ' 2. tblTurnosRegistro: datos del reporte de linea.
    '    fecha_hora_inicio y fecha_hora_fin definen el TD real del turno
    '    (horario anotado en el reporte); ya no son columnas informativas.
    Set lo = AsegurarTabla("11_Turnos", "tblTurnosRegistro", HeadersTurnosRegistro())
    If Not lo Is Nothing Then
        AsegurarColumna lo, "control_calidad", 14
        AsegurarColumna lo, "nro_registro_sopladora", 15
        AsegurarColumna lo, "registros_bolsones", 16
        AsegurarColumna lo, "nro_reporte", 4
        AsegurarColumna lo, "encargado", 8
        AsegurarColumna lo, "estado_revision", 9
        AsegurarColumna lo, "anulado", 10
        AsegurarColumna lo, "notas_sin_horario", 11
        AsegurarColumna lo, "observaciones", 12
        AsegurarColumna lo, "velocidad_informada", 13
        FormatearFechas lo
    End If

    ' 3. tblParadas: marca de parada interpretada.
    Set lo = AsegurarTabla("12_Paradas", "tblParadas", HeadersParadas())
    If Not lo Is Nothing Then
        AsegurarColumna lo, "interpretada", 14
        ValorPorDefecto lo, "interpretada", "No"
        FormatearFechas lo
    End If

    ' 4. Detalle opcional en rechazos y perdidas de insumo.
    Set lo = AsegurarTabla("14_Rechazos", "tblRechazos", HeadersRechazos())
    If Not lo Is Nothing Then
        AsegurarColumna lo, "detalle", 8
        FormatearFechas lo
    End If

    Set lo = AsegurarTabla("16_PerdidasInsumo", "tblPerdidasInsumo", HeadersPerdidasInsumo())
    If Not lo Is Nothing Then
        AsegurarColumna lo, "detalle", 7
        FormatearFechas lo
    End If

    ' 5. tblVelocidades: condicion de envase y unidad en el orden del Doc.
    Set lo = AsegurarTabla("04_Cat_Productos", "tblVelocidades", HeadersVelocidades())
    If Not lo Is Nothing Then
        AsegurarColumna lo, "condicion_envase", 3
        AsegurarColumna lo, "unidad", 5
        FormatearFechas lo
    End If

    ' 6. Validaciones de columnas de catalogos y de la columna activo (Si/No).
    ConfigurarValidaciones sep

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If Not silencioso Then
        If Len(mAdvertencias) = 0 Then
            MsgBox "Estructura completada correctamente." & vbCrLf & vbCrLf & _
                   "Tablas creadas: " & mTablasCreadas & vbCrLf & _
                   "Columnas agregadas: " & mColumnasAgregadas & vbCrLf & _
                   "Validaciones configuradas: " & mValidacionesAplicadas, _
                   vbInformation, "CompletarEstructura"
        Else
            MsgBox "La estructura se completo con advertencias." & vbCrLf & vbCrLf & _
                   "Tablas creadas: " & mTablasCreadas & vbCrLf & _
                   "Columnas agregadas: " & mColumnasAgregadas & vbCrLf & _
                   "Validaciones configuradas: " & mValidacionesAplicadas & vbCrLf & vbCrLf & _
                   "Detalle:" & vbCrLf & mAdvertencias, _
                   vbExclamation, "CompletarEstructura"
        End If
    End If
    Exit Sub

Fallo:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not silencioso Then
        MsgBox "No se pudo completar la estructura." & vbCrLf & vbCrLf & _
               Err.Description & IIf(Len(mAdvertencias) = 0, "", vbCrLf & vbCrLf & _
               "Advertencias:" & vbCrLf & mAdvertencias), _
               vbCritical, "CompletarEstructura"
    End If
End Sub

Private Sub ConfigurarValidaciones(ByVal sep As String)
    Dim tablas As Variant
    Dim i As Long

    ' Columna activo (Si/No) en todos los catalogos que la tienen.
    tablas = Array( _
        "tblTiposParada", "tblOrigenes", "tblMotivos", "tblTurnos", _
        "tblProductos", "tblVelocidades", "tblSeccionesRechazo", _
        "tblMotivosRechazo", "tblInsumos", "tblPuestos")
    For i = LBound(tablas) To UBound(tablas)
        ValidarSino CStr(tablas(i)), "activo", sep
    Next i

    ValidarSino "tblTurnosRegistro", "anulado", sep
    ValidarSino "tblParadas", "interpretada", sep

    ' Columnas que guardan el codigo de otro catalogo.
    ValidarCatalogo "tblParametros", "producto_codigo", "tblProductos", "codigo"

    ValidarCatalogo "tblMotivos", "origen", "tblOrigenes", "codigo"
    ValidarCatalogo "tblMotivos", "tipo_sugerido", "tblTiposParada", "codigo"

    ValidarCatalogo "tblVelocidades", "producto", "tblProductos", "codigo"
    ValidarCatalogo "tblMotivosRechazo", "seccion", "tblSeccionesRechazo", "codigo"
    ValidarCatalogo "tblLotes", "producto", "tblProductos", "codigo"

    ValidarCatalogo "tblTurnosRegistro", "turno_codigo", "tblTurnos", "codigo"
    ValidarCatalogo "tblTurnosRegistro", "id_lote", "tblLotes", "id_lote"

    ValidarCatalogo "tblParadas", "turno", "tblTurnos", "codigo"
    ValidarCatalogo "tblParadas", "id_lote", "tblLotes", "id_lote"
    ValidarCatalogo "tblParadas", "origen", "tblOrigenes", "codigo"
    ValidarCatalogo "tblParadas", "motivo", "tblMotivos", "codigo"
    ValidarCatalogo "tblParadas", "tipo", "tblTiposParada", "codigo"

    ValidarCatalogo "tblProduccion", "turno", "tblTurnos", "codigo"
    ValidarCatalogo "tblProduccion", "lote", "tblLotes", "id_lote"
    ValidarCatalogo "tblProduccion", "producto", "tblProductos", "codigo"

    ValidarCatalogo "tblRechazos", "turno", "tblTurnos", "codigo"
    ValidarCatalogo "tblRechazos", "lote", "tblLotes", "id_lote"
    ValidarCatalogo "tblRechazos", "seccion", "tblSeccionesRechazo", "codigo"
    ValidarCatalogo "tblRechazos", "motivo", "tblMotivosRechazo", "codigo"

    ValidarCatalogo "tblReducciones", "turno", "tblTurnos", "codigo"
    ValidarCatalogo "tblReducciones", "origen", "tblOrigenes", "codigo"
    ValidarCatalogo "tblReducciones", "motivo", "tblMotivos", "codigo"

    ValidarCatalogo "tblPerdidasInsumo", "turno", "tblTurnos", "codigo"
    ValidarCatalogo "tblPerdidasInsumo", "insumo", "tblInsumos", "codigo"

    ValidarCatalogo "tblPersonalTurno", "turno", "tblTurnos", "codigo"
    ValidarCatalogo "tblPersonalTurno", "puesto", "tblPuestos", "codigo"

    ' Listas fijas definidas en la seccion 5 del Doc.
    ValidarListaFija "tblOrigenes", "clase", ListaValidacion(sep, _
        "m" & ChrW(225) & "quina de la l" & ChrW(237) & "nea", _
        "servicio de apoyo", "otro")

    ValidarListaFija "tblTurnos", "equipo", ListaValidacion(sep, "SOP", "RES")
    ValidarListaFija "tblTurnosRegistro", "equipo", ListaValidacion(sep, "SOP", "RES")
    ValidarListaFija "tblPuestos", "equipo", ListaValidacion(sep, "SOP", "RES")

    ValidarListaFija "tblProductos", "formato", ListaValidacion(sep, _
        "botella", "bid" & ChrW(243) & "n")

    ValidarListaFija "tblVelocidades", "condicion_envase", _
        ListaValidacion(sep, "con envase", "sin envase")
    ValidarListaFija "tblVelocidades", "unidad", ListaValidacion(sep, _
        "botellas/min", "bidones/min", "botellas/hora", "bidones/hora")

    ValidarListaFija "tblInsumos", "unidad_medida", _
        ListaValidacion(sep, "unidades", "gramos", "peso")
End Sub

Private Sub ValidarSino(ByVal tabla As String, ByVal columna As String, ByVal sep As String)
    Dim lista As String

    lista = ListaValidacion(sep, "S" & ChrW(237), "No")
    ValidarListaTabla tabla, columna, lista, xlValidAlertStop, _
        "Valor no valido", "Seleccione S" & ChrW(237) & " o No."
End Sub

Private Sub ValidarCatalogo(ByVal tabla As String, ByVal columna As String, _
                            ByVal tablaCatalogo As String, ByVal columnaCatalogo As String)
    Dim loCat As ListObject
    Dim idxCat As Long
    Dim rngCat As Range
    Dim formula As String

    Set loCat = Utils.Tabla(tablaCatalogo)
    If loCat Is Nothing Then Exit Sub

    idxCat = Utils.IndiceColumna(loCat, columnaCatalogo)
    If idxCat = 0 Then Exit Sub

    Set rngCat = RangoCatalogo(loCat.ListColumns(idxCat))
    If rngCat Is Nothing Then Exit Sub

    formula = "='" & Replace(rngCat.Worksheet.Name, "'", "''") & "'!" & _
              rngCat.Address(True, True, xlA1)

    ValidarListaTabla tabla, columna, formula, xlValidAlertWarning, _
        "Valor fuera del catalogo", "Se permite como valor nuevo (RN-08)."
End Sub

Private Sub ValidarListaFija(ByVal tabla As String, ByVal columna As String, ByVal lista As String)
    ValidarListaTabla tabla, columna, lista, xlValidAlertWarning, _
        "Valor fuera de la lista", "Se permite como valor nuevo (RN-08)."
End Sub

Private Sub ValidarListaTabla(ByVal tabla As String, ByVal columna As String, _
                              ByVal formula As String, ByVal alerta As Long, _
                              ByVal titulo As String, ByVal mensaje As String)
    Dim lo As ListObject
    Dim idx As Long
    Dim rng As Range

    Set lo = Utils.Tabla(tabla)
    If lo Is Nothing Then Exit Sub

    idx = Utils.IndiceColumna(lo, columna)
    If idx = 0 Then Exit Sub

    Set rng = RangoValidacion(lo, idx)
    If rng Is Nothing Then Exit Sub

    If ValidarListaRango(rng, formula, alerta, titulo, mensaje) Then
        mValidacionesAplicadas = mValidacionesAplicadas + 1
    End If
End Sub

Private Function ValidarListaRango(ByVal rng As Range, ByVal formula As String, _
                                   ByVal alerta As Long, ByVal titulo As String, _
                                   ByVal mensaje As String) As Boolean
    If rng Is Nothing Then Exit Function

    ' El Resume Next es solo para borrar una validacion previa: si no existe,
    ' Excel devuelve error y hay que seguir para poder crearla.
    On Error Resume Next
    rng.Validation.Delete
    Err.Clear
    On Error GoTo Fallo

    With rng.Validation
        .Add xlValidateList, alerta, xlBetween, formula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = titulo
        .ErrorMessage = mensaje
    End With

    ValidarListaRango = True
    Exit Function

Fallo:
    RegistrarAdvertencia "Validacion no aplicada en " & _
        rng.Address(False, False) & ": " & Err.Description
    ValidarListaRango = False
End Function

Private Function RangoValidacion(ByVal lo As ListObject, ByVal indiceColumna As Long) As Range
    Dim ws As Worksheet
    Dim filaIni As Long
    Dim filaFin As Long
    Dim col As Long

    If lo Is Nothing Then Exit Function

    Set ws = lo.Parent
    filaIni = lo.Range.Row + 1
    filaFin = lo.Range.Row + lo.Range.Rows.Count + BUFFER_FILAS

    If filaFin > ws.Rows.Count Then filaFin = ws.Rows.Count
    If filaFin < filaIni Then Exit Function

    col = lo.Range.Column + indiceColumna - 1
    Set RangoValidacion = ws.Range(ws.Cells(filaIni, col), ws.Cells(filaFin, col))
End Function

Private Function RangoCatalogo(ByVal lc As ListColumn) As Range
    Dim rng As Range
    Dim i As Long
    Dim fila As Long
    Dim v As Variant

    If lc Is Nothing Then Exit Function

    Set rng = lc.DataBodyRange
    If rng Is Nothing Then Exit Function
    If rng.Rows.Count = 0 Then Exit Function

    For i = rng.Rows.Count To 1 Step -1
        v = rng.Cells(i, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
            If Len(Trim$(CStr(v))) > 0 Then
                fila = rng.Row + i - 1
                Set RangoCatalogo = rng.Worksheet.Range( _
                    rng.Cells(1, 1), rng.Worksheet.Cells(fila, rng.Column))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function AsegurarTabla(ByVal nombreHoja As String, ByVal nombreTabla As String, _
                               ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    Dim ws As Worksheet

    Set lo = Utils.Tabla(nombreTabla)
    If Not lo Is Nothing Then
        Set AsegurarTabla = lo
        Exit Function
    End If

    Set ws = HojaOCrear(nombreHoja)
    If ws Is Nothing Then
        RegistrarAdvertencia "No se pudo obtener la hoja " & nombreHoja & _
            " para la tabla " & nombreTabla & "."
        Exit Function
    End If

    Set lo = CrearTabla(ws, nombreTabla, headers)
    If Not lo Is Nothing Then mTablasCreadas = mTablasCreadas + 1
    Set AsegurarTabla = lo
End Function

Private Function CrearTabla(ByVal ws As Worksheet, ByVal nombre As String, _
                            ByVal headers As Variant) As ListObject
    Dim nCols As Long
    Dim i As Long
    Dim colIni As Long
    Dim rng As Range
    Dim lo As ListObject

    On Error GoTo Fallo

    nCols = UBound(headers) - LBound(headers) + 1
    If nCols <= 0 Then Exit Function

    colIni = PrimeraColumnaLibre(ws)

    For i = 0 To nCols - 1
        ws.Cells(1, colIni + i).Value = headers(LBound(headers) + i)
    Next i

    Set rng = ws.Range(ws.Cells(1, colIni), ws.Cells(2, colIni + nCols - 1))
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = nombre

    Set CrearTabla = lo
    Exit Function

Fallo:
    RegistrarAdvertencia "No se pudo crear la tabla " & nombre & ": " & Err.Description
    If Not lo Is Nothing Then
        On Error Resume Next
        lo.Delete
        On Error GoTo 0
    End If
    Set CrearTabla = Nothing
End Function

Private Function AsegurarColumna(ByVal lo As ListObject, ByVal nombre As String, _
                                 Optional ByVal posicion As Long = 0) As Boolean
    Dim idx As Long
    Dim total As Long
    Dim nueva As ListColumn

    If lo Is Nothing Then Exit Function

    idx = Utils.IndiceColumna(lo, nombre)
    If idx > 0 Then Exit Function

    On Error GoTo Fallo

    total = lo.ListColumns.Count
    If posicion < 1 Or posicion > total Then
        Set nueva = lo.ListColumns.Add
    Else
        Set nueva = lo.ListColumns.Add(posicion)
    End If
    nueva.Name = nombre

    mColumnasAgregadas = mColumnasAgregadas + 1
    AsegurarColumna = True
    Exit Function

Fallo:
    RegistrarAdvertencia "No se pudo agregar la columna " & nombre & _
        " en la tabla " & lo.Name & ": " & Err.Description
    If Not nueva Is Nothing Then
        On Error Resume Next
        nueva.Delete
        On Error GoTo 0
    End If
End Function

Private Sub ValorPorDefecto(ByVal lo As ListObject, ByVal columna As String, ByVal valor As Variant)
    Dim idx As Long
    Dim cId As Long
    Dim rng As Range
    Dim i As Long
    Dim v As Variant
    Dim incluir As Boolean

    If lo Is Nothing Then Exit Sub

    idx = Utils.IndiceColumna(lo, columna)
    If idx = 0 Then Exit Sub

    cId = Utils.IndiceColumna(lo, "id")
    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = 1 To rng.Rows.Count
        If cId = 0 Then
            incluir = True
        Else
            v = rng.Cells(i, cId).Value
            If IsError(v) Then
                incluir = False
            Else
                incluir = (Len(Trim$(CStr(v))) > 0)
            End If
        End If

        If incluir Then
            v = rng.Cells(i, idx).Value
            If IsEmpty(v) Or IsNull(v) Then
                rng.Cells(i, idx).Value = valor
            ElseIf Not IsError(v) Then
                If Len(Trim$(CStr(v))) = 0 Then rng.Cells(i, idx).Value = valor
            End If
        End If
    Next i
End Sub

Private Sub FormatearFechas(ByVal lo As ListObject)
    ' Excel en espanol usa "aaaa" para el ano.
    FormatearColumna lo, "fecha", "dd/mm/aaaa"
    FormatearColumna lo, "fecha_hora_inicio", "dd/mm/aaaa hh:mm"
    FormatearColumna lo, "fecha_hora_fin", "dd/mm/aaaa hh:mm"
    FormatearColumna lo, "vigencia_desde", "dd/mm/aaaa"
    FormatearColumna lo, "vigencia_hasta", "dd/mm/aaaa"
    FormatearColumna lo, "fecha_aprobacion", "dd/mm/aaaa"
End Sub

Private Sub FormatearColumna(ByVal lo As ListObject, ByVal nombre As String, ByVal formato As String)
    Dim idx As Long
    Dim rng As Range

    If lo Is Nothing Then Exit Sub

    idx = Utils.IndiceColumna(lo, nombre)
    If idx = 0 Then Exit Sub

    Set rng = lo.ListColumns(idx).DataBodyRange
    If rng Is Nothing Then Exit Sub

    rng.NumberFormat = formato
End Sub

Private Function HeadersTramos() As Variant
    HeadersTramos = Array("id_turno_lote", "fecha_hora_inicio", "fecha_hora_fin")
End Function

Private Function HeadersTurnosRegistro() As Variant
    HeadersTurnosRegistro = Array( _
        "id_turno", "id_turno_lote", "id_lote", "nro_reporte", "fecha", _
        "turno_codigo", "equipo", "encargado", "estado_revision", "anulado", _
        "notas_sin_horario", "observaciones", "velocidad_informada", _
        "fecha_hora_inicio", "fecha_hora_fin", "control_calidad", _
        "nro_registro_sopladora", "registros_bolsones")
End Function

Private Function HeadersParadas() As Variant
    HeadersParadas = Array( _
        "id", "fecha", "turno", "linea_id", "id_lote", "id_turno_lote", _
        "fecha_hora_inicio", "fecha_hora_fin", "duracion_min", "origen", "motivo", _
        "tipo", "codigo", "interpretada", "observaciones")
End Function

Private Function HeadersRechazos() As Variant
    HeadersRechazos = Array( _
        "fecha", "turno", "lote", "id_turno_lote", "seccion", "motivo", _
        "cantidad", "detalle")
End Function

Private Function HeadersPerdidasInsumo() As Variant
    HeadersPerdidasInsumo = Array( _
        "fecha", "turno", "insumo", "cantidad", "unidad", "motivo", "detalle")
End Function

Private Function HeadersVelocidades() As Variant
    HeadersVelocidades = Array( _
        "codigo", "producto", "condicion_envase", "velocidad", "unidad", _
        "vigencia_desde", "origen_dato", "activo")
End Function

Private Function HojaEstructura(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nombre, vbTextCompare) = 0 Then
            Set HojaEstructura = ws
            Exit Function
        End If
    Next ws
End Function

Private Function HojaOCrear(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    Dim antes As Worksheet

    Set ws = HojaEstructura(nombre)
    If Not ws Is Nothing Then
        Set HojaOCrear = ws
        Exit Function
    End If

    Set antes = HojaEstructura("20_Calculos")
    If Not antes Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=antes)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count))
    End If
    ws.Name = nombre

    Set HojaOCrear = ws
End Function

Private Function PrimeraColumnaLibre(ByVal ws As Worksheet) As Long
    Dim usado As Range

    Set usado = ws.UsedRange
    If Application.WorksheetFunction.CountA(usado) = 0 Then
        PrimeraColumnaLibre = 1
    Else
        PrimeraColumnaLibre = usado.Column + usado.Columns.Count + 1
    End If
End Function

Private Sub RegistrarAdvertencia(ByVal mensaje As String)
    Debug.Print "CompletarEstructura: " & mensaje
    If Len(mAdvertencias) > 0 Then mAdvertencias = mAdvertencias & vbCrLf
    mAdvertencias = mAdvertencias & "- " & mensaje
End Sub

Private Function ListaValidacion(ByVal sep As String, ParamArray valores() As Variant) As String
    Dim i As Long
    Dim s As String

    For i = LBound(valores) To UBound(valores)
        If i > LBound(valores) Then s = s & sep
        s = s & CStr(valores(i))
    Next i
    ListaValidacion = s
End Function

Private Function SeparadorLista() As String
    Dim sep As String

    sep = CStr(Application.International(xlListSeparator))
    If sep <> "," And sep <> ";" Then sep = ","
    SeparadorLista = sep
End Function
