Attribute VB_Name = "DatosDemo"
Option Explicit

' F2a - Bloque 2: dataset de demostracion para el dashboard (F3).
' Los CSV estan en data\ejemplo\ (los genera tools\generar_dataset_demo.py).
' Cada archivo es un turno completo (uno o mas reportes); la primera columna
' (seccion) clasifica cada fila: REPORTE, PARADA, PRODUCCION, RECHAZO, TRAMO
' o REDUCCION. Las fechas van como dd/mm/yyyy hh:mm y los textos no llevan
' comas (el separador de campo es la coma).
'
' CargarDatasetDemo:
'   1. Reporte + produccion + rechazos por GuardarReporteCompleto
'      (pasa por Validaciones y hace rollback si falla).
'   2. Paradas por CrearParada, tramos por CrearTramo y reducciones por
'      CrearReduccion (las tres APIs validan).
'   3. Si el REPORTE lleva anulado=Si, se anula al final porque las APIs
'      de carga rechazan reportes ya anulados.
' Todo lo ficticio lleva prefijo DEMO- en nro_reporte y "DATO DE DEMO" en
' observaciones; el 498 real (sin prefijo) no lo toca la limpieza.
'
' LimpiarDatasetDemo borra fisicamente los reportes DEMO- y todos sus hijos
' (patron de LimpiarDatosPrueba en DatosPrueba). Los lotes quedan: al volver
' a cargar, BuscarOCrearLote los reutiliza y no hay duplicados.

Private Const PREFIJO As String = "DEMO-"
Private Const NOTA_DEMO As String = "DATO DE DEMO"

Public Function CargarDatasetDemoTexto() As String
    ' Ejecuta la carga y devuelve el resultado como texto (para pruebas
    ' automatizadas y para el MsgBox del wrapper).
    Dim fso As Object
    Dim carpeta As Object
    Dim archivo As Object
    Dim cargados As Long
    Dim fallados As Long
    Dim detalle As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = CarpetaDemo(fso)
    If carpeta Is Nothing Then
        CargarDatasetDemoTexto = "No se encontro la carpeta data\ejemplo junto a la raiz del proyecto."
        Exit Function
    End If

    cargados = 0
    fallados = 0
    detalle = ""

    For Each archivo In carpeta.Files
        If LCase$(fso.GetExtensionName(archivo.Name)) = "csv" Then
            ProcesarArchivo archivo, cargados, fallados, detalle
        End If
    Next archivo

    CargarDatasetDemoTexto = "Reportes cargados: " & cargados & _
        " | Reportes con error: " & fallados & _
        IIf(Len(detalle) = 0, "", " | Detalle: " & Replace(detalle, vbCrLf, " // "))
End Function

Public Sub CargarDatasetDemo(Optional ByVal silencioso As Boolean = False)
    Dim resumen As String

    resumen = CargarDatasetDemoTexto()

    If Not silencioso Then
        If InStr(resumen, "error") > 0 Then
            MsgBox "Dataset de demo cargado con errores." & vbCrLf & vbCrLf & resumen, _
                   vbExclamation, "CargarDatasetDemo"
        Else
            MsgBox "Dataset de demo cargado." & vbCrLf & vbCrLf & resumen, _
                   vbInformation, "CargarDatasetDemo"
        End If
    End If
End Sub

Public Sub LimpiarDatasetDemo(Optional ByVal silencioso As Boolean = False)
    Dim ids As Object
    Dim n As Long

    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = 1

    n = ReunirIdsDemo(ids)

    ' Hijos primero (los ids ya estan reunidos); despues los reportes DEMO-.
    On Error Resume Next
    BorrarFilasPorIds "tblParadas", "id_turno_lote", ids
    BorrarFilasPorIds "tblProduccion", "id_turno_lote", ids
    BorrarFilasPorIds "tblRechazos", "id_turno_lote", ids
    BorrarFilasPorIds "tblTramosSinEnvase", "id_turno_lote", ids
    BorrarFilasPorIds "tblReducciones", "id_turno_lote", ids
    n = BorrarReportesDemo()
    On Error GoTo 0

    If Not silencioso Then
        MsgBox "Limpieza de dataset demo terminada." & vbCrLf & _
               "Reportes DEMO eliminados: " & n & vbCrLf & _
               "(el reporte 498 se conserva)", _
               vbInformation, "LimpiarDatasetDemo"
    End If
End Sub

' ==========================================================================
' Carga de un archivo CSV
' ==========================================================================

Private Sub ProcesarArchivo(ByVal archivo As Object, _
        ByRef cargados As Long, ByRef fallados As Long, ByRef detalle As String)

    Dim ts As Object
    Dim texto As String
    Dim lineas As Variant
    Dim enc As Object
    Dim mapa As Object
    Dim i As Long
    Dim p As Long
    Dim nRep As Long
    Dim nPar As Long
    Dim nTram As Long
    Dim nRed As Long
    Dim iRep As Long
    Dim iPar As Long
    Dim iTram As Long
    Dim iRed As Long
    Dim nroActual As String
    Dim s As String
    Dim fila As Variant
    Dim datosR() As TDatosReporte
    Dim prodR() As TDatosProduccion
    Dim colRech() As Collection
    Dim anulR() As Boolean
    Dim okRep() As Boolean
    Dim falloRep() As Boolean
    Dim idTL() As String
    Dim paradaNro() As String
    Dim paradaIni() As Date
    Dim paradaFin() As Date
    Dim paradaOrigen() As String
    Dim paradaMotivo() As String
    Dim paradaTipo() As String
    Dim paradaInterp() As String
    Dim tramoNro() As String
    Dim tramoIni() As Date
    Dim tramoFin() As Date
    Dim redNro() As String
    Dim redIni() As Date
    Dim redFin() As Date
    Dim redOrigen() As String
    Dim redMotivo() As String
    Dim redVel() As Double
    Dim colVacia As New Collection
    Dim msg As String
    Dim dRep As TDatosReporte
    Dim idParada As String
    Dim d As Object
    Dim r As Boolean

    Set ts = archivo.OpenAsTextStream(1)
    texto = ts.ReadAll
    ts.Close

    lineas = Split(texto, vbCrLf)
    If UBound(lineas) < 1 Then
        detalle = detalle & archivo.Name & ": archivo vacio." & vbCrLf
        Exit Sub
    End If

    Set enc = CreateObject("Scripting.Dictionary")
    enc.CompareMode = 1
    LlenaEncabezado lineas(0), enc

    ' Pass 1: conteo de filas por tipo para dimensionar los arrays.
    nRep = 0
    nPar = 0
    nTram = 0
    nRed = 0
    For i = 1 To UBound(lineas)
        s = Trim$(CStr(lineas(i)))
        If Len(s) > 0 Then
            fila = Split(s, ",")
            Select Case UCase$(V(fila, enc, "seccion"))
                Case "REPORTE": nRep = nRep + 1
                Case "PARADA": nPar = nPar + 1
                Case "TRAMO": nTram = nTram + 1
                Case "REDUCCION": nRed = nRed + 1
            End Select
        End If
    Next i

    If nRep = 0 Then
        detalle = detalle & archivo.Name & ": no tiene fila REPORTE." & vbCrLf
        Exit Sub
    End If

    ReDim datosR(1 To nRep)
    ReDim prodR(1 To nRep)
    ReDim colRech(1 To nRep)
    ReDim anulR(1 To nRep)
    ReDim okRep(1 To nRep)
    ReDim falloRep(1 To nRep)
    ReDim idTL(1 To nRep)
    If nPar > 0 Then
        ReDim paradaNro(1 To nPar)
        ReDim paradaIni(1 To nPar)
        ReDim paradaFin(1 To nPar)
        ReDim paradaOrigen(1 To nPar)
        ReDim paradaMotivo(1 To nPar)
        ReDim paradaTipo(1 To nPar)
        ReDim paradaInterp(1 To nPar)
    End If
    If nTram > 0 Then
        ReDim tramoNro(1 To nTram)
        ReDim tramoIni(1 To nTram)
        ReDim tramoFin(1 To nTram)
    End If
    If nRed > 0 Then
        ReDim redNro(1 To nRed)
        ReDim redIni(1 To nRed)
        ReDim redFin(1 To nRed)
        ReDim redOrigen(1 To nRed)
        ReDim redMotivo(1 To nRed)
        ReDim redVel(1 To nRed)
    End If

    ' Pass 2: parseo. Produccion y rechazos se adjudican al ultimo REPORTE
    ' leido; paradas, tramos y reducciones guardan su nro_reporte.
    iRep = 0
    iPar = 0
    iTram = 0
    iRed = 0
    nroActual = ""

    For i = 1 To UBound(lineas)
        s = Trim$(CStr(lineas(i)))
        If Len(s) > 0 Then
            fila = Split(s, ",")
            Select Case UCase$(V(fila, enc, "seccion"))

            Case "REPORTE"
                iRep = iRep + 1
                nroActual = V(fila, enc, "nro_reporte")
                Set colRech(iRep) = New Collection
                With datosR(iRep)
                    .nro_reporte = nroActual
                    .fecha = FechaHora(V(fila, enc, "fecha_reporte"))
                    .turno_codigo = V(fila, enc, "turno")
                    .equipo = V(fila, enc, "equipo")
                    .encargado = V(fila, enc, "encargado")
                    .anulado = "No"
                    .notas_sin_horario = V(fila, enc, "notas_sin_horario")
                    .observaciones = V(fila, enc, "observaciones")
                    .fecha_hora_inicio = FechaHora(V(fila, enc, "horario_inicio"))
                    .fecha_hora_fin = FechaHora(V(fila, enc, "horario_fin"))
                End With
                If IsNumeric(V(fila, enc, "velocidad_informada")) Then
                    If CDbl(V(fila, enc, "velocidad_informada")) > 0 Then
                        datosR(iRep).velocidad_informada = CDbl(V(fila, enc, "velocidad_informada"))
                    End If
                End If
                anulR(iRep) = EsAnuladoDatos(V(fila, enc, "anulado"))
                datosR(iRep).id_lote = AccesoHojas.BuscarOCrearLote( _
                    V(fila, enc, "lote"), V(fila, enc, "producto"), datosR(iRep).fecha)
                If Len(Trim$(datosR(iRep).id_lote)) = 0 Then
                    detalle = detalle & archivo.Name & ": lote " & _
                        V(fila, enc, "lote") & " invalido (" & nroActual & ")." & vbCrLf
                End If

            Case "PARADA"
                If iRep > 0 Then
                    iPar = iPar + 1
                    paradaNro(iPar) = nroActual
                    paradaIni(iPar) = FechaHora(V(fila, enc, "inicio"))
                    paradaFin(iPar) = FechaHora(V(fila, enc, "fin"))
                    paradaOrigen(iPar) = V(fila, enc, "origen")
                    paradaMotivo(iPar) = V(fila, enc, "motivo")
                    paradaTipo(iPar) = V(fila, enc, "tipo_parada")
                    paradaInterp(iPar) = V(fila, enc, "interpretada")
                End If

            Case "PRODUCCION"
                If iRep > 0 Then
                    prodR(iRep).unidades_producidas = NumFila(fila, enc, "unidades")
                    prodR(iRep).pallets = NumFila(fila, enc, "pallets")
                End If

            Case "RECHAZO"
                If iRep > 0 Then
                    Set d = CreateObject("Scripting.Dictionary")
                    d.CompareMode = 1
                    d("seccion") = V(fila, enc, "rechazo_seccion")
                    d("motivo") = V(fila, enc, "rechazo_motivo")
                    d("cantidad") = NumFila(fila, enc, "cantidad")
                    d("detalle") = V(fila, enc, "detalle")
                    colRech(iRep).Add d
                End If

            Case "TRAMO"
                If iRep > 0 Then
                    iTram = iTram + 1
                    tramoNro(iTram) = nroActual
                    tramoIni(iTram) = FechaHora(V(fila, enc, "inicio"))
                    tramoFin(iTram) = FechaHora(V(fila, enc, "fin"))
                End If

            Case "REDUCCION"
                If iRep > 0 Then
                    iRed = iRed + 1
                    redNro(iRed) = nroActual
                    redIni(iRed) = FechaHora(V(fila, enc, "inicio"))
                    redFin(iRed) = FechaHora(V(fila, enc, "fin"))
                    redOrigen(iRed) = V(fila, enc, "origen")
                    redMotivo(iRed) = V(fila, enc, "motivo")
                    redVel(iRed) = NumFila(fila, enc, "velocidad")
                End If

            End Select
        End If
    Next i

    ' Fase 1: reportes + produccion + rechazos por GuardarReporteCompleto.
    For i = 1 To nRep
        If Len(Trim$(datosR(i).id_lote)) = 0 Then
            okRep(i) = False
        Else
            msg = ""
            okRep(i) = AccesoHojas.GuardarReporteCompleto(datosR(i), prodR(i), _
                colRech(i), colVacia, msg)
            If Not okRep(i) Then
                detalle = detalle & archivo.Name & ": " & nroDe(i, datosR) & ": " & msg & vbCrLf
            End If
        End If
        If okRep(i) Then
            dRep = AccesoHojas.LeerReportePorNro(nroDe(i, datosR))
            If dRep.Encontrado Then idTL(i) = dRep.id_turno_lote
        End If
    Next i

    ' Mapa nro_reporte -> indice del reporte.
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = 1
    For i = 1 To nRep
        If Len(nroDe(i, datosR)) > 0 Then mapa(nroDe(i, datosR)) = i
    Next i

    ' Fase 2: paradas (CrearParada valida horario, solapamientos y suma).
    For p = 1 To nPar
        If mapa.Exists(paradaNro(p)) Then
            i = mapa(paradaNro(p))
            If okRep(i) And Not falloRep(i) Then
                msg = ""
                idParada = ""
                s = paradaInterp(p)
                If Len(Trim$(s)) = 0 Then
                    r = AccesoHojas.CrearParada(idTL(i), paradaIni(p), paradaFin(p), _
                        paradaOrigen(p), paradaMotivo(p), paradaTipo(p), _
                        idParada, "", , msg)
                Else
                    r = AccesoHojas.CrearParada(idTL(i), paradaIni(p), paradaFin(p), _
                        paradaOrigen(p), paradaMotivo(p), paradaTipo(p), _
                        idParada, "", s, msg)
                End If
                If Not r Then
                    falloRep(i) = True
                    detalle = detalle & archivo.Name & ": " & paradaNro(p) & _
                        " parada: " & msg & vbCrLf
                End If
            End If
        End If
    Next p

    ' Fase 3: tramos sin envase.
    For p = 1 To nTram
        If mapa.Exists(tramoNro(p)) Then
            i = mapa(tramoNro(p))
            If okRep(i) And Not falloRep(i) Then
                msg = ""
                If Not AccesoHojas.CrearTramo(idTL(i), tramoIni(p), tramoFin(p), msg) Then
                    falloRep(i) = True
                    detalle = detalle & archivo.Name & ": " & tramoNro(p) & _
                        " tramo: " & msg & vbCrLf
                End If
            End If
        End If
    Next p

    ' Fase 4: reducciones de velocidad.
    For p = 1 To nRed
        If mapa.Exists(redNro(p)) Then
            i = mapa(redNro(p))
            If okRep(i) And Not falloRep(i) Then
                msg = ""
                If Not AccesoHojas.CrearReduccion(idTL(i), redIni(p), redFin(p), _
                        redOrigen(p), redMotivo(p), redVel(p), msg) Then
                    falloRep(i) = True
                    detalle = detalle & archivo.Name & ": " & redNro(p) & _
                        " reduccion: " & msg & vbCrLf
                End If
            End If
        End If
    Next p

    ' Fase 5: anulacion al final de todo (las APIs de carga rechazan
    ' reportes anulados, asi que se anula una vez cargado completo).
    For i = 1 To nRep
        If okRep(i) And Not falloRep(i) And anulR(i) Then
            msg = ""
            If Not AccesoHojas.AnularReporte(idTL(i), msg) Then
                falloRep(i) = True
                detalle = detalle & archivo.Name & ": " & nroDe(i, datosR) & _
                    " anulacion: " & msg & vbCrLf
            End If
        End If
    Next i

    For i = 1 To nRep
        If okRep(i) And Not falloRep(i) Then
            cargados = cargados + 1
        Else
            fallados = fallados + 1
        End If
    Next i
End Sub

Private Function nroDe(ByVal indice As Long, ByRef datosR() As TDatosReporte) As String
    nroDe = datosR(indice).nro_reporte
End Function

' ==========================================================================
' Limpieza
' ==========================================================================

Private Function ReunirIdsDemo(ByVal ids As Object) As Long
    Dim lo As ListObject
    Dim rng As Range
    Dim cNro As Long
    Dim cId As Long
    Dim i As Long
    Dim sNro As String

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cNro = Utils.IndiceColumna(lo, "nro_reporte")
    cId = Utils.IndiceColumna(lo, "id_turno_lote")
    If cNro = 0 Or cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    ReunirIdsDemo = 0
    For i = 1 To rng.Rows.Count
        sNro = ValorCelda(rng.Cells(i, 1), cNro)
        If InStr(1, sNro, PREFIJO, vbTextCompare) = 1 Then
            ids(ValorCelda(rng.Cells(i, 1), cId)) = True
            ReunirIdsDemo = ReunirIdsDemo + 1
        End If
    Next i
End Function

Private Function BorrarReportesDemo() As Long
    Dim lo As ListObject
    Dim rng As Range
    Dim cNro As Long
    Dim i As Long
    Dim n As Long
    Dim sNro As String

    n = 0
    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cNro = Utils.IndiceColumna(lo, "nro_reporte")
    If cNro = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = rng.Rows.Count To 1 Step -1
        sNro = ValorCelda(rng.Cells(i, 1), cNro)
        If InStr(1, sNro, PREFIJO, vbTextCompare) = 1 Then
            lo.ListRows(i).Delete
            n = n + 1
        End If
    Next i

    BorrarReportesDemo = n
End Function

Private Sub BorrarFilasPorIds(ByVal nombreTabla As String, ByVal columna As String, _
        ByVal ids As Object)

    Dim lo As ListObject
    Dim rng As Range
    Dim c As Long
    Dim i As Long
    Dim s As String

    Set lo = Utils.Tabla(nombreTabla)
    If lo Is Nothing Then Exit Sub

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = rng.Rows.Count To 1 Step -1
        s = ValorCelda(rng.Cells(i, 1), c)
        If Len(s) > 0 Then
            If ids.Exists(s) Then lo.ListRows(i).Delete
        End If
    Next i
End Sub

' ==========================================================================
' Helpers
' ==========================================================================

Private Function CarpetaDemo(ByVal fso As Object) As Object
    Dim raiz As String
    Dim ruta As String

    raiz = fso.GetParentFolderName(ThisWorkbook.Path)
    If Len(raiz) = 0 Then Exit Function
    ruta = fso.BuildPath(raiz, "data\ejemplo")
    If fso.FolderExists(ruta) Then Set CarpetaDemo = fso.GetFolder(ruta)
End Function

Private Sub LlenaEncabezado(ByVal linea As String, ByVal enc As Object)
    Dim partes As Variant
    Dim i As Long

    enc.RemoveAll
    partes = Split(linea, ",")
    For i = LBound(partes) To UBound(partes)
        enc(Trim$(CStr(partes(i)))) = i
    Next i
End Sub

Private Function V(ByVal fila As Variant, ByVal enc As Object, _
        ByVal nombre As String) As String

    Dim idx As Long

    If enc.Exists(nombre) Then
        idx = enc(nombre)
        If idx <= UBound(fila) Then V = Trim$(CStr(fila(idx)))
    End If
End Function

Private Function NumFila(ByVal fila As Variant, ByVal enc As Object, _
        ByVal nombre As String) As Double

    Dim s As String

    s = V(fila, enc, nombre)
    If IsNumeric(s) Then NumFila = CDbl(s)
End Function

Private Function FechaHora(ByVal s As String) As Date
    s = Replace(s, ".", "/")
    If IsDate(s) Then FechaHora = CDate(s)
End Function

Private Function EsAnuladoDatos(ByVal s As String) As Boolean
    s = Trim$(s)
    EsAnuladoDatos = (StrComp(s, "S" & ChrW(237), vbTextCompare) = 0) _
                     Or (StrComp(s, "Si", vbTextCompare) = 0)
End Function

Private Function ValorCelda(ByVal rng As Range, ByVal columna As Long) As String
    Dim v As Variant

    If columna = 0 Then Exit Function
    v = rng.Cells(1, columna).Value
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    ValorCelda = Trim$(CStr(v))
End Function
