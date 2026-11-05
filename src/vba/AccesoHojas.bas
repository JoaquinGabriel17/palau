Attribute VB_Name = "AccesoHojas"
Option Explicit

Private Const LINEA_ID As Long = 1
Private Const ESTADO_OK As String = "OK"
Private Const ESTADO_PENDIENTE As String = "Pendiente de revisar"
Private Const ESTADO_ANULADO As String = "Anulado"

Public Type TDatosReporte
    id_turno As String
    id_turno_lote As String
    id_lote As String
    producto As String
    nro_reporte As String
    fecha As Date
    turno_codigo As String
    equipo As String
    encargado As String
    estado_revision As String
    anulado As String
    notas_sin_horario As String
    observaciones As String
    velocidad_informada As Variant
    control_calidad As String
    nro_registro_sopladora As String
    registros_bolsones As String
    fecha_hora_inicio As Date
    fecha_hora_fin As Date
    Encontrado As Boolean
End Type

Public Type TDatosParada
    id As String
    fecha As Date
    turno As String
    linea_id As Variant
    id_lote As String
    id_turno_lote As String
    fecha_hora_inicio As Date
    fecha_hora_fin As Date
    duracion_min As Double
    origen As String
    motivo As String
    tipo As String
    codigo As String
    interpretada As String
    observaciones As String
    Encontrado As Boolean
End Type

Public Type TDatosProduccion
    fecha As Date
    turno As String
    linea_id As Variant
    lote As String
    id_turno_lote As String
    producto As String
    unidades_producidas As Double
    pallets As Double
    litros As Double
    velocidad_nominal As Double
    Encontrado As Boolean
End Type

Public Type TDatosRechazo
    fecha As Date
    turno As String
    lote As String
    id_turno_lote As String
    seccion As String
    motivo As String
    cantidad As Double
    detalle As String
    Encontrado As Boolean
End Type

Public Type TDatosReduccion
    fecha As Date
    turno As String
    id_turno_lote As String
    origen As String
    motivo As String
    fecha_hora_inicio As Date
    fecha_hora_fin As Date
    velocidad_real As Double
    Encontrado As Boolean
End Type

Public Function CrearReporte( _
        ByVal fecha As Date, _
        ByVal turnoCodigo As String, _
        ByVal idLote As String, _
        ByVal ini As Date, _
        ByVal fin As Date, _
        ByRef idTurnoLote As String, _
        Optional ByVal nroReporte As String = "", _
        Optional ByVal equipo As String = "", _
        Optional ByVal encargado As String = "", _
        Optional ByVal controlCalidad As String = "", _
        Optional ByVal nroRegistroSopladora As String = "", _
        Optional ByVal notasSinHorario As String = "", _
        Optional ByVal observaciones As String = "", _
        Optional ByVal velocidadInformada As Variant, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim lo As ListObject
    Dim fila As ListRow
    Dim res As TResultadoValidacion
    Dim idTurno As String
    Dim idNuevo As String
    Dim numeroLote As String

    idTurnoLote = ""
    mensaje = ""

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblTurnosRegistro."
        Exit Function
    End If

    If Not ExisteLote(idLote) Then
        mensaje = "El lote " & idLote & " no existe."
        Exit Function
    End If

    numeroLote = NumeroLotePorId(idLote)
    res = ValidarLote(numeroLote)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    res = ValidarReporte(nroReporte, NrosReporteExistentes(lo), "", True)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    idTurno = BuscarIdTurno(lo, fecha, turnoCodigo, equipo)
    If Len(idTurno) = 0 Then idTurno = SiguienteId(lo, "id_turno", "T", fecha)
    idNuevo = SiguienteId(lo, "id_turno_lote", "T", fecha)

    Set fila = lo.ListRows.Add
    idTurnoLote = idNuevo

    SetValor lo, fila, "id_turno", idTurno
    SetValor lo, fila, "id_turno_lote", idNuevo
    SetValor lo, fila, "id_lote", idLote
    SetValor lo, fila, "nro_reporte", nroReporte
    SetValor lo, fila, "fecha", fecha
    SetValor lo, fila, "turno_codigo", turnoCodigo
    SetValor lo, fila, "equipo", equipo
    SetValor lo, fila, "encargado", encargado
    SetValor lo, fila, "estado_revision", ESTADO_OK
    SetValor lo, fila, "anulado", "No"
    SetValor lo, fila, "notas_sin_horario", notasSinHorario
    SetValor lo, fila, "observaciones", observaciones
    If Not IsMissing(velocidadInformada) Then SetValor lo, fila, "velocidad_informada", velocidadInformada
    SetValor lo, fila, "control_calidad", controlCalidad
    SetValor lo, fila, "nro_registro_sopladora", nroRegistroSopladora
    SetValor lo, fila, "fecha_hora_inicio", ini
    SetValor lo, fila, "fecha_hora_fin", fin

    CrearReporte = True
End Function

Public Function LeerReporte(ByVal idTurnoLote As String) As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim d As TDatosReporte

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "id_turno_lote", idTurnoLote)
    If fila Is Nothing Then Exit Function

    d.id_turno = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_turno"))
    d.id_turno_lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_turno_lote"))
    d.id_lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_lote"))
    d.nro_reporte = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "nro_reporte"))
    d.fecha = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha"))
    d.turno_codigo = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "turno_codigo"))
    d.equipo = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "equipo"))
    d.encargado = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "encargado"))
    d.estado_revision = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "estado_revision"))
    d.anulado = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "anulado"))
    d.notas_sin_horario = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "notas_sin_horario"))
    d.observaciones = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "observaciones"))
    d.velocidad_informada = ValorVariant(fila.Range, Utils.IndiceColumna(lo, "velocidad_informada"))
    d.control_calidad = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "control_calidad"))
    d.nro_registro_sopladora = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "nro_registro_sopladora"))
    d.registros_bolsones = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "registros_bolsones"))
    d.fecha_hora_inicio = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha_hora_inicio"))
    d.fecha_hora_fin = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha_hora_fin"))
    d.producto = ProductoPorLote(d.id_lote)
    d.Encontrado = True

    LeerReporte = d
End Function

Public Function LeerReportePorNro(ByVal nroReporte As String) As TDatosReporte
    Dim lo As ListObject
    Dim rng As Range
    Dim cNro As Long
    Dim cId As Long
    Dim i As Long

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cNro = Utils.IndiceColumna(lo, "nro_reporte")
    cId = Utils.IndiceColumna(lo, "id_turno_lote")
    If cNro = 0 Or cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cNro), nroReporte, vbTextCompare) = 0 Then
            LeerReportePorNro = LeerReporte(ValorTexto(rng.Cells(i, 1), cId))
            Exit Function
        End If
    Next i
End Function

Public Function ActualizarReporte(ByRef datos As TDatosReporte, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim lo As ListObject
    Dim fila As ListRow
    Dim res As TResultadoValidacion
    Dim numeroLote As String

    mensaje = ""

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblTurnosRegistro."
        Exit Function
    End If

    If Len(Trim$(datos.id_turno_lote)) = 0 Then
        mensaje = "Falta el id_turno_lote del reporte."
        Exit Function
    End If

    Set fila = BuscarFila(lo, "id_turno_lote", datos.id_turno_lote)
    If fila Is Nothing Then
        mensaje = "No existe el reporte " & datos.id_turno_lote & "."
        Exit Function
    End If

    If Not ExisteLote(datos.id_lote) Then
        mensaje = "El lote " & datos.id_lote & " no existe."
        Exit Function
    End If

    numeroLote = NumeroLotePorId(datos.id_lote)
    res = ValidarLote(numeroLote)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    res = ValidarReporte(datos.nro_reporte, NrosReporteExistentes(lo, datos.id_turno_lote), "", True)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    SetValor lo, fila, "id_lote", datos.id_lote
    SetValor lo, fila, "nro_reporte", datos.nro_reporte
    SetValor lo, fila, "fecha", datos.fecha
    SetValor lo, fila, "turno_codigo", datos.turno_codigo
    SetValor lo, fila, "equipo", datos.equipo
    SetValor lo, fila, "encargado", datos.encargado
    SetValor lo, fila, "estado_revision", datos.estado_revision
    If Len(Trim$(datos.anulado)) = 0 Then
        SetValor lo, fila, "anulado", "No"
    Else
        SetValor lo, fila, "anulado", datos.anulado
    End If
    SetValor lo, fila, "notas_sin_horario", datos.notas_sin_horario
    SetValor lo, fila, "observaciones", datos.observaciones
    If Not IsEmpty(datos.velocidad_informada) Then SetValor lo, fila, "velocidad_informada", datos.velocidad_informada
    SetValor lo, fila, "control_calidad", datos.control_calidad
    SetValor lo, fila, "nro_registro_sopladora", datos.nro_registro_sopladora
    SetValor lo, fila, "fecha_hora_inicio", datos.fecha_hora_inicio
    SetValor lo, fila, "fecha_hora_fin", datos.fecha_hora_fin

    ActualizarReporte = True
End Function

Public Function AnularReporte(ByVal idTurnoLote As String, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim lo As ListObject
    Dim fila As ListRow

    mensaje = ""

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblTurnosRegistro."
        Exit Function
    End If

    Set fila = BuscarFila(lo, "id_turno_lote", idTurnoLote)
    If fila Is Nothing Then
        mensaje = "No existe el reporte " & idTurnoLote & "."
        Exit Function
    End If

    SetValor lo, fila, "anulado", TextoSi()
    SetValor lo, fila, "estado_revision", ESTADO_ANULADO
    AnularReporte = True
End Function

Public Function ExisteReporte(ByVal idTurnoLote As String) As Boolean
    Dim lo As ListObject

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function
    ExisteReporte = Not (BuscarFila(lo, "id_turno_lote", idTurnoLote) Is Nothing)
End Function

Public Function ReporteAnulado(ByVal idTurnoLote As String) As Boolean
    Dim d As TDatosReporte

    d = LeerReporte(idTurnoLote)
    If d.Encontrado Then ReporteAnulado = EsActivo(d.anulado)
End Function

Public Function CrearParada( _
        ByVal idTurnoLote As String, _
        ByVal ini As Date, _
        ByVal fin As Date, _
        ByVal origen As String, _
        ByVal motivo As String, _
        ByVal tipo As String, _
        ByRef idParada As String, _
        Optional ByVal observaciones As String = "", _
        Optional ByVal interpretada As String = "No", _
        Optional ByRef mensaje As String = "") As Boolean

    Dim reporte As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim res As TResultadoValidacion
    Dim iniTurno As Date
    Dim finTurno As Date
    Dim td As Double
    Dim suma As Double
    Dim otras As Variant
    Dim dur As Double
    Dim tipoFinal As String
    Dim activo As Boolean
    Dim idNuevo As String
    Dim turnoOk As Boolean

    idParada = ""
    mensaje = ""

    If Not ValidarYLeerReporte(idTurnoLote, reporte, mensaje) Then Exit Function

    dur = DuracionMinutos(ini, fin)
    tipoFinal = Trim$(tipo)

    RangoTurno reporte.id_turno, iniTurno, finTurno
    If iniTurno > 0 And finTurno > iniTurno Then td = DateDiff("n", iniTurno, finTurno)

    If Len(tipoFinal) = 0 Then tipoFinal = SugerirTipoParada(dur)

    turnoOk = TipoParadaExisteActivo(tipoFinal, activo)
    If Not turnoOk Then
        mensaje = "El tipo de parada " & tipoFinal & " no existe."
        Exit Function
    End If

    otras = OtrasParadasTurno(reporte.id_turno, "", suma)

    res = ValidarParada(ini, fin, iniTurno, finTurno, tipoFinal, motivo, dur, "", otras, td, suma, activo)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblParadas."
        Exit Function
    End If

    idNuevo = SiguienteId(lo, "id", "P", reporte.fecha)
    Set fila = lo.ListRows.Add
    idParada = idNuevo

    SetValor lo, fila, "id", idNuevo
    SetValor lo, fila, "fecha", reporte.fecha
    SetValor lo, fila, "turno", reporte.id_turno
    SetValor lo, fila, "linea_id", LINEA_ID
    SetValor lo, fila, "id_lote", reporte.id_lote
    SetValor lo, fila, "id_turno_lote", idTurnoLote
    SetValor lo, fila, "fecha_hora_inicio", ini
    SetValor lo, fila, "fecha_hora_fin", fin
    SetValor lo, fila, "duracion_min", dur
    SetValor lo, fila, "origen", origen
    SetValor lo, fila, "motivo", motivo
    SetValor lo, fila, "tipo", tipoFinal
    SetValor lo, fila, "codigo", idNuevo
    SetValor lo, fila, "interpretada", ValorSino(interpretada)
    SetValor lo, fila, "observaciones", observaciones

    If res.Resultado = ResultadoValidacion.AVISA Then MarcarPendiente idTurnoLote
    CrearParada = True
End Function

Public Function LeerParada(ByVal idParada As String) As TDatosParada
    Dim lo As ListObject
    Dim fila As ListRow
    Dim d As TDatosParada

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "id", idParada)
    If fila Is Nothing Then Exit Function

    d.id = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id"))
    d.fecha = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha"))
    d.turno = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "turno"))
    d.linea_id = ValorVariant(fila.Range, Utils.IndiceColumna(lo, "linea_id"))
    d.id_lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_lote"))
    d.id_turno_lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_turno_lote"))
    d.fecha_hora_inicio = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha_hora_inicio"))
    d.fecha_hora_fin = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha_hora_fin"))
    d.duracion_min = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "duracion_min"))
    d.origen = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "origen"))
    d.motivo = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "motivo"))
    d.tipo = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "tipo"))
    d.codigo = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "codigo"))
    d.interpretada = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "interpretada"))
    If Len(Trim$(d.interpretada)) = 0 Then d.interpretada = "No"
    d.observaciones = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "observaciones"))
    d.Encontrado = True

    LeerParada = d
End Function

Public Function ActualizarParada( _
        ByVal idParada As String, _
        ByVal ini As Date, _
        ByVal fin As Date, _
        ByVal origen As String, _
        ByVal motivo As String, _
        ByVal tipo As String, _
        Optional ByVal observaciones As String = "", _
        Optional ByVal interpretada As String = "No", _
        Optional ByRef mensaje As String = "") As Boolean

    Dim parada As TDatosParada
    Dim reporte As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim res As TResultadoValidacion
    Dim iniTurno As Date
    Dim finTurno As Date
    Dim td As Double
    Dim suma As Double
    Dim otras As Variant
    Dim dur As Double
    Dim tipoFinal As String
    Dim activo As Boolean

    mensaje = ""

    parada = LeerParada(idParada)
    If Not parada.Encontrado Then
        mensaje = "No existe la parada " & idParada & "."
        Exit Function
    End If

    If Not ValidarYLeerReporte(parada.id_turno_lote, reporte, mensaje) Then Exit Function

    dur = DuracionMinutos(ini, fin)
    tipoFinal = Trim$(tipo)
    If Len(tipoFinal) = 0 Then tipoFinal = SugerirTipoParada(dur)

    If Not TipoParadaExisteActivo(tipoFinal, activo) Then
        mensaje = "El tipo de parada " & tipoFinal & " no existe."
        Exit Function
    End If

    RangoTurno reporte.id_turno, iniTurno, finTurno
    If iniTurno > 0 And finTurno > iniTurno Then td = DateDiff("n", iniTurno, finTurno)

    otras = OtrasParadasTurno(reporte.id_turno, idParada, suma)

    res = ValidarParada(ini, fin, iniTurno, finTurno, tipoFinal, motivo, dur, idParada, otras, td, suma, activo)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    Set lo = Utils.Tabla("tblParadas")
    Set fila = BuscarFila(lo, "id", idParada)
    If fila Is Nothing Then
        mensaje = "No existe la parada " & idParada & "."
        Exit Function
    End If

    SetValor lo, fila, "fecha_hora_inicio", ini
    SetValor lo, fila, "fecha_hora_fin", fin
    SetValor lo, fila, "duracion_min", dur
    SetValor lo, fila, "origen", origen
    SetValor lo, fila, "motivo", motivo
    SetValor lo, fila, "tipo", tipoFinal
    SetValor lo, fila, "interpretada", ValorSino(interpretada)
    SetValor lo, fila, "observaciones", observaciones

    If res.Resultado = ResultadoValidacion.AVISA Then MarcarPendiente parada.id_turno_lote
    ActualizarParada = True
End Function

Public Function BorrarParada(ByVal idParada As String, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim lo As ListObject
    Dim fila As ListRow

    mensaje = ""

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblParadas."
        Exit Function
    End If

    Set fila = BuscarFila(lo, "id", idParada)
    If fila Is Nothing Then
        mensaje = "No existe la parada " & idParada & "."
        Exit Function
    End If

    fila.Delete
    BorrarParada = True
End Function

Public Function CrearProduccion( _
        ByVal idTurnoLote As String, _
        ByVal unidades As Variant, _
        Optional ByVal pallets As Variant = 0, _
        Optional ByVal velocidadNominal As Variant, _
        Optional ByRef mensaje As String = "") As Boolean

    If IsMissing(velocidadNominal) Then
        CrearProduccion = GuardarProduccion(idTurnoLote, unidades, pallets, Empty, False, mensaje)
    Else
        CrearProduccion = GuardarProduccion(idTurnoLote, unidades, pallets, velocidadNominal, False, mensaje)
    End If
End Function

Public Function ActualizarProduccion( _
        ByVal idTurnoLote As String, _
        ByVal unidades As Variant, _
        Optional ByVal pallets As Variant = 0, _
        Optional ByVal velocidadNominal As Variant, _
        Optional ByRef mensaje As String = "") As Boolean

    If IsMissing(velocidadNominal) Then
        ActualizarProduccion = GuardarProduccion(idTurnoLote, unidades, pallets, Empty, True, mensaje)
    Else
        ActualizarProduccion = GuardarProduccion(idTurnoLote, unidades, pallets, velocidadNominal, True, mensaje)
    End If
End Function

Public Function LeerProduccion(ByVal idTurnoLote As String) As TDatosProduccion
    Dim lo As ListObject
    Dim fila As ListRow
    Dim d As TDatosProduccion

    Set lo = Utils.Tabla("tblProduccion")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "id_turno_lote", idTurnoLote)
    If fila Is Nothing Then Exit Function

    d.fecha = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha"))
    d.turno = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "turno"))
    d.linea_id = ValorVariant(fila.Range, Utils.IndiceColumna(lo, "linea_id"))
    d.lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "lote"))
    d.id_turno_lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_turno_lote"))
    d.producto = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "producto"))
    d.unidades_producidas = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "unidades_producidas"))
    d.pallets = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "pallets"))
    d.litros = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "litros"))
    d.velocidad_nominal = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "velocidad_nominal"))
    d.Encontrado = True

    LeerProduccion = d
End Function

Public Function BorrarProduccion(ByVal idTurnoLote As String, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim lo As ListObject
    Dim rng As Range
    Dim c As Long
    Dim i As Long
    Dim borrados As Long

    mensaje = ""

    Set lo = Utils.Tabla("tblProduccion")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblProduccion."
        Exit Function
    End If

    c = Utils.IndiceColumna(lo, "id_turno_lote")
    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = rng.Rows.Count To 1 Step -1
        If StrComp(Trim$(CStr(rng.Cells(i, c).Value)), idTurnoLote, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
            borrados = borrados + 1
        End If
    Next i

    If borrados = 0 Then
        mensaje = "No hay produccion para el reporte " & idTurnoLote & "."
        Exit Function
    End If

    BorrarProduccion = True
End Function

Public Function CrearRechazo( _
        ByVal idTurnoLote As String, _
        ByVal seccion As String, _
        ByVal motivo As String, _
        ByVal cantidad As Variant, _
        Optional ByVal detalle As String = "", _
        Optional ByRef mensaje As String = "") As Boolean

    CrearRechazo = GuardarRechazo(idTurnoLote, seccion, motivo, cantidad, detalle, False, mensaje)
End Function

Public Function ActualizarRechazo( _
        ByVal idTurnoLote As String, _
        ByVal seccion As String, _
        ByVal motivo As String, _
        ByVal cantidad As Variant, _
        Optional ByVal detalle As String = "", _
        Optional ByRef mensaje As String = "") As Boolean

    ActualizarRechazo = GuardarRechazo(idTurnoLote, seccion, motivo, cantidad, detalle, True, mensaje)
End Function

Public Function LeerRechazo(ByVal idTurnoLote As String, ByVal seccion As String, _
        ByVal motivo As String) As TDatosRechazo

    Dim lo As ListObject
    Dim fila As ListRow
    Dim d As TDatosRechazo

    Set lo = Utils.Tabla("tblRechazos")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFilaRechazo(lo, idTurnoLote, seccion, motivo)
    If fila Is Nothing Then Exit Function

    d.fecha = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha"))
    d.turno = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "turno"))
    d.lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "lote"))
    d.id_turno_lote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_turno_lote"))
    d.seccion = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "seccion"))
    d.motivo = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "motivo"))
    d.cantidad = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "cantidad"))
    d.detalle = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "detalle"))
    d.Encontrado = True

    LeerRechazo = d
End Function

Public Function BorrarRechazo(ByVal idTurnoLote As String, ByVal seccion As String, _
        ByVal motivo As String, Optional ByRef mensaje As String = "") As Boolean

    Dim lo As ListObject
    Dim rng As Range
    Dim i As Long
    Dim borrados As Long
    Dim sSeg As String
    Dim sMot As String
    Dim sId As String

    mensaje = ""

    Set lo = Utils.Tabla("tblRechazos")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblRechazos."
        Exit Function
    End If

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = rng.Rows.Count To 1 Step -1
        sId = ValorTexto(rng.Cells(i, 1), Utils.IndiceColumna(lo, "id_turno_lote"))
        sSeg = ValorTexto(rng.Cells(i, 1), Utils.IndiceColumna(lo, "seccion"))
        sMot = ValorTexto(rng.Cells(i, 1), Utils.IndiceColumna(lo, "motivo"))

        If StrComp(sId, idTurnoLote, vbTextCompare) = 0 _
           And StrComp(sSeg, seccion, vbTextCompare) = 0 _
           And StrComp(sMot, motivo, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
            borrados = borrados + 1
        End If
    Next i

    If borrados = 0 Then
        mensaje = "No se encontro el rechazo indicado."
        Exit Function
    End If

    BorrarRechazo = True
End Function

Public Function AgregarValorNuevo(ByVal tabla As String, ByVal valor As String, _
        ByVal columna As String, Optional ByVal columnasExtra As Variant) As String

    Dim lo As ListObject
    Dim fila As ListRow
    Dim cCodigo As Long
    Dim cTexto As Long
    Dim rng As Range
    Dim i As Long
    Dim s As String
    Dim prefijo As String
    Dim codigo As String
    Dim extras As Object

    valor = Trim$(valor)
    If Len(valor) = 0 Then Exit Function

    Set extras = CreateObject("Scripting.Dictionary")
    extras.CompareMode = 1
    CargarColumnasExtra extras, columnasExtra

    Set lo = Utils.Tabla(tabla)
    If lo Is Nothing Then Exit Function

    cCodigo = Utils.IndiceColumna(lo, "codigo")
    cTexto = Utils.IndiceColumna(lo, columna)
    If cCodigo = 0 Or cTexto = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            s = Trim$(CStr(rng.Cells(i, cTexto).Value))
            If StrComp(s, valor, vbTextCompare) = 0 Then
                AgregarValorNuevo = Trim$(CStr(rng.Cells(i, cCodigo).Value))
                Exit Function
            End If
        Next i
    End If

    prefijo = PrefijoCatalogo(tabla, lo)
    codigo = ProximoCodigo(lo, prefijo)

    Set fila = lo.ListRows.Add
    SetValor lo, fila, "codigo", codigo
    SetValor lo, fila, columna, valor
    SetValor lo, fila, "activo", TextoSi()
    AplicarColumnasExtra lo, fila, columna, extras

    AgregarValorNuevo = codigo
End Function

Private Sub CargarColumnasExtra(ByVal destino As Object, ByVal columnasExtra As Variant)
    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim k As Variant

    If IsEmpty(columnasExtra) Or IsNull(columnasExtra) Then Exit Sub

    If IsObject(columnasExtra) Then
        If TypeName(columnasExtra) = "Dictionary" Then
            For Each k In columnasExtra.Keys
                destino(CStr(k)) = columnasExtra(k)
            Next k
        End If
        Exit Sub
    End If

    If IsArray(columnasExtra) Then
        On Error GoTo Fin
        lb = LBound(columnasExtra, 1)
        ub = UBound(columnasExtra, 1)
        For i = lb To ub
            destino(CStr(columnasExtra(i, LBound(columnasExtra, 2)))) = _
                columnasExtra(i, LBound(columnasExtra, 2) + 1)
        Next i
    End If
    Exit Sub

Fin:
    On Error GoTo 0
End Sub

Private Sub AplicarColumnasExtra(ByVal lo As ListObject, ByVal fila As ListRow, _
        ByVal columnaPrincipal As String, ByVal extras As Object)

    Dim col As String
    Dim k As Variant

    If extras Is Nothing Then Exit Sub

    For Each k In extras.Keys
        col = Trim$(CStr(k))
        If EsColumnaExtraValida(col, columnaPrincipal) Then
            If Utils.IndiceColumna(lo, col) > 0 Then
                SetValor lo, fila, col, extras(k)
            End If
        End If
    Next k
End Sub

Private Function EsColumnaExtraValida(ByVal columna As String, ByVal columnaPrincipal As String) As Boolean
    If Len(columna) = 0 Then Exit Function
    If StrComp(columna, "codigo", vbTextCompare) = 0 Then Exit Function
    If StrComp(columna, columnaPrincipal, vbTextCompare) = 0 Then Exit Function
    EsColumnaExtraValida = True
End Function

Public Function ObtenerParametro(ByVal nombre As String, Optional ByVal fecha As Date = 0) As Variant
    Dim lo As ListObject
    Dim rng As Range
    Dim cParametro As Long
    Dim cNombre As Long
    Dim cId As Long
    Dim cValor As Long
    Dim cDesde As Long
    Dim cHasta As Long
    Dim i As Long
    Dim desde As Date
    Dim hasta As Date
    Dim mejorFecha As Date
    Dim mejorValor As Variant
    Dim encontrado As Boolean

    If fecha = 0 Then fecha = Date
    fecha = Int(fecha)

    Set lo = Utils.Tabla("tblParametros")
    If lo Is Nothing Then Exit Function

    cParametro = Utils.IndiceColumna(lo, "parametro")
    cNombre = Utils.IndiceColumna(lo, "nombre")
    cId = Utils.IndiceColumna(lo, "id")
    cValor = Utils.IndiceColumna(lo, "valor")
    cDesde = Utils.IndiceColumna(lo, "vigente_desde")
    cHasta = Utils.IndiceColumna(lo, "vigente_hasta")
    If cValor = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If CoincideParametro(rng, i, cParametro, cNombre, cId, nombre) Then
            desde = 0
            hasta = 0
            If cDesde > 0 Then desde = ValorFecha(rng.Cells(i, 1), cDesde)
            If cHasta > 0 Then hasta = ValorFecha(rng.Cells(i, 1), cHasta)

            If (desde = 0 Or fecha >= Int(desde)) And (hasta = 0 Or fecha <= Int(hasta)) Then
                If Not encontrado Or desde >= mejorFecha Then
                    mejorFecha = desde
                    mejorValor = rng.Cells(i, cValor).Value
                    encontrado = True
                End If
            End If
        End If
    Next i

    If encontrado Then ObtenerParametro = mejorValor
End Function

Public Function ListarParadas(ByVal idTurnoLote As String) As Collection
    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim i As Long
    Dim col As Collection

    Set col = New Collection
    Set ListarParadas = col

    If Len(Trim$(idTurnoLote)) = 0 Then Exit Function

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cSeg), idTurnoLote, vbTextCompare) = 0 Then
            col.Add ItemParada(rng.Cells(i, 1), lo)
        End If
    Next i
End Function

Private Function ItemParada(ByVal rng As Range, ByVal lo As ListObject) As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("id") = ValorTexto(rng, Utils.IndiceColumna(lo, "id"))
    d("fecha") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha"))
    d("turno") = ValorTexto(rng, Utils.IndiceColumna(lo, "turno"))
    d("linea_id") = ValorVariant(rng, Utils.IndiceColumna(lo, "linea_id"))
    d("id_lote") = ValorTexto(rng, Utils.IndiceColumna(lo, "id_lote"))
    d("id_turno_lote") = ValorTexto(rng, Utils.IndiceColumna(lo, "id_turno_lote"))
    d("fecha_hora_inicio") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha_hora_inicio"))
    d("fecha_hora_fin") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha_hora_fin"))
    d("duracion_min") = ValorNumero(rng, Utils.IndiceColumna(lo, "duracion_min"))
    d("origen") = ValorTexto(rng, Utils.IndiceColumna(lo, "origen"))
    d("motivo") = ValorTexto(rng, Utils.IndiceColumna(lo, "motivo"))
    d("tipo") = ValorTexto(rng, Utils.IndiceColumna(lo, "tipo"))
    d("codigo") = ValorTexto(rng, Utils.IndiceColumna(lo, "codigo"))
    d("interpretada") = ValorTexto(rng, Utils.IndiceColumna(lo, "interpretada"))
    d("observaciones") = ValorTexto(rng, Utils.IndiceColumna(lo, "observaciones"))
    Set ItemParada = d
End Function

Public Function ListarRechazos(ByVal idTurnoLote As String) As Collection
    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim i As Long
    Dim col As Collection

    Set col = New Collection
    Set ListarRechazos = col

    If Len(Trim$(idTurnoLote)) = 0 Then Exit Function

    Set lo = Utils.Tabla("tblRechazos")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cSeg), idTurnoLote, vbTextCompare) = 0 Then
            col.Add ItemRechazo(rng.Cells(i, 1), lo)
        End If
    Next i
End Function

Private Function ItemRechazo(ByVal rng As Range, ByVal lo As ListObject) As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("fecha") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha"))
    d("turno") = ValorTexto(rng, Utils.IndiceColumna(lo, "turno"))
    d("lote") = ValorTexto(rng, Utils.IndiceColumna(lo, "lote"))
    d("id_turno_lote") = ValorTexto(rng, Utils.IndiceColumna(lo, "id_turno_lote"))
    d("seccion") = ValorTexto(rng, Utils.IndiceColumna(lo, "seccion"))
    d("motivo") = ValorTexto(rng, Utils.IndiceColumna(lo, "motivo"))
    d("cantidad") = ValorNumero(rng, Utils.IndiceColumna(lo, "cantidad"))
    d("detalle") = ValorTexto(rng, Utils.IndiceColumna(lo, "detalle"))
    Set ItemRechazo = d
End Function

Public Function ListarReportes(ByVal fecha As Date, ByVal turno As String) As Collection
    Dim lo As ListObject
    Dim rng As Range
    Dim cId As Long
    Dim cFecha As Long
    Dim cTurno As Long
    Dim cAnulado As Long
    Dim i As Long
    Dim f As Date
    Dim col As Collection
    Dim reporte As TDatosReporte

    Set col = New Collection
    Set ListarReportes = col

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cId = Utils.IndiceColumna(lo, "id_turno_lote")
    cFecha = Utils.IndiceColumna(lo, "fecha")
    cTurno = Utils.IndiceColumna(lo, "turno_codigo")
    cAnulado = Utils.IndiceColumna(lo, "anulado")
    If cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If Not Utils.EsAnulado(ValorTexto(rng.Cells(i, 1), cAnulado)) Then
            f = 0
            If cFecha > 0 Then f = ValorFecha(rng.Cells(i, 1), cFecha)
            If (fecha = 0 Or (f > 0 And Int(f) = Int(fecha))) And _
               (Len(Trim$(turno)) = 0 Or StrComp(ValorTexto(rng.Cells(i, 1), cTurno), turno, vbTextCompare) = 0) Then
                reporte = LeerReporte(ValorTexto(rng.Cells(i, 1), cId))
                If reporte.Encontrado Then col.Add ItemReporte(reporte)
            End If
        End If
    Next i
End Function

Private Function ItemReporte(ByRef d As TDatosReporte) As Object
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = 1
    item("id_turno") = d.id_turno
    item("id_turno_lote") = d.id_turno_lote
    item("id_lote") = d.id_lote
    item("producto") = d.producto
    item("nro_reporte") = d.nro_reporte
    item("fecha") = d.fecha
    item("turno_codigo") = d.turno_codigo
    item("equipo") = d.equipo
    item("encargado") = d.encargado
    item("estado_revision") = d.estado_revision
    item("anulado") = d.anulado
    item("notas_sin_horario") = d.notas_sin_horario
    item("observaciones") = d.observaciones
    item("velocidad_informada") = d.velocidad_informada
    item("control_calidad") = d.control_calidad
    item("nro_registro_sopladora") = d.nro_registro_sopladora
    item("registros_bolsones") = d.registros_bolsones
    item("fecha_hora_inicio") = d.fecha_hora_inicio
    item("fecha_hora_fin") = d.fecha_hora_fin
    Set ItemReporte = item
End Function

Public Function ListarCatalogo(ByVal tabla As String, ByVal columnaTexto As String, _
        Optional ByVal filtroColumna As String = "", Optional ByVal filtroValor As String = "") As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cCodigo As Long
    Dim cTexto As Long
    Dim cFiltro As Long
    Dim cActivo As Long
    Dim i As Long
    Dim sCodigo As String
    Dim sTexto As String
    Dim sFiltro As String
    Dim col As Collection
    Dim item As Object

    Set col = New Collection
    Set ListarCatalogo = col

    Set lo = Utils.Tabla(tabla)
    If lo Is Nothing Then Exit Function

    cCodigo = Utils.IndiceColumna(lo, "codigo")
    If cCodigo = 0 Then cCodigo = Utils.IndiceColumna(lo, "id_lote")
    cTexto = Utils.IndiceColumna(lo, columnaTexto)
    If cCodigo = 0 Or cTexto = 0 Then Exit Function

    If Len(Trim$(filtroColumna)) > 0 Then
        cFiltro = Utils.IndiceColumna(lo, filtroColumna)
        If cFiltro = 0 Then Exit Function
    End If

    cActivo = Utils.IndiceColumna(lo, "activo")

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sTexto = ValorTexto(rng.Cells(i, 1), cTexto)
        If Len(Trim$(sTexto)) > 0 Then
            If cActivo = 0 Or EsActivo(ValorTexto(rng.Cells(i, 1), cActivo)) Then
                sFiltro = ""
                If cFiltro > 0 Then sFiltro = ValorTexto(rng.Cells(i, 1), cFiltro)
                If Len(Trim$(filtroValor)) = 0 Or StrComp(sFiltro, filtroValor, vbTextCompare) = 0 Then
                    sCodigo = ValorTexto(rng.Cells(i, 1), cCodigo)
                    If Len(Trim$(sCodigo)) > 0 Then
                        Set item = CreateObject("Scripting.Dictionary")
                        item.CompareMode = 1
                        item("codigo") = sCodigo
                        item("texto") = sTexto
                        col.Add item
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function BuscarOCrearLote(ByVal numeroLote As String, ByVal producto As String, _
        ByVal fecha As Date) As String

    Dim lo As ListObject
    Dim loProd As ListObject
    Dim filaProd As ListRow
    Dim fila As ListRow
    Dim rng As Range
    Dim cNum As Long
    Dim cIni As Long
    Dim cId As Long
    Dim cProd As Long
    Dim cActivo As Long
    Dim i As Long
    Dim sNum As String
    Dim sProd As String
    Dim dIni As Date
    Dim anio As Long
    Dim res As TResultadoValidacion
    Dim idNuevo As String

    numeroLote = Trim$(numeroLote)
    producto = Trim$(producto)

    If Len(numeroLote) = 0 Or Len(producto) = 0 Then Exit Function
    If fecha = 0 Then fecha = Date
    anio = Year(fecha)

    res = ValidarLote(numeroLote)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then Exit Function

    Set loProd = Utils.Tabla("tblProductos")
    If loProd Is Nothing Then Exit Function
    Set filaProd = BuscarFila(loProd, "codigo", producto)
    If filaProd Is Nothing Then Exit Function
    cActivo = Utils.IndiceColumna(loProd, "activo")
    If cActivo > 0 Then
        If Not EsActivo(ValorTexto(filaProd.Range, cActivo)) Then Exit Function
    End If

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Function

    cNum = Utils.IndiceColumna(lo, "numero_lote")
    cIni = Utils.IndiceColumna(lo, "fecha_hora_inicio")
    cId = Utils.IndiceColumna(lo, "id_lote")
    cProd = Utils.IndiceColumna(lo, "producto")
    If cNum = 0 Or cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            sNum = ValorTexto(rng.Cells(i, 1), cNum)
            If StrComp(sNum, numeroLote, vbTextCompare) = 0 Then
                dIni = 0
                If cIni > 0 Then dIni = ValorFecha(rng.Cells(i, 1), cIni)
                If dIni > 0 Then
                    If Year(dIni) = anio Then
                        sProd = ""
                        If cProd > 0 Then sProd = ValorTexto(rng.Cells(i, 1), cProd)
                        If StrComp(sProd, producto, vbTextCompare) = 0 Then
                            BuscarOCrearLote = ValorTexto(rng.Cells(i, 1), cId)
                        End If
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    idNuevo = SiguienteId(lo, "id_lote", "L", fecha)
    Set fila = lo.ListRows.Add
    SetValor lo, fila, "id_lote", idNuevo
    ' numero_lote como texto para conservar los ceros a la izquierda (ej. 09261).
    EscribirTexto lo, fila, "numero_lote", numeroLote
    SetValor lo, fila, "producto", producto
    SetValor lo, fila, "fecha_hora_inicio", fecha
    SetValor lo, fila, "fecha_hora_fin", Empty

    BuscarOCrearLote = idNuevo
End Function

Private Sub EscribirTexto(ByVal lo As ListObject, ByVal fila As ListRow, _
        ByVal columna As String, ByVal valor As String)

    Dim c As Long

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Sub
    With fila.Range.Cells(1, c)
        .NumberFormat = "@"
        .Value = valor
    End With
End Sub

Public Function ListarTramos(ByVal idTurnoLote As String) As Collection
    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim i As Long
    Dim col As Collection
    Dim item As Object

    Set col = New Collection
    Set ListarTramos = col

    If Len(Trim$(idTurnoLote)) = 0 Then Exit Function

    Set lo = Utils.Tabla("tblTramosSinEnvase")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    cIni = Utils.IndiceColumna(lo, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(lo, "fecha_hora_fin")
    If cSeg = 0 Or cIni = 0 Or cFin = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cSeg), idTurnoLote, vbTextCompare) = 0 Then
            Set item = CreateObject("Scripting.Dictionary")
            item.CompareMode = 1
            item("id_turno_lote") = idTurnoLote
            item("fecha_hora_inicio") = ValorFecha(rng.Cells(i, 1), cIni)
            item("fecha_hora_fin") = ValorFecha(rng.Cells(i, 1), cFin)
            col.Add item
        End If
    Next i
End Function

Public Function CrearTramo(ByVal idTurnoLote As String, ByVal inicio As Date, ByVal fin As Date, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim reporte As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim tramos As Collection
    Dim item As Variant
    Dim tIni As Date
    Dim tFin As Date

    mensaje = ""

    If Not ValidarYLeerReporte(idTurnoLote, reporte, mensaje) Then Exit Function

    If inicio <= 0 Or fin <= 0 Then
        mensaje = "El tramo debe tener hora de inicio y fin."
        Exit Function
    End If
    If fin = inicio Then
        mensaje = "La hora de fin del tramo debe ser posterior a la de inicio."
        Exit Function
    End If
    If fin < inicio Then fin = DateAdd("d", 1, fin)

    If reporte.fecha_hora_inicio > 0 And reporte.fecha_hora_fin > 0 Then
        If inicio < reporte.fecha_hora_inicio Or fin > reporte.fecha_hora_fin Then
            mensaje = "El tramo cae fuera del horario del reporte."
            Exit Function
        End If
    End If

    Set tramos = ListarTramos(idTurnoLote)
    For Each item In tramos
        If IsObject(item) Then
            If TypeName(item) = "Dictionary" Then
                tIni = item("fecha_hora_inicio")
                tFin = item("fecha_hora_fin")
                If tIni > 0 And tFin > 0 Then
                    If fin > tIni And inicio < tFin Then
                        mensaje = "El tramo se superpone con otro tramo sin envase del reporte."
                        Exit Function
                    End If
                End If
            End If
        End If
    Next item

    Set lo = Utils.Tabla("tblTramosSinEnvase")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblTramosSinEnvase."
        Exit Function
    End If

    Set fila = lo.ListRows.Add
    SetValor lo, fila, "id_turno_lote", idTurnoLote
    SetValor lo, fila, "fecha_hora_inicio", inicio
    SetValor lo, fila, "fecha_hora_fin", fin
    CrearTramo = True
End Function

Public Function GuardarReporteCompleto(ByRef datos As TDatosReporte, _
        ByRef produccion As TDatosProduccion, _
        ByRef rechazos As Collection, _
        ByRef paradas As Collection, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim idTurnoLote As String
    Dim idParada As String
    Dim reporteCreado As Boolean
    Dim produccionGuardada As Boolean
    Dim paradasGuardadas As Collection
    Dim rechazosGuardados As Collection
    Dim item As Variant
    Dim seccion As String
    Dim motivo As String
    Dim detalle As String
    Dim cantidad As Variant
    Dim iniP As Date
    Dim finP As Date
    Dim origen As String
    Dim tipo As String
    Dim interpretada As String
    Dim observaciones As String
    Dim i As Long
    Dim arrRechazo As Variant

    mensaje = ""
    idTurnoLote = ""
    idParada = ""
    Set paradasGuardadas = New Collection
    Set rechazosGuardados = New Collection

    On Error GoTo Rollback

    If Not CrearReporte(datos.fecha, datos.turno_codigo, datos.id_lote, _
            datos.fecha_hora_inicio, datos.fecha_hora_fin, idTurnoLote, _
            nroReporte:=datos.nro_reporte, equipo:=datos.equipo, _
            encargado:=datos.encargado, controlCalidad:=datos.control_calidad, _
            nroRegistroSopladora:=datos.nro_registro_sopladora, _
            notasSinHorario:=datos.notas_sin_horario, observaciones:=datos.observaciones, _
            velocidadInformada:=datos.velocidad_informada, mensaje:=mensaje) Then
        GoTo Rollback
    End If
    reporteCreado = True

    If HayProduccion(produccion) Then
        If Not CrearProduccion(idTurnoLote, produccion.unidades_producidas, _
                produccion.pallets, produccion.velocidad_nominal, mensaje) Then
            GoTo Rollback
        End If
        produccionGuardada = True
    End If

    If Not (rechazos Is Nothing) Then
        For Each item In rechazos
            seccion = ItemTexto(item, "seccion")
            motivo = ItemTexto(item, "motivo")
            detalle = ItemTexto(item, "detalle")
            cantidad = ItemValor(item, "cantidad")
            If Not CrearRechazo(idTurnoLote, seccion, motivo, cantidad, detalle, mensaje) Then
                GoTo Rollback
            End If
            rechazosGuardados.Add Array(seccion, motivo)
        Next item
    End If

    If Not (paradas Is Nothing) Then
        For Each item In paradas
            iniP = ItemFecha(item, "fecha_hora_inicio")
            finP = ItemFecha(item, "fecha_hora_fin")
            origen = ItemTexto(item, "origen")
            motivo = ItemTexto(item, "motivo")
            tipo = ItemTexto(item, "tipo")
            interpretada = ItemTexto(item, "interpretada")
            observaciones = ItemTexto(item, "observaciones")
            idParada = ""
            If Not CrearParada(idTurnoLote, iniP, finP, origen, motivo, tipo, _
                    idParada, observaciones, interpretada, mensaje) Then
                GoTo Rollback
            End If
            If Len(idParada) > 0 Then paradasGuardadas.Add idParada
        Next item
    End If

    On Error GoTo 0
    GuardarReporteCompleto = True
    Exit Function

Rollback:
    On Error Resume Next
    For i = paradasGuardadas.Count To 1 Step -1
        BorrarParada CStr(paradasGuardadas(i))
    Next i
    If produccionGuardada Then BorrarProduccion idTurnoLote
    For i = rechazosGuardados.Count To 1 Step -1
        arrRechazo = rechazosGuardados(i)
        BorrarRechazo idTurnoLote, CStr(arrRechazo(0)), CStr(arrRechazo(1))
    Next i
    If reporteCreado Then BorrarReporteFisico idTurnoLote
    On Error GoTo 0

    If Len(mensaje) = 0 Then
        mensaje = "No se pudo guardar el reporte completo."
    Else
        mensaje = "No se pudo guardar el reporte completo: " & mensaje
    End If
    GuardarReporteCompleto = False
End Function

Private Function HayProduccion(ByRef p As TDatosProduccion) As Boolean
    HayProduccion = (p.unidades_producidas <> 0) Or (p.pallets <> 0) _
                    Or (Len(Trim$(p.producto)) > 0)
End Function

Private Function ItemValor(ByVal item As Variant, ByVal clave As String) As Variant
    If IsObject(item) Then
        If TypeName(item) = "Dictionary" Then
            If item.Exists(clave) Then ItemValor = item(clave)
        End If
    End If
End Function

Private Function ItemTexto(ByVal item As Variant, ByVal clave As String) As String
    Dim v As Variant

    v = ItemValor(item, clave)
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    ItemTexto = CStr(v)
End Function

Private Function ItemFecha(ByVal item As Variant, ByVal clave As String) As Date
    Dim v As Variant

    v = ItemValor(item, clave)
    If IsDate(v) Then ItemFecha = CDate(v)
End Function

Private Sub BorrarReporteFisico(ByVal idTurnoLote As String)
    Dim lo As ListObject
    Dim fila As ListRow

    If Len(Trim$(idTurnoLote)) = 0 Then Exit Sub

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Sub

    Set fila = BuscarFila(lo, "id_turno_lote", idTurnoLote)
    If Not fila Is Nothing Then fila.Delete
End Sub

Private Function GuardarProduccion(ByVal idTurnoLote As String, ByVal unidades As Variant, _
        ByVal pallets As Variant, ByVal velocidadNominal As Variant, _
        ByVal debeExistir As Boolean, ByRef mensaje As String) As Boolean

    Dim reporte As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim res As TResultadoValidacion
    Dim palletsNum As Double
    Dim vel As Double
    Dim litros As Double
    Dim volumen As Double
    Dim unidadesNum As Double

    mensaje = ""

    If Not ValidarYLeerReporte(idTurnoLote, reporte, mensaje) Then Exit Function

    res = ValidarProduccion(unidades, "Las unidades producidas")
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    If Not IsEmpty(pallets) And Not IsNull(pallets) Then
        If Not IsNumeric(pallets) Then
            mensaje = "Los pallets deben ser un numero."
            Exit Function
        End If
        If CDbl(pallets) < 0 Or CDbl(pallets) <> Int(CDbl(pallets)) Then
            mensaje = "Los pallets deben ser un entero no negativo."
            Exit Function
        End If
        palletsNum = CDbl(pallets)
    End If

    Set lo = Utils.Tabla("tblProduccion")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblProduccion."
        Exit Function
    End If

    Set fila = BuscarFila(lo, "id_turno_lote", idTurnoLote)
    If fila Is Nothing Then
        If debeExistir Then
            mensaje = "No hay produccion para el reporte " & idTurnoLote & "."
            Exit Function
        End If
        Set fila = lo.ListRows.Add
    End If

    unidadesNum = CDbl(unidades)
    volumen = VolumenProducto(reporte.producto)
    litros = Calculos.CalcularLitros(unidadesNum, volumen)

    If Not IsEmpty(velocidadNominal) And Not IsNull(velocidadNominal) Then
        If IsNumeric(velocidadNominal) Then vel = CDbl(velocidadNominal)
    End If
    If vel <= 0 Then vel = VelocidadVigente(reporte.producto, reporte.fecha)

    SetValor lo, fila, "fecha", reporte.fecha
    SetValor lo, fila, "turno", reporte.id_turno
    SetValor lo, fila, "linea_id", LINEA_ID
    SetValor lo, fila, "lote", reporte.id_lote
    SetValor lo, fila, "id_turno_lote", idTurnoLote
    SetValor lo, fila, "producto", reporte.producto
    SetValor lo, fila, "unidades_producidas", unidadesNum
    SetValor lo, fila, "pallets", palletsNum
    SetValor lo, fila, "litros", litros
    SetValor lo, fila, "velocidad_nominal", vel

    If res.Resultado = ResultadoValidacion.AVISA Then MarcarPendiente idTurnoLote
    GuardarProduccion = True
End Function

Private Function GuardarRechazo(ByVal idTurnoLote As String, ByVal seccion As String, _
        ByVal motivo As String, ByVal cantidad As Variant, ByVal detalle As String, _
        ByVal debeExistir As Boolean, ByRef mensaje As String) As Boolean

    Dim reporte As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim res As TResultadoValidacion

    mensaje = ""

    If Not ValidarYLeerReporte(idTurnoLote, reporte, mensaje) Then Exit Function

    res = ValidarRechazo(seccion, motivo, cantidad)
    If res.Resultado = ResultadoValidacion.BLOQUEA Then
        mensaje = res.Mensaje
        Exit Function
    End If

    Set lo = Utils.Tabla("tblRechazos")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblRechazos."
        Exit Function
    End If

    Set fila = BuscarFilaRechazo(lo, idTurnoLote, seccion, motivo)
    If fila Is Nothing Then
        If debeExistir Then
            mensaje = "No se encontro el rechazo indicado."
            Exit Function
        End If
        Set fila = lo.ListRows.Add
    End If

    SetValor lo, fila, "fecha", reporte.fecha
    SetValor lo, fila, "turno", reporte.id_turno
    SetValor lo, fila, "lote", reporte.id_lote
    SetValor lo, fila, "id_turno_lote", idTurnoLote
    SetValor lo, fila, "seccion", seccion
    SetValor lo, fila, "motivo", motivo
    SetValor lo, fila, "cantidad", CDbl(cantidad)
    SetValor lo, fila, "detalle", detalle

    If res.Resultado = ResultadoValidacion.AVISA Then MarcarPendiente idTurnoLote
    GuardarRechazo = True
End Function

Private Function ValidarYLeerReporte(ByVal idTurnoLote As String, _
        ByRef reporte As TDatosReporte, ByRef mensaje As String) As Boolean

    reporte = LeerReporte(idTurnoLote)
    If Not reporte.Encontrado Then
        mensaje = "No existe el reporte " & idTurnoLote & "."
        Exit Function
    End If
    If EsActivo(reporte.anulado) Then
        mensaje = "El reporte " & idTurnoLote & " esta anulado."
        Exit Function
    End If
    ValidarYLeerReporte = True
End Function

Private Function OtrasParadasTurno(ByVal idTurno As String, ByVal idExcluir As String, _
        ByRef suma As Double) As Variant

    Dim lo As ListObject
    Dim loTurnos As ListObject
    Dim rng As Range
    Dim rngTurnos As Range
    Dim cTurno As Long
    Dim cId As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim cSegParada As Long
    Dim cIdTurno As Long
    Dim cSegTurno As Long
    Dim cAnulado As Long
    Dim i As Long
    Dim sTurno As String
    Dim sId As String
    Dim sSeg As String
    Dim dIni As Date
    Dim dFin As Date
    Dim arr() As Variant
    Dim anulados As Object

    suma = 0

    Set anulados = CreateObject("Scripting.Dictionary")
    anulados.CompareMode = 1

    Set loTurnos = Utils.Tabla("tblTurnosRegistro")
    If Not loTurnos Is Nothing Then
        cIdTurno = Utils.IndiceColumna(loTurnos, "id_turno")
        cSegTurno = Utils.IndiceColumna(loTurnos, "id_turno_lote")
        cAnulado = Utils.IndiceColumna(loTurnos, "anulado")
        Set rngTurnos = loTurnos.DataBodyRange
        If cIdTurno > 0 And cSegTurno > 0 And cAnulado > 0 Then
            If Not rngTurnos Is Nothing Then
                For i = 1 To rngTurnos.Rows.Count
                    If StrComp(ValorTexto(rngTurnos.Cells(i, 1), cIdTurno), idTurno, vbTextCompare) = 0 Then
                        If Utils.EsAnulado(rngTurnos.Cells(i, cAnulado).Value) Then
                            sSeg = ValorTexto(rngTurnos.Cells(i, 1), cSegTurno)
                            If Len(sSeg) > 0 Then anulados(sSeg) = True
                        End If
                    End If
                Next i
            End If
        End If
    End If

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then Exit Function

    cTurno = Utils.IndiceColumna(lo, "turno")
    cId = Utils.IndiceColumna(lo, "id")
    cIni = Utils.IndiceColumna(lo, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(lo, "fecha_hora_fin")
    cSegParada = Utils.IndiceColumna(lo, "id_turno_lote")
    If cTurno = 0 Or cIni = 0 Or cFin = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    ReDim arr(1 To rng.Rows.Count, 1 To 3)

    For i = 1 To rng.Rows.Count
        sTurno = ValorTexto(rng.Cells(i, 1), cTurno)
        If StrComp(sTurno, idTurno, vbTextCompare) = 0 Then
            sSeg = ""
            If cSegParada > 0 Then sSeg = ValorTexto(rng.Cells(i, 1), cSegParada)
            If Not anulados.Exists(sSeg) Then
                sId = ""
                If cId > 0 Then sId = ValorTexto(rng.Cells(i, 1), cId)
                If Len(idExcluir) = 0 Or StrComp(sId, idExcluir, vbTextCompare) <> 0 Then
                    dIni = ValorFecha(rng.Cells(i, 1), cIni)
                    dFin = ValorFecha(rng.Cells(i, 1), cFin)
                    If dIni > 0 And dFin > 0 Then
                        If dFin < dIni Then dFin = DateAdd("d", 1, dFin)
                        arr(i, 1) = dIni
                        arr(i, 2) = dFin
                        arr(i, 3) = sId
                        suma = suma + DateDiff("n", dIni, dFin)
                    End If
                End If
            End If
        End If
    Next i

    OtrasParadasTurno = arr
End Function

Private Sub RangoTurno(ByVal idTurno As String, ByRef ini As Date, ByRef fin As Date)
    Dim lo As ListObject
    Dim rng As Range
    Dim cId As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim i As Long
    Dim fIni As Date
    Dim fFin As Date

    ini = 0
    fin = 0

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Sub

    cId = Utils.IndiceColumna(lo, "id_turno")
    cIni = Utils.IndiceColumna(lo, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(lo, "fecha_hora_fin")
    If cId = 0 Or cIni = 0 Or cFin = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cId), idTurno, vbTextCompare) = 0 Then
            fIni = ValorFecha(rng.Cells(i, 1), cIni)
            fFin = ValorFecha(rng.Cells(i, 1), cFin)
            If fIni > 0 And fFin > 0 Then
                If ini = 0 Or fIni < ini Then ini = fIni
                If fin = 0 Or fFin > fin Then fin = fFin
            End If
        End If
    Next i
End Sub

Private Function SugerirTipoParada(ByVal dur As Double) As String
    Dim umbral As Variant

    umbral = ObtenerParametro("UMBRAL_MICROPARADA_MIN")
    If IsNumeric(umbral) Then
        If dur <= CDbl(umbral) Then
            SugerirTipoParada = "TP-05"
            Exit Function
        End If
    End If
    SugerirTipoParada = "TP-06"
End Function

Private Function TipoParadaExisteActivo(ByVal tipo As String, ByRef activo As Boolean) As Boolean
    Dim lo As ListObject
    Dim fila As ListRow

    activo = False

    Set lo = Utils.Tabla("tblTiposParada")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "codigo", tipo)
    If fila Is Nothing Then Exit Function

    TipoParadaExisteActivo = True
    activo = EsActivo(ValorTexto(fila.Range, Utils.IndiceColumna(lo, "activo")))
End Function

Private Function DuracionMinutos(ByVal ini As Date, ByVal fin As Date) As Double
    Dim d As Date

    If fin = ini Then Exit Function
    d = fin
    If d < ini Then d = DateAdd("d", 1, d)
    DuracionMinutos = DateDiff("n", ini, d)
End Function

Private Function SiguienteId(ByVal lo As ListObject, ByVal columna As String, _
        ByVal prefijo As String, ByVal fecha As Date) As String

    Dim rng As Range
    Dim c As Long
    Dim i As Long
    Dim s As String
    Dim base As String
    Dim n As Long
    Dim maxN As Long

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Function

    base = prefijo & "-" & Format$(fecha, "yyyymmdd") & "-"
    Set rng = lo.DataBodyRange
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            s = Trim$(CStr(rng.Cells(i, c).Value))
            If Len(s) >= Len(base) Then
                If StrComp(Left$(s, Len(base)), base, vbTextCompare) = 0 Then
                    n = Val(Mid$(s, Len(base) + 1))
                    If n > maxN Then maxN = n
                End If
            End If
        Next i
    End If

    SiguienteId = base & Format$(maxN + 1, "00")
End Function

Private Function BuscarIdTurno(ByVal lo As ListObject, ByVal fecha As Date, _
        ByVal turnoCodigo As String, ByVal equipo As String) As String

    Dim rng As Range
    Dim cFecha As Long
    Dim cTurno As Long
    Dim cEquipo As Long
    Dim cId As Long
    Dim i As Long
    Dim f As Date

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cTurno = Utils.IndiceColumna(lo, "turno_codigo")
    cEquipo = Utils.IndiceColumna(lo, "equipo")
    cId = Utils.IndiceColumna(lo, "id_turno")
    If cFecha = 0 Or cTurno = 0 Or cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        f = ValorFecha(rng.Cells(i, 1), cFecha)
        If f > 0 Then
            If Int(f) = Int(fecha) Then
                If StrComp(ValorTexto(rng.Cells(i, 1), cTurno), turnoCodigo, vbTextCompare) = 0 Then
                    If Len(Trim$(equipo)) = 0 Then
                        BuscarIdTurno = ValorTexto(rng.Cells(i, 1), cId)
                        Exit Function
                    ElseIf StrComp(ValorTexto(rng.Cells(i, 1), cEquipo), equipo, vbTextCompare) = 0 Then
                        BuscarIdTurno = ValorTexto(rng.Cells(i, 1), cId)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Function NrosReporteExistentes(ByVal lo As ListObject, _
        Optional ByVal excluirId As String = "") As Variant

    Dim rng As Range
    Dim cNro As Long
    Dim cId As Long
    Dim i As Long
    Dim n As Long
    Dim s As String
    Dim arr() As String

    cNro = Utils.IndiceColumna(lo, "nro_reporte")
    cId = Utils.IndiceColumna(lo, "id_turno_lote")
    If cNro = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    ReDim arr(0 To rng.Rows.Count - 1)

    For i = 1 To rng.Rows.Count
        If Len(excluirId) = 0 Then
            s = ValorTexto(rng.Cells(i, 1), cNro)
        ElseIf StrComp(ValorTexto(rng.Cells(i, 1), cId), excluirId, vbTextCompare) <> 0 Then
            s = ValorTexto(rng.Cells(i, 1), cNro)
        Else
            s = ""
        End If

        If Len(s) > 0 Then
            arr(n) = s
            n = n + 1
        End If
    Next i

    If n = 0 Then
        NrosReporteExistentes = Empty
    Else
        ReDim Preserve arr(0 To n - 1)
        NrosReporteExistentes = arr
    End If
End Function

Private Function PrefijoCatalogo(ByVal tabla As String, ByVal lo As ListObject) As String
    Dim rng As Range
    Dim c As Long
    Dim i As Long
    Dim s As String

    Select Case LCase$(tabla)
        Case "tbltiposparada": PrefijoCatalogo = "TP"
        Case "tblorigenes": PrefijoCatalogo = "OR"
        Case "tblmotivos": PrefijoCatalogo = "MO"
        Case "tblseccionesrechazo": PrefijoCatalogo = "SE"
        Case "tblmotivosrechazo": PrefijoCatalogo = "MR"
        Case "tblinsumos": PrefijoCatalogo = "IN"
        Case "tblpuestos": PrefijoCatalogo = "PU"
        Case "tblvelocidades": PrefijoCatalogo = "VEL"
    End Select
    If Len(PrefijoCatalogo) > 0 Then Exit Function

    c = Utils.IndiceColumna(lo, "codigo")
    If c > 0 Then
        Set rng = lo.DataBodyRange
        If Not rng Is Nothing Then
            For i = 1 To rng.Rows.Count
                s = Trim$(CStr(rng.Cells(i, c).Value))
                If InStr(1, s, "-") > 0 Then
                    PrefijoCatalogo = Left$(s, InStrRev(s, "-") - 1)
                    Exit Function
                End If
            Next i
        End If
    End If

    If Len(tabla) > 3 Then PrefijoCatalogo = UCase$(Left$(Mid$(tabla, 4), 2))
End Function

Private Function ProximoCodigo(ByVal lo As ListObject, ByVal prefijo As String) As String
    Dim rng As Range
    Dim c As Long
    Dim i As Long
    Dim s As String
    Dim n As Long
    Dim maxN As Long

    c = Utils.IndiceColumna(lo, "codigo")
    If c = 0 Or Len(prefijo) = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            s = Trim$(CStr(rng.Cells(i, c).Value))
            If Len(s) > 0 Then
                If InStr(1, s, prefijo & "-", vbTextCompare) = 1 Then
                    n = Val(Mid$(s, Len(prefijo) + 2))
                    If n > maxN Then maxN = n
                End If
            End If
        Next i
    End If

    ProximoCodigo = prefijo & "-" & Format$(maxN + 1, "00")
End Function

Private Function CoincideParametro(ByVal rng As Range, ByVal fila As Long, _
        ByVal cParametro As Long, ByVal cNombre As Long, ByVal cId As Long, _
        ByVal nombre As String) As Boolean

    If cParametro > 0 Then
        If StrComp(ValorTexto(rng.Cells(fila, 1), cParametro), nombre, vbTextCompare) = 0 Then
            CoincideParametro = True
            Exit Function
        End If
    End If
    If cNombre > 0 Then
        If StrComp(ValorTexto(rng.Cells(fila, 1), cNombre), nombre, vbTextCompare) = 0 Then
            CoincideParametro = True
            Exit Function
        End If
    End If
    If cId > 0 Then
        If StrComp(ValorTexto(rng.Cells(fila, 1), cId), nombre, vbTextCompare) = 0 Then
            CoincideParametro = True
            Exit Function
        End If
    End If
End Function

Private Function ExisteLote(ByVal idLote As String) As Boolean
    Dim lo As ListObject

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Function
    ExisteLote = Not (BuscarFila(lo, "id_lote", idLote) Is Nothing)
End Function

Private Function NumeroLotePorId(ByVal idLote As String) As String
    Dim lo As ListObject
    Dim fila As ListRow

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "id_lote", idLote)
    If fila Is Nothing Then Exit Function

    NumeroLotePorId = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "numero_lote"))
End Function

Private Function ProductoPorLote(ByVal idLote As String) As String
    Dim lo As ListObject
    Dim fila As ListRow

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "id_lote", idLote)
    If fila Is Nothing Then Exit Function

    ProductoPorLote = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "producto"))
End Function

Private Function VolumenProducto(ByVal producto As String) As Double
    Dim lo As ListObject
    Dim fila As ListRow

    Set lo = Utils.Tabla("tblProductos")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "codigo", producto)
    If fila Is Nothing Then Exit Function

    VolumenProducto = ValorNumero(fila.Range, Utils.IndiceColumna(lo, "volumen_litros"))
End Function

Private Sub MarcarPendiente(ByVal idTurnoLote As String)
    Dim lo As ListObject
    Dim fila As ListRow

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Sub

    Set fila = BuscarFila(lo, "id_turno_lote", idTurnoLote)
    If fila Is Nothing Then Exit Sub

    SetValor lo, fila, "estado_revision", ESTADO_PENDIENTE
End Sub

Private Function BuscarFila(ByVal lo As ListObject, ByVal columna As String, _
        ByVal valor As String) As ListRow

    Dim rng As Range
    Dim c As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If Len(valor) = 0 Then Exit Function

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), c), valor, vbTextCompare) = 0 Then
            Set BuscarFila = lo.ListRows(i)
            Exit Function
        End If
    Next i
End Function

Private Function BuscarFilaRechazo(ByVal lo As ListObject, ByVal idTurnoLote As String, _
        ByVal seccion As String, ByVal motivo As String) As ListRow

    Dim rng As Range
    Dim cId As Long
    Dim cSeccion As Long
    Dim cMotivo As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function

    cId = Utils.IndiceColumna(lo, "id_turno_lote")
    cSeccion = Utils.IndiceColumna(lo, "seccion")
    cMotivo = Utils.IndiceColumna(lo, "motivo")
    If cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cId), idTurnoLote, vbTextCompare) = 0 _
           And StrComp(ValorTexto(rng.Cells(i, 1), cSeccion), seccion, vbTextCompare) = 0 _
           And StrComp(ValorTexto(rng.Cells(i, 1), cMotivo), motivo, vbTextCompare) = 0 Then
            Set BuscarFilaRechazo = lo.ListRows(i)
            Exit Function
        End If
    Next i
End Function

Private Sub SetValor(ByVal lo As ListObject, ByVal fila As ListRow, _
        ByVal columna As String, ByVal valor As Variant)

    Dim c As Long

    c = Utils.IndiceColumna(lo, columna)
    If c > 0 Then fila.Range.Cells(1, c).Value = valor
End Sub

Private Function ValorTexto(ByVal rng As Range, ByVal col As Long) As String
    Dim v As Variant

    If col = 0 Then Exit Function
    v = rng.Cells(1, col).Value
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    ValorTexto = Trim$(CStr(v))
End Function

Private Function ValorVariant(ByVal rng As Range, ByVal col As Long) As Variant
    If col = 0 Then Exit Function
    ValorVariant = rng.Cells(1, col).Value
End Function

Private Function ValorFecha(ByVal rng As Range, ByVal col As Long) As Date
    Dim v As Variant

    If col = 0 Then Exit Function
    v = rng.Cells(1, col).Value
    If IsDate(v) Then ValorFecha = CDate(v)
End Function

Private Function ValorNumero(ByVal rng As Range, ByVal col As Long) As Double
    Dim v As Variant

    If col = 0 Then Exit Function
    v = rng.Cells(1, col).Value
    If IsNumeric(v) Then ValorNumero = CDbl(v)
End Function

Private Function EsActivo(ByVal v As String) As Boolean
    EsActivo = Utils.EsAnulado(v)
End Function

Private Function TextoSi() As String
    TextoSi = "S" & ChrW(237)
End Function

Private Function ValorSino(ByVal v As String) As String
    If Utils.EsAnulado(v) Then
        ValorSino = TextoSi()
    Else
        ValorSino = "No"
    End If
End Function

' ==========================================================================
' F2a - Reducciones de velocidad (tblReducciones)
' ============================================================================

Public Function CrearReduccion( _
        ByVal idTurnoLote As String, _
        ByVal ini As Date, _
        ByVal fin As Date, _
        ByVal origen As String, _
        ByVal motivo As String, _
        ByVal velocidadReal As Variant, _
        Optional ByRef mensaje As String = "") As Boolean

    Dim reporte As TDatosReporte
    Dim lo As ListObject
    Dim fila As ListRow
    Dim finAjustado As Date
    Dim idNuevo As String

    mensaje = ""

    If Not ValidarYLeerReporte(idTurnoLote, reporte, mensaje) Then Exit Function

    If Len(Trim$(origen)) = 0 Or Len(Trim$(motivo)) = 0 Then
        mensaje = "La reduccion debe tener origen y motivo."
        Exit Function
    End If

    If Not IsNumeric(velocidadReal) Then
        mensaje = "La velocidad real debe ser un numero."
        Exit Function
    End If
    If CDbl(velocidadReal) <= 0 Then
        mensaje = "La velocidad real debe ser mayor que cero."
        Exit Function
    End If

    If ini <= 0 Then
        mensaje = "La reduccion debe tener hora de inicio."
        Exit Function
    End If
    finAjustado = fin
    If finAjustado < ini Then finAjustado = DateAdd("d", 1, finAjustado)
    If finAjustado <= ini Then
        mensaje = "La hora de fin debe ser posterior a la de inicio."
        Exit Function
    End If

    Set lo = Utils.Tabla("tblReducciones")
    If lo Is Nothing Then
        mensaje = "No existe la tabla tblReducciones."
        Exit Function
    End If
    If Utils.IndiceColumna(lo, "id") = 0 Then
        mensaje = "La tabla tblReducciones no tiene columna id. Ejecutar CompletarEstructura."
        Exit Function
    End If

    idNuevo = SiguienteId(lo, "id", "R", reporte.fecha)
    Set fila = lo.ListRows.Add

    SetValor lo, fila, "id", idNuevo
    SetValor lo, fila, "fecha", reporte.fecha
    SetValor lo, fila, "turno", reporte.id_turno
    SetValor lo, fila, "id_turno_lote", idTurnoLote
    SetValor lo, fila, "origen", origen
    SetValor lo, fila, "motivo", motivo
    SetValor lo, fila, "fecha_hora_inicio", ini
    SetValor lo, fila, "fecha_hora_fin", finAjustado
    SetValor lo, fila, "velocidad_real", CDbl(velocidadReal)

    CrearReduccion = True
End Function

Private Function ItemReduccion(ByVal rng As Range, ByVal lo As ListObject) As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("id") = ValorTexto(rng, Utils.IndiceColumna(lo, "id"))
    d("fecha") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha"))
    d("turno") = ValorTexto(rng, Utils.IndiceColumna(lo, "turno"))
    d("id_turno_lote") = ValorTexto(rng, Utils.IndiceColumna(lo, "id_turno_lote"))
    d("origen") = ValorTexto(rng, Utils.IndiceColumna(lo, "origen"))
    d("motivo") = ValorTexto(rng, Utils.IndiceColumna(lo, "motivo"))
    d("fecha_hora_inicio") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha_hora_inicio"))
    d("fecha_hora_fin") = ValorFecha(rng, Utils.IndiceColumna(lo, "fecha_hora_fin"))
    d("velocidad_real") = ValorNumero(rng, Utils.IndiceColumna(lo, "velocidad_real"))
    Set ItemReduccion = d
End Function

Public Function ListarReducciones(ByVal idTurnoLote As String) As Collection
    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim i As Long
    Dim col As Collection

    Set col = New Collection
    Set ListarReducciones = col

    If Len(Trim$(idTurnoLote)) = 0 Then Exit Function

    Set lo = Utils.Tabla("tblReducciones")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(ValorTexto(rng.Cells(i, 1), cSeg), idTurnoLote, vbTextCompare) = 0 Then
            col.Add ItemReduccion(rng.Cells(i, 1), lo)
        End If
    Next i
End Function

' ==========================================================================
' F2a - Lecturas por periodo para el dashboard (F3).
' Todas excluyen los eventos de reportes anulados.
' ==========================================================================

Private Function AnuladosDelPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date) As Object

    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim cAnulado As Long
    Dim cEstado As Long
    Dim cFecha As Long
    Dim i As Long
    Dim anulados As Object
    Dim f As Date
    Dim enPeriodo As Boolean

    Set anulados = CreateObject("Scripting.Dictionary")
    anulados.CompareMode = 1
    Set AnuladosDelPeriodo = anulados

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    cAnulado = Utils.IndiceColumna(lo, "anulado")
    cEstado = Utils.IndiceColumna(lo, "estado_revision")
    cFecha = Utils.IndiceColumna(lo, "fecha")
    If cSeg = 0 Then Exit Function
    If cAnulado = 0 And cEstado = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If EsAnulado(ValorTexto(rng.Cells(i, 1), cAnulado)) _
           Or StrComp(ValorTexto(rng.Cells(i, 1), cEstado), ESTADO_ANULADO, vbTextCompare) = 0 Then
            enPeriodo = True
            If Not (fechaDesde = 0 And fechaHasta = 0) Then
                enPeriodo = False
                f = 0
                If cFecha > 0 Then f = ValorFecha(rng.Cells(i, 1), cFecha)
                If f > 0 Then
                    If (fechaDesde = 0 Or Int(f) >= Int(fechaDesde)) _
                       And (fechaHasta = 0 Or Int(f) <= Int(fechaHasta)) Then
                        enPeriodo = True
                    End If
                End If
            End If
            If enPeriodo Then
                anulados(ValorTexto(rng.Cells(i, 1), cSeg)) = True
            End If
        End If
    Next i
End Function

Private Function EnPeriodo(ByVal fecha As Date, ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date) As Boolean

    If fechaDesde = 0 And fechaHasta = 0 Then
        EnPeriodo = True
        Exit Function
    End If
    If fecha <= 0 Then Exit Function
    If fechaDesde > 0 And Int(fecha) < Int(fechaDesde) Then Exit Function
    If fechaHasta > 0 And Int(fecha) > Int(fechaHasta) Then Exit Function
    EnPeriodo = True
End Function

Private Function SegmentosDelPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date) As Object

    ' id_turno_lote -> True para los reportes NO anulados con fecha en periodo.
    ' Lo usa ListarTramosPorPeriodo: tblTramosSinEnvase no tiene columna fecha.
    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim cFecha As Long
    Dim i As Long
    Dim anulados As Object
    Dim segmentos As Object
    Dim sSeg As String

    Set segmentos = CreateObject("Scripting.Dictionary")
    segmentos.CompareMode = 1
    Set SegmentosDelPeriodo = segmentos

    Set anulados = AnuladosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    cFecha = Utils.IndiceColumna(lo, "fecha")
    If cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cSeg)
        If Len(sSeg) > 0 Then
            If Not anulados.Exists(sSeg) Then
                If EnPeriodo(ValorFecha(rng.Cells(i, 1), cFecha), fechaDesde, fechaHasta) Then
                    segmentos(sSeg) = True
                End If
            End If
        End If
    Next i
End Function

Public Function ListarReduccionesPorPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date, _
        Optional ByVal origenFiltro As String = "", _
        Optional ByVal motivoFiltro As String = "") As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cFecha As Long
    Dim cSeg As Long
    Dim i As Long
    Dim anulados As Object
    Dim col As Collection
    Dim sSeg As String
    Dim d As Object

    Set col = New Collection
    Set ListarReduccionesPorPeriodo = col

    Set anulados = AnuladosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblReducciones")
    If lo Is Nothing Then Exit Function

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cFecha = 0 Or cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cSeg)
        If Len(sSeg) > 0 Then
            If EnPeriodo(ValorFecha(rng.Cells(i, 1), cFecha), fechaDesde, fechaHasta) Then
                If Not anulados.Exists(sSeg) Then
                    Set d = ItemReduccion(rng.Cells(i, 1), lo)
                    If Len(origenFiltro) = 0 Or StrComp(CStr(d("origen")), origenFiltro, vbTextCompare) = 0 Then
                        If Len(motivoFiltro) = 0 Or StrComp(CStr(d("motivo")), motivoFiltro, vbTextCompare) = 0 Then
                            col.Add d
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function ListarProduccionPorPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date, _
        Optional ByVal productoFiltro As String = "") As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cFecha As Long
    Dim cSeg As Long
    Dim i As Long
    Dim anulados As Object
    Dim col As Collection
    Dim sSeg As String
    Dim d As Object

    Set col = New Collection
    Set ListarProduccionPorPeriodo = col

    Set anulados = AnuladosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblProduccion")
    If lo Is Nothing Then Exit Function

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cFecha = 0 Or cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cSeg)
        If Len(sSeg) > 0 Then
            If EnPeriodo(ValorFecha(rng.Cells(i, 1), cFecha), fechaDesde, fechaHasta) Then
                If Not anulados.Exists(sSeg) Then
                    Set d = CreateObject("Scripting.Dictionary")
                    d.CompareMode = 1
                    d("fecha") = ValorFecha(rng.Cells(i, 1), Utils.IndiceColumna(lo, "fecha"))
                    d("turno") = ValorTexto(rng.Cells(i, 1), Utils.IndiceColumna(lo, "turno"))
                    d("linea_id") = ValorVariant(rng.Cells(i, 1), Utils.IndiceColumna(lo, "linea_id"))
                    d("lote") = ValorTexto(rng.Cells(i, 1), Utils.IndiceColumna(lo, "lote"))
                    d("id_turno_lote") = sSeg
                    d("producto") = ValorTexto(rng.Cells(i, 1), Utils.IndiceColumna(lo, "producto"))
                    d("unidades_producidas") = ValorNumero(rng.Cells(i, 1), Utils.IndiceColumna(lo, "unidades_producidas"))
                    d("pallets") = ValorNumero(rng.Cells(i, 1), Utils.IndiceColumna(lo, "pallets"))
                    d("litros") = ValorNumero(rng.Cells(i, 1), Utils.IndiceColumna(lo, "litros"))
                    d("velocidad_nominal") = ValorNumero(rng.Cells(i, 1), Utils.IndiceColumna(lo, "velocidad_nominal"))
                    If Len(productoFiltro) = 0 Or StrComp(CStr(d("producto")), productoFiltro, vbTextCompare) = 0 Then
                        col.Add d
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function ListarParadasPorPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date, _
        Optional ByVal tipoFiltro As String = "", _
        Optional ByVal origenFiltro As String = "") As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cFecha As Long
    Dim cSeg As Long
    Dim i As Long
    Dim anulados As Object
    Dim col As Collection
    Dim sSeg As String
    Dim d As Object

    Set col = New Collection
    Set ListarParadasPorPeriodo = col

    Set anulados = AnuladosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then Exit Function

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cFecha = 0 Or cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cSeg)
        If Len(sSeg) > 0 Then
            If EnPeriodo(ValorFecha(rng.Cells(i, 1), cFecha), fechaDesde, fechaHasta) Then
                If Not anulados.Exists(sSeg) Then
                    Set d = ItemParada(rng.Cells(i, 1), lo)
                    If Len(tipoFiltro) = 0 Or StrComp(CStr(d("tipo")), tipoFiltro, vbTextCompare) = 0 Then
                        If Len(origenFiltro) = 0 Or StrComp(CStr(d("origen")), origenFiltro, vbTextCompare) = 0 Then
                            col.Add d
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function ListarRechazosPorPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date, _
        Optional ByVal seccionFiltro As String = "") As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cFecha As Long
    Dim cSeg As Long
    Dim i As Long
    Dim anulados As Object
    Dim col As Collection
    Dim sSeg As String
    Dim d As Object

    Set col = New Collection
    Set ListarRechazosPorPeriodo = col

    Set anulados = AnuladosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblRechazos")
    If lo Is Nothing Then Exit Function

    cFecha = Utils.IndiceColumna(lo, "fecha")
    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    If cFecha = 0 Or cSeg = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cSeg)
        If Len(sSeg) > 0 Then
            If EnPeriodo(ValorFecha(rng.Cells(i, 1), cFecha), fechaDesde, fechaHasta) Then
                If Not anulados.Exists(sSeg) Then
                    Set d = ItemRechazo(rng.Cells(i, 1), lo)
                    If Len(seccionFiltro) = 0 Or StrComp(CStr(d("seccion")), seccionFiltro, vbTextCompare) = 0 Then
                        col.Add d
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function ListarTramosPorPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date) As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cSeg As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim i As Long
    Dim col As Collection
    Dim segmentos As Object
    Dim item As Object
    Dim sSeg As String

    Set col = New Collection
    Set ListarTramosPorPeriodo = col

    Set segmentos = SegmentosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblTramosSinEnvase")
    If lo Is Nothing Then Exit Function

    cSeg = Utils.IndiceColumna(lo, "id_turno_lote")
    cIni = Utils.IndiceColumna(lo, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(lo, "fecha_hora_fin")
    If cSeg = 0 Or cIni = 0 Or cFin = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cSeg)
        If Len(sSeg) > 0 Then
            If segmentos.Exists(sSeg) Then
                Set item = CreateObject("Scripting.Dictionary")
                item.CompareMode = 1
                item("id_turno_lote") = sSeg
                item("fecha_hora_inicio") = ValorFecha(rng.Cells(i, 1), cIni)
                item("fecha_hora_fin") = ValorFecha(rng.Cells(i, 1), cFin)
                col.Add item
            End If
        End If
    Next i
End Function

Public Function ListarReportesPorPeriodo(ByVal fechaDesde As Date, _
        ByVal fechaHasta As Date, _
        Optional ByVal productoFiltro As String = "") As Collection

    Dim lo As ListObject
    Dim rng As Range
    Dim cId As Long
    Dim cFecha As Long
    Dim i As Long
    Dim anulados As Object
    Dim col As Collection
    Dim sSeg As String
    Dim reporte As TDatosReporte

    Set col = New Collection
    Set ListarReportesPorPeriodo = col

    Set anulados = AnuladosDelPeriodo(fechaDesde, fechaHasta)

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If lo Is Nothing Then Exit Function

    cId = Utils.IndiceColumna(lo, "id_turno_lote")
    cFecha = Utils.IndiceColumna(lo, "fecha")
    If cId = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        sSeg = ValorTexto(rng.Cells(i, 1), cId)
        If Len(sSeg) > 0 Then
            If Not anulados.Exists(sSeg) Then
                If EnPeriodo(ValorFecha(rng.Cells(i, 1), cFecha), fechaDesde, fechaHasta) Then
                    reporte = LeerReporte(sSeg)
                    If reporte.Encontrado Then
                        If Len(productoFiltro) = 0 Or StrComp(reporte.producto, productoFiltro, vbTextCompare) = 0 Then
                            col.Add ItemReporte(reporte)
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' ==========================================================================
' F2a - Lecturas de catalogos para el dashboard (F3)
' ==========================================================================

Public Function LeerLote(ByVal idLote As String) As Object
    Dim lo As ListObject
    Dim fila As ListRow
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set LeerLote = d

    If Len(Trim$(idLote)) = 0 Then Exit Function

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Function

    Set fila = BuscarFila(lo, "id_lote", idLote)
    If fila Is Nothing Then Exit Function

    d("id_lote") = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "id_lote"))
    d("numero_lote") = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "numero_lote"))
    d("producto") = ValorTexto(fila.Range, Utils.IndiceColumna(lo, "producto"))
    d("fecha_hora_inicio") = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha_hora_inicio"))
    d("fecha_hora_fin") = ValorFecha(fila.Range, Utils.IndiceColumna(lo, "fecha_hora_fin"))
    d("Encontrado") = True
End Function

Public Function ListarOrigenes() As Collection
    Dim lo As ListObject
    Dim rng As Range
    Dim cCodigo As Long
    Dim cNombre As Long
    Dim cClase As Long
    Dim cActivo As Long
    Dim i As Long
    Dim col As Collection
    Dim item As Object

    Set col = New Collection
    Set ListarOrigenes = col

    Set lo = Utils.Tabla("tblOrigenes")
    If lo Is Nothing Then Exit Function

    cCodigo = Utils.IndiceColumna(lo, "codigo")
    cNombre = Utils.IndiceColumna(lo, "nombre")
    cClase = Utils.IndiceColumna(lo, "clase")
    cActivo = Utils.IndiceColumna(lo, "activo")
    If cCodigo = 0 Or cNombre = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If cActivo = 0 Or EsActivo(ValorTexto(rng.Cells(i, 1), cActivo)) Then
            Set item = CreateObject("Scripting.Dictionary")
            item.CompareMode = 1
            item("codigo") = ValorTexto(rng.Cells(i, 1), cCodigo)
            item("nombre") = ValorTexto(rng.Cells(i, 1), cNombre)
            item("clase") = ValorTexto(rng.Cells(i, 1), cClase)
            If Len(Trim$(CStr(item("codigo")))) > 0 Then col.Add item
        End If
    Next i
End Function
