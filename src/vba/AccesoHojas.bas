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

    idTurnoLote = idNuevo
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

    idParada = idNuevo
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

    valor = Trim$(valor)
    If Len(valor) = 0 Then Exit Function

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
    AplicarColumnasExtra lo, fila, columna, columnasExtra

    AgregarValorNuevo = codigo
End Function

Private Sub AplicarColumnasExtra(ByVal lo As ListObject, ByVal fila As ListRow, _
        ByVal columnaPrincipal As String, ByVal columnasExtra As Variant)

    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim nCols As Long
    Dim col As String
    Dim k As Variant

    If IsEmpty(columnasExtra) Or IsNull(columnasExtra) Then Exit Sub

    If IsObject(columnasExtra) Then
        If TypeName(columnasExtra) = "Dictionary" Then
            For Each k In columnasExtra.Keys
                col = Trim$(CStr(k))
                If EsColumnaExtraValida(col, columnaPrincipal) Then
                    If Utils.IndiceColumna(lo, col) > 0 Then
                        SetValor lo, fila, col, columnasExtra(k)
                    End If
                End If
            Next k
        End If
        Exit Sub
    End If

    If IsArray(columnasExtra) Then
        On Error GoTo Fin
        lb = LBound(columnasExtra, 1)
        ub = UBound(columnasExtra, 1)
        nCols = UBound(columnasExtra, 2) - LBound(columnasExtra, 2) + 1
        If nCols < 2 Then Exit Sub

        For i = lb To ub
            col = Trim$(CStr(columnasExtra(i, LBound(columnasExtra, 2))))
            If EsColumnaExtraValida(col, columnaPrincipal) Then
                If Utils.IndiceColumna(lo, col) > 0 Then
                    SetValor lo, fila, col, columnasExtra(i, LBound(columnasExtra, 2) + 1)
                End If
            End If
        Next i
    End If
    Exit Sub

Fin:
    On Error GoTo 0
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
    Dim rng As Range
    Dim cTurno As Long
    Dim cId As Long
    Dim cIni As Long
    Dim cFin As Long
    Dim i As Long
    Dim sTurno As String
    Dim sId As String
    Dim dIni As Date
    Dim dFin As Date
    Dim arr() As Variant
    Dim total As Long

    suma = 0

    Set lo = Utils.Tabla("tblParadas")
    If lo Is Nothing Then Exit Function

    cTurno = Utils.IndiceColumna(lo, "turno")
    cId = Utils.IndiceColumna(lo, "id")
    cIni = Utils.IndiceColumna(lo, "fecha_hora_inicio")
    cFin = Utils.IndiceColumna(lo, "fecha_hora_fin")
    If cTurno = 0 Or cIni = 0 Or cFin = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    ReDim arr(1 To rng.Rows.Count, 1 To 3)
    total = rng.Rows.Count

    For i = 1 To rng.Rows.Count
        sTurno = ValorTexto(rng.Cells(i, 1), cTurno)
        If StrComp(sTurno, idTurno, vbTextCompare) = 0 Then
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
