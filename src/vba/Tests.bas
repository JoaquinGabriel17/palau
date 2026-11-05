Attribute VB_Name = "Tests"
Option Explicit

' Pruebas del motor OEE según Doc v0.8 (2026-11-04).
' Los cinco casos oficiales se cargan con DatosPrueba y se borran al final.
' Tolerancia de comparación: P-08 (0,05 puntos porcentuales).

Private Type TResultado
    Caso As String
    Metrica As String
    Esperado As Double
    Calculado As Double
    Diferencia As Double
    EsPorcentaje As Boolean
    Pasa As Boolean
End Type

Private mResultados() As TResultado
Private mCantidad As Long

Public Sub EjecutarPruebas()
    Dim ok As Boolean

    ok = RunAllTests()
    EscribirReporte

    If ok Then
        MsgBox "Todas las pruebas pasaron. Ver 90_Tests y el Depurador.", vbInformation, "Pruebas OEE"
    Else
        MsgBox "Hay pruebas con FALLA. Ver 90_Tests y el Depurador.", vbExclamation, "Pruebas OEE"
    End If
End Sub

Public Sub ReportePruebas()
    Dim ok As Boolean

    ok = RunAllTests()
    EscribirReporte
    Debug.Print "Reporte escrito en 90_Tests. Resultado global: " & IIf(ok, "PASA", "FALLA")
End Sub

Public Sub ReporteCaso498()
    ReportePruebas
End Sub

Public Function RunAllTests() As Boolean
    Dim tolerancia As Double
    Dim ok As Boolean

    mCantidad = 0
    Erase mResultados
    tolerancia = ToleranciaComparacionPP()

    On Error GoTo Fallo

    Debug.Print "=== Pruebas OEE - Doc v0.8 ==="
    Debug.Print "Tolerancia P-08: " & Format$(tolerancia, "0.00") & " pp"

    PrepararDatosPrueba

    ok = True
    ok = Test_Caso498(tolerancia) And ok
    ok = Test_Caso2(tolerancia) And ok
    ok = Test_Caso3(tolerancia) And ok
    ok = Test_Caso4(tolerancia) And ok
    ok = Test_Caso5(tolerancia) And ok
    ok = Test_RN15(tolerancia) And ok

    ok = Test_CeroUnidades() And ok
    ok = Test_CienRechazo() And ok
    ok = Test_TurnoSinParadas() And ok
    ok = Test_VigenciaParametros() And ok
    ok = Test_ParadaSolapada() And ok
    ok = Test_ParadaReglas() And ok
    ok = Test_Produccion() And ok
    ok = Test_Reporte() And ok
    ok = Test_LoteInvalido() And ok
    ok = Test_RechazoAvisaGuarda() And ok
    ok = Test_AgregarValorNuevo() And ok
    ok = Test_ObtenerParametro() And ok
    ok = Test_AnulacionReporte() And ok
    ok = Test_ParadaCruzadaTurno() And ok
    ok = Test_AccesoHojas() And ok

    LimpiarDatosPrueba

    Debug.Print "Resultados registrados: " & mCantidad
    RunAllTests = ok
    Exit Function

Fallo:
    Debug.Print "ERROR en RunAllTests: " & Err.Description
    On Error Resume Next
    LimpiarDatosPrueba
    RunAllTests = False
End Function

Private Function Test_Caso498(ByVal tolerancia As Double) As Boolean
    Dim detalle As TDetalleTurno
    Dim ok As Boolean
    Dim caso As String

    caso = "Caso 498 (D-01)"
    detalle = DetalleTurnoPorId("TST-498")

    If Not detalle.Encontrado Then
        RegistrarFallo caso, "Turno no encontrado"
        Exit Function
    End If

    ok = True
    ok = ValidarAbsoluto(caso, "TO (min)", detalle.TO, 330, 0.05) And ok
    ok = ValidarPorcentaje(caso, "Disponibilidad", Numero(detalle.disponibilidad), 0.815, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Rendimiento", Numero(detalle.rendimiento), 0.753, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Calidad", Numero(detalle.calidad), 0.9998, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "OEE", detalle.oee, 0.614, tolerancia) And ok

    Test_Caso498 = ok
End Function

Private Function Test_Caso2(ByVal tolerancia As Double) As Boolean
    Dim detalle As TDetalleTurno
    Dim ok As Boolean
    Dim caso As String

    caso = "Caso 2 (cambio de formato A y C)"
    detalle = DetalleTurnoPorId("TST-C2")

    If Not detalle.Encontrado Then
        RegistrarFallo caso, "Turno no encontrado"
        Exit Function
    End If

    ok = True
    ok = ValidarAbsoluto(caso, "TO (min)", detalle.TO, 382, 0.05) And ok
    ok = ValidarPorcentaje(caso, "Disponibilidad", Numero(detalle.disponibilidad), 0.858, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Rendimiento", Numero(detalle.rendimiento), 0.858, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Calidad", Numero(detalle.calidad), 0.9981, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "OEE", detalle.oee, 0.735, tolerancia) And ok

    Test_Caso2 = ok
End Function

Private Function Test_Caso3(ByVal tolerancia As Double) As Boolean
    Dim detalle As TDetalleTurno
    Dim ok As Boolean
    Dim caso As String

    caso = "Caso 3 (cruza medianoche)"
    detalle = DetalleTurnoPorId("TST-C3")

    If Not detalle.Encontrado Then
        RegistrarFallo caso, "Turno no encontrado"
        Exit Function
    End If

    ok = True
    ok = ValidarAbsoluto(caso, "TO (min)", detalle.TO, 353, 0.05) And ok
    ok = ValidarPorcentaje(caso, "Disponibilidad", Numero(detalle.disponibilidad), 0.861, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Rendimiento", Numero(detalle.rendimiento), 0.875, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Calidad", Numero(detalle.calidad), 0.9986, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "OEE", detalle.oee, 0.753, tolerancia) And ok

    Test_Caso3 = ok
End Function

Private Function Test_Caso4(ByVal tolerancia As Double) As Boolean
    Dim detalle As TDetalleTurno
    Dim ok As Boolean
    Dim caso As String

    caso = "Caso 4 (datos incompletos, 60/min)"
    detalle = DetalleTurnoPorId("TST-C4")

    If Not detalle.Encontrado Then
        RegistrarFallo caso, "Turno no encontrado"
        Exit Function
    End If

    ok = True
    ok = ValidarAbsoluto(caso, "TO (min)", detalle.TO, 395, 0.05) And ok
    ok = ValidarPorcentaje(caso, "Disponibilidad", Numero(detalle.disponibilidad), 0.888, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Rendimiento", Numero(detalle.rendimiento), 0.932, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Calidad", Numero(detalle.calidad), 0.9996, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "OEE", detalle.oee, 0.827, tolerancia) And ok

    Test_Caso4 = ok
End Function

Private Function Test_Caso5(ByVal tolerancia As Double) As Boolean
    Dim detalle As TDetalleTurno
    Dim ok As Boolean
    Dim caso As String

    caso = "Caso 5 (tramos con/sin envase)"
    detalle = DetalleTurnoPorId("TST-C5")

    If Not detalle.Encontrado Then
        RegistrarFallo caso, "Turno no encontrado"
        Exit Function
    End If

    ok = True
    ok = ValidarAbsoluto(caso, "TO (min)", detalle.TO, 415, 0.05) And ok
    ok = ValidarAbsoluto(caso, "Produccion teorica", detalle.produccionTeorica, 27925, 0.5) And ok
    ok = ValidarPorcentaje(caso, "Disponibilidad", Numero(detalle.disponibilidad), 0.933, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Rendimiento", Numero(detalle.rendimiento), 0.819, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Calidad", Numero(detalle.calidad), 0.9974, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "OEE", detalle.oee, 0.761, tolerancia) And ok

    Test_Caso5 = ok
End Function

Private Function Test_RN15(ByVal tolerancia As Double) As Boolean
    Dim detalle As TDetalleTurno
    Dim ok As Boolean
    Dim caso As String

    caso = "RN-15 (parada cruza tramo)"
    detalle = DetalleTurnoPorId("TST-RN15")

    If Not detalle.Encontrado Then
        RegistrarFallo caso, "Turno no encontrado"
        Exit Function
    End If

    ' 06:00-08:00 con envase: 120 min - 30 min de parada = 90 min * 90/min = 8.100.
    ' 08:00-10:00 sin envase: 120 min - 30 min de parada = 90 min * 2.500/60 = 3.750.
    ' Produccion teorica esperada: 11.850. Si no se reparte, daria 13.300 o 10.400.
    ok = True
    ok = ValidarAbsoluto(caso, "Produccion teorica", detalle.produccionTeorica, 11850, 0.5) And ok
    ok = ValidarAbsoluto(caso, "TO (min)", detalle.TO, 180, 0.05) And ok
    ok = ValidarPorcentaje(caso, "Disponibilidad", Numero(detalle.disponibilidad), 0.75, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Rendimiento", Numero(detalle.rendimiento), 1, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "Calidad", Numero(detalle.calidad), 1, tolerancia) And ok
    ok = ValidarPorcentaje(caso, "OEE", detalle.oee, 0.75, tolerancia) And ok

    Test_RN15 = ok
End Function

Public Function Test_CeroUnidades() As Boolean
    Dim oee As Double
    Dim calidad As Variant
    Dim rendimiento As Variant
    Dim ok As Boolean

    oee = CalcularOEE_DesdeTotales(480, 0, 0, 0, 0, 0, 0)
    calidad = CalcularCalidad(0, 0)
    rendimiento = CalcularRendimientoOEE(0, 70, 480)

    ok = Cerca(oee, 0) And IsEmpty(calidad) And Cerca(CDbl(rendimiento), 0)
    Debug.Print "Test_CeroUnidades: " & IIf(ok, "PASA", "FALLA") & _
                " (OEE=" & Format$(oee, "0.0000") & ")"
    Test_CeroUnidades = ok
End Function

Public Function Test_CienRechazo() As Boolean
    Dim oee As Double
    Dim calidad As Variant
    Dim ok As Boolean

    oee = CalcularOEE_DesdeTotales(480, 0, 0, 0, 0, 100, 100 / 70)
    calidad = CalcularCalidad(0, 100)

    ok = Cerca(CDbl(calidad), 0) And Cerca(oee, 0)
    Debug.Print "Test_CienRechazo: " & IIf(ok, "PASA", "FALLA") & _
                " (Calidad=" & Format$(CDbl(calidad), "0.00%") & _
                ", OEE=" & Format$(oee, "0.0000") & ")"
    Test_CienRechazo = ok
End Function

Public Function Test_TurnoSinParadas() As Boolean
    Dim tpl As Double
    Dim tOper As Double
    Dim disponibilidad As Variant
    Dim rendimiento As Variant
    Dim calidad As Variant
    Dim oee As Double
    Dim ok As Boolean

    tpl = 480
    tOper = CalcularTO(tpl, 0)
    disponibilidad = CalcularDisponibilidad(tOper, tpl)
    rendimiento = CalcularRendimientoOEE(1000, 70, tOper)
    calidad = CalcularCalidad(1000, 1000)
    oee = CalcularOEE(CDbl(disponibilidad), CDbl(rendimiento), CDbl(calidad))

    ok = Cerca(CDbl(disponibilidad), 1) _
        And Cerca(CDbl(calidad), 1) _
        And Cerca(CDbl(rendimiento), (1000 / 70) / tOper) _
        And Cerca(oee, CDbl(disponibilidad) * CDbl(rendimiento) * CDbl(calidad))

    Debug.Print "Test_TurnoSinParadas: " & IIf(ok, "PASA", "FALLA") & _
                " (Disponibilidad=" & Format$(CDbl(disponibilidad), "0.00%") & _
                ", Rendimiento=" & Format$(CDbl(rendimiento), "0.00%") & _
                ", Calidad=" & Format$(CDbl(calidad), "0.00%") & ")"
    Test_TurnoSinParadas = ok
End Function

Public Function Test_VigenciaParametros() As Boolean
    Dim lo As ListObject
    Dim v2025 As Double
    Dim v2026 As Double
    Dim oee2025 As Double
    Dim oee2026 As Double
    Dim ok As Boolean

    On Error GoTo Fallo

    Set lo = Utils.Tabla("tblVelocidades")
    If lo Is Nothing Then
        Debug.Print "Test_VigenciaParametros: FALLA (falta tblVelocidades)"
        Exit Function
    End If

    LimpiarVelocidadTest lo, "TST-VIG"
    AgregarVelocidadTest lo, "TST-VIG-A", 10, DateSerial(2025, 1, 1), DateSerial(2026, 1, 1)
    AgregarVelocidadTest lo, "TST-VIG-B", 20, DateSerial(2026, 1, 2), 0

    v2025 = VelocidadVigente("TST-VIG", DateSerial(2025, 12, 31))
    v2026 = VelocidadVigente("TST-VIG", DateSerial(2026, 1, 2))

    If v2025 > 0 And v2026 > 0 Then
        oee2025 = CalcularOEE_DesdeTotales(480, 0, 0, 0, 1000, 0, 1000 / v2025)
        oee2026 = CalcularOEE_DesdeTotales(480, 0, 0, 0, 1000, 0, 1000 / v2026)
        ok = Cerca(v2025, 10) And Cerca(v2026, 20) And (oee2025 > oee2026)
    Else
        ok = False
    End If

    Debug.Print "Test_VigenciaParametros: " & IIf(ok, "PASA", "FALLA") & _
                " (2025: vel=" & Format$(v2025, "0.00") & _
                ", OEE=" & Format$(oee2025, "0.0000") & _
                "; 2026: vel=" & Format$(v2026, "0.00") & _
                ", OEE=" & Format$(oee2026, "0.0000") & ")"
    Test_VigenciaParametros = ok

    LimpiarVelocidadTest lo, "TST-VIG"
    Exit Function

Fallo:
    Debug.Print "Test_VigenciaParametros: FALLA (" & Err.Description & ")"
    On Error Resume Next
    LimpiarVelocidadTest lo, "TST-VIG"
    Test_VigenciaParametros = False
End Function

Public Function Test_ParadaSolapada() As Boolean
    Dim f As Date
    Dim otras(1 To 1, 1 To 2) As Date
    Dim res As TResultadoValidacion
    Dim ok As Boolean

    f = DateSerial(2026, 11, 4)
    otras(1, 1) = f + TimeSerial(10, 0, 0)
    otras(1, 2) = f + TimeSerial(11, 0, 0)

    res = ValidarParada(f + TimeSerial(10, 30, 0), f + TimeSerial(11, 30, 0), _
                        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                        tipo:="TP-03", motivo:="prueba", otrasParadas:=otras)
    ok = ValidarLogico("F2.3 parada", "solapamiento bloquea", res.Resultado = ResultadoValidacion.BLOQUEA)

    res = ValidarParada(f + TimeSerial(12, 0, 0), f + TimeSerial(12, 10, 0), _
                        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                        tipo:="TP-03", motivo:="prueba", otrasParadas:=otras)
    ok = ValidarLogico("F2.3 parada", "sin solapamiento pasa", res.Resultado = ResultadoValidacion.OK) And ok

    Test_ParadaSolapada = ok
End Function

Public Function Test_ParadaReglas() As Boolean
    Dim f As Date
    Dim res As TResultadoValidacion
    Dim ok As Boolean
    Dim otras(1 To 1, 1 To 2) As Date

    f = DateSerial(2026, 11, 4)

    res = ValidarParada(Empty, f + TimeSerial(7, 0, 0), _
                        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                        tipo:="TP-03", motivo:="prueba")
    ok = ValidarLogico("F2.3 parada", "sin inicio bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA)

    res = ValidarParada(f + TimeSerial(8, 0, 0), f + TimeSerial(8, 0, 0), _
                        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                        tipo:="TP-03", motivo:="prueba")
    ok = ValidarLogico("F2.3 parada", "fin igual a inicio bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    res = ValidarParada(f + TimeSerial(5, 0, 0), f + TimeSerial(5, 30, 0), _
                        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                        tipo:="TP-03", motivo:="prueba")
    ok = ValidarLogico("F2.3 parada", "fuera del turno bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    otras(1, 1) = f + TimeSerial(6, 0, 0)
    otras(1, 2) = f + TimeSerial(10, 0, 0)
    res = ValidarParada(f + TimeSerial(11, 0, 0), f + TimeSerial(12, 0, 0), _
                        f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                        tipo:="TP-03", motivo:="prueba", otrasParadas:=otras, tdMinutos:=60)
    ok = ValidarLogico("F2.3 parada", "suma mayor al TD bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    Test_ParadaReglas = ok
End Function

Public Function Test_Produccion() As Boolean
    Dim res As TResultadoValidacion
    Dim ok As Boolean

    res = ValidarProduccion(-1)
    ok = ValidarLogico("F2.3 produccion", "negativa bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA)

    res = ValidarProduccion(1.5)
    ok = ValidarLogico("F2.3 produccion", "no entera bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    res = ValidarProduccion(0)
    ok = ValidarLogico("F2.3 produccion", "cero avisa", _
                       res.Resultado = ResultadoValidacion.AVISA) And ok

    res = ValidarProduccion(100)
    ok = ValidarLogico("F2.3 produccion", "positiva pasa", _
                       res.Resultado = ResultadoValidacion.OK) And ok

    Test_Produccion = ok
End Function

Public Function Test_Reporte() As Boolean
    Dim res As TResultadoValidacion
    Dim ok As Boolean
    Dim nros As Variant

    nros = Array("123", "456")

    res = ValidarReporte("123", nros)
    ok = ValidarLogico("F2.3 reporte", "numero duplicado bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA)

    res = ValidarReporte("789", nros, "TP-02", True)
    ok = ValidarLogico("F2.3 reporte", "TP-02 bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    res = ValidarReporte("789", nros, "TP-03", False)
    ok = ValidarLogico("F2.3 reporte", "tipo inactivo bloquea", _
                       res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    res = ValidarReporte("789", nros, "TP-03", True)
    ok = ValidarLogico("F2.3 reporte", "reporte valido pasa", _
                       res.Resultado = ResultadoValidacion.OK) And ok

    Test_Reporte = ok
End Function

Public Function Test_LoteInvalido() As Boolean
    Dim res As TResultadoValidacion
    Dim ok As Boolean

    res = ValidarLote("11261")
    ok = ValidarLogico("F2.3 lote", "11261 valido", res.Resultado = ResultadoValidacion.OK)

    res = ValidarLote("13261")
    ok = ValidarLogico("F2.3 lote", "mes 13 bloquea", res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    res = ValidarLote("1126")
    ok = ValidarLogico("F2.3 lote", "cuatro digitos bloquea", res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    res = ValidarLote("11A61")
    ok = ValidarLogico("F2.3 lote", "letra bloquea", res.Resultado = ResultadoValidacion.BLOQUEA) And ok

    Test_LoteInvalido = ok
End Function

Public Function Test_RechazoAvisaGuarda() As Boolean
    Dim res As TResultadoValidacion
    Dim ok As Boolean
    Dim mensaje As String
    Dim rechazo As TDatosRechazo
    Dim reporte As TDatosReporte
    Dim guardado As Boolean

    res = ValidarRechazo("", "", 5)
    ok = ValidarLogico("F2.3 rechazo", "sin motivo avisa", res.Resultado = ResultadoValidacion.AVISA)

    guardado = CrearRechazo("TST-498-TL", "", "", 5, "prueba F2.3", mensaje)
    ok = ValidarLogico("F2.3 rechazo", "se guarda pese al aviso", guardado) And ok

    rechazo = LeerRechazo("TST-498-TL", "", "")
    ok = ValidarLogico("F2.3 rechazo", "rechazo persistido", _
                       rechazo.Encontrado And rechazo.cantidad = 5) And ok

    reporte = LeerReporte("TST-498-TL")
    ok = ValidarLogico("F2.3 rechazo", "reporte pendiente", _
                       reporte.Encontrado And _
                       StrComp(reporte.estado_revision, "Pendiente de revisar", vbTextCompare) = 0) And ok

    Test_RechazoAvisaGuarda = ok
End Function

Public Function Test_AgregarValorNuevo() As Boolean
    Dim codigo As String
    Dim codigo2 As String
    Dim ok As Boolean
    Dim valor As String

    valor = "TST puesto F2.3"
    codigo = AgregarValorNuevo("tblPuestos", valor, "puesto")
    ok = ValidarLogico("F2.3 catalogo", "codigo generado", Len(codigo) > 0)
    ok = ValidarLogico("F2.3 catalogo", "prefijo PU", Left$(codigo, 3) = "PU-") And ok
    ok = ValidarLogico("F2.3 catalogo", "valor agregado", _
                       CatalogoContiene("tblPuestos", "puesto", valor)) And ok

    codigo2 = AgregarValorNuevo("tblPuestos", valor, "puesto")
    ok = ValidarLogico("F2.3 catalogo", "sin duplicar", _
                       StrComp(codigo, codigo2, vbTextCompare) = 0) And ok

    BorrarCatalogoTest "tblPuestos", "codigo", codigo
    Test_AgregarValorNuevo = ok
End Function

Public Function Test_ObtenerParametro() As Boolean
    Dim vAntes As Variant
    Dim vDespues As Variant
    Dim vHoy As Variant
    Dim ok As Boolean

    vAntes = ObtenerParametro("UMBRAL_MICROPARADA_MIN", DateSerial(2026, 10, 28))
    vDespues = ObtenerParametro("UMBRAL_MICROPARADA_MIN", DateSerial(2026, 10, 29))
    vHoy = ObtenerParametro("UMBRAL_MICROPARADA_MIN")

    ok = ValidarLogico("F2.3 parametro", "antes de vigencia vacio", IsEmpty(vAntes))
    ok = ValidarLogico("F2.3 parametro", "vigente devuelve 5", _
                       IsNumeric(vDespues) And CDbl(vDespues) = 5) And ok
    ok = ValidarLogico("F2.3 parametro", "hoy devuelve 5", _
                       IsNumeric(vHoy) And CDbl(vHoy) = 5) And ok

    Test_ObtenerParametro = ok
End Function

Public Function Test_AnulacionReporte() As Boolean
    Dim f As Date
    Dim ok As Boolean
    Dim idTurnoLote As String
    Dim idTurno As String
    Dim idParada As String
    Dim mensaje As String
    Dim oeeAntes As Double
    Dim oeeDespues As Double
    Dim oeeDia As Double

    f = DateSerial(2026, 11, 6)
    LimpiarReportePrueba "TST-RPT-ANUL", "TST-ANUL-L"

    AgregarLoteTest "TST-ANUL-L", "11271", "500SG", _
                    f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0)

    ok = ValidarLogico("F2.3 anulacion", "crear reporte", _
        CrearReporte(f, "RES-M", "TST-ANUL-L", f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                     idTurnoLote, nroReporte:="TST-RPT-ANUL", equipo:="RES", _
                     encargado:="prueba", mensaje:=mensaje))

    ok = ValidarLogico("F2.3 anulacion", "crear produccion", _
        CrearProduccion(idTurnoLote, 5000, 0, mensaje:=mensaje)) And ok
    ok = ValidarLogico("F2.3 anulacion", "crear parada", _
        CrearParada(idTurnoLote, f + TimeSerial(9, 0, 0), f + TimeSerial(9, 30, 0), _
                    "OR-02", "prueba", "TP-03", idParada, mensaje:=mensaje)) And ok

    idTurno = LeerReporte(idTurnoLote).id_turno
    oeeAntes = CalcularOEE_Turno(idTurno)
    ok = ValidarLogico("F2.3 anulacion", "OEE antes mayor a cero", oeeAntes > 0) And ok

    ok = ValidarLogico("F2.3 anulacion", "anular reporte", AnularReporte(idTurnoLote, mensaje)) And ok

    oeeDespues = CalcularOEE_Turno(idTurno)
    oeeDia = CalcularOEE_Dia(f)
    ok = ValidarLogico("F2.3 anulacion", "OEE turno anulado es cero", Cerca(oeeDespues, 0)) And ok
    ok = ValidarLogico("F2.3 anulacion", "OEE dia ignora anulado", Cerca(oeeDia, 0)) And ok

    LimpiarReportePrueba "TST-RPT-ANUL", "TST-ANUL-L"
    Test_AnulacionReporte = ok
End Function

Public Function Test_ParadaCruzadaTurno() As Boolean
    Dim f As Date
    Dim ok As Boolean
    Dim idTurnoLoteA As String
    Dim idTurnoLoteB As String
    Dim idTurnoA As String
    Dim idTurnoB As String
    Dim idParada As String
    Dim mensaje As String

    f = DateSerial(2026, 11, 7)
    LimpiarReportePrueba "TST-RPT-CRUZ-A", "TST-CRUZ-LA"
    LimpiarReportePrueba "TST-RPT-CRUZ-B", "TST-CRUZ-LB"

    AgregarLoteTest "TST-CRUZ-LA", "11272", "500SG", _
                    f + TimeSerial(6, 0, 0), f + TimeSerial(10, 0, 0)
    AgregarLoteTest "TST-CRUZ-LB", "11273", "500SG", _
                    f + TimeSerial(10, 0, 0), f + TimeSerial(14, 0, 0)

    ok = ValidarLogico("F2.3 turno cruzado", "crear reporte A", _
        CrearReporte(f, "RES-M", "TST-CRUZ-LA", f + TimeSerial(6, 0, 0), f + TimeSerial(10, 0, 0), _
                     idTurnoLoteA, nroReporte:="TST-RPT-CRUZ-A", equipo:="RES", _
                     encargado:="prueba", mensaje:=mensaje))

    ok = ValidarLogico("F2.3 turno cruzado", "crear reporte B", _
        CrearReporte(f, "RES-M", "TST-CRUZ-LB", f + TimeSerial(10, 0, 0), f + TimeSerial(14, 0, 0), _
                     idTurnoLoteB, nroReporte:="TST-RPT-CRUZ-B", equipo:="RES", _
                     encargado:="prueba", mensaje:=mensaje)) And ok

    idTurnoA = LeerReporte(idTurnoLoteA).id_turno
    idTurnoB = LeerReporte(idTurnoLoteB).id_turno
    ok = ValidarLogico("F2.3 turno cruzado", "mismo id_turno", _
                       StrComp(idTurnoA, idTurnoB, vbTextCompare) = 0) And ok
    ok = ValidarLogico("F2.3 turno cruzado", "segmentos distintos", _
                       StrComp(idTurnoLoteA, idTurnoLoteB, vbTextCompare) <> 0) And ok

    ok = ValidarLogico("F2.3 turno cruzado", "parada en reporte A", _
        CrearParada(idTurnoLoteA, f + TimeSerial(9, 50, 0), f + TimeSerial(10, 20, 0), _
                    "OR-02", "prueba", "TP-03", idParada, mensaje:=mensaje)) And ok

    ok = ValidarLogico("F2.3 turno cruzado", "parada en B solapada con A bloquea", _
        Not CrearParada(idTurnoLoteB, f + TimeSerial(10, 5, 0), f + TimeSerial(10, 25, 0), _
                        "OR-02", "prueba", "TP-03", idParada, mensaje:=mensaje)) And ok

    ok = ValidarLogico("F2.3 turno cruzado", "parada en B sin solape pasa", _
        CrearParada(idTurnoLoteB, f + TimeSerial(12, 0, 0), f + TimeSerial(12, 10, 0), _
                    "OR-02", "prueba", "TP-03", idParada, mensaje:=mensaje)) And ok

    LimpiarReportePrueba "TST-RPT-CRUZ-A", "TST-CRUZ-LA"
    LimpiarReportePrueba "TST-RPT-CRUZ-B", "TST-CRUZ-LB"
    Test_ParadaCruzadaTurno = ok
End Function

Public Function Test_AccesoHojas() As Boolean
    Dim f As Date
    Dim ok As Boolean
    Dim idTurnoLote As String
    Dim idParada As String
    Dim idDup As String
    Dim mensaje As String
    Dim reporte As TDatosReporte
    Dim parada As TDatosParada
    Dim produccion As TDatosProduccion
    Dim rechazo As TDatosRechazo

    f = DateSerial(2026, 11, 5)
    LimpiarReportePruebaF23

    AgregarLoteTest "TST-F23-L", "11261", "500SG", _
                    f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0)

    ok = ValidarLogico("F2.3 acceso", "crear reporte", _
        CrearReporte(f, "RES-M", "TST-F23-L", f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                     idTurnoLote, nroReporte:="TST-RPT-F23", equipo:="RES", _
                     encargado:="prueba", mensaje:=mensaje))

    reporte = LeerReporte(idTurnoLote)
    ok = ValidarLogico("F2.3 acceso", "reporte leido", _
        reporte.Encontrado And reporte.id_lote = "TST-F23-L" _
        And reporte.producto = "500SG" _
        And StrComp(reporte.anulado, "No", vbTextCompare) = 0) And ok

    ok = ValidarLogico("F2.3 acceso", "numero duplicado bloquea", _
        Not CrearReporte(f, "RES-M", "TST-F23-L", f + TimeSerial(6, 0, 0), f + TimeSerial(14, 0, 0), _
                         idDup, nroReporte:="TST-RPT-F23", equipo:="RES", mensaje:=mensaje)) And ok

    ok = ValidarLogico("F2.3 acceso", "crear parada rellena claves", _
        CrearParada(idTurnoLote, f + TimeSerial(9, 0, 0), f + TimeSerial(9, 10, 0), _
                    "OR-02", "prueba", "TP-03", idParada, mensaje:=mensaje)) And ok

    parada = LeerParada(idParada)
    ok = ValidarLogico("F2.3 acceso", "parada claves correctas", _
        parada.Encontrado And parada.id_lote = "TST-F23-L" _
        And StrComp(parada.turno, reporte.id_turno, vbTextCompare) = 0 _
        And Int(parada.fecha) = Int(f)) And ok

    ok = ValidarLogico("F2.3 acceso", "crear produccion rellena claves", _
        CrearProduccion(idTurnoLote, 1000, 2, mensaje:=mensaje)) And ok

    produccion = LeerProduccion(idTurnoLote)
    ok = ValidarLogico("F2.3 acceso", "produccion claves correctas", _
        produccion.Encontrado And produccion.lote = "TST-F23-L" _
        And produccion.producto = "500SG" _
        And produccion.unidades_producidas = 1000 _
        And Abs(produccion.litros - 500) < 0.001) And ok

    ok = ValidarLogico("F2.3 acceso", "crear rechazo rellena claves", _
        CrearRechazo(idTurnoLote, "SE-01", "MR-05", 5, "prueba", mensaje:=mensaje)) And ok

    rechazo = LeerRechazo(idTurnoLote, "SE-01", "MR-05")
    ok = ValidarLogico("F2.3 acceso", "rechazo claves correctas", _
        rechazo.Encontrado And rechazo.lote = "TST-F23-L" _
        And StrComp(rechazo.turno, reporte.id_turno, vbTextCompare) = 0) And ok

    ok = ValidarLogico("F2.3 acceso", "anular reporte", AnularReporte(idTurnoLote, mensaje)) And ok
    ok = ValidarLogico("F2.3 acceso", "reporte anulado", ReporteAnulado(idTurnoLote)) And ok
    ok = ValidarLogico("F2.3 acceso", "anulado no admite paradas", _
        Not CrearParada(idTurnoLote, f + TimeSerial(10, 0, 0), f + TimeSerial(10, 10, 0), _
                        "OR-02", "prueba", "TP-03", idParada, mensaje:=mensaje)) And ok

    LimpiarReportePruebaF23
    Test_AccesoHojas = ok
End Function

Private Sub AgregarLoteTest(ByVal idLote As String, ByVal numeroLote As String, _
                            ByVal producto As String, ByVal ini As Date, ByVal fin As Date)
    Dim lo As ListObject
    Dim fila As ListRow

    Set lo = Utils.Tabla("tblLotes")
    If lo Is Nothing Then Exit Sub

    Set fila = lo.ListRows.Add
    SetValorTest fila.Range, lo, "id_lote", idLote
    SetValorTest fila.Range, lo, "numero_lote", numeroLote
    SetValorTest fila.Range, lo, "producto", producto
    SetValorTest fila.Range, lo, "fecha_hora_inicio", ini
    SetValorTest fila.Range, lo, "fecha_hora_fin", fin
End Sub

Private Sub LimpiarReportePruebaF23()
    LimpiarReportePrueba "TST-RPT-F23", "TST-F23-L"
End Sub

Private Sub LimpiarReportePrueba(ByVal nroReporte As String, ByVal idLote As String)
    Dim lo As ListObject
    Dim rng As Range
    Dim cNro As Long
    Dim cId As Long
    Dim i As Long
    Dim idTurnoLote As String

    Set lo = Utils.Tabla("tblTurnosRegistro")
    If Not lo Is Nothing Then
        cNro = Utils.IndiceColumna(lo, "nro_reporte")
        cId = Utils.IndiceColumna(lo, "id_turno_lote")
        Set rng = lo.DataBodyRange
        If Not rng Is Nothing And cNro > 0 And cId > 0 Then
            For i = rng.Rows.Count To 1 Step -1
                If StrComp(Trim$(CStr(rng.Cells(i, cNro).Value)), nroReporte, vbTextCompare) = 0 Then
                    idTurnoLote = CStr(rng.Cells(i, cId).Value)
                    BorrarFilasPorValorTest "tblRechazos", "id_turno_lote", idTurnoLote
                    BorrarFilasPorValorTest "tblProduccion", "id_turno_lote", idTurnoLote
                    BorrarFilasPorValorTest "tblParadas", "id_turno_lote", idTurnoLote
                    lo.ListRows(i).Delete
                End If
            Next i
        End If
    End If

    BorrarFilasPorValorTest "tblLotes", "id_lote", idLote
End Sub

Private Sub BorrarFilasPorValorTest(ByVal nombreTabla As String, ByVal columna As String, _
                                    ByVal valor As String)
    Dim lo As ListObject
    Dim rng As Range
    Dim c As Long
    Dim i As Long

    Set lo = Utils.Tabla(nombreTabla)
    If lo Is Nothing Then Exit Sub

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = rng.Rows.Count To 1 Step -1
        If StrComp(Trim$(CStr(rng.Cells(i, c).Value)), valor, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
        End If
    Next i
End Sub

Private Function ValidarLogico(ByVal caso As String, ByVal metrica As String, _
                               ByVal pasa As Boolean) As Boolean
    Registrar caso, metrica, IIf(pasa, 1, 0), IIf(pasa, 1, 0), 0, False, pasa
    ValidarLogico = pasa
End Function

Private Function CatalogoContiene(ByVal nombreTabla As String, ByVal columna As String, _
                                  ByVal valor As String) As Boolean
    Dim lo As ListObject
    Dim rng As Range
    Dim c As Long
    Dim i As Long

    Set lo = Utils.Tabla(nombreTabla)
    If lo Is Nothing Then Exit Function

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Function

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Function

    For i = 1 To rng.Rows.Count
        If StrComp(Trim$(CStr(rng.Cells(i, c).Value)), valor, vbTextCompare) = 0 Then
            CatalogoContiene = True
            Exit Function
        End If
    Next i
End Function

Private Sub BorrarCatalogoTest(ByVal nombreTabla As String, ByVal columna As String, _
                               ByVal valor As String)
    Dim lo As ListObject
    Dim rng As Range
    Dim c As Long
    Dim i As Long

    Set lo = Utils.Tabla(nombreTabla)
    If lo Is Nothing Then Exit Sub

    c = Utils.IndiceColumna(lo, columna)
    If c = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = rng.Rows.Count To 1 Step -1
        If StrComp(Trim$(CStr(rng.Cells(i, c).Value)), valor, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
        End If
    Next i
End Sub

Private Function ValidarPorcentaje(ByVal caso As String, ByVal metrica As String, _
                                   ByVal calculado As Double, ByVal esperado As Double, _
                                   ByVal toleranciaPP As Double) As Boolean
    Dim dif As Double

    dif = (calculado - esperado) * 100
    ValidarPorcentaje = (Abs(dif) <= toleranciaPP)
    Registrar caso, metrica, esperado, calculado, dif, True, ValidarPorcentaje
End Function

Private Function ValidarAbsoluto(ByVal caso As String, ByVal metrica As String, _
                                 ByVal calculado As Double, ByVal esperado As Double, _
                                 ByVal tolerancia As Double) As Boolean
    Dim dif As Double

    dif = calculado - esperado
    ValidarAbsoluto = (Abs(dif) <= tolerancia)
    Registrar caso, metrica, esperado, calculado, dif, False, ValidarAbsoluto
End Function

Private Sub Registrar(ByVal caso As String, ByVal metrica As String, _
                      ByVal esperado As Double, ByVal calculado As Double, _
                      ByVal diferencia As Double, ByVal esPorcentaje As Boolean, _
                      ByVal pasa As Boolean)
    If mCantidad = 0 Then
        ReDim mResultados(0 To 0)
    Else
        ReDim Preserve mResultados(0 To mCantidad)
    End If

    With mResultados(mCantidad)
        .Caso = caso
        .Metrica = metrica
        .Esperado = esperado
        .Calculado = calculado
        .Diferencia = diferencia
        .EsPorcentaje = esPorcentaje
        .Pasa = pasa
    End With
    mCantidad = mCantidad + 1

    Debug.Print caso & " / " & metrica & ": " & IIf(pasa, "PASA", "FALLA") & _
                " (esperado " & Format$(esperado, "0.0000") & _
                ", calculado " & Format$(calculado, "0.0000") & _
                ", dif " & Format$(diferencia, "0.0000") & _
                IIf(esPorcentaje, " pp", "") & ")"
End Sub

Private Sub RegistrarFallo(ByVal caso As String, ByVal metrica As String)
    Registrar caso, metrica, 0, 0, 0, False, False
    Debug.Print caso & " / " & metrica & ": FALLA"
End Sub

Private Function Numero(ByVal v As Variant) As Double
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then Numero = CDbl(v)
End Function

Private Function Cerca(ByVal a As Double, ByVal b As Double) As Boolean
    Cerca = (Abs(a - b) <= 0.0005)
End Function

Private Function ToleranciaComparacionPP() As Double
    Dim v As Variant

    ToleranciaComparacionPP = 0.05
    v = ObtenerParametro("TOLERANCIA_COMPARACION_PP")
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then ToleranciaComparacionPP = CDbl(v)
    End If
End Function

Private Function ResultadoGlobal() As Boolean
    Dim i As Long

    ResultadoGlobal = True
    For i = 0 To mCantidad - 1
        If Not mResultados(i).Pasa Then
            ResultadoGlobal = False
            Exit Function
        End If
    Next i
End Function

Private Sub EscribirReporte()
    Dim ws As Worksheet
    Dim fila As Long
    Dim filaNota As Long
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("90_Tests")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Range("A1:J200").Clear

    ws.Cells(1, 1).Value = "Pruebas OEE - Doc v0.8 (2026-11-04)"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(2, 1).Value = "Tolerancia P-08: " & Format$(ToleranciaComparacionPP(), "0.00") & " pp"
    ws.Cells(3, 1).Value = "Resultado global: " & IIf(ResultadoGlobal(), "PASA", "FALLA")
    ws.Cells(3, 1).Font.Bold = True

    ws.Cells(5, 1).Value = "Caso"
    ws.Cells(5, 2).Value = "Metrica"
    ws.Cells(5, 3).Value = "Esperado"
    ws.Cells(5, 4).Value = "Calculado"
    ws.Cells(5, 5).Value = "Diferencia"
    ws.Cells(5, 6).Value = "Unidad"
    ws.Cells(5, 7).Value = "Resultado"
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 7)).Font.Bold = True
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 7)).Interior.Color = RGB(221, 235, 247)

    fila = 6
    For i = 0 To mCantidad - 1
        With mResultados(i)
            ws.Cells(fila, 1).Value = .Caso
            ws.Cells(fila, 2).Value = .Metrica
            ws.Cells(fila, 3).Value = .Esperado
            ws.Cells(fila, 4).Value = .Calculado
            ws.Cells(fila, 5).Value = .Diferencia
            ws.Cells(fila, 6).Value = IIf(.EsPorcentaje, "pp", "abs")

            If .EsPorcentaje Then
                ws.Cells(fila, 3).NumberFormat = "0.00%"
                ws.Cells(fila, 4).NumberFormat = "0.00%"
            Else
                ws.Cells(fila, 3).NumberFormat = "0.00"
                ws.Cells(fila, 4).NumberFormat = "0.00"
            End If
            ws.Cells(fila, 5).NumberFormat = "+0.00;-0.00;0.00"

            ws.Cells(fila, 7).Value = IIf(.Pasa, "PASA", "FALLA")
            If .Pasa Then
                ws.Cells(fila, 7).Interior.Color = RGB(198, 239, 206)
            Else
                ws.Cells(fila, 7).Interior.Color = RGB(255, 199, 206)
            End If
        End With
        fila = fila + 1
    Next i

    ws.Cells(fila + 1, 1).Value = "Unidad: pp = puntos porcentuales; abs = minutos o unidades segun la metrica."
    ws.Cells(fila + 1, 1).Font.Italic = True

    filaNota = fila + 3
    ws.Cells(filaNota, 1).Value = "RN-15 (parada que cruza tramos)"
    ws.Cells(filaNota, 1).Font.Bold = True
    ws.Cells(filaNota + 1, 1).Value = _
        "El turno se divide en tramos segun tblTramosSinEnvase; lo que no esta en un tramo es con envase (RN-14)."
    ws.Cells(filaNota + 2, 1).Value = _
        "Si una parada cruza el limite entre un tramo con envase y uno sin envase, se reparte por hora: " & _
        "a cada tramo solo se le descuentan los minutos de la parada que caen dentro de el (interseccion inicio/fin). " & _
        "No se asigna entera al tramo donde empezo."
    ws.Cells(filaNota + 3, 1).Value = _
        "Ejemplo TST-RN15: parada 07:30-08:30 y tramo sin envase 08:00-10:00 = 30 min con envase y 30 min sin envase; " & _
        "produccion teorica = 90 min * 90/min + 90 min * 2500/60 = 11.850."
    ws.Cells(filaNota + 1, 1).Font.Italic = True
    ws.Cells(filaNota + 2, 1).Font.Italic = True
    ws.Cells(filaNota + 3, 1).Font.Italic = True

    ws.Columns("A:G").AutoFit
End Sub

Private Sub AgregarVelocidadTest(ByVal lo As ListObject, ByVal codigo As String, _
                                 ByVal velocidad As Double, ByVal desde As Date, ByVal hasta As Date)
    Dim fila As ListRow
    Dim rng As Range

    Set fila = lo.ListRows.Add
    Set rng = fila.Range

    SetValorTest rng, lo, "codigo", codigo
    SetValorTest rng, lo, "producto", "TST-VIG"
    SetValorTest rng, lo, "velocidad", velocidad
    SetValorTest rng, lo, "vigencia_desde", desde
    If hasta > 0 Then
        SetValorTest rng, lo, "vigencia_hasta", hasta
    Else
        SetValorTest rng, lo, "vigencia_hasta", Empty
    End If
    SetValorTest rng, lo, "origen_dato", "prueba"
    SetValorTest rng, lo, "activo", "Si"
End Sub

Private Sub SetValorTest(ByVal rng As Range, ByVal lo As ListObject, ByVal columna As String, ByVal valor As Variant)
    Dim c As Long

    c = Utils.IndiceColumna(lo, columna)
    If c > 0 Then rng.Cells(1, c).Value = valor
End Sub



Private Sub LimpiarVelocidadTest(ByVal lo As ListObject, ByVal producto As String)
    Dim rng As Range
    Dim cProducto As Long
    Dim i As Long

    If lo Is Nothing Then Exit Sub
    cProducto = Utils.IndiceColumna(lo, "producto")
    If cProducto = 0 Then Exit Sub

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then Exit Sub

    For i = rng.Rows.Count To 1 Step -1
        If StrComp(CStr(rng.Cells(i, cProducto).Value), producto, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
        End If
    Next i
End Sub
