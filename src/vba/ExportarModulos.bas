Attribute VB_Name = "ExportarModulos"
Option Explicit

Public Sub ExportarModulos()
    Dim fso As Object
    Dim ruta As String
    Dim componentes As Object
    Dim comp As Object
    Dim ext As String
    Dim exportados As Long
    Dim fallos As String
    Dim mensaje As String
    Dim rutaArchivo As String

    If ThisWorkbook.Path = "" Then
        MsgBox "Guarde el libro antes de exportar.", vbExclamation, "Exportar modulos"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    ruta = fso.GetAbsolutePathName(ThisWorkbook.Path & "\..\src\vba")
    AsegurarCarpeta fso, ruta

    On Error GoTo SinAcceso
    Set componentes = ThisWorkbook.VBProject.VBComponents
    On Error GoTo 0

    BorrarArchivosExportados fso, ruta

    For Each comp In componentes
        Select Case comp.Type
            Case 1: ext = ".bas"
            Case 2: ext = ".cls"
            Case 3: ext = ".frm"
            Case Else: ext = ""
        End Select
        If ext <> "" Then
            rutaArchivo = fso.BuildPath(ruta, comp.Name & ext)
            mensaje = ""
            If ExportarComponente(comp, rutaArchivo, mensaje) Then
                exportados = exportados + 1
            Else
                fallos = fallos & vbCrLf & comp.Name & ": " & mensaje
            End If
        End If
    Next comp

    MsgBox "Exportados " & exportados & " componentes en:" & vbCrLf & ruta _
         & IIf(fallos = "", "", vbCrLf & vbCrLf & "Fallos:" & fallos), _
         vbInformation, "Exportar modulos"
    Exit Sub

SinAcceso:
    MsgBox "No se pudo acceder al proyecto VBA." & vbCrLf & _
           "Habilite: Opciones > Centro de confianza > Configuracion de macros > " & _
           "Confiar en el acceso al modelo de objetos de proyectos de VBA.", _
           vbCritical, "Exportar modulos"
End Sub

Private Function ExportarComponente(ByVal comp As Object, ByVal rutaArchivo As String, ByRef mensajeError As String) As Boolean
    On Error GoTo Manejador
    comp.Export rutaArchivo
    ExportarComponente = True
    Exit Function
Manejador:
    mensajeError = Err.Description
    ExportarComponente = False
End Function

Private Sub AsegurarCarpeta(ByVal fso As Object, ByVal ruta As String)
    If fso.FolderExists(ruta) Then Exit Sub
    AsegurarCarpeta fso, fso.GetParentFolderName(ruta)
    fso.CreateFolder ruta
End Sub

Private Sub BorrarArchivosExportados(ByVal fso As Object, ByVal ruta As String)
    Dim archivo As Object
    Dim pendientes As Collection
    Dim ext As String
    Dim i As Long

    Set pendientes = New Collection
    For Each archivo In fso.GetFolder(ruta).Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))
        Select Case ext
            Case "bas", "cls", "frm", "frx"
                pendientes.Add archivo.Name
        End Select
    Next archivo

    For i = 1 To pendientes.Count
        fso.DeleteFile fso.BuildPath(ruta, pendientes(i)), True
    Next i
End Sub
