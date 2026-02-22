Attribute VB_Name = "mod_Main"
Option Explicit

Public Sub EjecutarOfuscador()
    Dim rutaOriginal As String, rutaDestino As String, nombreFich As String
    Dim wb As Workbook, vbc As Object, dict As Object
    Dim i As Long, linea As String, lineaOfus As String
    Dim tMod As Long, tLin As Long, tBot As Long
    Dim claveUnica As String
    
    ' 1. SELECCIÓN Y VALIDACIÓN
    rutaOriginal = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm")
    If rutaOriginal = "Falso" Then Exit Sub
    
    ' Generar clave aleatoria única para esta sesión de ofuscación
    claveUnica = GenerarClaveAleatoria(16)
    
    rutaDestino = Left(rutaOriginal, InStrRev(rutaOriginal, ".") - 1) & "_OFUS.xlsm"
    nombreFich = Mid(rutaDestino, InStrRev(rutaDestino, "\") + 1)
    
    ' Control de errores de archivo
    On Error Resume Next
    If Dir(rutaDestino) <> "" Then Kill rutaDestino
    If Err.Number <> 0 Then: MsgBox "Cierra el archivo destino antes de continuar.", vbCritical: Exit Sub
    FileCopy rutaOriginal, rutaDestino
    If Err.Number <> 0 Then: MsgBox "Error al copiar archivo.", vbCritical: Exit Sub
    On Error GoTo 0

    ' 2. PROCESO DE OFUSCACIÓN
    Set wb = Workbooks.Open(rutaDestino)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Carga nombres de macros y módulos (mod_Collector debe estar presente)
    Call LlenarDiccionario(wb, dict)

    For Each vbc In wb.VBProject.VBComponents
        ' Renombrar Módulos
        If dict.Exists(vbc.Name) Then
            vbc.Name = dict(vbc.Name)
            tMod = tMod + 1
        End If

        ' Ofuscar Código Interior
        If vbc.CodeModule.CountOfLines > 0 Then
            Dim contenido As String: contenido = ""
            For i = 1 To vbc.CodeModule.CountOfLines
                linea = vbc.CodeModule.Lines(i, 1)
                ' Pasamos la clave única al Tokenizer
                lineaOfus = ProcesarLineaMaestra(linea, dict, claveUnica)
                
                If Trim(lineaOfus) <> "" Then
                    contenido = contenido & lineaOfus & vbCrLf
                    tLin = tLin + 1
                End If
            Next i
            vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
            vbc.CodeModule.AddFromString contenido
        End If
    Next vbc

    ' 3. SINCRONIZACIÓN Y AYUDANTE
    ' SincronizarBotones debe estar en tu proyecto para arreglar OnAction
    tBot = SincronizarBotones(wb, dict, nombreFich)
    Call InyectarModuloTraductor(wb, claveUnica)
    
    wb.Save
    MsgBox "--- OFUSCACIÓN COMPLETADA ---" & vbCrLf & _
           "Módulos: " & tMod & vbCrLf & _
           "Líneas: " & tLin & vbCrLf & _
           "Botones: " & tBot & vbCrLf & _
           "Clave: " & claveUnica, vbInformation
End Sub

Private Sub InyectarModuloTraductor(ByRef wb As Workbook, ByVal clave As String)
    Dim modT As Object
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("mod_Internal_Helper")
    On Error GoTo 0
    
    Set modT = wb.VBProject.VBComponents.Add(1)
    modT.Name = "mod_Internal_Helper"
    
    ' IMPORTANTE: Sincronización de índices (i Mod Len) + 1 para base 1 de Mid
    modT.CodeModule.AddFromString _
        "Public Function f_tr(ByVal s As String) As String" & vbCrLf & _
        "    If s = """" Then Exit Function" & vbCrLf & _
        "    Dim v, i, r, cl, k: cl = """ & clave & """: v = Split(s, "","")" & vbCrLf & _
        "    For i = 0 To UBound(v)" & vbCrLf & _
        "        k = (i Mod Len(cl)) + 1" & vbCrLf & _
        "        r = r & Chr(CInt(v(i)) Xor Asc(Mid(cl, k, 1)))" & vbCrLf & _
        "    Next" & vbCrLf & _
        "    f_tr = r" & vbCrLf & _
        "End Function"
End Sub

Private Function GenerarClaveAleatoria(ByVal n As Integer) As String
    Dim i As Integer, c As String, r As String
    c = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$"
    Randomize
    For i = 1 To n: r = r & Mid(c, Int(Len(c) * Rnd + 1), 1): Next
    GenerarClaveAleatoria = r
End Function

' ============================================================================
' SINCRONIZAR BOTONES
' Re-vincula las macros de los botones al nuevo nombre ofuscado
' ============================================================================
Public Function SincronizarBotones(ByRef wb As Workbook, ByRef dict As Object, ByVal nombreFich As String) As Long
    Dim ws As Worksheet
    Dim shp As Shape
    Dim macroActual As String
    Dim nombreMacro As String
    Dim contador As Long
    
    contador = 0
    
    For Each ws In wb.Worksheets
        For Each shp In ws.Shapes
            On Error Resume Next
            macroActual = shp.OnAction
            
            If macroActual <> "" Then
                ' 1. Extraer el nombre de la macro (después del '!')
                If InStr(macroActual, "!") > 0 Then
                    nombreMacro = Mid(macroActual, InStrRev(macroActual, "!") + 1)
                Else
                    nombreMacro = macroActual
                End If
                
                ' 2. Si el nombre original está en el diccionario, lo cambiamos
                If dict.Exists(nombreMacro) Then
                    ' El formato debe ser 'NombreArchivo.xlsm'!NuevaMacro
                    shp.OnAction = "'" & nombreFich & "'!" & dict(nombreMacro)
                    contador = contador + 1
                End If
            End If
            On Error GoTo 0
        Next shp
    Next ws
    
    SincronizarBotones = contador
End Function
