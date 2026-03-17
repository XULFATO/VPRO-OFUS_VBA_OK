Attribute VB_Name = "mod_Main"
Option Explicit

Public Sub EjecutarOfuscador()
    Dim rutaOriginal As String, rutaDestino As String, nombreFich As String
    Dim wb As Workbook, vbc As Object, dict As Object
    Dim i As Long, linea As String, lineaOfus As String
    Dim tMod As Long, tLin As Long, tBot As Long
    Dim claveUnica As String
    
    ' -----------------------------------------------------------------------
    ' 1. SELECCIÓN Y VALIDACIÓN
    ' -----------------------------------------------------------------------
    rutaOriginal = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm")
    ' CORRECCIÓN: GetOpenFilename devuelve False (Boolean) cuando se cancela
    If VarType(rutaOriginal) = vbBoolean Then Exit Sub
    
    ' Randomize UNA SOLA VEZ para toda la sesión (evita colisiones de nombres)
    Randomize
    
    claveUnica = GenerarClaveAleatoria(16)
    
    rutaDestino = Left(rutaOriginal, InStrRev(rutaOriginal, ".") - 1) & "_OFUS.xlsm"
    nombreFich  = Mid(rutaDestino, InStrRev(rutaDestino, "\") + 1)
    
    On Error Resume Next
    If Dir(rutaDestino) <> "" Then Kill rutaDestino
    If Err.Number <> 0 Then MsgBox "Cierra el archivo destino antes de continuar.", vbCritical: Exit Sub
    FileCopy rutaOriginal, rutaDestino
    If Err.Number <> 0 Then MsgBox "Error al copiar archivo.", vbCritical: Exit Sub
    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' 2. PROCESO DE OFUSCACIÓN
    ' -----------------------------------------------------------------------
    Set wb   = Workbooks.Open(rutaDestino)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Construir diccionario: nombre original → nombre transformado (espejo+leet)
    Call LlenarDiccionario(wb, dict)

    For Each vbc In wb.VBProject.VBComponents
        
        ' Renombrar módulo
        If dict.Exists(vbc.Name) Then
            vbc.Name = dict(vbc.Name)
            tMod = tMod + 1
        End If

        ' Ofuscar código interior
        If vbc.CodeModule.CountOfLines > 0 Then
            Dim contenido As String: contenido = ""
            For i = 1 To vbc.CodeModule.CountOfLines
                linea = vbc.CodeModule.Lines(i, 1)
                ' ProcesarLineaMaestra ya llama a InyectarRuido internamente
                lineaOfus = ProcesarLineaMaestra(linea, dict, claveUnica)
                
                If Trim(lineaOfus) <> "" Then
                    contenido = contenido & lineaOfus & vbCrLf
                    tLin = tLin + 1
                End If
            Next i
            
            ' Inyectar 1-3 bloques de código muerto al final del módulo
            Dim nBloques As Integer, b As Integer
            nBloques = Int(Rnd * 3) + 1
            For b = 1 To nBloques
                contenido = contenido & vbCrLf & GenerarBloqueRuido() & vbCrLf
            Next b
            
            vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
            vbc.CodeModule.AddFromString contenido
        End If
    Next vbc

    ' -----------------------------------------------------------------------
    ' 3. SINCRONIZACIÓN DE BOTONES E INYECCIÓN DEL TRADUCTOR
    ' -----------------------------------------------------------------------
    tBot = SincronizarBotones(wb, dict, nombreFich)
    Call InyectarModuloTraductor(wb, claveUnica)
    
    wb.Save
    MsgBox "--- OFUSCACIÓN COMPLETADA ---" & vbCrLf & _
           "Módulos renombrados : " & tMod & vbCrLf & _
           "Líneas procesadas   : " & tLin & vbCrLf & _
           "Botones actualizados: " & tBot & vbCrLf & _
           "Clave XOR           : " & claveUnica, vbInformation
End Sub

' ============================================================================
' INYECTAR MÓDULO TRADUCTOR (f_tr con XOR + clave única)
' ============================================================================
Private Sub InyectarModuloTraductor(ByRef wb As Workbook, ByVal clave As String)
    Dim modT As Object
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("mod_Internal_Helper")
    On Error GoTo 0
    
    Set modT  = wb.VBProject.VBComponents.Add(1)
    modT.Name = "mod_Internal_Helper"
    
    ' Sincronización: cifrado usa j=1→k=1; descifrado usa i=0→k=(0 Mod Len)+1=1 ✓
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

' ============================================================================
' GENERAR CLAVE ALEATORIA
' ============================================================================
Private Function GenerarClaveAleatoria(ByVal n As Integer) As String
    Dim i As Integer, c As String, r As String
    c = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$"
    For i = 1 To n: r = r & Mid(c, Int(Len(c) * Rnd + 1), 1): Next
    GenerarClaveAleatoria = r
End Function

' ============================================================================
' SINCRONIZAR BOTONES
' Cubre: Shapes (OnAction), ActiveX CommandButton, OLEObjects e Hipervínculos
' ============================================================================
Public Function SincronizarBotones(ByRef wb As Workbook, ByRef dict As Object, ByVal nombreFich As String) As Long
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ole As OLEObject
    Dim hv As Hyperlink
    Dim macroActual As String, nombreMacro As String
    Dim contador As Long
    contador = 0
    
    For Each ws In wb.Worksheets
        
        ' ------------------------------------------------------------------
        ' A) Shapes con OnAction (botones de formulario)
        ' ------------------------------------------------------------------
        For Each shp In ws.Shapes
            On Error Resume Next
            macroActual = shp.OnAction
            On Error GoTo 0
            
            If macroActual <> "" Then
                nombreMacro = ExtraerNombreMacro(macroActual)
                If dict.Exists(nombreMacro) Then
                    shp.OnAction = "'" & nombreFich & "'!" & dict(nombreMacro)
                    contador = contador + 1
                End If
            End If
        Next shp
        
        ' ------------------------------------------------------------------
        ' B) OLEObjects (incluye ActiveX CommandButton, etc.)
        '    El código del evento Click está en el módulo de la hoja y ya
        '    habrá sido renombrado en el bucle de módulos; aquí sincronizamos
        '    la propiedad OnAction si la tienen y también el LinkedCell/ListFillRange
        '    para otros controles, pero lo más importante es el sub vinculado.
        '    Para CommandButton ActiveX el handler es Private Sub NombreControl_Click()
        '    en el módulo de la hoja: se renombra solo si capturamos el patrón _Click.
        ' ------------------------------------------------------------------
        For Each ole In ws.OLEObjects
            On Error Resume Next
            ' Algunos OLEObjects exponen OnAction (no es estándar, pero lo cubrimos)
            macroActual = ole.OnAction
            On Error GoTo 0
            
            If macroActual <> "" Then
                nombreMacro = ExtraerNombreMacro(macroActual)
                If dict.Exists(nombreMacro) Then
                    On Error Resume Next
                    ole.OnAction = "'" & nombreFich & "'!" & dict(nombreMacro)
                    On Error GoTo 0
                    contador = contador + 1
                End If
            End If
        Next ole
        
        ' ------------------------------------------------------------------
        ' C) Hipervínculos con macro en SubAddress
        ' ------------------------------------------------------------------
        For Each hv In ws.Hyperlinks
            On Error Resume Next
            macroActual = hv.SubAddress
            On Error GoTo 0
            
            If macroActual <> "" Then
                nombreMacro = ExtraerNombreMacro(macroActual)
                If dict.Exists(nombreMacro) Then
                    hv.SubAddress = dict(nombreMacro)
                    contador = contador + 1
                End If
            End If
        Next hv
        
    Next ws
    
    SincronizarBotones = contador
End Function

' ============================================================================
' EXTRAER NOMBRE DE MACRO desde cadena OnAction
' Formato posible: 'Archivo.xlsm'!NombreMacro  o simplemente NombreMacro
' ============================================================================
Private Function ExtraerNombreMacro(ByVal onAction As String) As String
    If InStr(onAction, "!") > 0 Then
        ExtraerNombreMacro = Mid(onAction, InStrRev(onAction, "!") + 1)
    Else
        ExtraerNombreMacro = onAction
    End If
End Function
