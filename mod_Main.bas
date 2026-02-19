Attribute VB_Name = "mod_Main"

Option Explicit

' -----------------------------------------------------------------------
' EjecutarOfuscador
' Punto de entrada principal. Asigna este Sub a tu boton.
' -----------------------------------------------------------------------
Public Sub EjecutarOfuscador()
    Dim rutaOriginal As String
    Dim rutaDestino As String
    Dim wb As Workbook
    Dim dict As Object
    Dim excl As String

    ' 1. Seleccionar archivo origen
    rutaOriginal = SeleccionarArchivo()
    If rutaOriginal = "" Then
        MsgBox "Operacion cancelada.", vbInformation
        Exit Sub
    End If

    ' 2. Derivar ruta destino
    rutaDestino = Left(rutaOriginal, InStrRev(rutaOriginal, ".") - 1) & "_OFUS.xlsm"

    ' 3. Crear copia — NUNCA trabajamos sobre el original
    On Error Resume Next
    If Dir(rutaDestino) <> "" Then Kill rutaDestino
    FileCopy rutaOriginal, rutaDestino
    If Err.Number <> 0 Then
        MsgBox "Error al copiar el archivo." & vbCrLf & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 4. Abrir la copia
    Set wb = Workbooks.Open(rutaDestino)
    If wb Is Nothing Then
        MsgBox "No se pudo abrir la copia del archivo.", vbCritical
        Exit Sub
    End If

    ' 5. Verificar acceso al modelo VBA
    On Error Resume Next
    Dim testAccess As Long
    testAccess = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        MsgBox "Sin acceso al modelo VBA." & vbCrLf & _
               "Ve a Opciones de Excel > Centro de Confianza > " & _
               "Configuracion de macros > Confiar en el acceso al " & _
               "modelo de objetos de proyectos de VBA.", vbCritical
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0

    ' 6. Inicializar — semilla UNICA para toda la sesion
    Randomize Timer

    ' 7. Crear diccionario con CompareMode texto
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    excl = mod_Collector.ObtenerExclusiones()

    ' 8. Fase 1: Recolectar nombres a renombrar
    mod_Collector.LlenarDiccionario wb, dict

    If dict.Count = 0 Then
        MsgBox "No se encontraron identificadores para ofuscar.", vbInformation
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' 9. Fase 1.5: Renombrar modulos en el VBProject
    RenombrarModulos wb, dict, excl

    ' 10. Fase 2: Procesar cada componente
    Dim vbc As Object
    Dim i As Long
    Dim linea As String
    Dim nuevoCodigo As String
    Dim esContinuacion As Boolean

    For Each vbc In wb.VBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Or vbc.Type = 100 Then
            If vbc.CodeModule.CountOfLines > 0 Then
                nuevoCodigo = ""
                esContinuacion = False

                For i = 1 To vbc.CodeModule.CountOfLines
                    linea = vbc.CodeModule.Lines(i, 1)

                    Dim lineaProcesada As String

                    ' Si la linea anterior era continuacion, no tocar esta
                    If esContinuacion Then
                        lineaProcesada = linea
                    Else
                        lineaProcesada = mod_Tokenizer.ProcesarLineaMaestra(linea, dict)
                    End If

                    nuevoCodigo = nuevoCodigo & lineaProcesada & vbCrLf

                    ' Detectar si esta linea termina en continuacion
                    esContinuacion = (Right(RTrim(linea), 1) = "_")
                Next i

                ' Escribir codigo transformado
                vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
                vbc.CodeModule.AddFromString nuevoCodigo
            End If
        End If
    Next vbc

    ' 10. Actualizar OnAction de todos los botones
    ActualizarBotones wb, dict

    ' 11. Guardar y cerrar
    wb.Save
    wb.Close SaveChanges:=False

    MsgBox "Ofuscacion completada." & vbCrLf & vbCrLf & _
           "Identificadores renombrados: " & dict.Count & vbCrLf & _
           "Archivo guardado en: " & vbCrLf & rutaDestino, _
           vbInformation, "Ofuscador"
End Sub

' -----------------------------------------------------------------------
' InspectarArchivo
' Muestra todos los Subs publicos y botones del archivo sin modificarlo.
' Util para conocer el contenido antes de ofuscar.
' -----------------------------------------------------------------------
Public Sub InspectarArchivo()
    Dim rutaOriginal As String
    Dim wb As Workbook
    Dim informe As String

    rutaOriginal = SeleccionarArchivo()
    If rutaOriginal = "" Then Exit Sub

    Set wb = Workbooks.Open(rutaOriginal, ReadOnly:=True)
    If wb Is Nothing Then
        MsgBox "No se pudo abrir el archivo.", vbCritical
        Exit Sub
    End If

    informe = "=== INSPECCION: " & wb.Name & " ===" & vbCrLf & vbCrLf

    ' Botones con macro asignada
    informe = informe & "--- BOTONES Y MACROS ---" & vbCrLf
    Dim ws As Object
    Dim shp As Object
    Dim hayBotones As Boolean
    hayBotones = False

    For Each ws In wb.Worksheets
        For Each shp In ws.Shapes
            Dim mac As String
            mac = ""
            On Error Resume Next
            mac = shp.OnAction
            On Error GoTo 0
            If Len(mac) > 0 Then
                informe = informe & "Hoja: " & ws.Name & _
                          "  |  Boton: " & shp.Name & _
                          "  |  Macro: " & mac & vbCrLf
                hayBotones = True
            End If
        Next shp
    Next ws
    If Not hayBotones Then informe = informe & "(Ninguno encontrado)" & vbCrLf

    ' Subs y Functions publicas
    informe = informe & vbCrLf & "--- SUBS Y FUNCTIONS PUBLICAS ---" & vbCrLf
    Dim vbc As Object
    On Error Resume Next
    For Each vbc In wb.VBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Then
            Dim j As Long
            Dim ln As String
            For j = 1 To vbc.CodeModule.CountOfLines
                ln = Trim(vbc.CodeModule.Lines(j, 1))
                If ln Like "Public Sub *" Or ln Like "Public Function *" Then
                    informe = informe & "[" & vbc.Name & "] " & ln & vbCrLf
                End If
            Next j
        End If
    Next vbc
    On Error GoTo 0

    wb.Close SaveChanges:=False

    ' Mostrar o copiar al portapapeles si es largo
    If Len(informe) > 2000 Then
        On Error Resume Next
        Dim dObj As Object
        Set dObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dObj.SetText informe
        dObj.PutInClipboard
        On Error GoTo 0
        MsgBox "Resultado copiado al portapapeles." & vbCrLf & vbCrLf & _
               Left(informe, 1000) & "...", vbInformation, "Inspeccion"
    Else
        MsgBox informe, vbInformation, "Inspeccion: " & wb.Name
    End If
End Sub

' -----------------------------------------------------------------------
' RenombrarModulos
' Renombra los VBComponents en el proyecto destino.
' Registra el mapeo oldName->newName en el diccionario para que
' las referencias en codigo tambien queden actualizadas.
' No toca ThisWorkbook ni modulos de hoja (Type=100) porque
' sus nombres estan ligados a los nombres de hoja del libro.
' -----------------------------------------------------------------------
Private Sub RenombrarModulos(wb As Workbook, dict As Object, excl As String)
    Dim vbc As Object
    Dim oldName As String
    Dim newName As String

    For Each vbc In wb.VBProject.VBComponents
        ' Solo renombrar modulos estandar (1) y de clase (2)
        ' Los de documento (100) como ThisWorkbook y Hojas NO se tocan
        If vbc.Type = 1 Or vbc.Type = 2 Then
            oldName = vbc.Name

            ' Obtener o crear nombre ofuscado para este modulo
            If Not dict.Exists(oldName) Then
                If InStr(1, excl, "|" & oldName & "|", vbTextCompare) = 0 Then
                    dict.Add oldName, mod_Collector.CrearNombreUnico(dict, excl)
                End If
            End If

            If dict.Exists(oldName) Then
                newName = CStr(dict(oldName))
                On Error Resume Next
                vbc.Name = newName
                On Error GoTo 0
            End If
        End If
    Next vbc
End Sub

' -----------------------------------------------------------------------
' ActualizarBotones
' Recorre todas las hojas del workbook y actualiza el OnAction de
' cada boton/shape para que apunte al nuevo nombre ofuscado.
' Cubre: Form Controls, ActiveX, Shapes con macro asignada.
' -----------------------------------------------------------------------
Private Sub ActualizarBotones(wb As Workbook, dict As Object)
    Dim ws As Object
    Dim shp As Object
    Dim mac As String
    Dim macPuro As String
    Dim nuevoMac As String

    For Each ws In wb.Worksheets
        For Each shp In ws.Shapes
            mac = ""
            On Error Resume Next
            mac = shp.OnAction
            On Error GoTo 0

            If Len(mac) > 0 Then
                ' OnAction puede venir como:
                '   "MiSub"
                '   "'Libro.xlsm'!MiSub"
                '   "Modulo.MiSub"
                macPuro = ExtraerNombrePuro(mac)
                nuevoMac = mac

                If dict.Exists(macPuro) Then
                    ' Sustituir solo la parte del nombre, conservando prefijos
                    nuevoMac = Replace(mac, macPuro, CStr(dict(macPuro)), , , vbTextCompare)
                    On Error Resume Next
                    shp.OnAction = nuevoMac
                    On Error GoTo 0
                End If
            End If
        Next shp
    Next ws
End Sub

' -----------------------------------------------------------------------
' ExtraerNombrePuro
' Extrae el nombre del procedimiento de una referencia OnAction.
' Ejemplos:
'   "'Libro.xlsm'!MiSub"  ->  "MiSub"
'   "Modulo.MiSub"         ->  "MiSub"
'   "MiSub"                ->  "MiSub"
' -----------------------------------------------------------------------
Private Function ExtraerNombrePuro(ByVal mac As String) As String
    Dim resultado As String
    resultado = mac

    ' Quitar prefijo de workbook: 'Libro.xlsm'!
    Dim posExcl As Long
    posExcl = InStr(resultado, "!")
    If posExcl > 0 Then
        resultado = Mid(resultado, posExcl + 1)
    End If

    ' Quitar prefijo de modulo: Modulo.
    Dim posPunto As Long
    posPunto = InStrRev(resultado, ".")
    If posPunto > 0 Then
        resultado = Mid(resultado, posPunto + 1)
    End If

    ExtraerNombrePuro = Trim(resultado)
End Function

' -----------------------------------------------------------------------
' SeleccionarArchivo
' Abre dialogo de seleccion y devuelve la ruta o "" si se cancela.
' -----------------------------------------------------------------------
Private Function SeleccionarArchivo() As String
    Dim ruta As String
    ruta = Application.GetOpenFilename( _
        "Excel con Macros (*.xlsm),*.xlsm," & _
        "Excel Binario (*.xlsb),*.xlsb," & _
        "Complemento (*.xlam),*.xlam", _
        1, "Selecciona el archivo a ofuscar", , False)

    If ruta = "False" Then
        SeleccionarArchivo = ""
    Else
        SeleccionarArchivo = ruta
    End If
End Function


