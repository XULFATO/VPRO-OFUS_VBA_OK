Attribute VB_Name = "mod_Collector"
Option Explicit

' -----------------------------------------------------------------------
' LlenarDiccionario
' Escanea todos los componentes VBA del workbook y registra en dict
' los identificadores seguros para renombrar (solo declaraciones explicitas).
' -----------------------------------------------------------------------
Public Sub LlenarDiccionario(wb As Workbook, dict As Object)
    Dim vbc As Object
    Dim i As Long
    Dim linea As String
    Dim lineaTrim As String
    Dim excl As String

    excl = ObtenerExclusiones()

    For Each vbc In wb.VBProject.VBComponents
        If vbc.CodeModule.CountOfLines = 0 Then GoTo SiguienteModulo

        For i = 1 To vbc.CodeModule.CountOfLines
            linea = vbc.CodeModule.Lines(i, 1)
            lineaTrim = Trim(linea)

            ' Saltar lineas que no deben tocarse
            If lineaTrim = "" Then GoTo SiguienteLinea
            If lineaTrim Like "Declare *" Then GoTo SiguienteLinea
            If lineaTrim Like "Public Declare *" Then GoTo SiguienteLinea
            If lineaTrim Like "Private Declare *" Then GoTo SiguienteLinea
            If lineaTrim Like "#*" Then GoTo SiguienteLinea
            If lineaTrim Like "Attribute *" Then GoTo SiguienteLinea
            If lineaTrim Like "Option *" Then GoTo SiguienteLinea
            If Left(lineaTrim, 1) = "'" Then GoTo SiguienteLinea

            ' Solo analizar lineas de declaracion
            AnalizarDeclaracion lineaTrim, dict, excl

SiguienteLinea:
        Next i
SiguienteModulo:
    Next vbc
End Sub

' -----------------------------------------------------------------------
' AnalizarDeclaracion
' Solo registra nombres encontrados en lineas de declaracion explicita.
' -----------------------------------------------------------------------
Private Sub AnalizarDeclaracion(ByVal linea As String, dict As Object, excl As String)
    Dim esDeclaracion As Boolean
    esDeclaracion = False

    If linea Like "Dim *" Then esDeclaracion = True
    If linea Like "Private *" Then esDeclaracion = True
    If linea Like "Public *" Then esDeclaracion = True
    If linea Like "Friend *" Then esDeclaracion = True
    If linea Like "Static *" Then esDeclaracion = True
    If linea Like "Const *" Then esDeclaracion = True
    If linea Like "Sub *" Then esDeclaracion = True
    If linea Like "Function *" Then esDeclaracion = True
    If linea Like "Property *" Then esDeclaracion = True

    If Not esDeclaracion Then Exit Sub

    Dim nombre As String
    nombre = ExtraerNombreDeDeclaracion(linea)

    If Len(nombre) = 0 Then Exit Sub
    If Len(nombre) <= 2 Then Exit Sub  ' Ignorar nombres de 1-2 chars (i, j, k...)

    ' Verificar que no es una exclusion
    If InStr(1, excl, "|" & nombre & "|", vbTextCompare) > 0 Then Exit Sub

    ' Registrar en el diccionario si no existe
    If Not dict.Exists(nombre) Then
        dict.Add nombre, mod_Collector.CrearNombreUnico(dict, excl)
    End If
End Sub

' -----------------------------------------------------------------------
' ExtraerNombreDeDeclaracion
' Extrae el identificador real saltando modificadores de VBA.
' -----------------------------------------------------------------------
Private Function ExtraerNombreDeDeclaracion(ByVal linea As String) As String
    Dim partes() As String
    Dim i As Long
    Dim p As String
    Dim resultado As String

    partes = Split(linea, " ")

    For i = 0 To UBound(partes)
        p = LCase(Trim(partes(i)))

        ' Saltar modificadores y keywords de declaracion
        Select Case p
            Case "dim", "private", "public", "friend", "static", "const", _
                 "sub", "function", "property", "get", "let", "set", _
                 "withevents", ""
                ' Continuar al siguiente token
            Case Else
                ' El primer token que no sea keyword es el nombre
                resultado = Trim(partes(i))

                ' Limpiar sufijos: parentesis, dos puntos, tipo
                If InStr(resultado, "(") > 0 Then resultado = Split(resultado, "(")(0)
                If InStr(resultado, ":") > 0 Then resultado = Split(resultado, ":")(0)
                If InStr(resultado, "!") > 0 Then resultado = Split(resultado, "!")(0)
                If InStr(resultado, "#") > 0 Then resultado = Split(resultado, "#")(0)
                If InStr(resultado, "$") > 0 Then resultado = Split(resultado, "$")(0)
                If InStr(resultado, "%") > 0 Then resultado = Split(resultado, "%")(0)
                If InStr(resultado, "&") > 0 Then resultado = Split(resultado, "&")(0)
                If InStr(resultado, "@") > 0 Then resultado = Split(resultado, "@")(0)

                resultado = Trim(resultado)

                ' Validar que es un identificador VBA valido
                If EsIdentificadorValido(resultado) Then
                    ExtraerNombreDeDeclaracion = resultado
                End If
                Exit Function
        End Select
    Next i
End Function

' -----------------------------------------------------------------------
' EsIdentificadorValido
' Comprueba que el nombre cumple reglas basicas de identificador VBA.
' -----------------------------------------------------------------------
Private Function EsIdentificadorValido(ByVal nombre As String) As Boolean
    If Len(nombre) = 0 Then EsIdentificadorValido = False: Exit Function
    If Not (Left(nombre, 1) Like "[a-zA-Z]") Then EsIdentificadorValido = False: Exit Function
    Dim i As Long
    For i = 1 To Len(nombre)
        If Not (Mid(nombre, i, 1) Like "[a-zA-Z0-9_]") Then
            EsIdentificadorValido = False
            Exit Function
        End If
    Next i
    EsIdentificadorValido = True
End Function

' -----------------------------------------------------------------------
' CrearNombreUnico
' Genera un nombre ofuscado que no existe en el diccionario
' ni en la lista de exclusiones.
' -----------------------------------------------------------------------
Public Function CrearNombreUnico(dict As Object, excl As String) As String
    Dim nombre As String
    Dim i As Integer

    Do
        ' 3 letras iniciales garantizadas + 10 caracteres alfanumericos
        nombre = Chr(Int(26 * Rnd + 97)) & _
                 Chr(Int(26 * Rnd + 97)) & _
                 Chr(Int(26 * Rnd + 97))
        For i = 1 To 10
            nombre = nombre & Mid("abcdefghijklmnopqrstuvwxyz0123456789", _
                                  Int(36 * Rnd + 1), 1)
        Next i
    Loop While dict.Exists(nombre) Or _
               InStr(1, excl, "|" & nombre & "|", vbTextCompare) > 0

    CrearNombreUnico = nombre
End Function

' -----------------------------------------------------------------------
' ObtenerExclusiones
' Lista completa de palabras que NUNCA deben renombrarse.
' -----------------------------------------------------------------------
Public Function ObtenerExclusiones() As String
    Dim e As String

    ' Eventos de Workbook y Worksheet
    e = "|Workbook_Open|Workbook_BeforeClose|Workbook_BeforeSave|Workbook_AfterSave|"
    e = e & "Workbook_SheetChange|Workbook_SheetActivate|Workbook_NewSheet|"
    e = e & "Worksheet_Change|Worksheet_SelectionChange|Worksheet_Activate|"
    e = e & "Worksheet_Deactivate|Worksheet_BeforeDoubleClick|Worksheet_Calculate|"
    e = e & "UserForm_Initialize|UserForm_Terminate|UserForm_Click|"

    ' Objetos y propiedades de Excel
    e = e & "Workbook|Workbooks|Worksheet|Worksheets|Sheets|Range|Cells|Rows|Columns|"
    e = e & "ActiveWorkbook|ActiveSheet|ActiveCell|ThisWorkbook|Selection|"
    e = e & "Value|Name|Count|Address|Formula|Text|NumberFormat|Interior|Font|"
    e = e & "Offset|Resize|End|Row|Column|Width|Height|Left|Top|Visible|"

    ' Metodos de Excel
    e = e & "Open|Close|Save|SaveAs|Activate|Select|Copy|Paste|Delete|Insert|"
    e = e & "Add|Remove|Clear|ClearContents|AutoFit|Sort|Filter|Find|Replace|"
    e = e & "MsgBox|InputBox|Print|Debug|Application|"

    ' Tipos de datos VBA
    e = e & "String|Long|Integer|Single|Double|Boolean|Byte|Date|Object|Variant|"
    e = e & "Currency|LongLong|LongPtr|"

    ' Keywords VBA
    e = e & "Sub|Function|Property|End|If|Then|Else|ElseIf|For|Next|Do|Loop|"
    e = e & "While|Until|Select|Case|With|Each|In|To|Step|Exit|Return|GoTo|"
    e = e & "Dim|Set|Let|Get|New|Nothing|Empty|Null|True|False|And|Or|Not|Xor|"
    e = e & "Public|Private|Friend|Static|Const|Type|Enum|Declare|Lib|Alias|"
    e = e & "ByVal|ByRef|Optional|ParamArray|As|Is|Like|Mod|"
    e = e & "On|Error|Resume|Next|GoSub|Call|Implements|RaiseEvent|"

    ' Funciones built-in VBA
    e = e & "Len|Left|Right|Mid|InStr|InStrRev|UCase|LCase|Trim|LTrim|RTrim|"
    e = e & "Split|Join|Replace|String|Space|Chr|Asc|ChrW|AscW|"
    e = e & "CStr|CInt|CLng|CDbl|CBool|CDate|CVar|CByte|CSng|"
    e = e & "Int|Fix|Abs|Sgn|Sqr|Exp|Log|Sin|Cos|Tan|Atn|"
    e = e & "Rnd|Randomize|Timer|Now|Date|Time|DateAdd|DateDiff|DatePart|"
    e = e & "Format|FormatNumber|FormatDate|FormatCurrency|"
    e = e & "IsNumeric|IsDate|IsEmpty|IsNull|IsObject|IsArray|IsError|"
    e = e & "Array|UBound|LBound|ReDim|Preserve|Erase|"
    e = e & "MsgBox|InputBox|Shell|Environ|"
    e = e & "CreateObject|GetObject|TypeName|VarType|"
    e = e & "Dir|Kill|FileCopy|MkDir|RmDir|CurDir|ChDir|"
    e = e & "Open|Close|Print|Write|Line|Input|Get|Put|Seek|EOF|LOF|LOC|"
    e = e & "vbCrLf|vbCr|vbLf|vbTab|vbNullString|vbNullChar|"
    e = e & "vbOKOnly|vbOKCancel|vbYesNo|vbYesNoCancel|vbInformation|"
    e = e & "vbCritical|vbQuestion|vbExclamation|vbOK|vbCancel|vbYes|vbNo|"
    e = e & "vbTextCompare|vbBinaryCompare|"

    ObtenerExclusiones = e
End Function
