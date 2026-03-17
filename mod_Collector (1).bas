Attribute VB_Name = "mod_Collector"
Option Explicit

' ============================================================================
' LLENAR DICCIONARIO: Captura Módulos, Macros y Propiedades
' Estrategia de renombrado: Invertir + Leet (a→4,e→3,i→1,o→0,s→5) + minúsculas
' ============================================================================
Public Sub LlenarDiccionario(ByRef wb As Workbook, ByRef dict As Object)
    Dim vbc As Object
    Dim i As Long, linea As String, nombreCapturado As String
    
    On Error Resume Next
    
    ' ========================================================================
    ' 1. Capturar Nombres de MÓDULOS
    ' ========================================================================
    For Each vbc In wb.VBProject.VBComponents
        ' Tipo 1: Módulo estándar, Tipo 2: Clase, Tipo 3: Formulario
        If vbc.Type <= 3 Then
            If Not EsModuloProtegido(vbc.Name) Then
                If Not dict.Exists(vbc.Name) Then
                    dict.Add vbc.Name, TransformarNombre(vbc.Name)
                End If
            End If
            
            ' ================================================================
            ' 2. Capturar NOMBRES DE MACROS (Subs, Functions y Properties)
            ' ================================================================
            For i = 1 To vbc.CodeModule.CountOfLines
                linea = Trim(vbc.CodeModule.Lines(i, 1))
                
                nombreCapturado = ExtraerNombreDeLinea(linea)
                
                If nombreCapturado <> "" Then
                    If Not EsFuncionProtegida(nombreCapturado) Then
                        If Not dict.Exists(nombreCapturado) Then
                            dict.Add nombreCapturado, TransformarNombre(nombreCapturado)
                        End If
                    End If
                End If
            Next i
        End If
    Next vbc
    
    On Error GoTo 0
End Sub

' ============================================================================
' TRANSFORMAR NOMBRE: Invertir → Leet → Minúsculas
' Ejemplo: ContarNotas → satoNratnoC → 54t0Nr4tn0C → 54t0nr4tn0c
' ============================================================================
Public Function TransformarNombre(ByVal nombre As String) As String
    Dim i As Long, resultado As String, c As String
    
    ' Paso 1: Invertir
    resultado = ""
    For i = Len(nombre) To 1 Step -1
        resultado = resultado & Mid(nombre, i, 1)
    Next i
    
    ' Paso 2: Leet speak (a→4, e→3, i→1, o→0, s→5)
    resultado = Replace(resultado, "a", "4")
    resultado = Replace(resultado, "A", "4")
    resultado = Replace(resultado, "e", "3")
    resultado = Replace(resultado, "E", "3")
    resultado = Replace(resultado, "i", "1")
    resultado = Replace(resultado, "I", "1")
    resultado = Replace(resultado, "o", "0")
    resultado = Replace(resultado, "O", "0")
    resultado = Replace(resultado, "s", "5")
    resultado = Replace(resultado, "S", "5")
    
    ' Paso 3: Todo minúsculas
    resultado = LCase(resultado)
    
    ' Seguridad: si empieza por dígito (inválido en VBA), prefijo 'x'
    If resultado <> "" Then
        If resultado Like "[0-9]*" Then resultado = "x" & resultado
    End If
    
    TransformarNombre = resultado
End Function

' ============================================================================
' EXTRAER NOMBRE DE LÍNEA (Sub, Function, Property)
' ============================================================================
Private Function ExtraerNombreDeLinea(ByVal l As String) As String
    Dim partes() As String, i As Integer, p As String
    
    ' Limpiar paréntesis
    l = Replace(Replace(l, "(", " "), ")", " ")
    partes = Split(l, " ")
    
    For i = LBound(partes) To UBound(partes) - 1
        p = LCase(Trim(partes(i)))
        
        ' Detectar Sub, Function o Property Get/Let/Set
        If p = "sub" Or p = "function" Or p = "property" Then
            Dim nom As String
            nom = Trim(partes(i + 1))
            
            ' Si es Property, el nombre está en partes(i+2): "Property Get NombreProp"
            If p = "property" And i + 2 <= UBound(partes) Then
                nom = Trim(partes(i + 2))
            End If
            
            If nom <> "" And Not EsPalabraReservada(nom) Then
                ExtraerNombreDeLinea = nom
                Exit Function
            End If
        End If
    Next i
    
    ExtraerNombreDeLinea = ""
End Function

' ============================================================================
' VALIDACIONES DE PROTECCIÓN
' ============================================================================
Private Function EsModuloProtegido(ByVal nombre As String) As Boolean
    Dim protegidos As Variant
    protegidos = Array("mod_Main", "mod_Collector", "mod_Tokenizer", _
                       "mod_Garbage", "mod_Types", "mod_Traductor", _
                       "mod_Internal_Helper", "Comun_Constantes", _
                       "Comun_Buscador", "ThisWorkbook")
    
    Dim i As Long
    For i = LBound(protegidos) To UBound(protegidos)
        If LCase(nombre) = LCase(protegidos(i)) Then
            EsModuloProtegido = True
            Exit Function
        End If
    Next i
    
    EsModuloProtegido = False
End Function

Private Function EsFuncionProtegida(ByVal nombre As String) As Boolean
    Dim protegidas As Variant
    protegidas = Array("f_tr", "ProcesarLineaMaestra", "LlenarDiccionario", _
                       "TransformarNombre", "EjecutarOfuscador", _
                       "SincronizarBotones", "InyectarModuloTraductor", _
                       "Auto_Open", "Workbook_Open")
    
    Dim i As Long
    For i = LBound(protegidas) To UBound(protegidas)
        If LCase(nombre) = LCase(protegidas(i)) Then
            EsFuncionProtegida = True
            Exit Function
        End If
    Next i
    
    EsFuncionProtegida = False
End Function

Private Function EsPalabraReservada(ByVal palabra As String) As Boolean
    Dim reservadas As String
    reservadas = ",get,let,set,property,byval,byref,optional,paramarray," & _
                 "as,new,withevents,friend,private,public,static,"
    
    EsPalabraReservada = InStr(1, reservadas, "," & LCase(palabra) & ",", vbTextCompare) > 0
End Function
