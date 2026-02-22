Attribute VB_Name = "mod_Collector"
Option Explicit

' ============================================================================
' LLENAR DICCIONARIO: Captura Módulos y Macros
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
            ' Excluir módulos protegidos
            If Not EsModuloProtegido(vbc.Name) Then
                If Not dict.Exists(vbc.Name) Then
                    dict.Add vbc.Name, GenerarNombreAleatorio()
                End If
            End If
            
            ' ================================================================
            ' 2. Capturar NOMBRES DE MACROS (Subs y Functions)
            ' ================================================================
            For i = 1 To vbc.CodeModule.CountOfLines
                linea = Trim(vbc.CodeModule.Lines(i, 1))
                
                ' Extraer nombre de Sub/Function
                nombreCapturado = ExtraerNombreDeLinea(linea)
                
                If nombreCapturado <> "" Then
                    ' Excluir funciones protegidas
                    If Not EsFuncionProtegida(nombreCapturado) Then
                        If Not dict.Exists(nombreCapturado) Then
                            dict.Add nombreCapturado, GenerarNombreAleatorio()
                        End If
                    End If
                End If
            Next i
        End If
    Next vbc
    
    On Error GoTo 0
End Sub

' ============================================================================
' EXTRAER NOMBRE DE LÍNEA
' ============================================================================
Private Function ExtraerNombreDeLinea(ByVal l As String) As String
    Dim partes() As String, i As Integer, p As String, nom As String
    
    ' Limpiar paréntesis
    l = Replace(Replace(l, "(", " "), ")", " ")
    partes = Split(l, " ")
    
    For i = LBound(partes) To UBound(partes) - 1
        p = LCase(Trim(partes(i)))
        
        ' Detectar declaración de procedimiento
        If p = "sub" Or p = "function" Then
            nom = Trim(partes(i + 1))
            
            ' Filtrar palabras reservadas y decoradores
            If nom <> "" And Not EsPalabraReservada(nom) Then
                ExtraerNombreDeLinea = nom
                Exit Function
            End If
        End If
    Next i
    
    ExtraerNombreDeLinea = ""
End Function

' ============================================================================
' GENERAR NOMBRE ALEATORIO
' ============================================================================
Public Function GenerarNombreAleatorio() As String
    Dim i As Integer, n As String
    Const LETRAS As String = "abcdefghijklmnopqrstuvwxyz"
    Const ALFANUM As String = "abcdefghijklmnopqrstuvwxyz0123456789"
    
    Randomize
    
    ' Primer carácter: siempre una letra (requisito VBA)
    n = Mid(LETRAS, Int(26 * Rnd + 1), 1)
    
    ' Caracteres 2-10: letras o números
    For i = 2 To 10
        n = n & Mid(ALFANUM, Int(Len(ALFANUM) * Rnd + 1), 1)
    Next i
    
    GenerarNombreAleatorio = n
End Function

' ============================================================================
' VALIDACIONES DE PROTECCIÓN
' ============================================================================
Private Function EsModuloProtegido(ByVal nombre As String) As Boolean
    Dim protegidos As Variant
    protegidos = Array("mod_Main", "mod_Collector", "mod_Tokenizer", _
                       "mod_Garbage", "mod_Types", "mod_Traductor", _
                       "mod_Internal_Helper", "ThisWorkbook")
    
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
                       "GenerarNombreAleatorio", "EjecutarOfuscador", _
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
