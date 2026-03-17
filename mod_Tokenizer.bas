Attribute VB_Name = "mod_Tokenizer"
Option Explicit

Public Function ProcesarLineaMaestra(ByVal linea As String, ByRef dict As Object, ByVal clave As String) As String
    Dim k As Variant, resultado As String, partes() As String, i As Long
    Dim esConstante As Boolean
    
    resultado = linea   ' SIN NormalizarTexto: no tocamos el contenido del código fuente
    
    ' 1. Ignorar comentarios y detectar declaraciones de constantes
    If Left(Trim(resultado), 1) = "'" Then ProcesarLineaMaestra = "": Exit Function
    esConstante = (InStr(1, LCase(resultado), "const ") > 0)

    ' 2. Dividir por comillas
    partes = Split(resultado, """")
    
    For i = LBound(partes) To UBound(partes)
        If i Mod 2 <> 0 Then
            ' DENTRO DE COMILLAS: Cifrar XOR (excepto constantes y strings vacíos)
            If Len(partes(i)) > 0 And Not esConstante Then
                partes(i) = "f_tr(""" & CifrarTextoXOR(partes(i), clave) & """)"
            Else
                ' String vacío "" o constante: mantener comillas literales
                partes(i) = """" & partes(i) & """"
            End If
        Else
            ' FUERA DE COMILLAS: Ofuscar nombres de variables/subs
            For Each k In dict.Keys
                partes(i) = ReemplazarPalabraExacta(partes(i), CStr(k), CStr(dict(k)))
            Next k
        End If
    Next i
    
    ' Unimos sin separador (las comillas ya están dentro de cada segmento)
    resultado = Join(partes, "")
    
    ' 3. Inyectar ruido al 15% de probabilidad
    ProcesarLineaMaestra = InyectarRuido(resultado)
End Function

' ============================================================================
' CIFRADO XOR
' Sincronización: j=1 usa clave(1). En f_tr, i=0 usa k=(0 Mod Len)+1=1 ✓
' ============================================================================
Private Function CifrarTextoXOR(ByVal t As String, ByVal clave As String) As String
    Dim j As Long, k As Long, res As String, cLen As Integer
    cLen = Len(clave)
    For j = 1 To Len(t)
        k = ((j - 1) Mod cLen) + 1
        res = res & (Asc(Mid(t, j, 1)) Xor Asc(Mid(clave, k, 1))) & IIf(j < Len(t), ",", "")
    Next j
    CifrarTextoXOR = res
End Function

' ============================================================================
' REEMPLAZO CON LÍMITES DE PALABRA (RegEx \b)
' ============================================================================
Private Function ReemplazarPalabraExacta(ByVal texto As String, ByVal viejo As String, ByVal nuevo As String) As String
    Dim RegEx As Object, vEsc As String, c As Variant
    If InStr(1, texto, viejo, vbBinaryCompare) = 0 Then
        ReemplazarPalabraExacta = texto
        Exit Function
    End If
    
    ' Escapar caracteres especiales para RegEx
    vEsc = viejo
    For Each c In Array(".", "(", ")", "[", "]", "{", "}", "*", "+", "?", "^", "$", "|", "\")
        vEsc = Replace(vEsc, c, "\" & c)
    Next c
    
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .Global = True
        .IgnoreCase = False
        .Pattern = "\b" & vEsc & "\b"
    End With
    ReemplazarPalabraExacta = RegEx.Replace(texto, nuevo)
End Function
