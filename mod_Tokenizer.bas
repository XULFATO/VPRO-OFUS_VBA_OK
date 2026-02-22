Attribute VB_Name = "mod_Tokenizer"

Option Explicit

Public Function ProcesarLineaMaestra(ByVal linea As String, ByRef dict As Object, ByVal clave As String) As String
    Dim k As Variant, resultado As String, partes() As String, i As Long
    Dim esConstante As Boolean
    
    resultado = NormalizarTexto(linea)
    
    ' 1. Ignorar comentarios y detectar declaraciones de constantes
    If Left(Trim(resultado), 1) = "'" Then: ProcesarLineaMaestra = "": Exit Function
    esConstante = (InStr(1, LCase(resultado), "const ") > 0)

    ' 2. Dividir por comillas
    partes = Split(resultado, """")
    
    For i = LBound(partes) To UBound(partes)
        If i Mod 2 <> 0 Then
            ' DENTRO DE COMILLAS: Cifrar XOR (EXCEPTO si es Constante)
            If Len(partes(i)) > 0 And Not esConstante Then
                partes(i) = "f_tr(""" & CifrarTextoXOR(partes(i), clave) & """)"
            Else
                ' Si es constante o está vacío, mantenemos comillas literales
                partes(i) = """" & partes(i) & """"
            End If
        Else
            ' FUERA DE COMILLAS: Ofuscar nombres de variables/subs
            For Each k In dict.Keys
                partes(i) = ReemplazarPalabraExacta(partes(i), CStr(k), CStr(dict(k)))
            Next k
        End If
    Next i
    
    ' Unimos las partes. Las comillas ya fueron gestionadas dentro del bucle
    ProcesarLineaMaestra = Join(partes, "")
End Function

Private Function CifrarTextoXOR(ByVal t As String, ByVal clave As String) As String
    Dim j As Long, k As Long, res As String, cLen As Integer
    cLen = Len(clave)
    For j = 1 To Len(t)
        ' Sincronizado: j=1 usa clave(1). En f_tr, i=0 usará clave(1).
        k = ((j - 1) Mod cLen) + 1
        res = res & (Asc(Mid(t, j, 1)) Xor Asc(Mid(clave, k, 1))) & IIf(j < Len(t), ",", "")
    Next j
    CifrarTextoXOR = res
End Function

Private Function ReemplazarPalabraExacta(ByVal texto As String, ByVal viejo As String, ByVal nuevo As String) As String
    Dim RegEx As Object, vEsc As String
    If InStr(1, texto, viejo, vbBinaryCompare) = 0 Then: ReemplazarPalabraExacta = texto: Exit Function
    
    ' Escapar caracteres especiales para RegEx
    vEsc = viejo: Dim c: For Each c In Array(".", "(", ")", "[", "]", "{", "}", "*", "+", "?", "^", "$", "|", "\")
        vEsc = Replace(vEsc, c, "\" & c)
    Next
    
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx: .Global = True: .IgnoreCase = False: .Pattern = "\b" & vEsc & "\b": End With
    ReemplazarPalabraExacta = RegEx.Replace(texto, nuevo)
End Function

Private Function NormalizarTexto(ByVal t As String) As String
    ' Aquí puedes mantener tu función de ADODB.Stream o los Replace manuales
    NormalizarTexto = Replace(t, "Ã³", "ó") ' Ejemplo simplificado
End Function
