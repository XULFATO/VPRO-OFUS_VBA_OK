Attribute VB_Name = "mod_Tokenizer"
Option Explicit

' -----------------------------------------------------------------------
' ProcesarLineaMaestra
' Punto de entrada principal. Protege lineas criticas y orquesta el proceso.
' -----------------------------------------------------------------------
Public Function ProcesarLineaMaestra(ByVal linea As String, dict As Object) As String
    Dim lineaTrim As String
    lineaTrim = Trim(linea)

    ' Proteger lineas que no deben tocarse nunca
    If lineaTrim = "" Then
        ProcesarLineaMaestra = linea
        Exit Function
    End If
    If lineaTrim Like "Declare *" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "Public Declare *" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "Private Declare *" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "#If *" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "#ElseIf *" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "#Else*" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "#End If*" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "#End*" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "Attribute *" Then ProcesarLineaMaestra = linea: Exit Function
    If lineaTrim Like "Option *" Then ProcesarLineaMaestra = linea: Exit Function

    ' Tokenizar, ofuscar y reensamblar
    Dim segmentos() As Segmento
    segmentos = TokenizarLinea(linea)

    Dim lineaProcesada As String
    lineaProcesada = OfuscarSegmentos(segmentos, dict)

    ' Inyectar ruido sobre la linea completa y limpia
    If Len(Trim(lineaProcesada)) > 0 Then
        lineaProcesada = mod_Garbage.InyectarRuido(lineaProcesada)
    End If

    ProcesarLineaMaestra = lineaProcesada
End Function

' -----------------------------------------------------------------------
' TokenizarLinea
' Descompone la linea en segmentos: codigo o string literal.
' Descarta comentarios. Maneja comillas dobles escapadas "".
' -----------------------------------------------------------------------
Public Function TokenizarLinea(ByVal linea As String) As Segmento()
    Dim segmentos() As Segmento
    Dim i As Long
    Dim n As Long
    Dim char As String
    Dim sig As String
    Dim acumulador As String
    Dim enString As Boolean

    ReDim segmentos(0 To 0)
    n = 0
    i = 1
    enString = False
    acumulador = ""

    Do While i <= Len(linea)
        char = Mid(linea, i, 1)
        sig = ""
        If i < Len(linea) Then sig = Mid(linea, i + 1, 1)

        If Not enString Then
            If char = """" Then
                ' Guardar codigo acumulado antes del string
                GuardarSegmento segmentos, n, acumulador, True
                acumulador = """"
                enString = True
            ElseIf char = "'" Then
                ' Inicio de comentario: guardar codigo y terminar
                GuardarSegmento segmentos, n, acumulador, True
                acumulador = ""
                Exit Do
            Else
                acumulador = acumulador & char
            End If
        Else
            If char = """" Then
                If sig = """" Then
                    ' Comilla escapada "" dentro de string
                    acumulador = acumulador & """"""
                    i = i + 1
                Else
                    ' Fin del string
                    acumulador = acumulador & """"
                    GuardarSegmento segmentos, n, acumulador, False
                    acumulador = ""
                    enString = False
                End If
            Else
                acumulador = acumulador & char
            End If
        End If

        i = i + 1
    Loop

    ' Guardar lo que quede en el acumulador
    If Len(acumulador) > 0 Then
        GuardarSegmento segmentos, n, acumulador, True
    End If

    TokenizarLinea = segmentos
End Function

' -----------------------------------------------------------------------
' GuardarSegmento
' Añade un segmento al array dinamico.
' -----------------------------------------------------------------------
Private Sub GuardarSegmento(ByRef segs() As Segmento, ByRef n As Long, _
                             texto As String, esCod As Boolean)
    If Len(texto) = 0 Then Exit Sub
    If n > 0 Then ReDim Preserve segs(0 To n)
    segs(n).Contenido = texto
    segs(n).EsCodigo = esCod
    n = n + 1
End Sub

' -----------------------------------------------------------------------
' OfuscarSegmentos
' Aplica el diccionario de renombrado solo en segmentos de codigo.
' -----------------------------------------------------------------------
Public Function OfuscarSegmentos(ByRef segs() As Segmento, dict As Object) As String
    Dim numSegs As Long
    Dim i As Long
    Dim j As Long
    Dim resultado As String
    Dim claves() As String

    ' Calculo seguro del numero de segmentos
    On Error Resume Next
    numSegs = UBound(segs) - LBound(segs) + 1
    On Error GoTo 0

    If numSegs <= 0 Then
        OfuscarSegmentos = ""
        Exit Function
    End If

    If dict.Count > 0 Then
        claves = ObtenerClavesOrdenadas(dict)
        For i = LBound(segs) To UBound(segs)
            If segs(i).EsCodigo Then
                For j = LBound(claves) To UBound(claves)
                    segs(i).Contenido = ReemplazarIdentificador( _
                        segs(i).Contenido, claves(j), CStr(dict(claves(j))))
                Next j
            End If
            resultado = resultado & segs(i).Contenido
        Next i
    Else
        For i = LBound(segs) To UBound(segs)
            resultado = resultado & segs(i).Contenido
        Next i
    End If

    OfuscarSegmentos = resultado
End Function

' -----------------------------------------------------------------------
' ObtenerClavesOrdenadas
' Devuelve las claves del diccionario ordenadas por longitud descendente.
' Evita reemplazos parciales (ej: "Calc" dentro de "CalcTotal").
' -----------------------------------------------------------------------
Private Function ObtenerClavesOrdenadas(dict As Object) As String()
    Dim claves() As String
    Dim k As Variant
    Dim i As Long
    Dim j As Long
    Dim temp As String

    ReDim claves(0 To dict.Count - 1)
    i = 0
    For Each k In dict.Keys
        claves(i) = CStr(k)
        i = i + 1
    Next k

    ' Bubble sort descendente por longitud
    For i = LBound(claves) To UBound(claves) - 1
        For j = i + 1 To UBound(claves)
            If Len(claves(i)) < Len(claves(j)) Then
                temp = claves(i)
                claves(i) = claves(j)
                claves(j) = temp
            End If
        Next j
    Next i

    ObtenerClavesOrdenadas = claves
End Function

' -----------------------------------------------------------------------
' ReemplazarIdentificador
' Sustituye "viejo" por "nuevo" solo cuando aparece como palabra completa.
' Usa vbTextCompare para cubrir diferencias de mayusculas/minusculas.
' -----------------------------------------------------------------------
Private Function ReemplazarIdentificador(texto As String, _
                                          viejo As String, _
                                          nuevo As String) As String
    Dim pos As Long
    Dim longViejo As Long
    Dim charPrev As String
    Dim charPost As String

    longViejo = Len(viejo)
    pos = 1

    Do
        pos = InStr(pos, texto, viejo, vbTextCompare)
        If pos = 0 Then Exit Do

        If pos > 1 Then
            charPrev = Mid(texto, pos - 1, 1)
        Else
            charPrev = " "
        End If

        If pos + longViejo <= Len(texto) Then
            charPost = Mid(texto, pos + longViejo, 1)
        Else
            charPost = " "
        End If

        If EsDelimitador(charPrev) And EsDelimitador(charPost) Then
            texto = Left(texto, pos - 1) & nuevo & Mid(texto, pos + longViejo)
            pos = pos + Len(nuevo)
        Else
            pos = pos + longViejo
        End If
    Loop

    ReemplazarIdentificador = texto
End Function

' -----------------------------------------------------------------------
' EsDelimitador
' Devuelve True si el caracter NO forma parte de un identificador VBA.
' -----------------------------------------------------------------------
Private Function EsDelimitador(char As String) As Boolean
    EsDelimitador = Not (char Like "[a-zA-Z0-9_]")
End Function
