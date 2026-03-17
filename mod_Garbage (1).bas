Attribute VB_Name = "mod_Garbage"
Option Explicit

' ============================================================================
' INYECCIÓN DE RUIDO EN LÍNEAS
' NOTA: Llamar Randomize una sola vez en EjecutarOfuscador antes de usar esto
' ============================================================================
Public Function InyectarRuido(ByVal linea As String) As String
    Dim lineaTrim As String
    lineaTrim = Trim(linea)
    
    ' No modificar líneas de cierre de bloque
    If lineaTrim Like "End Sub*"      Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Function*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Property*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End If*"       Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End With*"     Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Select*"   Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Type*"     Then InyectarRuido = linea: Exit Function
    If lineaTrim = ""                 Then InyectarRuido = linea: Exit Function
    
    ' No inyectar si ya tiene comentario o es línea de continuación
    If InStr(linea, "'") > 0        Then InyectarRuido = linea: Exit Function
    If Right(lineaTrim, 1) = "_"    Then InyectarRuido = linea: Exit Function
    
    ' 15% de probabilidad de inyectar comentario falso
    If Rnd > 0.85 Then
        InyectarRuido = linea & " " & GenerarComentario()
    Else
        InyectarRuido = linea
    End If
End Function

' ============================================================================
' GENERACIÓN DE COMENTARIOS FALSOS
' ============================================================================
Private Function GenerarComentario() As String
    Dim acciones(9)     As String
    Dim conceptos(9)    As String
    Dim complementos(9) As String
    
    acciones(0) = "Aqui se ordena"
    acciones(1) = "Ahora contamos"
    acciones(2) = "Desde aqui entramos en la tabla de"
    acciones(3) = "Descartamos que el"
    acciones(4) = "Solamente entramos cuando"
    acciones(5) = "En este punto validamos"
    acciones(6) = "Recorremos el bloque de"
    acciones(7) = "Aqui se acumula el"
    acciones(8) = "Comprobamos que el"
    acciones(9) = "Se filtra por"
    
    conceptos(0) = "ID de denominacion"
    conceptos(1) = "rubrica contable"
    conceptos(2) = "enlace PAC"
    conceptos(3) = "contador de release"
    conceptos(4) = "cod tabla cliente"
    conceptos(5) = "num cuenta de enlace"
    conceptos(6) = "tipo de concepto"
    conceptos(7) = "cod enlace de rubrica"
    conceptos(8) = "importe acumulado"
    conceptos(9) = "registro de denominacion"
    
    complementos(0) = "sea mayor que la columna de referencia"
    complementos(1) = "antes de pasar al siguiente bloque"
    complementos(2) = "si el PAC esta activo"
    complementos(3) = "por tipo de cliente y release"
    complementos(4) = "cuando el cod enlace coincide"
    complementos(5) = "segun la rubrica del periodo"
    complementos(6) = "por num cuenta y denominacion"
    complementos(7) = "si el contador supera el umbral"
    complementos(8) = "para el concepto vigente"
    complementos(9) = "hasta completar el ciclo de tabla"
    
    GenerarComentario = "' " & acciones(Int(Rnd * 10)) & " " & _
                                conceptos(Int(Rnd * 10)) & " " & _
                                complementos(Int(Rnd * 10))
End Function

' ============================================================================
' GENERACIÓN DE BLOQUES DE CÓDIGO MUERTO
' ============================================================================
Public Function GenerarBloqueRuido() As String
    Dim bloque As String
    Dim v1 As String, v2 As String, v3 As String
    Dim n1 As Long, n2 As Long
    
    v1 = NombreOpaco()
    v2 = NombreOpaco()
    v3 = NombreOpaco()
    n1 = Int(Rnd * 9000) + 1000
    n2 = Int(Rnd * 900) + 100
    
    Select Case Int(Rnd * 6)
        
        Case 0
            bloque = "Private Sub " & v1 & "()" & vbCrLf
            bloque = bloque & "    Dim " & v2 & " As Long" & vbCrLf
            bloque = bloque & "    Dim " & v3 & " As Long" & vbCrLf
            bloque = bloque & "    " & v2 & " = " & n1 & vbCrLf
            bloque = bloque & "    " & v3 & " = " & n2 & vbCrLf
            bloque = bloque & "    ' Validamos que el cod enlace este dentro del rango PAC" & vbCrLf
            bloque = bloque & "    If " & v2 & " > " & v3 & " Then" & vbCrLf
            bloque = bloque & "        " & v2 & " = " & v3 & vbCrLf
            bloque = bloque & "    End If" & vbCrLf
            bloque = bloque & "End Sub"
        
        Case 1
            bloque = "Private Function " & v1 & "() As Long" & vbCrLf
            bloque = bloque & "    Dim " & v2 & " As Long" & vbCrLf
            bloque = bloque & "    Dim " & v3 & " As Long" & vbCrLf
            bloque = bloque & "    " & v3 & " = 0" & vbCrLf
            bloque = bloque & "    ' Recorremos el bloque de denominaciones hasta encontrar el PAC" & vbCrLf
            bloque = bloque & "    For " & v2 & " = 1 To " & n2 & vbCrLf
            bloque = bloque & "        " & v3 & " = " & v3 & " + " & v2 & vbCrLf
            bloque = bloque & "        If " & v3 & " > " & n1 & " Then Exit For" & vbCrLf
            bloque = bloque & "    Next " & v2 & vbCrLf
            bloque = bloque & "    " & v1 & " = " & v3 & vbCrLf
            bloque = bloque & "End Function"
        
        Case 2
            bloque = "Private Sub " & v1 & "()" & vbCrLf
            bloque = bloque & "    Dim " & v2 & " As Double" & vbCrLf
            bloque = bloque & "    Dim " & v3 & " As Double" & vbCrLf
            bloque = bloque & "    ' Aqui se acumula el importe por rubrica contable y num cuenta" & vbCrLf
            bloque = bloque & "    " & v2 & " = " & n1 & " * 0." & CStr(Int(Rnd * 9000 + 1000)) & vbCrLf
            bloque = bloque & "    " & v3 & " = " & v2 & " + " & n2 & vbCrLf
            bloque = bloque & "    ' Descartamos que el acumulado sea mayor que el umbral de release" & vbCrLf
            bloque = bloque & "    If " & v3 & " < 0 Then " & v3 & " = 0" & vbCrLf
            bloque = bloque & "End Sub"
        
        Case 3
            bloque = "Private Function " & v1 & "() As Boolean" & vbCrLf
            bloque = bloque & "    Dim " & v2 & " As Long" & vbCrLf
            bloque = bloque & "    ' Solamente entramos cuando el tipo cliente coincide con cod tabla" & vbCrLf
            bloque = bloque & "    " & v2 & " = " & n1 & vbCrLf
            bloque = bloque & "    Select Case " & v2 & vbCrLf
            bloque = bloque & "        Case 1 To " & n2 & vbCrLf
            bloque = bloque & "            " & v1 & " = True" & vbCrLf
            bloque = bloque & "        Case Else" & vbCrLf
            bloque = bloque & "            " & v1 & " = False" & vbCrLf
            bloque = bloque & "    End Select" & vbCrLf
            bloque = bloque & "End Function"
        
        Case 4
            bloque = "Private Sub " & v1 & "()" & vbCrLf
            bloque = bloque & "    Dim " & v2 & " As Long" & vbCrLf
            bloque = bloque & "    Dim " & v3 & " As String" & vbCrLf
            bloque = bloque & "    ' Desde aqui contamos los conceptos del periodo PAC" & vbCrLf
            bloque = bloque & "    " & v2 & " = 0" & vbCrLf
            bloque = bloque & "    Do While " & v2 & " < " & n2 & vbCrLf
            bloque = bloque & "        " & v2 & " = " & v2 & " + 1" & vbCrLf
            bloque = bloque & "        " & v3 & " = ""COD-"" & CStr(" & v2 & ")" & vbCrLf
            bloque = bloque & "        If " & v2 & " Mod 10 = 0 Then Exit Do" & vbCrLf
            bloque = bloque & "    Loop" & vbCrLf
            bloque = bloque & "End Sub"
        
        Case 5
            bloque = "Private Function " & v1 & "() As String" & vbCrLf
            bloque = bloque & "    Dim " & v2 & " As Long" & vbCrLf
            bloque = bloque & "    Dim " & v3 & " As String" & vbCrLf
            bloque = bloque & "    On Error Resume Next" & vbCrLf
            bloque = bloque & "    ' En este punto validamos el enlace contable del release" & vbCrLf
            bloque = bloque & "    " & v2 & " = " & n1 & vbCrLf
            bloque = bloque & "    " & v3 & " = ""ENL-"" & Format(" & v2 & ", ""00000"")" & vbCrLf
            bloque = bloque & "    If Len(" & v3 & ") = 0 Then " & v3 & " = ""ENL-00000""" & vbCrLf
            bloque = bloque & "    On Error GoTo 0" & vbCrLf
            bloque = bloque & "    " & v1 & " = " & v3 & vbCrLf
            bloque = bloque & "End Function"
        
    End Select
    
    GenerarBloqueRuido = bloque
End Function

' ============================================================================
' NOMBRES OPACOS para código muerto (no necesitan seguir convención leet)
' ============================================================================
Private Function NombreOpaco() As String
    Dim n As String, i As Integer
    Const ALFANUM As String = "abcdefghijklmnopqrstuvwxyz0123456789"
    
    n = Chr(Int(26 * Rnd + 97)) & Chr(Int(26 * Rnd + 97)) & Chr(Int(26 * Rnd + 97))
    For i = 1 To 10
        n = n & Mid(ALFANUM, Int(36 * Rnd + 1), 1)
    Next i
    
    NombreOpaco = n
End Function
