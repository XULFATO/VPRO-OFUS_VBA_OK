Attribute VB_Name = "mod_Garbage"
Option Explicit

' -----------------------------------------------------------------------
' InyectarRuido
' Añade un comentario falso al final de la linea con 15% de probabilidad.
' Se llama sobre la linea completa ya reconstruida, nunca sobre fragmentos.
' No actua sobre lineas de cierre (End Sub, End Function, etc.)
' -----------------------------------------------------------------------
Public Function InyectarRuido(ByVal linea As String) As String
    Dim lineaTrim As String
    lineaTrim = Trim(linea)

    ' No saturar lineas de cierre de procedimiento
    If lineaTrim Like "End Sub*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Function*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Property*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End If*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End With*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Select*" Then InyectarRuido = linea: Exit Function
    If lineaTrim Like "End Type*" Then InyectarRuido = linea: Exit Function
    If lineaTrim = "" Then InyectarRuido = linea: Exit Function

    ' No inyectar si la linea ya tiene comentario o es continuacion
    If InStr(linea, "'") > 0 Then InyectarRuido = linea: Exit Function
    If Right(Trim(linea), 1) = "_" Then InyectarRuido = linea: Exit Function

    ' 15% de probabilidad (Randomize ya fue llamado una vez en mod_Main)
    If Rnd > 0.85 Then
        InyectarRuido = linea & " " & GenerarComentario()
    Else
        InyectarRuido = linea
    End If
End Function

' -----------------------------------------------------------------------
' GenerarComentario
' Devuelve un comentario falso con aspecto de logica de negocio narrativa.
' -----------------------------------------------------------------------
Private Function GenerarComentario() As String
    Dim acciones(9) As String
    Dim conceptos(9) As String
    Dim complementos(9) As String

    ' Acciones narrativas
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

    ' Conceptos de negocio
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

    ' Complementos
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

    Dim idx1 As Integer
    Dim idx2 As Integer
    Dim idx3 As Integer
    idx1 = Int(Rnd * 10)
    idx2 = Int(Rnd * 10)
    idx3 = Int(Rnd * 10)

    GenerarComentario = "' " & acciones(idx1) & " " & conceptos(idx2) & " " & complementos(idx3)
End Function

' -----------------------------------------------------------------------
' GenerarLineaRuido
' Genera una linea completa de codigo muerto para insertar entre
' procedimientos (constantes privadas que nunca se usan).
' -----------------------------------------------------------------------
Public Function GenerarLineaRuido() As String
    Dim sufijo As String
    Dim i As Integer

    ' Nombre opaco de 10 chars
    sufijo = Chr(Int(26 * Rnd + 97))
    For i = 1 To 9
        sufijo = sufijo & Mid("abcdefghijklmnopqrstuvwxyz0123456789", _
                              Int(36 * Rnd + 1), 1)
    Next i

    Select Case Int(Rnd * 3)
        Case 0
            GenerarLineaRuido = "Private Const " & sufijo & " As Long = " & _
                                 CStr(Int(Rnd * 99999) + 1)
        Case 1
            GenerarLineaRuido = "Private Const " & sufijo & " As Double = " & _
                                 Format(Rnd * 9999, "0.0000")
        Case 2
            GenerarLineaRuido = "Private Const " & sufijo & " As String = " & _
                                 Chr(34) & Hex(CLng(Rnd * &HFFFF&)) & Chr(34)
    End Select
End Function


