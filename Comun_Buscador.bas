Attribute VB_Name = "Comun_Buscador"
Option Explicit

' ============================================================================
' FUNCIÓN DE BÚSQUEDA/TRADUCCIÓN (VERSIÓN DESARROLLO)
' Esta versión NO se usa en archivos ofuscados
' ============================================================================
Public Function f_tr(ByVal s As String) As String
    On Error Resume Next
    Dim v As Variant, i As Long, r As String
    
    r = ""
    If s = "" Then Exit Function
    
    v = Split(s, ",")
    For i = LBound(v) To UBound(v)
        ' Desplazamiento de -7 para revertir la ofuscación simple
        ' NOTA: En archivos ofuscados se usa XOR con clave dinámica
        r = r & Chr(CInt(v(i)) - 7)
    Next i
    
    f_tr = r
End Function
