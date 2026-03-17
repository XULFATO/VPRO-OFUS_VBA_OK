Attribute VB_Name = "mod_Traductor"
Option Explicit

' ============================================================================
' FUNCIÓN DE TRADUCCIÓN — VERSIÓN DESARROLLO
' IMPORTANTE: Esta versión NO se usa en archivos ofuscados.
' El ofuscador inyecta dinámicamente mod_Internal_Helper con XOR + clave única.
' Esta función está aquí solo como referencia para pruebas locales.
' ============================================================================

Public Function f_tr(ByVal s As String) As String
    On Error Resume Next
    Dim v As Variant, i As Long, r As String
    
    If s = "" Then Exit Function
    
    r = ""
    v = Split(s, ",")
    
    For i = LBound(v) To UBound(v)
        r = r & Chr(CInt(v(i)) - 7)
    Next i
    
    f_tr = r
End Function
