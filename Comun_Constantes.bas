Attribute VB_Name = "Comun_Constantes"
Option Explicit

' ============================================================================
' VARIABLES PÚBLICAS DEL OFUSCADOR
' ============================================================================
Public HOJA_HOME As String
Public HOJA_ESP As String

' ============================================================================
' FUNCIÓN DE TRADUCCIÓN (VERSIÓN PARA DESARROLLO)
' Esta versión NO se usa en archivos ofuscados
' El archivo ofuscado recibe su propia versión con clave embebida
' ============================================================================
Public Function f_tr(ByVal s As String) As String
    On Error Resume Next
    Dim v As Variant, i As Long, r As String
    
    If s = "" Then Exit Function
    
    v = Split(s, ",")
    For i = LBound(v) To UBound(v)
        ' Desplazamiento simple para desarrollo (NO SE USA EN PRODUCCIÓN)
        r = r & Chr(CInt(v(i)) - 7)
    Next i
    
    f_tr = r
End Function
