Attribute VB_Name = "mod_Types"
Option Explicit

' ============================================================================
' TIPO DE DATO: SEGMENTO
' Representa un fragmento de línea que puede ser código ejecutable o string
' ============================================================================
Public Type Segmento
    contenido As String
    EsCodigo  As Boolean
End Type
