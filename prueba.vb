Option Compare Database
Option Explicit

Public VEp1, SQL, SQL1, SQL2 As String
Public VEp2 As String
Public VEp3 As String
Public VEp4 As String
Public Cmb As New Collection
Public Txt As New Collection
Public CODIGO, Piel, Ep As String
Public Vk As Double
Public OpNuevo As Boolean

Public Function Proximonumero(Dtll As String) As String
Dim Srt As Integer
Dim Cdg1 As String
Dim Vc9 As Variant
Vc9 = DLookup("Serie", "Consecutivos", "Prefijo='" & Dtll & "'")
If IsNull(Vc9) Then
    MsgBox "Este consecutivo no está creado" & Chr(13) & "Verifique en Creación de Proyecto", vbCritical
    Exit Function
Else
    Srt = Vc9 + 1
    Cdg1 = Dtll & LTrim(Format(Srt, "0000"))
    Proximonumero = Cdg1
End If
End Function