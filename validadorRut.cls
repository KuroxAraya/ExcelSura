VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

If Not Intersect(Target, Columns("C")) Is Nothing Then
    
    rute = Target.Value
    
    arut = UCase(rute)
    rut = Replace("0000" & rut, ".", "", 1)
    If InStr(1, rut, "-") > 0 Then rut = Left(rute, InStr(1, rute, "-") - 1)
    rut = Right(rute, 8)
    suma = 0
    For i = 1 To 8
    suma = suma + Val(Mid(rute, i, 1)) * Val(Mid("32765432", i, 1))
    Next i
    dv = 11 - (suma Mod 11)
    If dv = 10 Then dv = "K"
    If dv = 11 Then dv = 0
    
    'If Right(arut, 1) = CStr(dv) Then MsgBox ("Rut " & rute & " Correcto")
    If Right(arut, 1) <> CStr(dv) Then MsgBox ("Rut " & rute & " Incorrecto")
    If rute = "" Then MsgBox ("Campo borrado")
    
End If
End Sub
