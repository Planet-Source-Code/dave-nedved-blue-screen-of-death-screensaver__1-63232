Attribute VB_Name = "Resmod"
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Sub ColorTORGB(r As Integer, g As Integer, b As Integer, ByVal c As Integer)
r = 0
g = 0
b = 0
If c < 256 Then
    r = 255
    g = c
ElseIf c < 512 Then
    r = 512 - c
    g = 255
ElseIf c < 768 Then
    g = 255
    b = c - 512
ElseIf c < 1024 Then
    g = 1024 - c
    b = 255
    Else
    r = 255
    b = 1235 - c
    End If
End Sub
