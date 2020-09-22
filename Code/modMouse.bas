Attribute VB_Name = "modMouse"
Rem // This code will Hide and Show the Mouse
Rem // Go to the user32, and find the function to show and hide the mouse
Option Explicit
 Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Function CursorShow()
    On Error Resume Next
    Rem // Show the mouse
    ShowCursor True
End Function

Function CursorHide()
    On Error Resume Next
    Rem // Hide the mouse
    ShowCursor False
End Function

