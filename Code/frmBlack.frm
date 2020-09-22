VERSION 5.00
Begin VB.Form frmBlack 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmBlack.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmBlack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // This form is only a black background, so you cant see the desktop.
Rem // Declare function needed to set window on top of other Windows
Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
On Error Resume Next
Rem // set the Form's Position to 0, 0 and make the form Verry Large
Rem // I do this incase the person is running a dual screen, like me.. it gets anoying
Rem // When screensavers only work on one screen, or where the other scren is normal ;)
Me.Top = 0
Me.Left = 0
Me.Width = 99999999
Me.Height = 999999999
Rem // Set this windows on top of all other Windows
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub
