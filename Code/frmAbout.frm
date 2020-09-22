VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4425
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   8145
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   4425
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // This form is only a black background, so you cant see the desktop.
Rem // Declare function needed to set window on top of other Windows
Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Click()
Rem // Unload the About Form
Unload Me
End Sub

Private Sub Form_Load()
Rem // Set my icon to be the same as the Options Icon
Me.Icon = frmOptions.Icon
Rem // Set this windows on top of all other Windows
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub
