VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "There are no options to be Set!"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2925
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // This project dosn't need Options, but i decided to include a options form anyway, so ppl
Rem // get the general idea of how the options screen is called etc...

Rem // Declare Functions Needed
Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub cmdAbout_Click()
On Error Resume Next
Rem // Show the About Screen
frmAbout.Show
Unload Me
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
Rem // Close the Options Form
Unload frmAbout
Unload Me
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Rem // Disable the "X" Button, but Keep the Icon
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

Rem // Set this form on top of The Rest Of Windows
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

