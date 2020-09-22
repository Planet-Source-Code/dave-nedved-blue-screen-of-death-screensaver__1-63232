VERSION 5.00
Begin VB.Form frmScreen 
   BackColor       =   &H00840000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11910
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmScreen.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMessage1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   3120
   End
   Begin VB.Timer tmrErrAdd 
      Interval        =   700
      Left            =   3000
      Top             =   3120
   End
   Begin VB.Label lblErrorCount 
      BackColor       =   &H00840000&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   885
   End
   Begin VB.Label lblError 
      AutoSize        =   -1  'True
      BackColor       =   &H00840000&
      Caption         =   "*** STOP: 0x00000019 (0xC00E0FF0, 0xC00E0FF0, 0xC0000000)"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6285
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // ------------------------------------------------------------------------------
Rem // | I hope you enjoy this little Project Screensaver                           |
Rem // | it has lessons like how to Hide the mouse etc...                           |
Rem // | I hope you learn from this, and all the little features                    |
Rem // | That i did not need to add, but i did so people can learn                  |
Rem // |                                                                            |
Rem // | The Blue Screen of Death is by David Nedved                                |
Rem // | em. dnedved@datosoftware.com                                               |
Rem // | ws. www.datosoftware.com                                                   |
Rem // |                                                                            |
Rem // | Please leave your votes and comments for this little screensaver           |
Rem // | As i sat up late into the night programming this screensaver beacuse       |
Rem // | I was half bored, and half wanted to shop ppl how to use                   |
Rem // | Basic Screensaver Functions.                                               |
Rem // |                                                                            |
Rem // | Feel free to Mod, Package, Deploy ... the screensaver                      |
Rem // | Just give credit to me, and link to the website.                           |
Rem // |                                                                            |
Rem // | Have Fun!!!                                                                |
Rem // ------------------------------------------------------------------------------




Rem // Find function in user32 to make the screensaver on top of all other windows

Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

 Rem // Dim the functions needed within the Project
 Rem // I am a lil bad, i used baby Commands like i, ii, mm, ct.. this is beacuse it was currently 1:30am and i a but tired... and / or couldnt be bothered to do it properly :)
 Dim i As Double
 Dim quick As Boolean
 Dim TextLeft As Boolean
 Dim TEXTTOP As Boolean
 Dim ct As Integer
 Dim MM As Boolean
 Dim ii As Integer

Private Sub Form_Activate()
On Error Resume Next
Rem // Set the form to be always on top, when ever the form is activated
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_Click()
On Error Resume Next
Rem // Exit Screensaver
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Rem // Exit Screensaver
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Rem // Set up the screensaver Pos, Text, Timers ...

ii = 1


Me.lblError.Caption = "*** STOP: 0x00000019 (0xC00E0FF0, 0xC00E0FF0, 0xC0000000)" & vbNewLine & "SYSTEM_STOR_ERROR" & vbNewLine & vbNewLine & "SYSVER 0x0000568" & vbNewLine & vbNewLine & "Address dword dump Build [1381]" & vbNewLine & vbNewLine & _
vbNewLine & App.hInstance & " - " & App.hInstance / 0.2 & " - System Stop Error_0xFFFE2D60" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x0000FHG0" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x00C34GH9" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x00C34GV9" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0xG0C34GV3" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x0000JJ12" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0xDFF1200H" & _
vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0xFEC32DB0"

Me.Show

DoEvents
lblErrorCount.Width = Me.Width - 120 - 120
DoEvents
Me.Refresh
lblErrorCount.Height = Me.Height - Me.lblErrorCount.Top

Me.lblError.FontSize = 17
Me.lblErrorCount.FontSize = 17
frmBlack.Show
Me.Show

lblErrorCount.Top = Me.lblError.Top + Me.lblError.Height + 120
frmBlack.Enabled = False
CursorHide

Me.tmrErrAdd.Enabled = False
Me.tmrMessage1.Enabled = True
ii = 25
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Rem // When the mouse is moved it needs to be moved for 10ms otherwise the screensaver will move
Rem // this is caled 'mouse sensativity' if thats how you spell it.
Rem // [P.S. I & most programmers i no cant spell if there life depended on it :):)]
Static ct As Integer
If ct > 10 Then
Unload Me
Else
ct = ct + 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Rem // When the form  is unloaded the mouse should be shown again
CursorShow
End
End Sub

Private Sub lblError_Click()
On Error Resume Next
Rem // this is so i dont have to double code, when you click this i will refer to the code under 'Form_Click'
Form_Click
End Sub

Private Sub lblError_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Rem // This is so i do not have to double code, when the mouse is moved it will be refered to the code under 'Form_MouseMove'
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblErrorCount_Click()
On Error Resume Next
Rem // this is so i dont have to double code, when you click this i will refer to the code under 'Form_Click'
Form_Click
End Sub

Private Sub lblErrorCount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Rem // This is so i do not have to double code, when the mouse is moved it will be refered to the code under 'Form_MouseMove'
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub tmrErrAdd_Timer()
Rem // this timer will create a fake memory dump.

On Error Resume Next
If Not Me.lblErrorCount.Caption = "" Then Me.lblErrorCount.Caption = Me.lblErrorCount.Caption & "0xFFD" & App.hInstance * ii & ", "
If Me.lblErrorCount.Caption = "" Then Me.lblErrorCount.Caption = "0xFFD" & App.hInstance * ii & ", "

ii = ii + 1
If ii = 40 Then
 Me.lblError.Caption = ""
 Me.lblErrorCount.Caption = ""
 Me.tmrErrAdd.Enabled = False
 Me.tmrMessage1.Enabled = True
 Me.lblError.AutoSize = False
 Me.lblError.Width = Me.ScaleWidth - 120 - 120
 Me.lblError.Height = Me.ScaleHeight - 120 - 120
 Me.lblErrorCount.Visible = False
 ii = 0
End If
End Sub

Private Sub tmrMessage1_Timer()
On Error Resume Next
Rem // This timer is used to create the little message,
Rem // i could go on, but it is getting late, and i made this screensaver
Rem // cause i was bord with all the other windows screensaver

ii = ii + 1
If ii = 1 Then Me.lblError.Caption = "A Problem has been detected and this life has been shut down to prevent damage to this soul." & vbNewLine & vbNewLine
If ii = 4 Then Me.lblError.Caption = Me.lblError.Caption & "BRAIN_ANEURYSM" & vbNewLine & vbNewLine
If ii = 5 Then Me.lblError.Caption = Me.lblError.Caption & "If this is the first time you've seen this Life-ending screen, it will probably be the last. If this screen appears again follow these steps:" & vbNewLine & vbNewLine
If ii = 8 Then Me.lblError.Caption = Me.lblError.Caption & "Check to make sure any organs or memories are properly installed. If this is a new installation, This Life might have needed additional gestation time." & vbNewLine & vbNewLine
If ii = 11 Then Me.lblError.Caption = Me.lblError.Caption & "If Problems continue, disable or remove any newly installed organs or memories. Disable life-extending optiors such as a respirator or a feeding tube. If you need to use Genetic Engineering Mode to remove or disable components, restart this Life, Select IVF Options, and then select Genetic Engineering." & vbNewLine & vbNewLine
If ii = 16 Then Me.lblError.Caption = Me.lblError.Caption & "Technical information:" & vbNewLine & "*** STOP: 0x0000004e (0x00000099, 0x00000000, 0x00000000)" & vbNewLine & vbNewLine
If ii = 18 Then Me.lblError.Caption = Me.lblError.Caption & "Beginning dump of memories." & vbNewLine & "Physical memory dump complete" & vbNewLine & "Contact your general practitioner, advisor, or professional psychic for futher assistance."
If ii = 25 Then
 Me.lblError.Caption = ""
 Me.lblErrorCount.Caption = ""
 Me.lblError.AutoSize = True
 Me.lblErrorCount.Visible = True
End If

If ii = 26 Then
 Me.lblError.Caption = "*** STOP: 0x00000019 (0xC00E0FF0, 0xC00E0FF0, 0xC0000000)" & vbNewLine & "SYSTEM_STOR_ERROR" & vbNewLine & vbNewLine & "SYSVER 0x0000568" & vbNewLine & vbNewLine & "Address dword dump Build [1381]" & vbNewLine & vbNewLine & _
 vbNewLine & App.hInstance & " - " & App.hInstance / 0.2 & " - System Stop Error_0xFFFE2D60" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x0000FHG0" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x00C34GH9" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x00C34GV9" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0xG0C34GV3" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0x0000JJ12" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0xDFF1200H" & _
 vbNewLine & App.hInstance / 0.8 & " - " & App.hInstance / 0.2 & " - System Stop Error_0xFEC32DB0"

 DoEvents
 lblErrorCount.Width = Me.ScaleWidth - 120 - 120
 DoEvents
 lblErrorCount.Height = Me.ScaleHeight - Me.lblErrorCount.Top

 Me.lblError.FontSize = 17
 Me.lblErrorCount.FontSize = 17

 lblErrorCount.Top = Me.lblError.Top + Me.lblError.Height + 120
 
 Me.tmrMessage1.Enabled = False
 Me.tmrErrAdd.Enabled = True
End If
End Sub

