Attribute VB_Name = "MainMode"
Rem // This sub is used as the 'loader' so that the program can choose if
Rem // to display the Screensaver, Options etc...
Rem // This is found out by a Shell Command, that the program picks up.

Public Sub Main()
On Error Resume Next
Rem // If the program is running then exit this Instance of the Program
If App.PrevInstance Then
 End
Rem // If the command is '/c' then show the options screen
ElseIf Left(LCase(Command()), 2) = "/c" Then
 frmOptions.Show
Rem // If the command is '/s' then show the Screensaver Screen
ElseIf Left(LCase(Command()), 2) = "/s" Then
 frmScreen.Show
 frmScreen.WindowState = vbMaximized
Rem // If the command is '/p' then show the screensaver Preview
Rem // (I did not code this in beacuse the Screensaver is set to work within a full screen enviroment)
'ElseIf Left(LCase(Command()), 2) = "/p" Then
 'frmScreen.Show
End If
End Sub
