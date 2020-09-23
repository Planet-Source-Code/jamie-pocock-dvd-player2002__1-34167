VERSION 5.00
Begin VB.Form FrmParentalLevels 
   BackColor       =   &H00808000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parental Levels Enabled"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   Icon            =   "FrmParentalLevels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1170
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Password"
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   1770
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   1770
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1170
      TabIndex        =   0
      Top             =   990
      Width           =   2625
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   $"FrmParentalLevels.frx":030A
      ForeColor       =   &H00E0E0E0&
      Height          =   825
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "New Password"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   30
      TabIndex        =   5
      Top             =   1350
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Password"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1020
      Width           =   690
   End
End
Attribute VB_Name = "FrmParentalLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'frmDocument.DVD.DVDAdm.ChangePassword '(sUserName, sOld, sNew)
'frmDocument.DVD.DVDAdm.SaveParentalLevel '(iLevel, sUserName, sPassword)
'frmDocument.DVD.SelectParentalCountry
'frmDocument.DVD.DVDAdm.SaveParentalCountry
'frmDocument.DVD.DVDAdm.ConfirmPassword '(sUserName, sPassword)

If Text1.Visible = True Then

    If Len(Text1) < 1 Then
        MsgBox "The password can not be changed until a new password has been entered", vbOKOnly, "Note!"
        Text1.SetFocus
        Exit Sub
    End If

    If frmDocument.DVD.DVDAdm.ConfirmPassword("", Text2.Text) = True Then
            SaveSetting App.Title, "Settings", "ParentalLevel", frmDocument.DVD.DVDAdm.ChangePassword("", Text2.Text, Text1.Text)
            MsgBox "Password changed", vbOKOnly, "Done!"
            'FrmParentalLevels.Visible = False
            frmOptions.ParentialUserAndPasswordOk
            SaveSetting App.Title, "Settings", "hack", "0"
            frmOptions.Text2.Text = FrmParentalLevels.Text2.Text
            Unload FrmParentalLevels
            Exit Sub
    End If
End If


If frmDocument.DVD.DVDAdm.ConfirmPassword("", Text2.Text) = True Then
FrmParentalLevels.Visible = False
frmOptions.ParentialUserAndPasswordOk
frmOptions.Text2.Text = FrmParentalLevels.Text2.Text
SaveSetting App.Title, "Settings", "hack", "0"
Unload FrmParentalLevels
Else
If Len(GetSetting(App.Title, "Settings", "hack")) = 0 Then
SaveSetting App.Title, "Settings", "hack", 1
Else
a = GetSetting(App.Title, "Settings", "hack") + 1
SaveSetting App.Title, "Settings", "hack", a
End If

If a = 1 Then ab = a & " st"
If a = 2 Then ab = a & " nd"
If a = 3 Then ab = a & " rd"
If a >= 4 Then ab = a & " th"

MsgBox "Wrong Password the application has logged the failed attempt, when the correct password is entered the log will be set to 0", vbOKOnly, "Warning " & ab & " attempt to change password"
End If

End Sub

Private Sub Command2_Click()
FrmParentalLevels.Refresh
Text1.Visible = True
Label1.Visible = True
Label2.Caption = "Old password"
Text2.SetFocus

End Sub
