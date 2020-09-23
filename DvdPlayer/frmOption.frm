VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DVD 202 Options"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      Height          =   3840
      Index           =   1
      Left            =   60
      Picture         =   "frmOption.frx":030A
      ScaleHeight     =   3780
      ScaleWidth      =   7230
      TabIndex        =   2
      Top             =   360
      Width           =   7290
      Begin VB.Frame Frame1 
         Caption         =   "DVD Screen behavier"
         Height          =   855
         Left            =   4080
         TabIndex        =   40
         Top             =   1320
         Width           =   3135
         Begin VB.OptionButton OptAspectRatioOff 
            Caption         =   "Free Form"
            Height          =   225
            Left            =   150
            TabIndex        =   42
            Top             =   510
            Width           =   1545
         End
         Begin VB.OptionButton OptAspectRatioOn 
            Caption         =   "Keep Aspect Ratio"
            Height          =   225
            Left            =   150
            TabIndex        =   41
            Top             =   270
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.Image Image4 
            Height          =   705
            Left            =   30
            Top             =   120
            Width           =   3045
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Message Box Tips"
         Height          =   855
         Left            =   4080
         TabIndex        =   34
         Top             =   2490
         Width           =   3135
         Begin VB.OptionButton Optionmesagetipson 
            Caption         =   "On"
            Height          =   225
            Left            =   150
            TabIndex        =   36
            Top             =   270
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton Optionmesagetipsoff 
            Caption         =   "Off"
            Height          =   225
            Left            =   150
            TabIndex        =   35
            Top             =   510
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   705
            Left            =   30
            Top             =   120
            Width           =   3045
         End
         Begin VB.Label Label3 
            Caption         =   "Turn on/of non critical user messages"
            Height          =   465
            Left            =   1050
            TabIndex        =   37
            Top             =   300
            Width           =   1905
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "DVD control ""Auto relocate the Navigator on resize"
         Height          =   1155
         Left            =   30
         TabIndex        =   30
         Top             =   1320
         Width           =   4035
         Begin VB.CheckBox Check1 
            Caption         =   "Do not auto align after resize events"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   33
            Top             =   810
            Width           =   2985
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Always align to centa, bottom of parent window"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   32
            Top             =   540
            Width           =   3765
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Snap to centa, bottom of DVD screen"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   31
            Top             =   270
            Width           =   3045
         End
         Begin VB.Image Image2 
            Height          =   1125
            Left            =   0
            Top             =   0
            Width           =   4005
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "DVD control ""Search speeds"""
         Height          =   1245
         Left            =   30
         TabIndex        =   26
         Top             =   2490
         Width           =   4035
         Begin VB.CheckBox Check1 
            Caption         =   "Forward and rewind at 2 times play speed"
            Height          =   285
            Index           =   3
            Left            =   150
            TabIndex        =   29
            Top             =   270
            Width           =   3315
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Forward and rewind at 4 times play speed"
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   28
            Top             =   570
            Width           =   3315
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Forward and rewind at 8 times play speed"
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   27
            Top             =   870
            Width           =   3255
         End
         Begin VB.Image Image3 
            Height          =   1125
            Left            =   0
            Top             =   120
            Width           =   4005
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   1170
         TabIndex        =   38
         Top             =   120
         Width           =   4875
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      Height          =   3090
      Index           =   0
      Left            =   60
      Picture         =   "frmOption.frx":2717
      ScaleHeight     =   3030
      ScaleWidth      =   4230
      TabIndex        =   1
      Top             =   360
      Width           =   4290
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000A&
         Caption         =   "Parental Level"
         Height          =   2445
         Left            =   60
         TabIndex        =   10
         Top             =   1290
         Width           =   7125
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "Password for all"
            Enabled         =   0   'False
            Height          =   255
            Index           =   9
            Left            =   2250
            TabIndex        =   39
            Top             =   1650
            Width           =   1545
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Failed password record"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   2070
            Width           =   2385
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Password"
            Height          =   375
            Left            =   5790
            TabIndex        =   20
            Top             =   1980
            Width           =   1245
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "1 Suitable for children"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1845
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "2"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   630
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "3"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   900
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "4"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   1170
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "5"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   2250
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "6"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   2250
            TabIndex        =   14
            Top             =   630
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "7"
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   2250
            TabIndex        =   13
            Top             =   900
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "8 (18 years +)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   2250
            TabIndex        =   12
            Top             =   1170
            Width           =   1305
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H8000000A&
            Caption         =   "No Password"
            Enabled         =   0   'False
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   1650
            Width           =   1245
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   22
            Top             =   2100
            Width           =   2205
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOption.frx":4B24
            ForeColor       =   &H00000000&
            Height          =   1305
            Index           =   1
            Left            =   3660
            TabIndex        =   25
            Top             =   390
            Width           =   3345
         End
         Begin VB.Label Label2 
            Caption         =   "Password"
            Height          =   195
            Left            =   510
            TabIndex        =   23
            Top             =   2100
            Width           =   1245
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Some DVDs may be of an adult nature containing swear words, violance, drugs, sex or nudity.  "
         ForeColor       =   &H8000000B&
         Height          =   465
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         Top             =   330
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6210
      TabIndex        =   9
      Tag             =   "1055"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4920
      TabIndex        =   8
      Tag             =   "1056"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3690
      TabIndex        =   7
      Tag             =   "1057"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3015
      Index           =   5
      Left            =   900
      ScaleHeight     =   2955
      ScaleWidth      =   3390
      TabIndex        =   6
      Top             =   1050
      Width           =   3450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3030
      Index           =   4
      Left            =   600
      ScaleHeight     =   2970
      ScaleWidth      =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   3420
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   2790
      Index           =   3
      Left            =   450
      ScaleHeight     =   2730
      ScaleWidth      =   3210
      TabIndex        =   4
      Top             =   720
      Width           =   3270
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   2730
      Index           =   2
      Left            =   330
      ScaleHeight     =   2670
      ScaleWidth      =   3060
      TabIndex        =   3
      Top             =   630
      Width           =   3120
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   7541
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Perental Levels"
            Key             =   "picture1p"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player Settings"
            Key             =   "picture2p"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 3"
            Key             =   "picture3p"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 4"
            Key             =   "picture4p"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 5"
            Key             =   "picture5p"
            Object.Tag             =   "5"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1p()
End Sub
Private Sub Picture2p()
End Sub
Private Sub Picture3p()
End Sub
Private Sub Picture4p()
End Sub
Private Sub Picture5p()
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Caption = "Enter or change the parental level password "

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Caption = "Keep track of how many times the wrong password has been entered, once the correct password is entered the count resets"
End Sub


Private Sub Form_Load()
TabStrip1.Tabs("picture1p").Selected = True

    If GetSetting(App.Title, "Settings", "msgtips") = "on" Then
    Optionmesagetipson.Value = True
    Else
    Optionmesagetipsoff.Value = True
    End If

If (Val(GetSetting(App.Title, "Settings", "OptAspectRatio"))) = 1 Then
OptAspectRatioOn.Value = True
Else
OptAspectRatioOff.Value = True
End If

Check1(Val(GetSetting(App.Title, "Settings", "NaviSnapFunc"))).Value = 1
Check1(Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed"))).Value = 1
    If frmDocument.DVD.DVDAdm.GetParentalLevel = -1 Then
    Option5(8).Value = True
    Option5(8).Enabled = True
    Else
    Option5(Val(frmDocument.DVD.DVDAdm.GetParentalLevel - 1)).Value = 1
    Option5(Val(frmDocument.DVD.DVDAdm.GetParentalLevel - 1)).Enabled = True
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "When you use certain functions a meassage box may appear that informs you of how certain features work, by selecting the ""off"" check box will prevent these meassages from appearing"
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "When you resize a window the player responds by moving, choose how you wish it to respond here."
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Use these options to select how fast the fastforward and rewind should be."
End Sub

Private Sub TabStrip1_Click()
On Error GoTo 10
Picture1(TabStrip1.SelectedItem.Index - 1).Move TabStrip1.ClientLeft _
, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight

Picture1(TabStrip1.SelectedItem.Index - 1).ZOrder
Select Case TabStrip1.SelectedItem.Index

    Case 1
        Picture1p
    Case 2
        Picture2p
    Case 3
        Picture3p
    Case 4
        Picture4p
    Case 5
        Picture5p
End Select
10: Exit Sub
End Sub
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
Check1(0).Value = 1
Check1(1).Value = 0
Check1(2).Value = 0
Exit Sub
End If
If Index = 1 Then
Check1(0).Value = 0
Check1(1).Value = 1
Check1(2).Value = 0
Exit Sub
End If
If Index = 2 Then
Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 1
Exit Sub
End If
If Index = 3 Then
Check1(3).Value = 1
Check1(4).Value = 0
Check1(5).Value = 0
Exit Sub
End If
If Index = 4 Then
Check1(3).Value = 0
Check1(4).Value = 1
Check1(5).Value = 0
Exit Sub
End If
If Index = 5 Then
Check1(3).Value = 0
Check1(4).Value = 0
Check1(5).Value = 1
Exit Sub
End If
End Sub

Private Sub Command1_Click()
FrmParentalLevels.Show
End Sub

Function ParentialUserAndPasswordOk()

Text2.Enabled = True

Text2.BackColor = vbWhite
    For i = 0 To Option5.Count - 1
    Option5(i).Enabled = True
    Next i
    MsgBox "Select the level you wish to use and click Apply to set it", vbOKOnly, "Parental Level Change"
End Function


Private Sub Command2_Click()
If Val(GetSetting(App.Title, "Settings", "hack")) >= 1 Then
MsgBox "There has been " & Val(GetSetting(App.Title, "Settings", "hack")) & " failed atempts at entering the password", vbOKCancel, "Atempt Count"
Else
MsgBox "There has been no atempts to alter the password", vbOKOnly, "Atempt Count"
End If

End Sub



Private Sub cmdApply_Click()

    If Optionmesagetipson.Value = True Then
    SaveSetting App.Title, "Settings", "msgtips", "on"
    Else
    SaveSetting App.Title, "Settings", "msgtips", "off"
    End If
    For i = 0 To Option5.Count - 2
    If Option5(i).Value = True Then
    SaveSetting App.Title, "Settings", "ParentalLevel", frmDocument.DVD.DVDAdm.SaveParentalLevel(Option5(i + 1).Index, "", Text2.Text)
    End If
    Next i
    If Option5(8).Value = True Then
    SaveSetting App.Title, "Settings", "ParentalLevel", frmDocument.DVD.DVDAdm.SaveParentalLevel(-1, "", Text2.Text)
    End If
    
If Check1(0).Value = 1 Then SaveSetting App.Title, "Settings", "NaviSnapFunc", "0"
If Check1(1).Value = 1 Then SaveSetting App.Title, "Settings", "NaviSnapFunc", "1"
If Check1(2).Value = 1 Then SaveSetting App.Title, "Settings", "NaviSnapFunc", "2"
If Check1(3).Value = 1 Then SaveSetting App.Title, "Settings", "NaviSearchSpeed", "3" '2
If Check1(4).Value = 1 Then SaveSetting App.Title, "Settings", "NaviSearchSpeed", "4" '4
If Check1(5).Value = 1 Then SaveSetting App.Title, "Settings", "NaviSearchSpeed", "5" '8
If OptAspectRatioOn.Value = True Then
SaveSetting App.Title, "Settings", "OptAspectRatio", "1" '8
Else
SaveSetting App.Title, "Settings", "OptAspectRatio", "0" '8
End If
    For i = 0 To Option5.Count - 1
    If Option5(i).Value = False Then
    Option5(i).Enabled = False
    End If
    Next i
'MSWebDVD.NotifyParentalLevelChange(bNotify)                    'en/disables the event handling for temporary parental management level commands.
'MSWebDVD.AcceptParentalLevelChange(bAccept)                    'accepts or rejects the new temporary parental management level.
'MSWebDVD.SelectParentalCountry(iCountry, sUserName, sPassword) 'sets the specified parental country/region for subsequent playback
'MSWebDVD.SelectParentalLevel(iLevel, sUserName, sPassword)     'sets the specified parental level for subsequent playback
'DVD.DVDAdm.ConfirmPassword(sUserName, sPassword)               'tests whether the specified password matches the previously saved password
'DVD.DVDAdm.SaveParentalCountry( iCountry, sUserName, sPassword)'saves the application's new parental country/region to the registry
'DVD.DVDAdm.SaveParentalLevel(iLevel, sUserName, sPassword)     'saves a new default parental level to the registry.
'DVD.DVDAdm.ChangePassword(sUserName, sOld, sNew)               'saves a new application password in the registry
'[ iCountryCode = ] MSWebDVD.GetPlayerParentalCountry()         'retrieves the current country/region as set in the MSWebDVD object.
'[ iLevels = ] MSWebDVD.GetTitleParentalLevels(iTitle)          'retrieves the parental management levels for the specified title
'[ iParentalLevel = ] DVD.DVDAdm.GetParentalLevel();            'retrieves the parental level that was last saved to the registry
'[ iLevel = ] MSWebDVD.GetPlayerParentalLevel()                 'determine if level is greater than level set by SelectParentalLevel
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub




Private Sub Option5_Click(Index As Integer)

'frmDocument.DVD.SelectParentalLevel(7, FrmParentalLevels.Text1, FrmParentalLevels.Text2)

End Sub



Function Subtitles(stream)
  Select Case Switch(stream)
    Case "0"
      DVD.CurrentSubpictureStream = 0
      DVD.SubpictureOn = 1
    Case "1"
      DVD.CurrentSubpictureStream = 1
      DVD.SubpictureOn = 1
    Case "2"
      DVD.CurrentSubpictureStream = 2
      DVD.SubpictureOn = 1
    Case "3"
      DVD.CurrentSubpictureStream = 3
      DVD.SubpictureOn = 1
    Case "4"
      DVD.CurrentSubpictureStream = 4
      DVD.SubpictureOn = 1
    Case "5"
      DVD.CurrentSubpictureStream = 5
      DVD.SubpictureOn = 1
    Case "6"
      DVD.CurrentSubpictureStream = 6
      DVD.SubpictureOn = 1
    Case "7"
      DVD.CurrentSubpictureStream = 7
      DVD.SubpictureOn = 1
    Case "8"
      DVD.SubpictureOn = 0
End Select
End Function
Function Audio(stream)
  Select Case Switch(stream)
    Case "0"
      DVD.CurrentAudioStream = 0
    Case "1"
      DVD.CurrentAudioStream = 1
    Case "2"
      DVD.CurrentAudioStream = 2
    Case "3"
      DVD.CurrentAudioStream = 3
    Case "4"
      DVD.CurrentAudioStream = 4
    Case "5"
      DVD.CurrentAudioStream = 5
    Case "6"
      DVD.CurrentAudioStream = 6
    End Select
End Function


