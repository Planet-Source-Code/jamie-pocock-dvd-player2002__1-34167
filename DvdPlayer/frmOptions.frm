VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Tag             =   "1050"
   Begin VB.PictureBox picOption 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   3795
      Index           =   1
      Left            =   90
      ScaleHeight     =   3856.209
      ScaleMode       =   0  'User
      ScaleWidth      =   6109.48
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   6045
      Begin VB.Frame Frame1 
         Caption         =   "1"
         Height          =   3645
         Left            =   60
         TabIndex        =   15
         Tag             =   "1051"
         Top             =   90
         Width           =   5940
         Begin VB.Frame Frame7 
            Caption         =   "DVD control ""Search speeds"""
            Height          =   1245
            Left            =   300
            TabIndex        =   42
            Top             =   2250
            Width           =   5295
            Begin VB.CheckBox Check1 
               Caption         =   "Forward and rewind at 8 times play speed"
               Height          =   255
               Index           =   5
               Left            =   150
               TabIndex        =   45
               Top             =   870
               Width           =   4905
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Forward and rewind at 4 times play speed"
               Height          =   255
               Index           =   4
               Left            =   150
               TabIndex        =   44
               Top             =   570
               Width           =   4905
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Forward and rewind at 2 times play speed"
               Height          =   285
               Index           =   3
               Left            =   150
               TabIndex        =   43
               Top             =   270
               Width           =   3315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "DVD control ""Auto relocate the Navigator on resize"
            Height          =   1155
            Left            =   300
            TabIndex        =   37
            Top             =   990
            Width           =   5295
            Begin VB.CheckBox Check1 
               Caption         =   "Snap to centa, bottom of DVD screen"
               Height          =   255
               Index           =   2
               Left            =   150
               TabIndex        =   40
               Top             =   270
               Width           =   3045
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Always align to centa, bottom of parent window"
               Height          =   255
               Index           =   1
               Left            =   150
               TabIndex        =   39
               Top             =   540
               Width           =   4905
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Do not auto align after resize events"
               Height          =   255
               Index           =   0
               Left            =   150
               TabIndex        =   38
               Top             =   810
               Width           =   4125
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Message Box Tips"
            Height          =   645
            Left            =   300
            TabIndex        =   20
            Top             =   240
            Width           =   5325
            Begin VB.OptionButton Optionmesagetipsoff 
               Caption         =   "Off"
               Height          =   225
               Left            =   870
               TabIndex        =   22
               Top             =   270
               Width           =   585
            End
            Begin VB.OptionButton Optionmesagetipson 
               Caption         =   "On"
               Height          =   225
               Left            =   150
               TabIndex        =   21
               Top             =   270
               Value           =   -1  'True
               Width           =   555
            End
            Begin VB.Label Label3 
               Caption         =   "This will turn of non critical user messages, such as the prompt when using the Zoom feature."
               Height          =   435
               Left            =   1650
               TabIndex        =   41
               Top             =   150
               Width           =   3555
            End
         End
      End
   End
   Begin VB.PictureBox picOption 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   2520
      Index           =   2
      Left            =   6810
      ScaleHeight     =   2560.645
      ScaleMode       =   0  'User
      ScaleWidth      =   1713.08
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1890
      Width           =   1695
      Begin VB.Frame Frame2 
         Caption         =   "2"
         Height          =   3645
         Left            =   30
         TabIndex        =   17
         Tag             =   "1051"
         Top             =   60
         Width           =   5580
      End
   End
   Begin VB.PictureBox picOption 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   3810
      Index           =   0
      Left            =   150
      ScaleHeight     =   3871.452
      ScaleMode       =   0  'User
      ScaleWidth      =   6048.84
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5985
      Begin VB.Frame fraSample1 
         Caption         =   "Perental control"
         Height          =   3705
         Left            =   60
         TabIndex        =   4
         Tag             =   "1051"
         Top             =   60
         Width           =   5880
         Begin VB.Frame Frame5 
            Caption         =   "Parental Level"
            Height          =   1965
            Left            =   150
            TabIndex        =   23
            Top             =   1590
            Width           =   5295
            Begin VB.CommandButton Command2 
               Caption         =   "Failed password record"
               Height          =   285
               Left            =   2790
               TabIndex        =   36
               Top             =   1590
               Width           =   2385
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H80000004&
               Enabled         =   0   'False
               Height          =   255
               Left            =   2820
               TabIndex        =   34
               Top             =   1620
               Width           =   2205
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Password"
               Height          =   375
               Left            =   3900
               TabIndex        =   33
               Top             =   180
               Width           =   1245
            End
            Begin VB.OptionButton Option5 
               Caption         =   "1 Suitable for children"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   1845
            End
            Begin VB.OptionButton Option5 
               Caption         =   "2"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   31
               Top             =   510
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "3"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   30
               Top             =   780
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "4"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   29
               Top             =   1050
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "5"
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   2280
               TabIndex        =   28
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "6"
               Enabled         =   0   'False
               Height          =   255
               Index           =   5
               Left            =   2280
               TabIndex        =   27
               Top             =   510
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "7"
               Enabled         =   0   'False
               Height          =   255
               Index           =   6
               Left            =   2280
               TabIndex        =   26
               Top             =   780
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "8 (18 years +)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   7
               Left            =   2280
               TabIndex        =   25
               Top             =   1050
               Width           =   1305
            End
            Begin VB.OptionButton Option5 
               Caption         =   "No Password"
               Enabled         =   0   'False
               Height          =   255
               Index           =   8
               Left            =   3900
               TabIndex        =   24
               Top             =   1050
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Password"
               Height          =   195
               Left            =   2970
               TabIndex        =   35
               Top             =   1650
               Width           =   1245
            End
         End
         Begin VB.Label Label1 
            Caption         =   $"frmOptions.frx":08CA
            Height          =   825
            Index           =   1
            Left            =   150
            TabIndex        =   13
            Top             =   690
            Width           =   5235
         End
         Begin VB.Label Label1 
            Caption         =   "Some DVDs may be of an adult nature containing swear words, violance, drugs, sex or nudity.  "
            Height          =   615
            Index           =   0
            Left            =   150
            TabIndex        =   12
            Top             =   270
            Width           =   5385
         End
      End
   End
   Begin VB.PictureBox picOption 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1230
      Index           =   3
      Left            =   6750
      ScaleHeight     =   1249.839
      ScaleMode       =   0  'User
      ScaleWidth      =   1227.96
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
      Begin VB.Frame Frame3 
         Caption         =   "3"
         Height          =   3645
         Left            =   60
         TabIndex        =   19
         Tag             =   "1051"
         Top             =   60
         Width           =   5580
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2340
      TabIndex        =   1
      Tag             =   "1057"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3570
      TabIndex        =   3
      Tag             =   "1056"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4830
      TabIndex        =   5
      Tag             =   "1055"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   11
         Tag             =   "1054"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   10
         Tag             =   "1053"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   8
         Tag             =   "1052"
         Top             =   305
         Width           =   2033
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parental Control"
            Key             =   "picOption1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 2"
            Key             =   "picOption2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 3"
            Key             =   "picOption3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 4"
            Key             =   "picOption4"
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
Private Sub picOption0()
End Sub
Private Sub picOption1()
End Sub
Private Sub picOption2()
End Sub
Private Sub picOption3()
End Sub
Private Sub picOption4()
End Sub
Private Sub picOption5()
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

Private Sub Form_Load()
tbsOptions.Tabs("picOption1").Selected = True
    If GetSetting(App.Title, "Settings", "msgtips") = "on" Then
    Optionmesagetipson.Value = True
    Else
    Optionmesagetipsoff.Value = True
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
       
picOption(tbsOptions.SelectedItem.Index - 1).Move tbsOptions.ClientLeft _
, tbsOptions.ClientTop, tbsOptions.ClientWidth, tbsOptions.ClientHeight

picOption(tbsOptions.SelectedItem.Index - 1).ZOrder
    
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
  
End Sub


Private Sub Option5_Click(Index As Integer)

'frmDocument.DVD.SelectParentalLevel(7, FrmParentalLevels.Text1, FrmParentalLevels.Text2)

End Sub

Private Sub tbsOptions_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOption(i).Left = 210
            picOption(i).Enabled = True

        End If
    Next
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
