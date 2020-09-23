VERSION 5.00
Begin VB.Form DirectorySort 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectorySort"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   FillStyle       =   0  'Solid
   ForeColor       =   &H00400000&
   Icon            =   "DirectorySort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5175
   Begin VB.CommandButton Command2 
      Caption         =   "Use this folder"
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   2340
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create new folder"
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   2700
      Width           =   1815
   End
   Begin VB.TextBox FileName 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   2370
      Width           =   3075
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00808000&
      ForeColor       =   &H00E0E0E0&
      Height          =   1890
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Your computers directory."
      Top             =   405
      Width           =   3120
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00808000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"DirectorySort.frx":08CA
      ForeColor       =   &H00E0E0E0&
      Height          =   1965
      Left            =   3300
      TabIndex        =   3
      Top             =   180
      Width           =   1725
   End
End
Attribute VB_Name = "DirectorySort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DirectorySort.Command1.Caption = "Exit" Then
DirectorySort.Visible = False
DirectorySort.Cls
Unload DirectorySort
Exit Sub
End If
On Error GoTo errors

    If MsgBox("This will create a folder named DVDDump at the following location : " & Dir1.Path & vbNewLine & "Click ""OK"" to continue or ""Cancel"" to stop", vbOKCancel, "Create Dump Folder") = vbOK Then
        If Len(Dir1.Path) <= 3 Then
                    MkDir Dir1.Path & "DVD Dump"
                    SaveSetting App.Title, "Settings", "Dumpfolder", Dir1.Path & "DVD Dump"
        Else
                    MkDir Dir1.Path & "\DVD Dump"
                    SaveSetting App.Title, "Settings", "Dumpfolder", Dir1.Path & "\DVD Dump"
        End If
                    Dir1.Refresh
                    
    Else
            Exit Sub
    End If
FrmRecord.File2.Path = GetSetting(App.Title, "Settings", "Dumpfolder")
FrmRecord.Label1.Caption = FrmRecord.File2.Path
FrmRecord.File2.Refresh
DirectorySort.Command1.Caption = "Exit"
Exit Sub

errors:
If MsgBox("The dump folder named DVDDump at the following location : " & Dir1.Path & " already exsists " & vbNewLine & "To continue click ""OK"" or click ""Cancel"" to stop", vbOKCancel, "Create Dump Folder") = vbOK Then
        If Len(Dir1.Path) <= 3 Then
                    RmDir Dir1.Path & "DVD Dump"
            Else
                    RmDir Dir1.Path & "\DVD Dump"
        End If
End If
Dir1.Refresh


End Sub

Private Sub Command2_Click()

                    SaveSetting App.Title, "Settings", "Dumpfolder", FileName.Text
                    
FrmRecord.File2.Path = GetSetting(App.Title, "Settings", "Dumpfolder")
FrmRecord.Label1.Caption = FrmRecord.File2.Path
FrmRecord.File2.Refresh

DirectorySort.Visible = False
DirectorySort.Cls
Unload DirectorySort

End Sub



Private Sub Dir1_Change()
    File1 = Dir1
    ChDir Dir1
    FileName.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
  'On Error Resume Next
    Dir1 = Drive1
    ChDrive Drive1
End Sub

Private Sub Form_Load()
FileName.Text = Dir1.Path
End Sub

