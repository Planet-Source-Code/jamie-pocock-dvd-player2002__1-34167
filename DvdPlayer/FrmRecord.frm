VERSION 5.00
Begin VB.Form FrmRecord 
   BackColor       =   &H00808000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DVD Riper"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   6315
   Icon            =   "FrmRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      Caption         =   "Compare files to original"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   36
      Top             =   2580
      Width           =   1815
   End
   Begin VB.CommandButton cmdFileOp 
      Caption         =   "&Copy file/s"
      Height          =   315
      Index           =   0
      Left            =   4410
      TabIndex        =   17
      ToolTipText     =   "Enter Source and Destination."
      Top             =   3780
      Width           =   1785
   End
   Begin VB.TextBox txtSource 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1470
      TabIndex        =   15
      ToolTipText     =   "Enter the full path of the file or folder to copy, rename, move or delete. Wildcards are allowed."
      Top             =   2910
      Width           =   3375
   End
   Begin VB.TextBox txtDestination 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1470
      TabIndex        =   14
      ToolTipText     =   "Enter the full path of the destination file or folder.  Wildcards are allowed."
      Top             =   3210
      Width           =   3375
   End
   Begin VB.CheckBox chkSilent 
      BackColor       =   &H00808000&
      Caption         =   "Display a progress dialog box"
      Height          =   255
      Left            =   6690
      TabIndex        =   10
      Top             =   1230
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkYesToAll 
      BackColor       =   &H00808000&
      Caption         =   "Respond with 'Yes to all' for any dialog box that is displayed"
      Height          =   255
      Left            =   6690
      TabIndex        =   9
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CheckBox chkRename 
      BackColor       =   &H00808000&
      Height          =   285
      Left            =   6720
      TabIndex        =   8
      Top             =   930
      Width           =   5535
   End
   Begin VB.CheckBox chkDir 
      BackColor       =   &H00808000&
      Caption         =   "Do not confirm if a new directory is created"
      Height          =   255
      Left            =   6690
      TabIndex        =   7
      Top             =   1650
      Width           =   5565
   End
   Begin VB.CheckBox chkFilesOnly 
      BackColor       =   &H00808000&
      Caption         =   "Perform the operation only on files if a wildcard filename (*.*) is specified"
      Height          =   225
      Left            =   6690
      TabIndex        =   6
      Top             =   1890
      Width           =   5535
   End
   Begin VB.FileListBox File2 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   4650
      TabIndex        =   2
      Top             =   330
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   90
      Pattern         =   "*.vob"
      TabIndex        =   1
      Top             =   330
      Width           =   1605
   End
   Begin VB.CommandButton cmdFileOp 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   11
      ToolTipText     =   "Enter Source only."
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdFileOp 
      Caption         =   "&Move"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Enter Source and Destination."
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdFileOp 
      Caption         =   "&Rename"
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   13
      ToolTipText     =   "Enter Source and Destination."
      Top             =   1740
      Width           =   855
   End
   Begin VB.CommandButton cmdconvert 
      Caption         =   "&Convert to AVI"
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   3780
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FrmRecord.frx":08CA
      Height          =   1665
      Left            =   1740
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   3750
      X2              =   4500
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   1800
      X2              =   2550
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Label lblsizematch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   2820
      TabIndex        =   37
      Top             =   2310
      Width           =   630
   End
   Begin VB.Label lblmatch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   2820
      TabIndex        =   35
      Top             =   2100
      Width           =   630
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   1830
      X2              =   4500
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Index           =   0
      Left            =   3210
      TabIndex        =   34
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Index           =   3
      Left            =   1740
      TabIndex        =   33
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label lblleft1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1740
      TabIndex        =   32
      Top             =   1740
      Width           =   630
   End
   Begin VB.Label lblright1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   3210
      TabIndex        =   31
      Top             =   1740
      Width           =   630
   End
   Begin VB.Label lblifo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   3210
      TabIndex        =   30
      Top             =   1560
      Width           =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   3120
      X2              =   3120
      Y1              =   300
      Y2              =   2070
   End
   Begin VB.Label lblIfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1740
      TabIndex        =   29
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lbl2a2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   3210
      TabIndex        =   28
      Top             =   330
      Width           =   630
   End
   Begin VB.Label lbl2b2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   3210
      TabIndex        =   27
      Top             =   510
      Width           =   630
   End
   Begin VB.Label lbl2c2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   3210
      TabIndex        =   26
      Top             =   720
      Width           =   630
   End
   Begin VB.Label lbl2d2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   3210
      TabIndex        =   25
      Top             =   900
      Width           =   630
   End
   Begin VB.Label LblinfoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1740
      TabIndex        =   23
      Top             =   900
      Width           =   630
   End
   Begin VB.Label LblinfoA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1740
      TabIndex        =   22
      Top             =   720
      Width           =   630
   End
   Begin VB.Label LblinfoB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1740
      TabIndex        =   21
      Top             =   510
      Width           =   630
   End
   Begin VB.Label LblinfoC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....................."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1740
      TabIndex        =   20
      Top             =   330
      Width           =   630
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Source File :"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   -450
      TabIndex        =   19
      Top             =   2910
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destination File :"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   -450
      TabIndex        =   18
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Folder"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4920
      TabIndex        =   5
      Top             =   3270
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Folder"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4890
      TabIndex        =   4
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Directory"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   4650
      TabIndex        =   3
      Top             =   90
      Width           =   1290
   End
   Begin VB.Label lblDVDPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DVD Directory"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   915
   End
   Begin VB.Menu mfile 
      Caption         =   "File"
      Begin VB.Menu mnStartCopy 
         Caption         =   "Start Copy"
      End
      Begin VB.Menu mnuFileBarsmal1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSmalMove 
         Caption         =   "Move"
      End
      Begin VB.Menu MnuSmallRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu MnuSmallDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuFileBarsmal2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSmallQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu MnuSmallHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuSmallHelp2 
         Caption         =   "Ripper Help"
      End
      Begin VB.Menu MnuSmallAbout 
         Caption         =   "About DVD Ripper"
      End
   End
End
Attribute VB_Name = "FrmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_ALLOWUNDO = &H40
Const FOF_SILENT = &H4
Const FOF_NOCONFIRMATION = &H10
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMMKDIR = &H200
Const FOF_FILESONLY = &H80
Dim oFSO As Object
Private Type SHFILEOPSTRUCT
    hwnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long



Private Sub Check1_Click()
lblmatch.Caption = ""
lblsizematch.Caption = ""
End Sub

Private Sub File1_Click()
GetTheData File1.FileName
txtSource = File1.Path & "\" & File1.FileName
txtDestination = File2.Path & "\" & File1.FileName
End Sub

Private Sub File2_Click()
'On Error Resume Next
lblmatch.Caption = ""
lblsizematch.Caption = ""
GetTheData2 File2.FileName
If Check1 Then match File2.FileName
End Sub

Private Sub Form_Load()
''On Error Resume Next
Label1.Caption = GetSetting(App.Title, "Settings", "Dumpfolder")

'File1.Pattern = "*.vob"
File1.Path = frmDocument.DVD.DVDDirectory
Dim oDrive  As Object
lblIfo.Caption = File1.ListCount & " files"
lblifo2.Caption = File2.ListCount & " files"
lblleft1.Caption = FileDateTime(File1.Path)
lblright1.Caption = FileDateTime(File2.Path)
cmdFldrDetails
If Val(GetSetting(App.Title, "Settings", "Dumpfolder")) >= 1 Then
lblDVDPath.Caption = lblDVDPath.Caption & " = " & frmDocument.DVD.DVDDirectory
End If
If Len(GetSetting(App.Title, "Settings", "Dumpfolder")) <= 0 Then GoTo SKIP
File2.Path = GetSetting(App.Title, "Settings", "Dumpfolder")
SKIP:

Set oDrive = Nothing
End Sub

Function cmdFldrDetails()
'On Error Resume Next
Dim oFolder As Object
Dim oFile   As Object

If Right$(File1.Path, 1) <> "\" Then File1.Path = File1.Path & "\"

Set oFSO = CreateObject("Scripting.FileSystemObject")

If oFSO.FolderExists(File1.Path) Then
    Set oFolder = oFSO.GetFolder(File1.Path)
    With oFolder
        lblInfo(3) = "Folder size " & Mid(.Size / 1048576, 1, 4) & " MB"
    End With
End If
    Set oFolder = Nothing
    Set oFile = Nothing
    
Set oFSO = CreateObject("Scripting.FileSystemObject")
If oFSO.FolderExists(File2.Path) Then
    Set oFolder = oFSO.GetFolder(File2.Path)
    With oFolder
        lblInfo(0) = "Folder size " & Mid(.Size / 1048576, 1, 4) & " MB"
    End With
End If

    Set oFolder = Nothing
    Set oFile = Nothing
End Function

Function match(FileName As String)
'On Error Resume Next
Dim i As Integer
For i = 0 To File1.ListCount - 1
File1.Selected(i) = True
If File1.FileName = File2.FileName Then
lblmatch.Caption = "File name 100% match"

If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) = FileLen(frmDocument.DVD.DVDDirectory & "\" & File2.FileName) Then
lblsizematch.Caption = "File size 100% match"
Else
lblsizematch.Caption = "WARNING file size mismatch"
End If

Exit Function
End If
Next i
End Function



Function GetTheData2(FileName As String)
'On Error Resume Next
lbl2d2.Caption = FileDateTime(frmDocument.DVD.DVDDirectory & "\" & FileName) '
lbl2b2.Caption = FileName
lbl2c2.Caption = "File attr " & GetAttr(frmDocument.DVD.DVDDirectory & "\" & FileName)

If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) < 999 Then
lbl2a2.Caption = "File size " & FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) & " bytes"
End If
If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) >= 1000 And FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) < 999999 Then
lbl2a2.Caption = "File size " & Mid(FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) / 1024, 1, 6) & " KB"
End If
If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) >= 1000000 Then
lbl2a2.Caption = "File size " & Mid(FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) / 1048576, 1, 5) & " MB"
End If

End Function



Function GetTheData(FileName As String)
'On Error Resume Next
LblinfoA.Caption = FileDateTime(frmDocument.DVD.DVDDirectory & "\" & FileName) '
LblinfoB1.Caption = FileName
LblinfoA1.Caption = "File attr " & GetAttr(frmDocument.DVD.DVDDirectory & "\" & FileName)

If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) < 999 Then
LblinfoC1.Caption = "File size " & FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) & " bytes"
End If
If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) >= 1000 And FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) < 999999 Then
LblinfoC1.Caption = "File size " & Mid(FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) / 1024, 1, 6) & " KB"
End If
If FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) >= 1000000 Then
LblinfoC1.Caption = "File size " & Mid(FileLen(frmDocument.DVD.DVDDirectory & "\" & FileName) / 1048576, 1, 5) & " MB"
End If

End Function

Private Sub Label4_Click()
'On Error Resume Next
DirectorySort.Show
End Sub

Private Sub cmdFileOp_Click(Index As Integer)
'On Error Resume Next
Label6.Visible = True
DoEvents
Dim lFileOp  As Long
Dim lresult  As Long
Dim lFlags   As Long
Dim SHFileOp As SHFILEOPSTRUCT

Screen.MousePointer = vbHourglass
Select Case Index
    Case 0
        lFileOp = FO_COPY
    Case 1
        lFileOp = FO_MOVE
    Case 2
        lFileOp = FO_RENAME
    Case 3
        lFileOp = FO_DELETE
End Select
If chkSilent Then lFlags = lFlags Or FOF_SILENT
If chkYesToAll Then lFlags = lFlags Or FOF_NOCONFIRMATION
If chkRename Then lFlags = lFlags Or FOF_RENAMEONCOLLISION
If chkDir Then lFlags = lFlags Or FOF_NOCONFIRMMKDIR
If chkFilesOnly Then lFlags = lFlags Or FOF_FILESONLY
'
' NOTE:  By adding the FOF_ALLOWUNDO flag you can move
' a file to the Recycle Bin instead of deleting it.

With SHFileOp
    .wFunc = lFileOp
    .pFrom = txtSource.Text & vbNullChar & vbNullChar
    .pTo = txtDestination & vbNullChar & vbNullChar
    .fFlags = lFlags
End With
lresult = SHFileOperation(SHFileOp)
'
' If User hit Cancel button while operation is in progress,
' the fAborted parameter will be true
'
Screen.MousePointer = vbDefault

If lresult <> 0 Or SHFileOp.fAborted Then
MsgBox "Operation Has Failed To Complete,", vbInformation, "File Operations"
File2.Refresh
Label6.Caption = "Attempting to fix, please wait"
frmDocument.DVD.Play
frmDocument.DVD.Stop
Exit Sub
End If
File2.Refresh
lblIfo.Caption = File1.ListCount & " files"
lblifo2.Caption = File2.ListCount & " files"
lblleft1.Caption = FileDateTime(File1.Path)
lblright1.Caption = FileDateTime(File2.Path)
Label6.Visible = False
MsgBox "Operation Complete", vbInformation, "File Operations"
End Sub


Private Sub mnStartCopy_Click()
cmdFileOp_Click (0)
End Sub

Private Sub MnuSmallDelete_Click()
cmdFileOp_Click (3)
End Sub

Private Sub MnuSmallQuit_Click()
Unload Me
End Sub

Private Sub MnuSmallRename_Click()
cmdFileOp_Click (2)
End Sub

Private Sub MnuSmalMove_Click()
cmdFileOp_Click (1)
End Sub








