VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H00000000&
   Caption         =   "DVD 202"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9060
   Icon            =   "frmMain.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "Load Book Mark"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save Book Mark"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "&Controls"
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play DVD"
      End
      Begin VB.Menu MNUSTOP 
         Caption         =   "&Stop DVD"
      End
      Begin VB.Menu MNUPAUSE 
         Caption         =   "P&ause DVD"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu fowardmenu 
         Caption         =   "Forward"
         Begin VB.Menu mnuforx2 
            Caption         =   "Forward x 2"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuforx4 
            Caption         =   "Forward x 4"
         End
         Begin VB.Menu mnuforx8 
            Caption         =   "Forward x 8"
         End
      End
      Begin VB.Menu RewindMenu 
         Caption         =   "Rewind"
         Begin VB.Menu mnurewx2 
            Caption         =   "Rewind x 2"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnurewx4 
            Caption         =   "Rewind x 4"
         End
         Begin VB.Menu mnurewx8 
            Caption         =   "Rewind x 8"
         End
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStiiloff 
         Caption         =   "Still Off"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnustepdvd 
         Caption         =   "Step"
         Begin VB.Menu mnustep1 
            Caption         =   "Step x 1"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mnustep2 
            Caption         =   "Step x 2"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnustep4 
            Caption         =   "Step x 4"
            Shortcut        =   +{F3}
         End
      End
      Begin VB.Menu mnubreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnusnapshot 
         Caption         =   "Take snapshot"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuVeiwsnapshot 
         Caption         =   "Veiw Snap Shot"
      End
   End
   Begin VB.Menu DVDMenus 
      Caption         =   "&Menus"
      Begin VB.Menu RootMenu 
         Caption         =   "&Root Menu"
      End
      Begin VB.Menu TitleMenu 
         Caption         =   "&Title Menu"
      End
      Begin VB.Menu AudioMenu 
         Caption         =   "&Audio Menu"
      End
      Begin VB.Menu AngleMenu 
         Caption         =   "A&ngle Menu"
      End
      Begin VB.Menu ChapterMenu 
         Caption         =   "&Chapter Menu"
      End
      Begin VB.Menu SubpictureMenu 
         Caption         =   "&Subpicture Menu"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "DVD Player"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "Switch to FULLSCREEN"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "Search For Help On"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub AngleMenu_Click()
On Error GoTo ErrLinee
Call frmDocument.DVD.ShowMenu(6)  'Angle
Exit Sub
ErrLinee:
Call MsgBox(err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): err.Clear
Exit Sub
End Sub

Private Sub AudioMenu_Click()
On Error GoTo ErrLinef
Call frmDocument.DVD.ShowMenu(5)  'Audio
Exit Sub
ErrLinef:
Call MsgBox(err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): err.Clear
Exit Sub
End Sub

Private Sub ChapterMenu_Click()
On Error GoTo ErrLineg
Call frmDocument.DVD.ShowMenu(7)  'Chapter
Exit Sub
ErrLineg:
Call MsgBox(err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): err.Clear
Exit Sub
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
If Len(GetSetting(App.Title, "Settings", "msgtips")) <= 1 Then
SaveSetting App.Title, "Settings", "msgtips", "on"
SaveSetting App.Title, "Settings", "NaviSnapFunc", "2"
SaveSetting App.Title, "Settings", "NaviSearchSpeed", "3"
SaveSetting App.Title, "Settings", "MainLeft", "100"
SaveSetting App.Title, "Settings", "MainTop", "100"
SaveSetting App.Title, "Settings", "MainWidth", Me.Width
SaveSetting App.Title, "Settings", "MainHeight", Me.Height
SaveSetting App.Title, "Settings", "dvdLeft", "10"
SaveSetting App.Title, "Settings", "dvdWidth", frmDocument.Width
SaveSetting App.Title, "Settings", "dvdHeight", frmDocument.Height
SaveSetting App.Title, "Settings", "OptAspectRatio", "1"
End If
Me.Left = Val(GetSetting(App.Title, "Settings", "MainLeft"))
Me.Top = Val(GetSetting(App.Title, "Settings", "MainTop"))
Me.Width = Val(GetSetting(App.Title, "Settings", "MainWidth"))
Me.Height = Val(GetSetting(App.Title, "Settings", "MainHeight"))
    
frmDocument.Move 10, 10, Me.ScaleWidth - 200, Me.ScaleHeight - 900

    DoEvents
    MDIForm_Resize
    'Form1.Show
End Sub

Private Sub MDIForm_Resize()
    ShapedForm.Show 'Me.Width / 8
If Val(GetSetting(App.Title, "Settings", "OptAspectRatio")) = 1 Then frmDocument.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - ShapedForm.Height - 60
If Val(GetSetting(App.Title, "Settings", "NaviSnapFunc")) = 2 Then ShapedForm.Move (frmDocument.Left + frmDocument.Width / 2) - ShapedForm.Width / 2, frmDocument.Top + frmDocument.Height  '* 1.65
If Val(GetSetting(App.Title, "Settings", "NaviSnapFunc")) = 1 Then ShapedForm.Move (frmDocument.Left + frmDocument.Width / 2) - ShapedForm.Width / 2, Me.Height - ShapedForm.Height * 1.65

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "dvdLeft", frmDocument.Left
        SaveSetting App.Title, "Settings", "dvdWidth", frmDocument.Width
        SaveSetting App.Title, "Settings", "dvdHeight", frmDocument.Height
    End If
End Sub

Private Sub mnuFileNew_Click()

            If frmDocument.DVD.CurrentDomain = 1 Then GoTo SKIP
            If frmDocument.DVD.CurrentDomain = 2 Then GoTo SKIP
            If frmDocument.DVD.CurrentDomain = 3 Then GoTo SKIP
            If frmDocument.DVD.CurrentDomain = 4 Then GoTo SKIP
                    CallDomain 'ensure DVD Navigator is in a valid domain
            Exit Sub

SKIP:
On Error GoTo SKIPBOOK
frmDocument.DVD.RestoreBookmark
Exit Sub
SKIPBOOK:
MsgBox "Either there is no book mark stored or the book mark stored is not for this title", vbOKOnly, "Can not restore"
End Sub

Private Sub mnuforx2_Click()
frmDocument.DVD.PlayForwards 2
End Sub

Private Sub mnuforx4_Click()
frmDocument.DVD.PlayForwards 4
End Sub

Private Sub mnuforx8_Click()
frmDocument.DVD.PlayForwards 8
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        'On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If err Then
            MsgBox err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        'On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If err Then
            MsgBox err.Description
        End If
    End If

End Sub

Function GetVideoSize() As DVDRect
MsgBox DVDRect
End Function

Private Sub MNUPAUSE_Click()
frmDocument.DVD.Pause
End Sub

Private Sub mnuPlay_Click()
frmDocument.DVD.Play
End Sub

Private Sub mnurewx2_Click()
frmDocument.DVD.PlayBackwards 2
End Sub

Private Sub mnurewx4_Click()
frmDocument.DVD.PlayBackwards 4
End Sub

Private Sub mnurewx8_Click()
frmDocument.DVD.PlayBackwards 8
End Sub

Private Sub mnustep1_Click()
frmDocument.DVD.Step 1
End Sub

Private Sub mnustep2_Click()
frmDocument.DVD.Step 2
End Sub

Private Sub mnustep4_Click()
frmDocument.DVD.Step 4
End Sub

Private Sub mnuStiiloff_Click()
frmDocument.DVD.StillOff
End Sub

Private Sub MNUSTOP_Click()
frmDocument.DVD.Stop
End Sub

Private Sub mnuWindowMenu_Click()
On Error GoTo 1
frmDocument.DVD.ShowMenu (1)
Exit Sub
1:
On Error GoTo 2
'frmDocument.DVD.ShowMenu (2)
Exit Sub
2:
On Error GoTo 3
frmDocument.DVD.ShowMenu (3)
Exit Sub
3:
End Sub

Private Sub mnuWindowNewWindow_Click()
If frmDocument.DVD.CurrentDomain = 4 Then
frmDocument.DVD.DisableAutoMouseProcessing = True
MsgBox "Press the windows key to minamize screen"
ShapedForm.ZOrder 1
frmDocument.DVD.FullScreenMode = True
Else
CallDomain
End If
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    frmDocument.DVD.ShowMenu (3)
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
     ExitPlayer
    Unload Me

End Sub


Private Sub mnuFileSave_Click()
On Error GoTo SKIP
frmDocument.DVD.DeleteBookmark
DoEvents
frmDocument.DVD.SaveBookmark
MsgBox "To return to this book mark use the file menu Load Book Mark", vbOKOnly, "Book mark saved"
Exit Sub
SKIP:
MsgBox "An error occured durring save, please try while the DVD is running"
End Sub

Private Sub RootMenu_Click()

Dim i As Integer
    For i = 0 To 5
        If frmDocument.DVD.CurrentDomain = i Then
            If frmDocument.DVD.CurrentDomain = 3 Then GoTo SKIP
                    CallDomain 'ensure DVD Navigator is in a valid domain
            Exit Sub
        End If
    Next i
SKIP:

On Error GoTo ErrLinec
Call frmDocument.DVD.ShowMenu(3)  'Root
Exit Sub
ErrLinec:
Call MsgBox(err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): err.Clear
Exit Sub
End Sub

Private Sub SubpictureMenu_Click()
On Error GoTo ErrLineb
Call frmDocument.DVD.ShowMenu(4)  'Subpicture
Exit Sub
ErrLineb:
Call MsgBox(err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): err.Clear
Exit Sub
End Sub

Private Sub TitleMenu_Click()
On Error GoTo ErrLinea
Call frmDocument.DVD.ShowMenu(2)  'Title
Exit Sub
ErrLinea:
Call MsgBox(err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): err.Clear
Exit Sub
End Sub
