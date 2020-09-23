VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "MSWEBDVD.DLL"
Begin VB.Form frmDocument 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin MSWEBDVDLibCtl.MSWebDVD DVD 
      Height          =   4065
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6405
      _cx             =   11298
      _cy             =   7170
      DisableAutoMouseProcessing=   0   'False
      BackColor       =   1048592
      EnableResetOnStop=   0   'False
      ColorKey        =   1048592
      WindowlessActivation=   0   'False
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_Resize()
frmDocument.Left = Me.ScaleLeft
frmDocument.Width = Me.Width
DVD.Move 5, 5, Me.ScaleWidth - 10, Me.ScaleHeight - 10
If Val(GetSetting(App.Title, "Settings", "NaviSnapFunc")) = 2 Then ShapedForm.Move (frmDocument.Left + frmDocument.Width / 2) - ShapedForm.Width / 2, frmDocument.Top + frmDocument.Height  '* 1.65
End Sub

Private Sub DVD_DVDNotify(ByVal lEventCode As Long, ByVal lParam1 As Variant, ByVal lParam2 As Variant)
DVD.NotifyParentalLevelChange (True) 'application is notified it encounters video segments with a rating more restrictive than the overall rating for the disc
Dim t As String
ShapedForm.Label7.ForeColor = &HE0E0E0
ShapedForm.Label8.ForeColor = &HE0E0E0
ShapedForm.lblTimeTracker.Caption = ""
            On Local Error GoTo ErrLine
If 282 = lEventCode Then '282 is the event code for the time event, pass in param1 to get you the current time-convert to hh:mm:ss:ff format with DVDTimeCode2BSTR API
    If ShapedForm.lblTimeTrackerValue.Caption <> CStr(DVD.DVDTimeCode2bstr(lParam1)) Then ShapedForm.lblTimeTrackerValue.Caption = "Time " & CStr(DVD.DVDTimeCode2bstr(lParam1))
        ShapedForm.Lbllang.Caption = "Language " & DVD.GetLangFromLangID(0)
        ShapedForm.LblChapter.Caption = "Chapter " & DVD.CurrentChapter & " of " & DVD.GetNumberOfChapters(1)
        ShapedForm.lbltotaltime.Caption = "TotalTime " & DVD.TotalTitleTime
        DoTimes
End If
EnableDisableDomains
            Exit Sub
ErrLine:
            err.Clear
End Sub

Function DoTimes()
  CURTIME = DVD.CurrentTime 'CURRENT TIME
        DVDCurFrame = Total_Frames(CURTIME)
'lblNumFrames.Caption = "Total Frames : " & CaptionString
        t = DVD.TotalTitleTime ' TOTAL TIME
        DVDtotFrames = Total_Frames(t)
        If ShapedForm.lbltotaltime <> t Then 'UPDATE TO TOTAL TIME
            ShapedForm.lbltotaltime = t
            ShapedForm.UserControl11.Max = DVDtotFrames 'MAX TEH SLIDER SET TOT HE MAX FRAMES
        End If
        ShapedForm.lblFramesCur.Caption = DVDCurFrame  'CAL THE TOTAL FRAMES
        ShapedForm.lblNumFrames = DVDtotFrames
        ShapedForm.UserControl11.NavValue Val(DVDCurFrame)
        If ShapedForm.lblNumFrames <> DVDtotFrames Then 'CHECK TO UPDATE FRAMES TOTAL
            DVDtotFrames = Total_Frames(DVD.TotalTitleTime)
            ShapedForm.lblNumFrames = DVDtotFrames
        End If
        If DVD.CurrentChapter <> Last_Chapter Then
                t = DVD.TotalTitleTime
                DVDtotFrames = Total_Frames(t)
        End If
        tmp = DVDCurFrame Mod 25
        If tmp > 15 Then ShapedForm.UserControl11 = DVDCurFrame
End Function

Function EnableDisableDomains()
If frmDocument.DVD.CurrentDomain = 4 Then
ShapedForm.UserControl11.Visible = True
Else
ShapedForm.UserControl11.Visible = False
End If
'1   First Play
'2   Video Manager Menu
'3   Video Title Set Menu
'4   Title
'5   Stop
Select Case frmDocument.DVD.CurrentDomain

Case 1
disabled
ShapedForm.Label7.Visible = True 'play
Case 2
disabled
ShapedForm.Image6.Visible = True 'menu
ShapedForm.Label8.Visible = True 'Stop
ShapedForm.Label7.Visible = True 'play
Case 3
disabled
ShapedForm.Image6.Visible = True 'menu
ShapedForm.Label8.Visible = True 'Stop
ShapedForm.Label7.Visible = True 'play
Case 4

ShapedForm.Image7.Visible = True 'Pause
ShapedForm.Label8.Visible = True 'Stop
ShapedForm.Image1.Visible = True 'ffoward
ShapedForm.Image2.Visible = True 'rewind
ShapedForm.Image3.Visible = True 'prev
ShapedForm.Image4.Visible = True 'next
ShapedForm.Image5.Visible = True 'Step
ShapedForm.Image8.Visible = True 'nothing
Case 5
ShapedForm.Image6.Visible = True
disabled

End Select
End Function

Function disabled()
ShapedForm.Image7.Visible = False 'Pause
ShapedForm.Label8.Visible = False 'Stop
ShapedForm.Image1.Visible = False 'ffoward
ShapedForm.Image2.Visible = False 'rewind
ShapedForm.Image3.Visible = False 'prev
ShapedForm.Image4.Visible = False 'next
ShapedForm.Image5.Visible = False 'Step
ShapedForm.Image8.Visible = False 'nothing
End Function

Function GetTheTitleParentalLevel()
Dim ilevels As Variant

If DVD.GetPlayerParentalLevel = 1 Then result = 1
If DVD.GetPlayerParentalLevel = 2 Then result = 2
If DVD.GetPlayerParentalLevel = 3 Then result = 3
If DVD.GetPlayerParentalLevel = 4 Then result = 4
If DVD.GetPlayerParentalLevel = 5 Then result = 5
If DVD.GetPlayerParentalLevel = 6 Then result = 6
If DVD.GetPlayerParentalLevel = 7 Then result = 7
If DVD.GetPlayerParentalLevel = 8 Then result = 8
SKIP:
MsgBox result

End Function

'Type tagDVD_PARENTAL_LEVEL
'DVD_PARENTAL_LEVEL_8 = 0x8000
'DVD_PARENTAL_LEVEL_7 = 0x4000
'DVD_PARENTAL_LEVEL_6 = 0x2000
'DVD_PARENTAL_LEVEL_5 = 0x1000
'DVD_PARENTAL_LEVEL_4 = 0x0800
'DVD_PARENTAL_LEVEL_3 = 0x0400
'DVD_PARENTAL_LEVEL_2 = 0x0200
'DVD_PARENTAL_LEVEL_1 = 0x0100
'End Type
