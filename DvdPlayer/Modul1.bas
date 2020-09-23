Attribute VB_Name = "Exit"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public fMainForm As frmMain
Public lFileLen As Long
Public Function Total_Frames(ByVal s As String) As String
'"hh:mm:ss:ff"
X = Split(s, ":")
DVDHours = X(0)
DVDMins = X(1)
DVDSecs = X(2)
DVDFf = X(3)
fps = 25 ''  TabForm.DVD.FramesPerSecond
tmph = DVDHours * 3600 * fps
tmps = DVDMins * 60 * fps
tmpss = DVDSecs * fps
tmpDVDtotFrames = tmph + tmps + tmpss + DVDFf
''ShapedForm.lblFrames = "Total Frames: " & DVDtotFrames

Total_Frames = tmpDVDtotFrames
'
End Function
Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub
Function CallDomain() 'Call this method to ensure the DVD Navigator is in a valid domain
' for the method you are about to call. For example, before calling PlayTitle,
' check the CurrentDomain property to make sure that the DVD Navigator is not in the Stop or First Play domain.

Select Case frmDocument.DVD.CurrentDomain
Case 1
MsgBox "This operation is not permitted at this time, please wait ", vbOKOnly, "First play"
Case 2
MsgBox "This operation is not permitted from the Video Manager Menu", vbOKOnly, "Video Manager Menu"
Case 3
MsgBox "This operation is not permitted from the Video Title Set Menu", vbOKOnly, "Video Title Menu"
Case 4
MsgBox "This operation is not permitted while the disc is playing, click stop first ", vbOKOnly, "Disc Playing"
Case 5
MsgBox "This operation is not permitted while the disc is stopped, click play first ", vbOKOnly, "Disc Stopped"
End Select

End Function
Public Function ExitPlayer()
If frmDocument.Visible = True Then
    Dim xc As String
    Dim yc As String
    Dim xs As String
    Dim ys As String
    
    'Get position of contral panel and screen
    xc = frmDocument.Left
    yc = frmDocument.Top
    'get position of screen
    xs = frmDocument.Left
    ys = frmDocument.Top
    'unload frames
    Unload frmDocument
    'End player
    End
    End If
End Function

Public Function Framesto_Times(ByVal timein As Long) As String
's=160000

    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String

    'timein = timein / 25 'FRAMES

    remTime = timein / 25
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    
    If conH > 0 Then
        strRetTime = "0" & Trim(Str(conH)) & ":"
    Else
        strRetTime = "00" & ":"
    End If
    
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(Str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conM))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    strRetTime = strRetTime & ":"
    
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(Str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    
   Framesto_Times = strRetTime
End Function

