VERSION 5.00
Object = "{820DD9F2-236E-4DF3-8763-6D74E6507251}#2.0#0"; "DVDNAV.OCX"
Begin VB.Form ShapedForm 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   1200
   ClientLeft      =   6000
   ClientTop       =   7065
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "Player.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Player.frx":030A
   ScaleHeight     =   1200
   ScaleWidth      =   7200
   Begin dvdNavigator.UserControl1 UserControl11 
      Height          =   135
      Left            =   1170
      TabIndex        =   18
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   238
      ForeColor       =   255
      BackColor       =   0
      Max             =   100
      Mode            =   0
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Frame"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   150
      Index           =   1
      Left            =   3780
      TabIndex        =   17
      Top             =   390
      Width           =   765
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Frames"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   150
      Index           =   0
      Left            =   3840
      TabIndex        =   16
      Top             =   600
      Width           =   690
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000004&
      BorderColor     =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   1170
      Top             =   840
      Width           =   4305
   End
   Begin VB.Label lblFramesCur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   4590
      TabIndex        =   15
      Top             =   390
      Width           =   765
   End
   Begin VB.Label lblNumFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   4590
      TabIndex        =   14
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   150
      Index           =   3
      Left            =   2010
      TabIndex        =   13
      Top             =   870
      Width           =   30
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      MouseIcon       =   "Player.frx":2717
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   420
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6540
      MouseIcon       =   "Player.frx":2869
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   420
      Width           =   270
   End
   Begin VB.Label lblTimeTracker 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Event"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   3000
      TabIndex        =   3
      Top             =   90
      Width           =   1155
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   6420
      MouseIcon       =   "Player.frx":29BB
      MousePointer    =   99  'Custom
      Top             =   780
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image7 
      Height          =   525
      Left            =   6150
      MouseIcon       =   "Player.frx":2B0D
      MousePointer    =   99  'Custom
      ToolTipText     =   "Pause"
      Top             =   270
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   6390
      MouseIcon       =   "Player.frx":2C5F
      MousePointer    =   99  'Custom
      ToolTipText     =   "Show Menu"
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   525
      Left            =   6930
      MouseIcon       =   "Player.frx":2DB1
      MousePointer    =   99  'Custom
      ToolTipText     =   "Frame step"
      Top             =   270
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   270
      MouseIcon       =   "Player.frx":2F03
      MousePointer    =   99  'Custom
      ToolTipText     =   "Next Chapter"
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   270
      MouseIcon       =   "Player.frx":3055
      MousePointer    =   99  'Custom
      ToolTipText     =   "Previous Chapter"
      Top             =   810
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      MouseIcon       =   "Player.frx":31A7
      MousePointer    =   99  'Custom
      ToolTipText     =   "Rewind"
      Top             =   270
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   810
      MouseIcon       =   "Player.frx":32F9
      MousePointer    =   99  'Custom
      ToolTipText     =   "Fast Forward"
      Top             =   300
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbltotaltime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   150
      Left            =   2340
      TabIndex        =   10
      Top             =   390
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      Index           =   3
      X1              =   3690
      X2              =   3690
      Y1              =   390
      Y2              =   780
   End
   Begin VB.Label LblChapter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chapter"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   150
      Left            =   1230
      TabIndex        =   9
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Lbllang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Language"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   150
      Left            =   2340
      TabIndex        =   8
      Top             =   600
      Width           =   510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   5460
      X2              =   1170
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      Index           =   0
      X1              =   2280
      X2              =   2280
      Y1              =   390
      Y2              =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5490
      MouseIcon       =   "Player.frx":344B
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   570
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   5940
      MouseIcon       =   "Player.frx":359D
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   570
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   150
      Index           =   0
      Left            =   5610
      TabIndex        =   5
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lblTimeTrackerValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time 00:00:00:00"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   150
      Left            =   1230
      TabIndex        =   4
      Top             =   390
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   4
      X1              =   7080
      X2              =   7080
      Y1              =   1170
      Y2              =   1050
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   5700
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   5670
      TabIndex        =   2
      Top             =   90
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   5
      X1              =   1140
      X2              =   6060
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label LabMute 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mute"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   165
      Left            =   5580
      MouseIcon       =   "Player.frx":36EF
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   5910
      TabIndex        =   0
      Top             =   90
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   5880
      Top             =   90
      Width           =   195
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   5670
      Top             =   90
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   5550
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000004&
      BorderColor     =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   0
      Left            =   1170
      Top             =   360
      Width           =   4305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   1
      X1              =   990
      X2              =   6240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image10 
      Height          =   795
      Left            =   1260
      Top             =   300
      Width           =   4935
   End
   Begin VB.Image Image9 
      Height          =   1095
      Left            =   1080
      MouseIcon       =   "Player.frx":3841
      MousePointer    =   99  'Custom
      Top             =   60
      Width           =   4905
   End
End
Attribute VB_Name = "ShapedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 6.2"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long
Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)


'!0,30,370,411,71,0,0,1
    ObjectRegion = CreateRectRgn(370 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 411 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
'!0,30,320,361,71,0,0,1
    ObjectRegion = CreateRectRgn(320 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 361 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,30,270,311,71,0,0,1
    ObjectRegion = CreateRectRgn(270 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 311 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,30,220,261,71,0,0,1
    ObjectRegion = CreateRectRgn(220 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 261 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,30,170,211,71,0,0,1
    ObjectRegion = CreateRectRgn(170 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 211 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,30,120,161,71,0,0,1
    ObjectRegion = CreateRectRgn(120 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 161 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,30,70,111,71,0,0,1
    ObjectRegion = CreateRectRgn(70 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 30 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 111 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,0,70,411,21,0,0,1
    ObjectRegion = CreateRectRgn(70 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 0 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 411 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 21 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,19,99,111,40,0,0,1
    ObjectRegion = CreateRectRgn(99 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 19 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 111 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 40 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,20,370,381,33,0,0,1
    ObjectRegion = CreateRectRgn(370 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 20 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 381 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 33 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,19,136,146,32,0,0,1
    ObjectRegion = CreateRectRgn(136 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 19 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 146 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 32 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,20,336,347,33,0,0,1
    ObjectRegion = CreateRectRgn(336 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 20 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 347 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 33 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,17,270,290,32,0,0,1
    ObjectRegion = CreateRectRgn(270 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 17 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 290 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 32 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,19,190,211,31,0,0,1
    ObjectRegion = CreateRectRgn(190 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 19 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 211 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 31 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,19,231,251,35,0,0,1
    ObjectRegion = CreateRectRgn(231 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 19 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 251 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 35 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!3,-1,411,480,70,0,0,1
    ObjectRegion = CreateEllipticRgn(411 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, -1 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 480 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 70 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!3,0,0,71,73,0,0,1
    ObjectRegion = CreateEllipticRgn(0 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 0 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 71 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 73 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,70,10,61,81,0,0,1
    ObjectRegion = CreateRectRgn(10 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 70 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 61 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 81 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,69,420,472,81,0,0,1
    ObjectRegion = CreateRectRgn(420 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 69 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 472 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 81 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,71,79,400,81,0,0,1
    ObjectRegion = CreateRectRgn(79 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 71 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 400 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 81 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,76,70,75,77,0,0,1
    ObjectRegion = CreateRectRgn(70 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 76 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 75 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 77 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,18,70,410,70,0,0,1
    ObjectRegion = CreateRectRgn(70 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 18 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 410 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 70 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,14,402,411,70,0,0,1
    ObjectRegion = CreateRectRgn(402 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 14 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 411 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 70 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
'!0,64,104,381,75,0,0,1
    ObjectRegion = CreateRectRgn(104 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 64 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 381 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 75 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function



Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 0, 0), True)
    UserControl11.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HE0E0E0
Label8.ForeColor = &HE0E0E0
lblTimeTracker.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub

Private Sub Image1_Click()
            If frmDocument.DVD.CurrentDomain = 4 Then
            If Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed")) = 3 Then frmDocument.DVD.PlayForwards 2
            If Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed")) = 4 Then frmDocument.DVD.PlayForwards 4
            If Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed")) = 5 Then frmDocument.DVD.PlayForwards 8
            Exit Sub
            End If
            CallDomain
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image1.ToolTipText
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = ""
End Sub




Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = "Loop"
End Sub

Private Sub Image2_Click()
            If frmDocument.DVD.CurrentDomain = 4 Then
            frmDocument.DVD.PlayBackwards 2 'GoTo SKIP
            If Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed")) = 3 Then frmDocument.DVD.PlayBackwards 2
            If Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed")) = 4 Then frmDocument.DVD.PlayBackwards 4
            If Val(GetSetting(App.Title, "Settings", "NaviSearchSpeed")) = 5 Then frmDocument.DVD.PlayBackwards 8
            Exit Sub
            End If
            CallDomain
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image2.ToolTipText
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image3.ToolTipText
End Sub
Private Sub Image3_Click()
            If frmDocument.DVD.CurrentDomain = 4 Then
            frmDocument.DVD.PlayPrevChapter
            Exit Sub
            End If
            CallDomain
End Sub
Private Sub Image4_Click()
            If frmDocument.DVD.CurrentDomain = 4 Then
            frmDocument.DVD.PlayNextChapter
            Exit Sub
            End If
            CallDomain
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image4.ToolTipText
End Sub

Private Sub Image5_Click()
frmDocument.DVD.Step 1
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image5.ToolTipText
End Sub

Private Sub Image6_Click()
Dim i As Integer
    For i = 0 To 5
        If frmDocument.DVD.CurrentDomain = i Then
            If frmDocument.DVD.CurrentDomain = 3 Then GoTo SKIP
            If frmDocument.DVD.CurrentDomain = 4 Then GoTo SKIP
                    CallDomain 'ensure DVD Navigator is in a valid domain
            Exit Sub
        End If
    Next i
Exit Sub
SKIP:
On Error Resume Next
frmDocument.DVD.ShowMenu (3)
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image6.ToolTipText
End Sub

Private Sub Image7_Click()
Dim i As Integer
    For i = 0 To 5
        If frmDocument.DVD.CurrentDomain = i Then
            If frmDocument.DVD.CurrentDomain = 4 Then GoTo SKIP
                    CallDomain 'ensure DVD Navigator is in a valid domain
            Exit Sub
        End If
    Next i
SKIP:
frmDocument.DVD.Pause
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image7.ToolTipText
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Image8.ToolTipText
End Sub


Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub



Private Sub Label1_Click()
ExitPlayer
End Sub

Private Sub Label2_Click()
frmDocument.DVD.FullScreenMode = True
End Sub

Private Sub Label5_Click()
If Val(Label5.Tag) = 0 Then
If GetSetting(App.Title, "Settings", "msgtips") = "on" Then MsgBox "Click and hold the left mouse button on the video to move around", vbOKOnly, "Goto options to turn message's off"
Label5.Tag = 1
Label6.Tag = 1
End If
frmDocument.DVD.Zoom 360, 270, 2
End Sub

Private Sub Label6_Click()
'SaveSetting App.Title, "Settings", "msgtips", "on"
If Val(Label6.Tag) = 0 Then
If GetSetting(App.Title, "Settings", "msgtips") = "on" Then MsgBox "Click and hold the left mouse button on the video to move around", vbOKOnly, "Goto options to turn message's off"
Label6.Tag = 1
Label5.Tag = 1
End If
frmDocument.DVD.Zoom 360, 270, 0.5
End Sub

Private Sub Label7_Click()
On Error GoTo err
Dim Ibutton As Integer
If frmDocument.DVD.CurrentDomain = 3 Then
Ibutton = frmDocument.DVD.CurrentButton
If Val(Ibutton) >= 1 Then
frmDocument.DVD.SelectAndActivateButton (Ibutton)
End If
Exit Sub
End If
frmDocument.DVD.Play
DoEvents
Exit Sub
err: MsgBox "Please insert a DVD into the machine", vbOKOnly, "No DVD in the drive"
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Label7.Caption
Label7.ForeColor = vbYellow
End Sub

Private Sub Label8_Click()

frmDocument.DVD.Stop
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTimeTracker.Caption = Label8.Caption
Label8.ForeColor = vbYellow
End Sub

Private Sub LabMute_Click()
frmDocument.DVD.Mute = Not (frmDocument.DVD.Mute)
Shape2.Visible = Not Shape2.Visible
End Sub

Private Sub LabMute_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabMute.ForeColor = &H8000000B
End Sub

Private Sub LabMute_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabMute.ForeColor = &H8000000E
End Sub


Private Sub UserControl11_click(Value As Double)
Dim pos As String
    'If Button = Pgleft Then
        pos = Framesto_Times(Str(Int(Value)))
        frmDocument.DVD.PlayAtTime (pos & ":25")
    'End If

End Sub

'''''''''''The below section deals with the navigatorbar'''''''''''''''''''''''''''''''''''''''

'Private Sub UserControl11_MouseHover(Button As PgButton, X As Single, Y As Single, Value As Double)


'End Sub
