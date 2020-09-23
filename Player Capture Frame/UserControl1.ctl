VERSION 5.00
Begin VB.UserControl MPXControl 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   FillStyle       =   0  'Solid
   LockControls    =   -1  'True
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   Begin VB.Timer Timerscroll 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10200
      Top             =   0
   End
   Begin VB.PictureBox Volume 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4320
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   7
      Top             =   0
      Width           =   1185
      Begin VB.Shape VolumeBar 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         FillColor       =   &H00C0C0C0&
         Height          =   135
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox Position 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   6
      Top             =   0
      Width           =   4125
      Begin VB.Shape Shapescroll 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         FillColor       =   &H00C0C0C0&
         Height          =   135
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox PicVid 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   15
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   368
      TabIndex        =   5
      Top             =   345
      Width           =   5520
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9840
      Top             =   0
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      LargeChange     =   10
      Left            =   3990
      Max             =   100
      TabIndex        =   2
      Top             =   7080
      Value           =   100
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      Left            =   2265
      Max             =   100
      TabIndex        =   1
      Top             =   7080
      Width           =   1710
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   195
      LargeChange     =   500
      Left            =   975
      Max             =   5000
      Min             =   -5000
      SmallChange     =   50
      TabIndex        =   0
      Top             =   7065
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00000000&
      Caption         =   "Balance: 0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   45
      TabIndex        =   10
      Top             =   4380
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   10785
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblVolume 
      BackColor       =   &H00000000&
      Caption         =   "Volume: 100"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4365
      TabIndex        =   8
      Top             =   135
      Width           =   1305
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H00000000&
      Caption         =   "Duration: 00:00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2205
      TabIndex        =   4
      Top             =   135
      Width           =   2145
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00000000&
      Caption         =   "Position: 00:00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   135
      Width           =   2145
   End
End
Attribute VB_Name = "MPXControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim File As String
Dim File1 As String

Private m_FileName As String
Private Const m_def_FileName As String = ""

Private m_Play As Boolean
Private Const m_def_Play As Boolean = False

Dim m_objBasicAudio As IBasicAudio
Dim m_objMediaControl As IMediaControl
Dim m_objMediaPosition As IMediaPosition
Dim m_objMediaEvent As IMediaEvent


Dim gGraph As IMediaControl
Dim gRegFilters As Object
Dim gCapStill As VBGrabber

Dim gVideofenster As IVideoWindow

Dim m_dblRate As Double
Dim m_dblRunLength As Double
Dim m_dblStartPosition As Double

Private m_dblFPS As Double
Private m_boolDirty As Boolean
Private m_nFrameCount As Long
Private m_boolHasAudio As Boolean
Private m_bstrFileName As String

Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public num As Long
Public menu As Boolean
Dim Scrolly As Integer
Private CD As CommonDialog

Dim DontMaintainRatio As Boolean

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Mode As Long) As Long
Private Declare Sub DeleteDC Lib "GDI32" (ByVal hDC As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal Count As Long)
    
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgsabound(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long


Private Sub lblExit_Click()
Unload UserControl.Parent
End Sub

Private Sub PicVid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
Call cmdTogglePlay
End If
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If UserControl.Enabled = False Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    m_FileName = m_def_FileName
    m_Play = m_def_Play
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace
            RaiseEvent Click
            Exit Sub
        Case vbKeyRight, vbKeyDown
            SendKeys "{Tab}" 'Shift The TabStop Forward
        Case vbKeyLeft, vbKeyUp
            SendKeys "+{Tab}" 'Shift The TabStop Backward
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = (x * Screen.TwipsPerPixelX)
    y = (y * Screen.TwipsPerPixelY)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = (x * Screen.TwipsPerPixelX)
    y = (y * Screen.TwipsPerPixelY)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ((x >= 0 And y >= 0) And (x < UserControl.ScaleWidth And y < UserControl.ScaleHeight)) Then
        If Button = vbLeftButton Then
            RaiseEvent Click
            Call cmdPlay
            If Button = 2 Then
                If menu = True Then
                    'Call PopupMenu(mnuMain)
                End If
            End If
        End If
        x = (x * Screen.TwipsPerPixelX)
        y = (y * Screen.TwipsPerPixelY)
        RaiseEvent MouseUp(Button, Shift, x, y)
    End If
End Sub

Private Sub PicVid_Paint()
    'Call Me.cmdCAspect
End Sub

Private Sub UserControl_Paint()
    'Call clDShow.DS_UpdateMovie
    'Call Me.cmdCAspect
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    m_Play = PropBag.ReadProperty("IsPlay", m_def_Play)
End Sub

Private Sub UserControl_Terminate()
    On Local Error Resume Next
    Timerscroll.Enabled = False
    tmrTimer.Enabled = False
    UnloadActiveMovieControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("IsPlay", m_Play, m_def_Play)
End Sub


Private Sub Volume_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim VolVal As Double
    On Error Resume Next
    Timerscroll.Enabled = False
    UserControl.VolumeBar.Width = x
    VolVal = (1000 / 100) * (UserControl.VolumeBar.Width / UserControl.Volume.Width) * 10
    SetVolume (VolVal)
End Sub
Private Sub Volume_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim VolVal As Double
On Error Resume Next
If Screen.MousePointer <> 0 Then Screen.MousePointer = 0

If Button = 1 Then
    VolVal = (1000 / 100) * (UserControl.VolumeBar.Width / UserControl.Volume.Width) * 10
    SetVolume (VolVal)
    UserControl.VolumeBar.Width = x
End If

End Sub
Private Sub Volume_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    UserControl.VolumeBar.Width = (m_objBasicAudio.Volume / 100) * UserControl.Volume.Width
    Timerscroll.Enabled = True
End Sub '


Private Sub Position_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
        UserControl.Timerscroll.Enabled = False
        Me.cmdPause
        UserControl.Shapescroll.Width = x
        m_objMediaPosition.CurrentPosition = (m_objMediaPosition.Duration / 100) * (UserControl.Shapescroll.Width / UserControl.Position.Width * 100)
End Sub
Private Sub Position_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Screen.MousePointer <> 0 Then Screen.MousePointer = 0

If Button = 1 Then
        UserControl.Shapescroll.Width = x
        m_objMediaPosition.CurrentPosition = (m_objMediaPosition.Duration / 100) * (UserControl.Shapescroll.Width / UserControl.Position.Width * 100)
End If
End Sub
Private Sub Position_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Timerscroll.Enabled = True
    Me.cmdPlay
End Sub


Private Sub Timerscroll_Timer()
On Error Resume Next

If UserControl.VolumeBar.Width <> (GetVolume / 100) * UserControl.Volume.Width Then
UserControl.VolumeBar.Width = (GetVolume / 100) * UserControl.Volume.Width
End If

If UserControl.Shapescroll.Width <> (Me.Position / Me.Duration) * (UserControl.Position.Width) Then
UserControl.Shapescroll.Width = (Me.Position / Me.Duration) * (UserControl.Position.Width)
End If

End Sub

Public Sub UserControl_Initialize()
    Set gGraph = New FilgraphManager
    Set gRegFilters = gGraph.RegFilterCollection
    num = 0
    Set CD = New CommonDialog
    CD.filter = "Supported Media Files|*.avi;*.asf;*.mpg;*.mpeg;*.wmv;*.divx;*.dat;*.mpx;*.mov;*.vob;*.flv|Divx/Avi files (divx avi)|*.avi;*.divx|Mpeg/MPG/Dat files (mpg mpeg dat)|*.mpg;*.mpeg;*.dat|Windows Media files (asf wmv)|*.asf;*.wmv|QuickTime file's (*.mov)|*.mov|DvD file (vob)|*.vob|Flash Video (flv)|*.flv"
    CD.DialogTitle = "Choose media to Play:"
End Sub


Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Position.Left = 0
'Position.Top = Me.Height - 270
UserControl.Shapescroll.Height = UserControl.Position.ScaleHeight
UserControl.Shapescroll.Top = 0
UserControl.Shapescroll.Left = 0
'Position.Top = Me.Height - 190
UserControl.Position.Width = UserControl.ScaleWidth - (60 + 5) - 0
UserControl.Volume.Left = UserControl.ScaleWidth - (60 + 8)
'Position.Top = 0
'Volume.Top = Me.Height - 190
UserControl.Volume.Width = UserControl.ScaleWidth - (UserControl.Position.ScaleWidth + 0) - UserControl.lblExit.Width
UserControl.VolumeBar.Width = UserControl.Volume.Width

UserControl.lblExit.Left = UserControl.ScaleWidth - UserControl.lblExit.Width

PicVid.Height = UserControl.ScaleHeight - 26 '/ Screen.TwipsPerPixelY
PicVid.Top = UserControl.ScaleTop + 26
PicVid.Width = UserControl.ScaleWidth '/ Screen.TwipsPerPixelX

End Sub

Private Sub UserControl_Unload(Cancel As Integer)

End Sub

Public Sub OpenFile(Optional DontMaintainRatio As Boolean = True)

    On Error GoTo ErrorHandler
    
    CD.filter = "Supported Media Files|*.avi;*.asf;*.mpg;*.mpeg;*.wmv;*.divx;*.dat;*.mpx;*mkv;*.mov;*.vob|Divx/Avi files (divx avi)|*.avi;*.divx|Mpeg/MPG/Dat files (mpg mpeg dat)|*.mpg;*.mpeg;*.dat|Windows Media files (asf wmv)|*.asf;*.wmv|Movie Player X Playlist (mpx)|*.mpx|Matroska (mkv)|*.mkv|Quick Time (mov)|*.mov|DvD file (vob)|*.vob"
    CD.DialogTitle = "Choose media to Play:"
    CD.ShowOpen
    File = CD.FileName
    
    If File = "" Then Exit Sub: If File = File1 Then Exit Sub: If m_Play = True Then cmdStop
    
    Set gGraph = Nothing
    Set gCapStill = Nothing
    Set gGraph = New FilgraphManager
    Set gRegFilters = gGraph.RegFilterCollection
    
    ' add the grabber including vb wrapper and default props
    Dim filter As IRegFilterInfo
    Dim fGrab As IFilterInfo
    For Each filter In gRegFilters
        If filter.Name = "SampleGrabber" Then
            filter.filter fGrab
            Set gCapStill = New VBGrabber
                gCapStill.FilterInfo = fGrab
            Exit For
        End If
    Next filter
    
    Dim fSrc As IFilterInfo
    
    gGraph.AddSourceFilter File, fSrc
    
    Dim pinOut As IPinInfo
    For Each pinOut In fSrc.Pins
        If pinOut.Direction = 1 Then
            Exit For
        End If
    Next pinOut
    
    ' find first input on grabber and connect
    Dim pinIn As IPinInfo
    For Each pinIn In fGrab.Pins
        If pinIn.Direction = 0 Then
                pinOut.Connect pinIn
            Exit For
        End If
    Next pinIn
    
    ' find grabber output pin and render
    For Each pinOut In fGrab.Pins
        If pinOut.Direction = 1 Then
                pinOut.Render
            Exit For
        End If
    Next pinOut
    
    Dim sScale As Double
    Dim topMod As Long
    
    Set gVideofenster = gGraph
    
    sScale = gVideofenster.Height / gVideofenster.Width
    
    If Not (DontMaintainRatio) Then
        gVideofenster.Height = PicVid.Height * sScale
        topMod = (PicVid.Height - gVideofenster.Height) / 2
    Else
        gVideofenster.Height = PicVid.Height
    End If
    
    gVideofenster.Top = topMod
    gVideofenster.Left = 0
    gVideofenster.Width = PicVid.Width
    gVideofenster.WindowStyle = CLng(&H6000000)
    
    gVideofenster.Owner = PicVid.hWnd
    
    Set m_objMediaControl = gGraph
    Set m_objMediaEvent = m_objMediaControl
    Set m_objMediaPosition = m_objMediaControl
    
    m_objMediaPosition.Rate = 1
    m_dblRate = m_objMediaPosition.Rate
    m_dblRunLength = Round(m_objMediaPosition.Duration, 2)
    
    HScroll2.Max = m_objMediaPosition.Duration
    
    Set m_objMediaControl = gGraph
    Call m_objMediaControl.RenderFile(File)
    
    Set gVideofenster = gGraph
    
    gVideofenster.Owner = PicVid.hWnd
    
    Set m_objBasicAudio = m_objMediaControl
    
    m_objBasicAudio.Volume = 0
    m_objBasicAudio.Balance = 0
    
    gGraph.Run
    
    PicVid.Visible = True
    
    If HScroll1.Value < 100 Then m_objBasicAudio.Volume = HScroll1.Value
    
    lblDuration.Caption = "Duration: " & SecToTime(m_objMediaPosition.Duration)
    lblPosition.Caption = "Position: " & SecToTime(m_objMediaPosition.CurrentPosition)
    lblVolume.Caption = "Volume: " & GetVolume
    
    Form1.cmbFrom.AddItem SecToTime(m_objMediaPosition.Duration)
    
    File1 = File
    tmrTimer.Enabled = True
    m_FileName = File: PropertyChanged "FileName"
    m_Play = True: PropertyChanged "IsPlay"
    Timerscroll.Enabled = True
    
    
    Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub
Private Sub PicVid_DblClick()
    num = num + 1
End Sub

Public Function Duration() As Long
    On Local Error Resume Next
    If m_Play = True Then _
    Duration = m_objMediaPosition.Duration
End Function

Public Function Position() As Long
    On Local Error Resume Next
    If m_Play = True Then _
    Position = m_objMediaPosition.CurrentPosition
End Function

Private Sub tmrTimer_Timer()
On Error GoTo ErrEnd
    lblDuration.Caption = "Duration: " & SecToTime(m_objMediaPosition.Duration)
    lblPosition.Caption = "Position: " & SecToTime(m_objMediaPosition.CurrentPosition)
    lblVolume.Caption = "Volume: " & GetVolume
    HScroll1.Value = m_objBasicAudio.Volume
ErrEnd:
End Sub

Public Function cmdTogglePlay()
If m_Play = False Then
    m_objMediaControl.Run
    m_Play = True
    PropertyChanged "IsPlay"
    Exit Function
End If

If m_Play = True Then
    m_objMediaControl.Pause
    m_Play = False
    PropertyChanged "IsPlay"
End If
End Function

Public Function cmdPlay()
    m_objMediaControl.Run
    m_Play = True
    PropertyChanged "IsPlay"
    Timerscroll.Enabled = True
End Function

Public Sub cmdPause()
    m_objMediaControl.Pause
    m_Play = False
    Timerscroll.Enabled = False
    PropertyChanged "IsPlay"
End Sub

Public Sub cmdStop()
    m_objMediaControl.Stop
    m_Play = False
    Timerscroll.Enabled = False
    PicVid.Cls
    PropertyChanged "IsPlay"
End Sub

Public Sub cmdSubVol()
    m_objBasicAudio.Volume = m_objBasicAudio.Volume - 10
End Sub

Public Sub cmdAddVol()
    'raise Volume
    m_objBasicAudio.Volume = m_objBasicAudio.Volume + 10
End Sub

Public Sub cmdBalCen()
    m_objBasicAudio.Balance = 0
End Sub

Public Sub cmdBalLft()
    'balance sound to left only
    m_objBasicAudio.Balance = -10000
End Sub

Public Sub cmdBalRgt()
    'balance sound to right only
    m_objBasicAudio.Balance = 10000
End Sub


Private Sub HScroll1_Change()
'Volume
    m_objBasicAudio.Volume = CStr(Trim(Str(HScroll1.Value)))
End Sub
Private Sub HScroll2_Scroll()
'Seek
On Local Error Resume Next
m_objMediaPosition.CurrentPosition = CStr(Trim(Str(HScroll2.Value)))
End Sub
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal newFileName As String)
    m_FileName = newFileName
    PropertyChanged "FileName"
End Property

Public Property Get IsPlay() As Boolean
   IsPlay = m_Play
End Property

Public Property Let IsPlay(ByVal NewPlay As Boolean)
    m_Play = NewPlay
    PropertyChanged "IsPlay"
End Property

Public Function Capture(sFile As String)
    gCapStill.FileName = sFile
    gCapStill.CaptureStill
    Timerscroll.Enabled = True
End Function

Public Function SecToTime(NewSec As Double) As String
    On Error Resume Next
        Dim Secx, MinX, Hourx
        NewSec = Int(NewSec)
        If NewSec < 1 Then SecToTime = "00:00:00": Exit Function
            Secx = NewSec - Int(NewSec / 60) * 60
            MinX = Int((NewSec - Int(NewSec / 3600) * 3600) / 60)
            Hourx = Int(NewSec / 3600)
        If Int(Hourx) > 24 Then
            SecToTime = "24:59:59"
        Else
            SecToTime = Format(Str(Hourx) & ":" & Str(MinX) & ":" & Str(Secx), "hh:mm:ss")
        End If
End Function

Private Function SetVolume(newVolume As Long)
    On Error GoTo ErrEnd
    If newVolume > 100 Then newVolume = 100
    If newVolume < 0 Then newVolume = 0
    If ObjPtr(m_objBasicAudio) > 0 Then m_objBasicAudio.Volume = (newVolume * 100) - 10000
Exit Function

ErrEnd:
End Function

Private Function GetVolume() As Long
    On Error GoTo ErrEnd
    If ObjPtr(m_objBasicAudio) > 0 Then GetVolume = (m_objBasicAudio.Volume + 10000) / 100
    Timerscroll.Enabled = True
ErrEnd:
End Function

Public Function PreRollTime() As Long
    On Local Error Resume Next
    If m_Play = True Then _
    PreRollTime = m_objMediaPosition.PreRollTime
End Function

Public Function Rate() As Long
    On Local Error Resume Next
    If m_Play = True Then _
    Rate = m_objMediaPosition.Rate
End Function

Public Function UnloadActiveMovieControl()
    On Local Error GoTo ErrLine
    
    DoEvents

    If Not m_objMediaControl Is Nothing Then
        m_objMediaControl.Stop
    End If
    
    If Not m_objBasicAudio Is Nothing Then Set m_objBasicAudio = Nothing
    If Not m_objMediaControl Is Nothing Then Set m_objMediaControl = Nothing
    If Not m_objMediaPosition Is Nothing Then Set m_objMediaPosition = Nothing
    If Not m_objMediaEvent Is Nothing Then Set m_objMediaEvent = Nothing
    Exit Function
            
ErrLine:
    Err.Clear
End Function

Public Function GoToFrame(ByVal newFrame As Long)
    On Local Error Resume Next
    m_objMediaPosition.CurrentPosition = newFrame
End Function
