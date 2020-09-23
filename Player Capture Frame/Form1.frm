VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movie Player and Frame's Capture"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcaptureFromTo 
      Caption         =   "Capture"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8655
      TabIndex        =   22
      ToolTipText     =   "Capture frame >From >To"
      Top             =   5265
      Width           =   900
   End
   Begin VB.ComboBox cmbTo 
      Height          =   330
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5280
      Width           =   1200
   End
   Begin VB.ComboBox cmbFrom 
      Height          =   330
      Left            =   5625
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5295
      Width           =   1305
   End
   Begin VB.CommandButton cmdRandomFrame 
      Caption         =   "Capture 6 Random Frame"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6075
      TabIndex        =   16
      ToolTipText     =   "Capture Random Frame..."
      Top             =   4455
      Width           =   2700
   End
   Begin VB.CheckBox checkNumFrame 
      Caption         =   "Use number of frame in the file's"
      Height          =   270
      Left            =   75
      TabIndex        =   15
      Top             =   5370
      Width           =   3900
   End
   Begin VB.Frame Frame2 
      Caption         =   "Captured Frames"
      Height          =   4335
      Left            =   6060
      TabIndex        =   7
      Top             =   60
      Width           =   3585
      Begin Project1.ShowImage imgs 
         Height          =   1080
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   330
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1905
         BackColor       =   -2147483640
      End
      Begin Project1.ShowImage imgs 
         Height          =   1095
         Index           =   1
         Left            =   1830
         TabIndex        =   9
         Top             =   330
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor       =   -2147483640
      End
      Begin Project1.ShowImage imgs 
         Height          =   1095
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   1470
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor       =   -2147483640
      End
      Begin Project1.ShowImage imgs 
         Height          =   1095
         Index           =   3
         Left            =   1830
         TabIndex        =   11
         Top             =   1470
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor       =   -2147483640
      End
      Begin Project1.ShowImage imgs 
         Height          =   1095
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   2625
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor       =   -2147483640
      End
      Begin Project1.ShowImage imgs 
         Height          =   1095
         Index           =   5
         Left            =   1830
         TabIndex        =   13
         Top             =   2625
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor       =   -2147483640
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "n.a"
         Height          =   510
         Left            =   120
         TabIndex        =   14
         Top             =   3780
         Width           =   3345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   75
      TabIndex        =   5
      Top             =   75
      Width           =   5955
      Begin Project1.MPXControl MPXControl1 
         Height          =   4005
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   7064
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3315
      TabIndex        =   4
      ToolTipText     =   "Play Movie..."
      Top             =   4455
      Width           =   870
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2115
      TabIndex        =   3
      ToolTipText     =   "Pause Movie..."
      Top             =   4455
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Media"
      Height          =   315
      Left            =   225
      TabIndex        =   2
      ToolTipText     =   "Open supported Media file's..."
      Top             =   4455
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Capture Frame"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4350
      TabIndex        =   1
      ToolTipText     =   "Capture this Frame of Movie..."
      Top             =   4455
      Width           =   1605
   End
   Begin VB.Label Label4 
      Caption         =   "To:"
      Height          =   240
      Left            =   6960
      TabIndex        =   21
      Top             =   5340
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Capture From:"
      Height          =   240
      Left            =   4140
      TabIndex        =   19
      Top             =   5340
      Width           =   1440
   End
   Begin VB.Label lblTotalFrame 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   240
      Left            =   6165
      TabIndex        =   17
      Top             =   4830
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "n.a"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4830
      Width           =   5925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Dim count_s As Integer

Private Sub cmdcaptureFromTo_Click()
    Dim fFrame As String: Dim dib As Long: Dim bOK As Long
    Dim imgFormat As FREE_IMAGE_FORMAT
    Dim sFileName As String: Dim fFileName As String
    Dim rFrame As Long: Dim x, n As Long: Dim i As Integer
    Dim currentFrame As Long
    
    On Local Error GoTo ErrorHandler
    
    If MPXControl1.FileName <> Empty And MPXControl1.IsPlay = True Then
    
    '// Memorize current Frame
    currentFrame = MPXControl1.Position
    
    
    
    Else
        MsgBox "Sorry, if the movie is in play the program crash!!!", vbInformation, App.Title
    End If
Exit Sub
ErrorHandler:
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub cmdRandomFrame_Click()
    Dim fFrame As String: Dim dib As Long: Dim bOK As Long
    Dim imgFormat As FREE_IMAGE_FORMAT
    Dim sFileName As String: Dim fFileName As String
    Dim rFrame As Long: Dim x, n As Long: Dim i As Integer
    Dim currentFrame As Long
    
    On Local Error GoTo ErrorHandler
    
    If MPXControl1.FileName <> Empty And MPXControl1.IsPlay = True Then
    
    '// Memorize current Frame
    currentFrame = MPXControl1.Position
    
    For i = 0 To 5
    
    '// Randomize Duration = Frames of the Movie
    x = RandomNumbers(lblTotalFrame.Caption, 1, 1, True)
    For n = LBound(x) To UBound(x)
      rFrame = x(n)
    Next n
    
    Call MPXControl1.GoToFrame(rFrame)
    
    '// START CAPTURE RANDOM FRAME ;)
    
    If count_s > 5 Then count_s = 0
    
    If checkNumFrame.Value = 1 Then
        fFrame = MPXControl1.Position
    Else
        fFrame = count_s
    End If
    
    '// Extract only the FileName
    sFileName = GetFilePath(MPXControl1.FileName, Only_FileName_no_Extension)
    
    '// Save picture as BMP
    Call MPXControl1.Capture(App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp")
    imgs(count_s).loadimg App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp"
            
    '// The complete FileName of picture
    fFileName = App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp"
            
    '/// Load a picture to convert
    dib = FreeImage_LoadEx(fFileName)
            
    '// Convert picture BMP to JPG
    Call FreeImage_SaveEx(dib, App.path + "\Frames\f_" + sFileName + "_" + fFrame, 2, FISO_BMP_DEFAULT, , , , False)
            
    '// Unload DLL FreeImage
    Call FreeImage_Unload(dib)
            
    '// Delete the BMP picture
    If Dir$(App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp") <> Empty Then _
    Kill (App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp")
            
    '// Write Filename to Tag Image of this Form
    imgs(count_s).Tag = App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".jpg"
    imgs(count_s).ToolTipText = "f_" + sFileName + "_" + fFrame & ".jpg"
            
    count_s = count_s + 1
    Label2.Caption = "Captured frame > " & fFrame & " /" & count_s
    '// END CAMPTURE RANDOM FRAME
    
    Next i
    
    '// Back to Init frame
    Call MPXControl1.GoToFrame(currentFrame)
    
    'MsgBox "Captured {Random} frame ok!", vbInformation, App.Title
    
    Else
        MsgBox "Sorry, if the movie is in play the program crash!!!", vbInformation, App.Title
    End If
Exit Sub
ErrorHandler:
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub Command1_Click()
    Dim fFrame As String: Dim dib As Long: Dim bOK As Long: Dim imgFormat As FREE_IMAGE_FORMAT
    Dim sFileName As String: Dim fFileName As String
    
    On Local Error GoTo ErrorHandler
    
    If MPXControl1.FileName <> "" Then
        If MPXControl1.IsPlay = True Then
        
            If count_s > 5 Then count_s = 0
            
            If checkNumFrame.Value = 1 Then
                fFrame = MPXControl1.Position
            Else
                fFrame = count_s
            End If
            
            '// Convert picture as JPG
            sFileName = GetFilePath(MPXControl1.FileName, Only_FileName_no_Extension)
            
            '// Save picture as BMP
            Call MPXControl1.Capture(App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp")
            imgs(count_s).loadimg App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp"
            
            '// The complete FileName of picture
            fFileName = App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp"
            
            '/// Load a picture to convert
            dib = FreeImage_LoadEx(fFileName)
            
            '// Convert picture BMP to JPG
            Call FreeImage_SaveEx(dib, App.path + "\Frames\f_" + sFileName + "_" + fFrame, 2, FISO_BMP_DEFAULT, , , , False)
            
            '// Unload DLL FreeImage
            Call FreeImage_Unload(dib)
            
            '// Delete the BMP picture
            If Dir$(App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp") <> Empty Then _
            Kill (App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".bmp")
            
            '// Write Filename to Tag Image of this Form
            imgs(count_s).Tag = App.path + "\Frames\f_" + sFileName + "_" + fFrame & ".jpg"
            imgs(count_s).ToolTipText = "f_" + sFileName + "_" + fFrame & ".jpg"
            
            count_s = count_s + 1
            Label2.Caption = "Captured frame > " & fFrame & " /" & count_s
            
        ElseIf MPXControl1.IsPlay = False Then
            MsgBox "Sorry, if the movie is in play the program crash!!!", vbInformation, App.Title
        End If
    End If
Exit Sub
ErrorHandler:
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Dim x As Double: x = 0
    On Local Error GoTo ErrorHandler
    Call MPXControl1.OpenFile(True)
    If MPXControl1.FileName <> "" Then
        Label1.Caption = GetFilePath(MPXControl1.FileName, Only_FileName_and_Extension)
        Command3.Enabled = True
        Command4.Enabled = True
        Command1.Enabled = True
        cmdRandomFrame.Enabled = True
        cmdcaptureFromTo.Enabled = True
        Command4.Caption = "Stop"
        
        cmbFrom.Clear
        cmbTo.Clear
        
        '// Adding Duration
        For i = 0 To MPXControl1.Duration
            cmbFrom.AddItem MPXControl1.SecToTime(x)
            cmbTo.AddItem MPXControl1.SecToTime(x)
            x = x + 1
        Next i
        
        If cmbFrom.ListCount > 0 Then cmbFrom.ListIndex = 0
        If cmbTo.ListCount > 0 Then cmbTo.ListIndex = 0
        
        lblTotalFrame.Caption = MPXControl1.Duration
        
    End If
Exit Sub
ErrorHandler:
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub Command3_Click()
    If MPXControl1.IsPlay = True Then
        MPXControl1.cmdPause
        cmdRandomFrame.Enabled = False
        Command4.Enabled = False
        Command1.Enabled = False
        cmdcaptureFromTo.Enabled = False
    ElseIf MPXControl1.IsPlay = False Then
        If MPXControl1.FileName <> "" Then
            MPXControl1.cmdPlay
            cmdRandomFrame.Enabled = True
            Command4.Enabled = True
            Command1.Enabled = True
            cmdcaptureFromTo.Enabled = True
        End If
    End If
End Sub

Private Sub Command4_Click()
    If MPXControl1.IsPlay = True Then
        
        If Command4.Caption = "Stop" Then
            Command4.Caption = "Play"
            MPXControl1.cmdPause
            cmdRandomFrame.Enabled = False
            Command3.Enabled = False
            Command1.Enabled = False
            cmdcaptureFromTo.Enabled = False
        ElseIf Command4.Caption = "Play" Then
            MPXControl1.cmdPlay
            cmdRandomFrame.Enabled = True
            Command3.Enabled = True
            Command1.Enabled = True
            cmdcaptureFromTo.Enabled = True
        End If
    ElseIf MPXControl1.IsPlay = False Then
        Command4.Caption = "Stop"
        MPXControl1.cmdPlay
        cmdRandomFrame.Enabled = True
        Command3.Enabled = True
        Command1.Enabled = True
        cmdcaptureFromTo.Enabled = True
    End If
End Sub
Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub MPXControl1_Click()
    Call FormMove(Form1)
End Sub


