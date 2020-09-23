Attribute VB_Name = "modFunction"
Option Explicit

Public Enum Extract
    [Only_Extension] = 0
    [Only_FileName_and_Extension] = 1
    [Only_FileName_no_Extension] = 2
    [Only_Path] = 3
End Enum

Public stripMyString As String

Public MsgSizing

Public cdOpenFile As CommonDialog

Public gGraph As FilgraphManager
Public gRegFilters As Object
Public gCapStill As VBGrabber

Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hBitmap As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Mode As Long) As Long
Public Declare Sub DeleteDC Lib "GDI32" (ByVal hDC As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal Count As Long)
    
Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Public Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgsabound(0 To 1) As SAFEARRAYBOUND
End Type

Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long


Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Sub FormMove(ByVal TheObject As Object)
If MsgSizing <> vbNull Then
    ReleaseCapture
    SendMessage TheObject.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Function RandomNumber(Upper As Integer, Lower As Integer) As Integer
    Randomize: RandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
End Function

Public Function RandomNumbers(Upper As Integer, _
                    Optional Lower As Integer = 1, _
                    Optional HowMany As Integer = 1, _
                    Optional Unique As Boolean = True) As Variant
        Dim x As Integer: Dim n As Integer
        Dim arrNums() As Variant: Dim colNumbers As New Collection
        
        On Error GoTo ErrorHandler
    
    If HowMany > ((Upper + 1) - (Lower - 1)) Then Exit Function

    ReDim arrNums(HowMany - 1)
        With colNumbers

    For x = Lower To Upper
        .Add x
    Next x

    For x = 0 To HowMany - 1
        n = RandomNumber(0, colNumbers.Count + 1)
        arrNums(x) = colNumbers(n)

        If Unique Then
            colNumbers.Remove n
        End If
    Next x
    End With
    
    Set colNumbers = Nothing
        RandomNumbers = arrNums
Exit Function
ErrorHandler:
    RandomNumbers = ""
End Function

Public Function GetFilePath(ByVal FileName As String, strExtract As Extract) As String
    Select Case strExtract
        'Extract only extension of File
        Case 0
            GetFilePath = Mid$(FileName, InStrRev(FileName, ".", , vbTextCompare) + 1)
        'Extract only Filename and Extension
        Case 1
            GetFilePath = Mid$(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
        'Extract only FileName
        Case 2
            GetFilePath = StripUndo(Mid$(FileName, InStrRev(FileName, "\", , vbTextCompare) + 1))
        'Extract only Path
        Case 3
            GetFilePath = Mid$(FileName, 1, InStrRev(FileName, "\", , vbTextCompare) - 1)
        End Select
End Function

Private Function StripUndo(ByVal FileName As String) As String
    Dim i As Integer
    Dim stmp As String
    On Error Resume Next
stmp = Mid(FileName, i + 1, Len(FileName))
    For i = 1 To Len(stmp)
      If Mid(stmp, i, 1) = "." Then
        Exit For
    Else
        stripMyString = Mid(FileName, i + 2, Len(FileName))
    End If
Next
     StripUndo = Left(stmp, i - 1)
End Function
