Attribute VB_Name = "modMediaInfo"
Option Explicit

'/// Open/Save File Dialog
'/// *****************************************************************
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'/// Public Enum for MediaInfo.dll
'/// *****************************************************************
Public Enum stream_t_C
  Stream_General
  Stream_Video
  Stream_Audio
  Stream_Text
  Stream_Chapters
  Stream_Image
  Stream_Max
End Enum

Public Enum info_t_C
  Info_Name
  Info_Text
  Info_Measure
  Info_Options
  Info_Name_Text
  Info_Measure_Text
  Info_Info
  Info_HowTo
  Info_Max
End Enum

Public Enum infooptions_t_C
  InfoOption_ShowInInform
  InfoOption_Support
  InfoOption_ShowInSupported
  InfoOption_TypeOfValue
  InfoOption_Max
End Enum

Public Enum informoptions_t_C
  InformOption_Nothing
  InformOption_Custom
  InformOption_HTML
  InformOption_Max
End Enum

'/// Public Declaration for MediaInfo.dll
'/// *****************************************************************
Public Declare Sub MediaInfo_Close Lib "MediaInfo.dll" (ByVal Handle As Long)
Public Declare Function MediaInfo_Count_Get Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal StreamKind As stream_t_C, ByVal StreamNumber As Long) As Long
Public Declare Function MediaInfo_Get Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal StreamKind As stream_t_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal InfoKind As info_t_C, ByVal SearchKind As info_t_C) As Long
Public Declare Function MediaInfo_GetI Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal StreamKind As stream_t_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal InfoKind As info_t_C) As Long
Public Declare Function MediaInfo_Inform Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal Options As informoptions_t_C) As Long
Public Declare Function MediaInfo_Open Lib "MediaInfo.dll" (ByVal File As Long) As Long
Public Declare Function MediaInfo_Open_Buffer Lib "MediaInfo.dll" (Begin As Any, ByVal Begin_Size As Long, End_ As Any, ByVal End_Size As Long) As Long
Public Declare Function MediaInfo_Option Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal Option_ As Long, ByVal Value As Long) As Long
Public Declare Function MediaInfo_Save Lib "MediaInfo.dll" (ByVal Handle As Long) As Long
Public Declare Function MediaInfo_Set Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal ToSet As Long, ByVal StreamKind As stream_t_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal OldParameter As Long) As Long
Public Declare Function MediaInfo_SetI Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal ToSet As Long, ByVal StreamKind As stream_t_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal OldParameter As Long) As Long
Public Declare Function MediaInfo_State_Get Lib "MediaInfo.dll" (ByVal Handle As Long) As Long

'/// Service Declaration APIs
'/// *****************************************************************
Private Declare Function lstrlenW Lib "kernel32" (ByVal pStr As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bLen As Long)

Public Function StripStrinCtoVB(ptr As Long) As String
  Dim l As Long
  On Local Error Resume Next
  '/// Convert C value to VB value
  '/// *****************************************************************
  l = lstrlenW(ptr)
  StripStrinCtoVB = String$(l, vbNullChar)
  RtlMoveMemory ByVal StrPtr(StripStrinCtoVB), ByVal ptr, l * 2
End Function

Public Function DialogOpenFile(strFilter As String, Optional InitialDir As String = "C:\") As String
    Dim ofn As OPENFILENAME: Dim a
    On Local Error GoTo ErrorHandler
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = strFilter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitialDir
    ofn.lpstrTitle = "Select Media File:"
    ofn.flags = 0
    a = GetOpenFileName(ofn)
    If (a) Then
        DialogOpenFile = Trim$(ofn.lpstrFile)
    Else
        DialogOpenFile = ""
    End If
Exit Function
ErrorHandler:
        DialogOpenFile = ""
    Err.Clear
End Function

Public Function DialogSaveAs(Optional InitialDir As String = "C:\", Optional strFileName As String = "") As String
    Dim ofn As OPENFILENAME: Dim a
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "File Log (*.log)" & Chr$(0) & "*.log" & Chr$(0) + "File Txt (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) _
                                    + "File Note (*.ntd)" + Chr$(0) + "*.ntd" + Chr$(0)
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitialDir
    ofn.lpstrTitle = "Save file as:"
    ofn.flags = 6
    ofn.lpstrFileTitle = strFileName
    ofn.lpstrDefExt = ".txt"
    a = GetSaveFileName(ofn)
    If (a) Then
        DialogSaveAs = Trim$(ofn.lpstrFile)
    Else
        DialogSaveAs = ""
    End If
Exit Function
ErrorHandler:
        DialogSaveAs = ""
    Err.Clear
End Function
