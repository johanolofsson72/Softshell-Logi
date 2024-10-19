Attribute VB_Name = "modGetSpecialFolderLocation"
Option Explicit

Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10

Private Declare Function SHGetFolderPath Lib "SHFolder" Alias "SHGetFolderPathA" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, _
ByVal dwFlags As Long, ByVal sPath As String) As Long


Private Const S_OK = &H0

Public Function FindNeedFolders(CSIDL As Long) As String

    Dim sPath As String * 255, l As Long
    Dim SW As Long

    FindNeedFolders = ""

    SW = SHGetFolderPath(0, CSIDL, 0&, 0&, sPath)
    
    If SW = S_OK Then
        l = InStr(sPath, Chr$(0))
        If (l > 0) And (l <= 255) Then FindNeedFolders = Left$(sPath, l - 1)
    End If

End Function
