Attribute VB_Name = "modShellExec"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As Any, _
    ByVal lpDirectory As Any, _
    ByVal nShowCmd As Long) As Long

Public Sub ShellStart(Path_File As String)
    Dim Xx As Long
    
    Xx = ShellEx(0, "open", Path_File, "", "", 1)
End Sub


