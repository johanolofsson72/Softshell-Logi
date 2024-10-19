Attribute VB_Name = "modErrFix"
Option Explicit

'**************************************************
'Thank´s to Jeffrey C Tatum for the "Ini read and Ini Write"
'**************************************************
'INI Read and Write
'**************************************************

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
    As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpRetunedString As String, ByVal nSize As Long, _
    ByVal lpFilename As String) As Long


Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Public Function INIRead(iAppName As String, iKeyName As String, iFileName As String) As String
    'Example:
    'x = INIRead("boot", "shell", "C:\WINDOW
    '     S\system.ini")
    Dim iStr As String
    iStr = String(255, Chr(0))
    INIRead = Left(iStr, GetPrivateProfileString(iAppName, ByVal iKeyName, "", iStr, Len(iStr), iFileName))
End Function


Public Function INIWrite(iAppName As String, iKeyName As String, iKeyString As String, iFileName As String)
    'Example:
    'x = INIWrite("boot", "shell", "Explorer
    '     .exe", "C:\WINDOWS\system.ini")
    Dim R As Long
    R = WritePrivateProfileString(iAppName, iKeyName, iKeyString, iFileName)
End Function

Public Sub ErrFixing()
MsgBox "Fel"
End Sub


