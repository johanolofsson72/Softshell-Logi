Attribute VB_Name = "modStartUp"


Option Explicit
Public StartmenuPath          As String
Public DesktopPath             As String
Public FavoritesPath            As String
Public Mydocuments            As String
Public QuickmenuPath         As String

Public DoublePath               As Boolean
Public DoublepathString      As String
Public SubMenuShow          As Boolean
Public QuickMenuShow        As Boolean

Public FormFade                 As String

Public Aleft                        As Long
Public Aindex                     As Long
Public Const Abottom         As Long = 400
Public Acaption                  As String


Public MouseOn                 As Boolean

Sub Main()

    HideExplorerTaskBar
    
    frmTaskbar.Show
    
End Sub

