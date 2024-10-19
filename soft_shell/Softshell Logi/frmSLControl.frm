VERSION 5.00
Begin VB.Form frmSLControl 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer MsgBoxTimer 
      Left            =   3240
      Top             =   1080
   End
   Begin VB.PictureBox picBlackArrow 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Picture         =   "frmSLControl.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   960
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox picWhiteArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2880
      Picture         =   "frmSLControl.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   240
      Visible         =   0   'False
   End
   Begin VB.PictureBox picPro 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2880
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox PicPro2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   600
      Width           =   300
   End
   Begin VB.PictureBox PicSub 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   300
      Width           =   2600
   End
   Begin VB.PictureBox PicSub 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   600
      Width           =   2600
   End
   Begin VB.PictureBox PicSub 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   900
      Width           =   2600
   End
   Begin VB.PictureBox PicSub 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   1200
      Width           =   2600
   End
   Begin VB.PictureBox picPublicity 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   1500
      Width           =   2600
   End
   Begin VB.PictureBox picHead 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2600
   End
End
Attribute VB_Name = "frmSLControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurY As Long
Private CurYIcon As Long
Private OldIndex As Long
Private CurIndex As Long

Private Type PicSubIcon
    Iconpath As String
    Dirtype As Boolean
End Type

Private PicSubI(4) As PicSubIcon
Public MsgTitle As String


Private Sub Form_Load()
  
  OldIndex = 0
  Call PaintSubMenu
  
End Sub

Public Sub DrawI(size As Long, DestPic As PictureBox, path As String)

picPro.Cls
PicPro2.Cls
picPro.Picture = LoadPicture(path)
BitBlt PicPro2.hdc, 2, 2, 20, 20, picPro.hdc, 0, 0, vbSrcCopy
BitBlt DestPic.hdc, 0, CurYIcon, 20, 20, PicPro2.hdc, 0, 0, vbSrcCopy

End Sub

Private Sub PaintSubMenu()

CurY = ((PicSub(0).Height) - PicSub(0).TextHeight("|")) / 2

'//--Control Panel
PicSub(0).CurrentY = CurY
PicSub(0).Tag = Space(8) & "ControlPanel" 'GetSetting("Softshell Logi", "Startmenu Settings", "Run", "Run...")
PicSub(0).Print PicSub(0).Tag
PicSubI(0).Dirtype = False
PicSubI(0).Iconpath = App.path & "\Icons\ConfigIcon.ico"
Call DrawI(CurYIcon, PicSub(0), PicSubI(0).Iconpath)
'//--SLControl Settings
PicSub(1).CurrentY = CurY
PicSub(1).Tag = Space(8) & "SLControl Settings" 'GetSetting("Softshell Logi", "Startmenu Settings", "Control Devices", "Control Devices...")
PicSub(1).Print PicSub(1).Tag
PicSubI(1).Dirtype = False
PicSubI(1).Iconpath = App.path & "\Icons\ConfigIcon.ico"
Call DrawI(CurYIcon, PicSub(1), PicSubI(1).Iconpath)
'//--Update Softshell Logi
PicSub(2).CurrentY = CurY
PicSub(2).Tag = Space(8) & "Update Softshell Logi" 'GetSetting("Softshell Logi", "Startmenu Settings", "Recycled", "Recycled...")
PicSub(2).Print PicSub(2).Tag
PicSubI(2).Dirtype = False
PicSubI(2).Iconpath = App.path & "\Icons\ConfigIcon.ico"
Call DrawI(CurYIcon, PicSub(2), PicSubI(2).Iconpath)
'//--Return to Explorer
PicSub(3).CurrentY = CurY
PicSub(3).Tag = Space(8) & "Return to Explorer" 'GetSetting("Softshell Logi", "Startmenu Settings", "Shut-Down", "Shut-Down...")
PicSub(3).Print PicSub(3).Tag
PicSubI(3).Dirtype = False
PicSubI(3).Iconpath = App.path & "\Icons\ConfigIcon.ico"
Call DrawI(CurYIcon, PicSub(3), PicSubI(3).Iconpath)
        


  '//--Set the Value for Head and Publicity-------------------------------------------------------------------------
  
    picHead.Tag = GetSetting("Softshell Logi", "Startmenu Settings", "Control Devices", "Control Devices")

    picPublicity.Tag = GetSetting("Softshell Logi", "Startmenu Settings", "SubPublicityCaption", "®" & Space(1) & "Softworld" & Space(1) & "™")

    
    picHead.CurrentY = CurY
    picHead.CurrentX = 10
    picHead.Print Space(2) & picHead.Tag
    picHead.Top = 0
    picHead.Visible = True
    picHead.Width = Me.ScaleWidth
    
    picPublicity.CurrentY = CurY
    picPublicity.CurrentX = ((picHead.Width) / 2) - 650
    picPublicity.Print picPublicity.Tag
    picPublicity.Top = (PicSub(3).Top + PicSub(3).Height)
    picPublicity.Visible = True
    picPublicity.Width = Me.ScaleWidth


'//--Seperator lines on the Menu-------------------------------------------------------------------------------
    '//--Line Upper...........
    
    picHead.ForeColor = vb3DDKShadow
    picHead.Line (1, picHead.Top + (picHead.Height - 30))-(picHead.Width - 1, picHead.Top + (picHead.Height - 30))
    picHead.ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    picHead.Line (1, picHead.Top + (picHead.Height - 20))-(picPublicity.Width - 1, picHead.Top + (picHead.Height - 20))
    
    '//--Line Down...........
    picPublicity.ScaleMode = 3
    picPublicity.ForeColor = vb3DDKShadow
    picPublicity.Line (1, 0)-(picPublicity.Width - 1, 0)
    picPublicity.ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    picPublicity.Line (1, 1)-(picPublicity.Width - 1, 1)
    picPublicity.ScaleMode = 1
'//-------------------------------------------------------------------------------------------

           
End Sub

Public Sub KillSLControl()
    Call s_Playsound("select")
    frmSLControl.Hide
    Unload frmSLControl
End Sub


Private Sub MsgBoxTimer_Timer()
    Dim hWnd As Long
        MsgBoxTimer.Enabled = False
        hWnd = FindWindow(vbNullString, MsgTitle)
    Call SendMessage(hWnd, WM_CLOSE, 0, ByVal 0&)
End Sub

Private Sub PicSub_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Va As Long
 
    Select Case Index
    
        Case 0 '//--ControlPanel
            Call frmSubmenu.KillSubMenu
            Call KillSLControl
            Va = Shell("rundll32.exe shell32.dll,Control_RunDLL", vbNormalFocus)
            
        Case 1 '//--SLControl Settings
            Call frmSubmenu.KillSubMenu
            Call KillSLControl
            frmControl.Show
        Case 2 '//--Update Softshell Logi
            Call frmSubmenu.KillSubMenu
            Call KillSLControl
            Call frmTaskbar.UnloadLogi
    
            MsgTitle = "Update and Reboot..."
            With MsgBoxTimer
                .Interval = 2000
                .Enabled = True
            End With
            MsgBox "This Operation will only take a few seconds.", , MsgTitle
            Call modShellExec.ShellStart(App.path & "\Logi.exe")
        Case 3 '//--Return to Explorer
            Call frmSubmenu.KillSubMenu
            Call KillSLControl
            frmSwap.Show
    End Select
End Sub

Private Sub picsub_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index <> OldIndex Then

    Call s_Playsound("hover")


    
    '//--Reset the old selection ----------------------------------------------------------------------------
    picPro.BackColor = vbButtonFace: PicPro2.BackColor = vbButtonFace
    PicSub(OldIndex).Cls
    PicSub(OldIndex).BackColor = vbButtonFace
    PicSub(OldIndex).ForeColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
    PicSub(OldIndex).CurrentY = CurY
    PicSub(OldIndex).Print PicSub(OldIndex).Tag
    If PicSubI(OldIndex).Dirtype = False Then
        Call DrawI(CurYIcon, PicSub(OldIndex), PicSubI(OldIndex).Iconpath)
    Else
        BitBlt PicSub(OldIndex).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
        Call DrawI(CurYIcon, PicSub(OldIndex), PicSubI(OldIndex).Iconpath)
    End If
    
        '//--Highlite new selection ---
    picPro.BackColor = vbHighlight: PicPro2.BackColor = vbHighlight
    PicSub(Index).Cls
    PicSub(Index).BackColor = vbHighlight
    PicSub(Index).Line (18, 1)-(PicSub(Index).ScaleWidth, 1), GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    PicSub(Index).Line (18, PicSub(Index).ScaleHeight - 20)-(PicSub(Index).ScaleWidth, PicSub(Index).ScaleHeight - 20), vbButtonShadow
    PicSub(Index).CurrentX = 0
    PicSub(Index).ForeColor = vbHighlightText
    PicSub(Index).CurrentY = CurY
    PicSub(Index).Print PicSub(Index).Tag
    
    If PicSubI(Index).Dirtype = False Then
        Call DrawI(CurYIcon, PicSub(Index), PicSubI(Index).Iconpath)
    Else
        BitBlt PicSub(Index).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 16, 16, picWhiteArrow.hdc, 0, 0, vbSrcCopy
        Call DrawI(CurYIcon, PicSub(Index), PicSubI(Index).Iconpath)
    End If

    '//--Draw the Buttonface on the icon----------------------------------------------------------
    PicPro2.ScaleMode = 3
    PicPro2.ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    PicPro2.Line (0, 0)-(20, 0) 'Upper
    PicPro2.Line (0, 0)-(0, 19) 'Left
    PicPro2.ForeColor = vbButtonShadow
    PicPro2.Line (0, 19)-(20, 19) 'Buttom
    'PicPro2.Line (19, 0)-(19, 19)'Right
    BitBlt PicSub(Index).hdc, 0, CurYIcon, 20, 20, PicPro2.hdc, 0, 0, vbSrcCopy
    PicPro2.ScaleMode = 1
    

    OldIndex = Index
    Call Delay1(300)
    If OldIndex <> Index Then Exit Sub
    
    CurIndex = Index

End If
OldIndex = Index

End Sub

