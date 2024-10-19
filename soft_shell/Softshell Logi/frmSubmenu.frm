VERSION 5.00
Begin VB.Form frmSubmenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2490
   ControlBox      =   0   'False
   Icon            =   "frmSubmenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   14
      Top             =   0
      Width           =   2300
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
      Left            =   200
      ScaleHeight     =   270
      ScaleWidth      =   2295
      TabIndex        =   13
      Top             =   3000
      Width           =   2300
   End
   Begin VB.PictureBox PicPro2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2520
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   1440
      Width           =   300
   End
   Begin VB.PictureBox picWhiteArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2640
      Picture         =   "frmSubmenu.frx":08CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   240
      Visible         =   0   'False
   End
   Begin VB.PictureBox picBlackArrow 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Picture         =   "frmSubmenu.frx":0E54
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   1920
      Width           =   225
      Visible         =   0   'False
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
      Index           =   8
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   9
      Top             =   2700
      Width           =   2300
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
      Index           =   7
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   8
      Top             =   2400
      Width           =   2300
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
      Index           =   6
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   7
      Top             =   2100
      Width           =   2300
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
      Index           =   5
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   6
      Top             =   1800
      Width           =   2300
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
      Index           =   4
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   5
      Top             =   1500
      Width           =   2300
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
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   4
      Top             =   1200
      Width           =   2300
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
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   3
      Top             =   900
      Width           =   2300
   End
   Begin VB.PictureBox picPro 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2520
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   120
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
      Index           =   1
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   600
      Width           =   2300
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
      Left            =   200
      ScaleHeight     =   300
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   300
      Width           =   2300
   End
   Begin VB.Image Image1 
      Height          =   3270
      Left            =   0
      Picture         =   "frmSubmenu.frx":13DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
End
Attribute VB_Name = "frmSubmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurY As Long
Private CurYIcon As Long
Private OldIndex As Long
Private CurIndex As Long

Private Type PicSubIcon
    Iconpath As String
    Dirtype As Boolean
End Type

Private PicSubI(8) As PicSubIcon

Private Sub Form_Load()
  '//--Change index to one of items below the dir´s on startmenu...
  CurIndex = 4
  OldIndex = 4
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

'//--Programs
PicSub(0).CurrentY = CurY
PicSub(0).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Programs", "Programs")
PicSub(0).Print PicSub(0).Tag
PicSubI(0).Dirtype = True
BitBlt PicSub(0).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
PicSubI(0).Iconpath = App.path & "\Icons\ProgramIcon.ico"
Call DrawI(CurYIcon, PicSub(0), PicSubI(0).Iconpath)
'//--Favorites
PicSub(1).CurrentY = CurY
PicSub(1).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Favorites", "Favorites")
PicSub(1).Print PicSub(1).Tag
PicSubI(1).Dirtype = True
BitBlt PicSub(1).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
PicSubI(1).Iconpath = App.path & "\Icons\FavoritesIcon.ico"
Call DrawI(CurYIcon, PicSub(1), PicSubI(1).Iconpath)
'//--Desktop
PicSub(2).CurrentY = CurY
PicSub(2).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Desktop", "Desktop")
PicSub(2).Print PicSub(2).Tag
PicSubI(2).Dirtype = True
BitBlt PicSub(2).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
PicSubI(2).Iconpath = App.path & "\Icons\DesktopIcon.ico"
Call DrawI(CurYIcon, PicSub(2), PicSubI(2).Iconpath)
'//--My Documents
PicSub(3).CurrentY = CurY
PicSub(3).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "My Documents", "My Documents")
PicSub(3).Print PicSub(3).Tag
PicSubI(3).Dirtype = True
BitBlt PicSub(3).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
PicSubI(3).Iconpath = App.path & "\Icons\MydocIcon.ico"
Call DrawI(CurYIcon, PicSub(3), PicSubI(3).Iconpath)
'//--Find
PicSub(4).CurrentY = CurY
PicSub(4).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Find", "Find...")
PicSub(4).Print PicSub(4).Tag
PicSubI(4).Dirtype = False
PicSubI(4).Iconpath = App.path & "\Icons\findIcon.ico"
Call DrawI(CurYIcon, PicSub(4), PicSubI(4).Iconpath)
'//--Run
PicSub(5).CurrentY = CurY
PicSub(5).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Run", "Run...")
PicSub(5).Print PicSub(5).Tag
PicSubI(5).Dirtype = False
PicSubI(5).Iconpath = App.path & "\Icons\runIcon.ico"
Call DrawI(CurYIcon, PicSub(5), PicSubI(5).Iconpath)
'//--Control Devices
PicSub(6).CurrentY = CurY
PicSub(6).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Control Devices", "Control Devices")
PicSub(6).Print PicSub(6).Tag
PicSubI(6).Dirtype = True
BitBlt PicSub(6).hdc, PicSub(0).ScaleWidth / Screen.TwipsPerPixelY - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
PicSubI(6).Iconpath = App.path & "\Icons\ConfigIcon.ico"
Call DrawI(CurYIcon, PicSub(6), PicSubI(6).Iconpath)
'//--Recycled
PicSub(7).CurrentY = CurY
PicSub(7).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Recycled", "Recycled...")
PicSub(7).Print PicSub(7).Tag
PicSubI(7).Dirtype = False
PicSubI(7).Iconpath = App.path & "\Icons\RecycledEmptyIcon.ico"
Call DrawI(CurYIcon, PicSub(7), PicSubI(7).Iconpath)
'//--Shut-Down
PicSub(8).CurrentY = CurY
PicSub(8).Tag = Space(8) & GetSetting("Softshell Logi", "Startmenu Settings", "Shut-Down", "Shut-Down...")
PicSub(8).Print PicSub(8).Tag
PicSubI(8).Dirtype = False
PicSubI(8).Iconpath = App.path & "\Icons\ShutdownIcon.ico"
Call DrawI(CurYIcon, PicSub(8), PicSubI(8).Iconpath)
        


  '//--Set the Value for Head and Publicity-------------------------------------------------------------------------
  
    picHead.Tag = GetSetting("Softshell Logi", "Startmenu Settings", "SubHeadCaption", "Softshell Logi")

    picPublicity.Tag = GetSetting("Softshell Logi", "Startmenu Settings", "SubPublicityCaption", "®" & Space(1) & "Softworld" & Space(1) & "™")

    
    picHead.CurrentY = CurY
    picHead.CurrentX = 10
    picHead.Print Space(2) & picHead.Tag & "..."
    picHead.Top = 0
    picHead.Visible = True
    picHead.Width = Me.ScaleWidth
    
    picPublicity.CurrentY = CurY
    picPublicity.CurrentX = ((picHead.Width) / 2) - 650
    picPublicity.Print picPublicity.Tag
    picPublicity.Top = (PicSub(8).Top + PicSub(8).Height)
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

Public Sub KillSubMenu()
    
    Call s_Playsound("select")
    Unload frmMenuSystem
    frmSubmenu.Hide
    Unload frmSubmenu
    Unload frmSLControl
    frmTaskbar.StartButton.BorderStyle = 1
    frmTaskbar.QuickButton.BorderStyle = 1
    SubMenuShow = False
End Sub

Private Sub PicSub_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Va As Long
    Select Case Index
    
        Case 0 '//--Programs
            '//--Nothing to do..
            
        Case 1 '//--Favorites
            '//--Nothing to do..
            
        Case 2 '//--Desktop
            '//--Nothing to do..
            
        Case 3 '//--My Documents
            '//--Nothing to do..
            
        Case 4 '//--Find
            Call KillSubMenu
            
            'MsgTitle = "Show the find dialog..."
            'With frmSLControl.MsgBoxTimer
            '    .Interval = 2000
            '    .Enabled = True
            'End With
            'MsgBox "This will only work if you run Softshell with Explorer!!!!.", , MsgTitle

            Call keybd_event(VK_LWIN, 0, 0, 0): Call keybd_event(VK_ACTION, 0, 0, 0): Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
        
        Case 5 '//--Run
            Call KillSubMenu
            DisplayRunDialog 0, "Softshell Logi Run-Device...", "Ok..Type the Name of the program you wish to start-up or use the browse-button to find it...", True, True
        
        Case 6 '//--Control Devices
            '//--Nothing to do..
            
        Case 7 '//--Recycled
            Call KillSubMenu
            Call modShellExec.ShellStart(Left$(App.path, 1) & ":\recycled") '//--Modul
        
        Case 8 '//--Shut-down
            Call KillSubMenu
            '//--Show the form with Effect´s...
            Call MakeFormEffect("win98", frmAutoExitWindows, 10, 1)
        
    End Select
End Sub

Private Sub picsub_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index <> OldIndex Then

    If Index = 4 Or Index = 5 Or Index = 7 Or Index = 8 Then Call s_Playsound("hover")
    If Index <> 6 Then Call frmSLControl.Hide

    
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
    '//--Set the new HeadCaption---------------------------------------------------------------------------------
    Acaption = Space(2) & Trim$(PicSub(Index).Tag) & "..."

    

    OldIndex = Index
    Call Delay1(300)
    If OldIndex <> Index Then Exit Sub
'//--Making a new menu-----------------------------------------------------------------------------------------------------
    If CurIndex <> Index Then
        If Not (frmMenuSystem Is Nothing) Then frmMenuSystem.Hide: Unload frmMenuSystem
    
        If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Then
            FormFade = "side"
            Load frmMenuSystem
            frmMenuSystem.Top = Me.Top + (PicSub(Index).Top - PicSub(Index).Height)
            frmMenuSystem.Left = Me.Left + Me.Width - 50
            Call s_Playsound("open")
        End If
    
        Select Case Index
            Case 0 '//Programs
                frmMenuSystem.GetMenu modStartUp.StartmenuPath
            Case 1 '//Favorites
                frmMenuSystem.GetMenu modStartUp.FavoritesPath
            Case 2 '//Desktop
                frmMenuSystem.GetMenu modStartUp.DesktopPath
            Case 3 '//My Documents
                frmMenuSystem.GetMenu modStartUp.Mydocuments
            Case 4 '//Find
                '//--Nothing to do..
            Case 5 '//Run
                '//--Nothing to do..
            Case 6 '//Control Devices
                '//--Show the form with Effect´s...
                SetWindowPos frmSLControl.hWnd, -1, (PicSub(6).Left + PicSub(6).Width) / Screen.TwipsPerPixelX, (((Screen.Height - 400) - PicSub(5).Top) / Screen.TwipsPerPixelX), 0, 0, SWP_NOSIZE

                Call MakeFormEffect("side", frmSLControl, 0, 1)
                Call s_Playsound("open")

            Case 7 '//Recycled
                '//--Nothing to do..
            Case 8 '//Shut-down
                '//--Nothing to do..
        End Select
    End If
    
    CurIndex = Index

End If
OldIndex = Index

End Sub
