VERSION 5.00
Object = "{54E91B3E-3171-11D3-977A-96567B857403}#2.0#0"; "AMCLABEL.OCX"
Begin VB.Form frmTaskbar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      Height          =   375
      Index           =   0
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   3360
      Width           =   495
   End
   Begin VB.Timer TaskbarUpdate 
      Interval        =   500
      Left            =   1800
      Top             =   2280
   End
   Begin VB.Timer DateTime1 
      Interval        =   250
      Left            =   1800
      Top             =   2760
   End
   Begin VB.Timer DateTime2 
      Interval        =   250
      Left            =   1800
      Top             =   3240
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Visible         =   0   'False
   End
   Begin PAMCLabel.AMCLabel QuickButton 
      Height          =   255
      Left            =   1800
      ToolTipText     =   "Quick-Menu, use this to make fast start´s...."
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      Caption         =   "Q"
      BackColor       =   12632256
      BorderStyle     =   1
      CaptionAlign    =   1
      CaptionShadowColor=   8421504
      CaptionBehaviour=   3
      CaptionHighLightColor=   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PAMCLabel.AMCLabel StartButton 
      Height          =   375
      Left            =   0
      ToolTipText     =   "Click on this to do thing´s öhöhöhöhöhöh"
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "AMCLabel1"
      BackColor       =   255
      BorderStyle     =   2
      CaptionShadowColor=   8421504
      CaptionHighLightColor=   65535
      CaptionMouseOverColor=   -2147483634
      Picture         =   "frmTaskbar.frx":0000
      PictureOffset   =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PAMCLabel.AMCLabel DatumTidRam 
      Height          =   855
      Left            =   120
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      Caption         =   "DatumTidRam"
      BackColor       =   13160660
      CaptionAlign    =   1
      CaptionShadowColor=   8421504
      CaptionHighLightColor=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PAMCLabel.AMCLabel Exevisare 
      Height          =   375
      Index           =   0
      Left            =   0
      Tag             =   "0"
      Top             =   3840
      Width           =   2175
      Visible         =   0   'False
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Exevisare(0)"
      BackColor       =   12632256
      CaptionShadowColor=   8421504
      CaptionHighLightColor=   16777215
      PictureDisableAction=   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgIcon 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image ImgLink 
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   230
      Visible         =   0   'False
   End
   Begin VB.Shape ShapeList 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   120
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Line Lines5 
      BorderWidth     =   5
      X1              =   0
      X2              =   2160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Lines4 
      BorderWidth     =   5
      X1              =   0
      X2              =   2160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Lines3 
      BorderWidth     =   5
      X1              =   0
      X2              =   2160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Lines2 
      BorderWidth     =   5
      X1              =   0
      X2              =   2160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Lines1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   2160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Lines6 
      BorderWidth     =   5
      X1              =   0
      X2              =   2160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Lines7 
      BorderWidth     =   5
      X1              =   0
      X2              =   2160
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Chan As Boolean '//--Växel till Datumtidram
Private Pwidth As Long


Private Sub Form_Load()
On Error GoTo Errfix
'//--ExevisareWidth...
Pwidth = GetSetting("Softshell Logi", "General Settings", "ExeVisareWidth", 2500)

'//--Reset the button´s show array´s-----------
QuickMenuShow = False
SubMenuShow = False

'//--Set the modStartUp.() to it´s value------------------------------------------
modStartUp.DoublepathString = FindNeedFolders(CSIDL_STARTMENU)
modStartUp.StartmenuPath = FindNeedFolders(CSIDL_PROGRAMS)
modStartUp.DesktopPath = FindNeedFolders(CSIDL_DESKTOPDIRECTORY)
modStartUp.FavoritesPath = FindNeedFolders(CSIDL_FAVORITES)
modStartUp.Mydocuments = FindNeedFolders(CSIDL_PERSONAL)

'//--Check if Quickmenu Dir exists--------------
   DirN = Dir$(App.path & "\", 16)
    Do While DirN <> ""
        If DirN = "Quickmenu" Then
            DirF = True
            Exit Do
        End If
        DirN = Dir$
    Loop
    
    If Not DirF = True Then
        MkDir App.path & "\Quickmenu"
    End If
   
   modStartUp.QuickmenuPath = App.path & "\Quickmenu"

'//----------------------------------------------------------------------

Exevisare(0).Visible = False: Exevisare(0).Left = -10000

SetWindowPos Me.hWnd, (HWND_TOPMOST), (0), (Screen.Height - 400) / Screen.TwipsPerPixelX, ((Screen.Width) / Screen.TwipsPerPixelX), (400 / Screen.TwipsPerPixelX), SWP_NOACTIVATE '//SWP_NOREPOSITION '// & SWP_NOSIZE

Call KontrollList
Exit Sub
Errfix:
Call ErrFixing
End Sub

Private Sub Exevisare_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Call frmSubmenu.KillSubMenu
    If Exevisare(Index).BorderStyle = 1 Then
        pSetForegroundWindow List1.ItemData(Index - 1)
    Else
        ShowWindow List1.ItemData(Index - 1), SW_MINIMIZE
    End If
Else
    EXIndex = Index
    ExevisareRightClickMenu.PopupMenu _
    ExevisareRightClickMenu.mnuFile
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Call frmSubmenu.KillSubMenu
Else
    
End If

End Sub

Public Sub UnloadLogi()

    Dim Fo As Integer
    For Fo = (Forms.Count - 1) To 0 Step -1
        Unload Forms(Fo)
        DoEvents
    Next Fo
    ShowExplorerTaskBar
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Fo As Integer
For Fo = (Forms.Count - 1) To 0 Step -1
    Unload Forms(Fo)
Next Fo

End Sub

Private Sub PaintDesktop_Timer()
'
End Sub

Public Sub KontrollList()

ImgLink.Visible = False
ImgLink.Picture = LoadPicture(App.path & "\softlogoSB.bmp")
ImgIcon(0).Visible = False
'//--Taskbar...
With ShapeList
    .Top = 0
    .Left = 0
    .Height = 400
    .Width = Screen.Width
    .BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarBackcolor", &HC0C0C0)
End With
'//--Startknappen...
With StartButton
    .Top = 25
    .Left = 25
    .Height = 350
    .Width = 850
    .BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarBackcolor", &HC0C0C0)
    .BorderStyle = 1
    .Caption = GetSetting("Softshell Logi", "General Settings", "StartbuttonCaption", "Start ")
    .CaptionAlign = 2
    .CaptionShadowStyle = 1
    .CaptionShadowColor = GetSetting("Softshell Logi", "Color Settings", "StartbuttonCaptionShadowColor", &HE0E0E0)
    .CaptionBehaviour = 2
    .ForeColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
End With
'//--Avskiljare 1 efter startknappen...
With Lines6
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonShadowColor", vb3DShadow)
    .BorderWidth = 1
    .X1 = 925
    .X2 = 925
    .Y1 = 25
    .Y2 = 375
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Avskiljare 2 efter startknappen...
With Lines7
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    .BorderWidth = 1
    .X1 = 940
    .X2 = 940
    .Y1 = 25
    .Y2 = 375
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--QuickButton...
With QuickButton
    .Left = 1010
    .Top = 25
    .Height = 350
    .Width = 200
    .CaptionShadowColor = GetSetting("Softshell Logi", "Color Settings", "StartbuttonCaptionShadowColor", &HE0E0E0)
    .ForeColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
    .BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarBackcolor", &HC0C0C0)
    .CaptionMouseOverColor = vbHighlightText
    .CaptionBehaviour = 2
    .CaptionShadowStyle = 1
    .ZOrder vbBringToFront
    .Visible = True
End With
'//--Avskiljare 3 efter startknappen...
With Lines1
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonShadowColor", vb3DShadow)
    .BorderWidth = 1
    .X1 = 1225
    .X2 = 1225
    .Y1 = 25
    .Y2 = 375
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Avskiljare 4 efter startknappen...
With Lines3
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    .BorderWidth = 1
    .X1 = 1240
    .X2 = 1240
    .Y1 = 25
    .Y2 = 375
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Linjen i överkanten på kontroll listen...
With Lines2
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    .BorderWidth = 1
    .X1 = 0
    .X2 = Screen.Width
    .Y1 = 0
    .Y2 = 0
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Ram för datum och tid till höger på kontrolllisten.
With DatumTidRam
    .BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarBackcolor", &HC0C0C0)
    .ForeColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
    .BorderStyle = 2
    .CaptionAlign = 1
    .CaptionBehaviour = 0
    .ToolTipText = Date
    .Height = 350
    .Width = 800
    .Left = Screen.Width - (DatumTidRam.Width + 25)
    .Top = 25
    .Caption = ""
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Avskiljare 1 före DatumTidRam...
With Lines4
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonShadowColor", vb3DShadow)
    .BorderWidth = 1
    .X1 = DatumTidRam.Left - 90
    .X2 = DatumTidRam.Left - 90
    .Y1 = 25
    .Y2 = 375
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Avskiljare 2 före DatumTidRam...
With Lines5
    .BorderColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    .BorderWidth = 1
    .X1 = DatumTidRam.Left - 75
    .X2 = DatumTidRam.Left - 75
    .Y1 = 25
    .Y2 = 375
    .Visible = True
    .ZOrder vbBringToFront
End With
'//--Set the DateTime_Timers...
DateTime1.Enabled = True
DateTime2.Enabled = False

End Sub

Private Sub DatumTidRam_Click()

If Chan = False Then
 DatumTidRam.Width = 800
 Call UppDateraDatumTidRam
 DateTime1.Enabled = True
 DateTime2.Enabled = False
 Chan = True
 Exit Sub
End If
DatumTidRam.Width = 1650
Call UppDateraDatumTidRam
DateTime1.Enabled = False
DateTime2.Enabled = True
Chan = False

End Sub

Private Sub DatumTidRam_DblClick()
    '//--Show the Windows Clock--------------------------------------
    Shell ("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Public Sub UppDateraDatumTidRam()

DatumTidRam.Left = frmTaskbar.Width - (DatumTidRam.Width + 25)
 Lines4.X1 = DatumTidRam.Left - 90
 Lines4.X2 = DatumTidRam.Left - 90
 Lines5.X1 = DatumTidRam.Left - 75
 Lines5.X2 = DatumTidRam.Left - 75
 
End Sub

Private Sub StartButton_Click()
'//--Kill Quick menu----------------------
If QuickMenuShow = True Then
    frmTaskbar.QuickButton.BorderStyle = 1
    Call frmMenuSystem.UnloadAll
    QuickMenuShow = False
End If
'//--Kill Start menu------------------------------
If SubMenuShow = True Then
     
    Unload frmSubmenu
    Unload frmSLControl
    frmTaskbar.StartButton.BorderStyle = 1
    Call frmMenuSystem.UnloadAll
    SubMenuShow = False
    Exit Sub
End If
'//--Make Start menu-----------------------------------------------

FormFade = "up"
SetWindowPos frmSubmenu.hWnd, -1, 0, (((Screen.Height - 400) - frmSubmenu.Height) / Screen.TwipsPerPixelX), 0, 0, SWP_NOSIZE

SubMenuShow = True
StartButton.BorderStyle = 2

'//--Show the form with Effect´s...
Call MakeFormEffect(FormFade, frmSubmenu, 0, 200)

End Sub

Private Sub StartButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

StartButton.BorderStyle = 2

End Sub

Private Sub StartButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

StartButton.BorderStyle = 1

End Sub

Private Sub QuickButton_Click()
            
'//--Kill Start menu------------------------------
If SubMenuShow = True Then
    Unload frmSubmenu
    Unload frmSLControl
    frmTaskbar.StartButton.BorderStyle = 1
    Call frmMenuSystem.UnloadAll
    SubMenuShow = False
End If
'//--Kill Quick menu------------------------
If QuickMenuShow = True Then
     
    frmTaskbar.QuickButton.BorderStyle = 1
    Call frmMenuSystem.UnloadAll
    QuickMenuShow = False
    Exit Sub
End If
'//--Make Quick menu----------------------

QuickMenuShow = True
Acaption = Space(2) & "Quick Menu..."
QuickButton.BorderStyle = 2
FormFade = "up"
DoublePath = True ': DoublepathString = "c:\windows\start-menyn"
Load frmMenuSystem
frmMenuSystem.Top = Me.Top
frmMenuSystem.Left = QuickButton.Left
frmMenuSystem.GetMenu modStartUp.QuickmenuPath
'//--------------------------------------------
End Sub

Private Sub QuickButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

QuickButton.BorderStyle = 2

End Sub

Private Sub QuickButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

QuickButton.BorderStyle = 1

End Sub

Private Sub DateTime1_Timer()

DatumTidRam.Caption = Time

End Sub

Private Sub DateTime2_Timer()


DatumTidRam.Caption = Date & " " & " " & Time

End Sub


Private Sub TaskbarUpdate_Timer()
'************************************************************************
'   Thank´s to Brian for the Basic idea´s to this Update.................
'************************************************************************
'   Check if the mainwindow is to big for the screen, if so fix it........
'************************************************************************

    Dim FGWhwnd As Long
    Dim FGWrect As RECT
    
    FGWhwnd = GetForegroundWindow
    GetWindowRect FGWhwnd, FGWrect
    
    If FGWhwn <> frmMenuSystem.hWnd And FGWhwnd <> frmSubmenu.hWnd And FGWhwnd <> frmTaskbar.hWnd Then
        
        '//--ForegroundWindow.height stop at Taskbar.top-------------------------------------------------------
        If FGWrect.Bottom > (Screen.Height - 300) / Screen.TwipsPerPixelX Then
             SetWindowPos FGWhwnd, 0, FGWrect.Left, FGWrect.Top, FGWrect.Right - FGWrect.Left, ((Screen.Height - 400) / Screen.TwipsPerPixelX) - (FGWrect.Top), SWP_NOACTIVATE '// Or SWP_NOSIZE
        End If
        
    End If

'************************************************************************
'   Updating the taskbar and Exevisare.......
'************************************************************************
Dim xt As Integer
Dim Pleft As Long
Dim Ecap As Variant
Dim ExeVisareDivisionNumber As Long
    ExeVisareDivisionNumber = GetSetting("Softshell Logi", "General Settings", "ExeVisareDivisionNumber", 75)
Call modDevice1.fEnumWindows(List1)

fEnumWindows List1

Pleft = 1300

If List1.ListCount > Exevisare(0).Tag Then
    For xt = Exevisare(0).Tag + 1 To List1.ListCount
        Load Exevisare(xt)
        '//Load picIcon(xt)
        Exevisare(0).Tag = xt
    Next xt
End If

If List1.ListCount < Exevisare(0).Tag Then
    For xt = Exevisare(0).Tag To List1.ListCount + 1 Step -1
        Unload Exevisare(xt)
        '//Unload picIcon(xt)
        Exevisare(0).Tag = xt - 1
    Next xt
End If

'//--If Exevisarna can have there original size...
If (GetSetting("Softshell Logi", "General Settings", "ExeVisareWidth", 2500) * Exevisare.Count - 1) < _
    (DatumTidRam.Left - 1300) Then
    Pwidth = GetSetting("Softshell Logi", "General Settings", "ExeVisareWidth", 2500)
End If

For xt = 1 To List1.ListCount
    Ecap = Left$(List1.List(xt - 1), InStr(List1.List(xt - 1), "*Softshell_Logi*") - 1)
    
    If Len(Ecap) > Pwidth / ExeVisareDivisionNumber Then
        Exevisare(xt).Caption = Space(2) & Left$(Ecap, Pwidth / ExeVisareDivisionNumber) & "..."
    Else
        Exevisare(xt).Caption = Space(2) & Ecap
    End If
    
    If GetForegroundWindow = List1.ItemData(xt - 1) Then
        Exevisare(xt).BorderStyle = 2
           
        Exevisare(xt).BackColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    Else
        Exevisare(xt).BorderStyle = 1
            
        Exevisare(xt).BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarBackcolor", &HC0C0C0)
    End If
    
    With Exevisare(xt)
        
        .ForeColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
        .CaptionShadowStyle = 1
        .CaptionShadowColor = GetSetting("Softshell Logi", "Color Settings", "StartbuttonCaptionShadowColor", &HE0E0E0)
        .Font = StartmenyFontname
        .Font.size = 8
        .Visible = True
        .Top = 30
        .Height = 350
        .Left = Pleft
        .ZOrder vbBringToFront
        
    End With
    
    '//With picIcon(xt)
    '//    .BackColor = Exevisare(xt).BackColor
    '//    .BorderStyle = 0
    '//    .Top = Exevisare(xt).Top + 60
    '//    .Width = Exevisare(xt).Height - 120
    '//    .Height = Exevisare(xt).Height - 120
    '//    .Picture = ImgLink.Picture
    '//    .Left = Pleft + 60
    '//    .ZOrder vbBringToFront
    '//    .Visible = True
    '//End With
    
    
    If Exevisare(xt).Left + Exevisare(xt).Width > Lines4.X1 - 25 Then
        Call ControlSizeExeVisare
    Else
        Exevisare(xt).Width = Pwidth
    End If
    
    Pleft = Pleft + Exevisare(xt).Width + 30
   
Next xt

     
End Sub

Public Sub ControlSizeExeVisare()

'//--Anpassa storleken på ExeVisarna...
Dim CSEV As Integer
'Dim Pwidth As Long
Pwidth = Exevisare(xt).Width

Do
 
  For CSEV = 1 To Exevisare(0).Tag
  '//--Om den den sista ExeVisaren får plats så bryts loopen...
   If CSEV = Exevisare(0).Tag Then
    If Exevisare(Exevisare(0).Tag).Left + Exevisare(Exevisare(0).Tag).Width _
     <= Lines4.X1 - 25 Then Exit Do
   End If
    '//--Förminska ExeVisaren...
    Pwidth = Pwidth - 10
    
    Exevisare(CSEV).Width = Pwidth
    If Exevisare(0).Tag > CSEV Then
      Exevisare(CSEV + 1).Left = _
      Exevisare(CSEV).Left + Exevisare(CSEV).Width
    
    End If
  Next CSEV
  
Loop

End Sub


