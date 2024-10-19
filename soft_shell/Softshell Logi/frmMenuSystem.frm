VERSION 5.00
Object = "{54E91B3E-3171-11D3-977A-96567B857403}#2.0#0"; "AMCLABEL.OCX"
Begin VB.Form frmMenuSystem 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PAMCLabel.AMCLabel c2 
      Height          =   300
      Left            =   120
      Top             =   2280
      Width           =   1215
      Visible         =   0   'False
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "c2"
      BackColor       =   13160660
      BorderStyle     =   1
      CaptionAlign    =   1
      CaptionShadowColor=   8421504
      CaptionHighLightColor=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PAMCLabel.AMCLabel c1 
      Height          =   300
      Left            =   120
      Top             =   1800
      Width           =   1215
      Visible         =   0   'False
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "c1"
      BackColor       =   13160660
      BorderStyle     =   1
      CaptionAlign    =   1
      CaptionShadowColor=   8421504
      CaptionHighLightColor=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   3240
      Width           =   3975
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
      ScaleWidth      =   3975
      TabIndex        =   8
      Top             =   0
      Width           =   3975
   End
   Begin VB.FileListBox File2 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.DirListBox Dir2 
      Height          =   345
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   2760
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.PictureBox picItem 
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
      ScaleWidth      =   4035
      TabIndex        =   4
      Top             =   360
      Width           =   4035
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   1920
      Width           =   330
      Visible         =   0   'False
   End
   Begin VB.PictureBox picBlackArrow 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Picture         =   "frmMenuSystem.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   1440
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox picWhiteArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      Picture         =   "frmMenuSystem.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   240
      Visible         =   0   'False
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1275
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMenuSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************
'**     Thank´s to BOS Team and Brian for the ground idea to this form.     **
'**                                                                                                        **
'**     And thank´s to Bup for he´s ideas and solution´s  .                         **
'***************************************************************

Option Explicit
Private CurIndex            As Long
Private OldIndex            As Integer
Private MaxLen               As Long
Private fChild                 As frmMenuSystem
Private mParent             As frmMenuSystem
Private CurY                  As Long
Private CurYIcon            As Long
Private PublicityCaption   As String
Private Q                       As Integer

Private Sub c1_MouseEnter()
Dim Opi As Long
MouseOn = True
c1.BorderStyle = 2
c2.Visible = False
    Do Until picPublicity.Top <= Me.Height / Screen.TwipsPerPixelY - picPublicity.Height
        picHead.Top = picHead.Top - picItem(0).Height
        picPublicity.Top = picPublicity.Top - picItem(0).Height
        For Opi = 0 To picItem.Count - 1
            picItem(Opi).Top = picItem(Opi).Top - picItem(0).Height
        Next Opi
        DoEvents
        If MouseOn = False Then
            With c2
                .Visible = True
                .Top = Me.Top
                .Width = Me.ScaleWidth
                .Left = picItem(0).Left
                .Caption = "Down"
            End With
            Exit Do
        End If
    Loop
    Me.Refresh
    If picPublicity.Top <= Me.Height / Screen.TwipsPerPixelY - picPublicity.Height Then
        c1.Visible = False
        With c2
            .Visible = True
            .Top = Me.Top
            .Width = Me.ScaleWidth
            .Left = picItem(0).Left
            .Caption = "Down"
        End With
    End If

End Sub

Private Sub c1_MouseExit()
MouseOn = False
c1.BorderStyle = 1
End Sub

Private Sub c2_MouseEnter()
Dim Opi As Long
MouseOn = True
c2.BorderStyle = 2
c1.Visible = False
    Do Until picHead.Top >= Me.Top
        picHead.Top = picHead.Top + picItem(0).Height
        picPublicity.Top = picPublicity.Top + picItem(0).Height
        For Opi = 0 To picItem.Count - 1
            picItem(Opi).Top = picItem(Opi).Top + picItem(0).Height
        Next Opi
        DoEvents
        If MouseOn = False Then
            With c1
                .Visible = True
                .Top = (Screen.Height / Screen.TwipsPerPixelY) - (33 + c1.Height)
                .Width = Me.ScaleWidth
                .Left = picItem(0).Left
                .Caption = "Up"
            End With
            Exit Do
        End If
    Loop
    Me.Refresh
    If picHead.Top >= Me.Top Then
        c2.Visible = False
        With c1
            .Visible = True
            .Top = (Screen.Height / Screen.TwipsPerPixelY) - (33 + c1.Height)
            .Width = Me.ScaleWidth
            .Left = picItem(0).Left
            .Caption = "Up"
        End With
    End If

End Sub

Private Sub c2_MouseExit()
MouseOn = False
c2.BorderStyle = 1

End Sub

Private Sub Form_Load()
  
  '©®™
  PublicityCaption = GetSetting("Softshell Logi", "Startmenu Settings", "SubPublicityCaption", "®" & Space(1) & "Softworld" & Space(1) & "™")

  Me.AutoRedraw = False 'faster
  Me.ScaleMode = 3 'pixel
  OldIndex = -1
  CurIndex = Dir1.ListCount + File1.ListCount + 1
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  If Not (fChild Is Nothing) Then fChild.Hide: Unload fChild

End Sub

Public Sub UnloadAll()

  If Not (mParent Is Nothing) Then mParent.Hide: mParent.UnloadAll
  Unload Me
  
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Button = vbLeftButton Then
   
    '//--This is for the Orginal Menu´s---------------------------------------------------------
    If Index > Dir1.ListCount - 1 And Index < Dir1.ListCount + File1.ListCount Then
        
        ShellExecute Me.hWnd, "open", Dir1.path & "\" & Right(picItem(Index).Tag, Len(picItem(Index).Tag) - 8), "", "", 1
        If Not (mParent Is Nothing) Then mParent.Hide: mParent.UnloadAll
        Unload Me
        frmSubmenu.Hide
        frmTaskbar.StartButton.BorderStyle = 1
        frmTaskbar.QuickButton.BorderStyle = 1
        SubMenuShow = False
        Call s_Playsound("select")
    '//--And this is for the DoublePath menu´s---------------------------------------------
    ElseIf Index > Dir1.ListCount + File1.ListCount + Dir2.ListCount - 2 Then
        
        ShellExecute Me.hWnd, "open", Dir2.path & "\" & Right(picItem(Index).Tag, Len(picItem(Index).Tag) - 8), "", "", 1
        If Not (mParent Is Nothing) Then mParent.Hide: mParent.UnloadAll
        Unload Me
        frmSubmenu.Hide
        frmTaskbar.StartButton.BorderStyle = 1
        frmTaskbar.QuickButton.BorderStyle = 1
        SubMenuShow = False
        Call s_Playsound("select")
    End If
    '//-------------------------------------------------------------------------------------------
  ElseIf Button = vbRightButton Then
  
  End If

End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Dir1.ListCount + File1.ListCount = 0 Then Exit Sub 'if there is nothing to show
  If Index <> OldIndex Then 'Over(Index) = False Then
    If OldIndex = -1 Then OldIndex = 0
    '//--Reset the old selection ---
    picTemp.BackColor = vbButtonFace
    picItem(OldIndex).Cls
    picItem(OldIndex).BackColor = vbButtonFace
    picItem(OldIndex).ForeColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
    picItem(OldIndex).CurrentY = CurY
    If Left(Right(picItem(OldIndex).Tag, 4), 1) = "." Then
        picItem(OldIndex).Print Left(picItem(OldIndex).Tag, Len(picItem(OldIndex).Tag) - 4)
    Else
        picItem(OldIndex).Print picItem(OldIndex).Tag
    End If
    '//--Just for the DoublePath-------------------------------------------------------------------------------------------
    If OldIndex > Dir1.ListCount + File1.ListCount + Dir2.ListCount - 2 Then
        DrawIcon Dir2.path & "\" & File2.List(OldIndex - Dir2.ListCount - File1.ListCount + 1), OldIndex
    ElseIf OldIndex > Dir1.ListCount + File1.ListCount - 1 Then
        BitBlt picItem(OldIndex).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
        DrawIcon Dir2.List(OldIndex - Dir2.ListCount + 2), OldIndex
    '//--Continue with the rest of the menu or start here if ther is no DoublePath....--------------------------------
    ElseIf OldIndex > Dir1.ListCount - 1 Then
        DrawIcon Dir1.path & "\" & File1.List(OldIndex - Dir1.ListCount), OldIndex
    ElseIf OldIndex < Dir1.ListCount Then
        BitBlt picItem(OldIndex).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
        DrawIcon Dir1.List(OldIndex), OldIndex
    End If
    '//------------------------------------------------------------------------------------------------------------------------
    '//---Highlite new selection ---
    picTemp.BackColor = vbHighlight
    picItem(Index).Cls
    picItem(Index).BackColor = vbHighlight
    picItem(Index).Line (18, 1)-(picItem(Index).ScaleWidth, 1), GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    picItem(Index).Line (18, picItem(Index).ScaleHeight - 20)-(picItem(Index).ScaleWidth, picItem(Index).ScaleHeight - 20), vbButtonShadow
    picItem(Index).ForeColor = vbHighlightText
    picItem(Index).CurrentX = 0
    picItem(Index).CurrentY = CurY
    
    If Left(Right(picItem(Index).Tag, 4), 1) = "." Then
        picItem(Index).Print Left(picItem(Index).Tag, Len(picItem(Index).Tag) - 4)
    Else
        picItem(Index).Print picItem(Index).Tag
    End If
    '//--Just for the DoublePath-------------------------------------------------------------------------------------------
    If Index > Dir1.ListCount + File1.ListCount + Dir2.ListCount - 2 Then
        DrawIcon Dir2.path & "\" & File2.List(Index - Dir2.ListCount - File1.ListCount + 1), Index
    ElseIf Index > Dir1.ListCount + File1.ListCount - 1 Then
        BitBlt picItem(Index).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picWhiteArrow.hdc, 0, 0, vbSrcCopy
        DrawIcon Dir2.List(Index - Dir2.ListCount + 2), Index
    '//--Continue with the rest of the menu or start here if there is no DoublePath....--------------------------------
    ElseIf Index > Dir1.ListCount - 1 Then
        
        DrawIcon Dir1.path & "\" & File1.List(Index - Dir1.ListCount), Index
    ElseIf Index < Dir1.ListCount Then
        
        BitBlt picItem(Index).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picWhiteArrow.hdc, 0, 0, vbSrcCopy
        DrawIcon Dir1.List(Index), Index
    End If
    '//------------------------------------------------------------------------------------------------------------------------
    '//If Index >= Dir1.ListCount Then  'If not a directory then
    '//  DrawIcon Dir1.path & "\" & File1.List(Index - Dir1.ListCount), Index ', False
    '//Else 'index < Dir1.ListCount
    '//  BitBlt picItem(Index).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 16, 16, picWhiteArrow.hdc, 0, 0, vbSrcCopy
    '//  DrawIcon Dir1.List(Index), Index ', False
    '//End If
    
      '//--Show new child menu ---
      Timer1.Interval = 300
      Aindex = Index
    
    
    picTemp.ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    picTemp.Line (0, 0)-(20, 0) 'Upper
    picTemp.Line (0, 0)-(0, 19) 'Left
    picTemp.ForeColor = vbButtonShadow
    picTemp.Line (0, 19)-(20, 19) 'Bottom
    'picTemp.Line (19, 0)-(19, 19) 'Right
    BitBlt picItem(Index).hdc, 0, CurYIcon, 20, 20, picTemp.hdc, 0, 0, vbSrcCopy
    
    
    OldIndex = Index
    DoEvents
    If OldIndex <> Index Then Exit Sub
    
    If Index > Dir1.ListCount - 1 And Index < Dir1.ListCount + File1.ListCount Or _
       Index >= Dir1.ListCount + File1.ListCount + Dir2.ListCount - 1 Then
        Call s_Playsound("hover")
    End If
    
  End If

End Sub

Public Sub GetMenu(path As String, Optional Parent As frmMenuSystem = Nothing)
  
  Dim i As Long
  Dim lTemp As Long
  
  Set mParent = Parent
  MaxLen = 0
  picItem(0).Top = picItem(0).Height
  Dir1.path = path
  File1.path = path
'//--Just for the (Double) Quick menu----------------------------------------------------------------------------------
If DoublePath = True Then
    Dir2.path = DoublepathString
    File2.path = DoublepathString
    
    If File1.ListCount + Dir1.ListCount + File2.ListCount + Dir2.ListCount = 0 Then
        picItem(0).CurrentY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
        picItem(0).Print Space(5) & "[ Empty ]"
        MaxLen = picItem(0).TextWidth("[ Empty ]")
    Else
        For i = 1 To Dir1.ListCount + File1.ListCount + Dir2.ListCount + File2.ListCount - 2 '//1 if program will be with the others.
            Load picItem(i)
            picItem(i).Visible = True
            picItem(i).Top = picItem(0).Height * (i + 1)
        Next
        CurYIcon = ((picItem(0).Height) - 20) / 2
        CurY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
        
        For i = 0 To Dir1.ListCount - 1
            DrawIcon Dir1.List(i), i
            picItem(i).CurrentY = CurY
            picItem(i).Tag = "        " & ExtractFileName(Dir1.List(i))
            picItem(i).Print picItem(i).Tag
            lTemp = picItem(i).TextWidth(picItem(i).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next
        For i = 0 To File1.ListCount - 1
            picTemp.BackColor = vbButtonFace
            DrawIcon Dir1.path & "\" & File1.List(i), i + Dir1.ListCount
            picItem(i + Dir1.ListCount).CurrentY = CurY
            picItem(i + Dir1.ListCount).Tag = "        " & Left(File1.List(i), Len(File1.List(i))) '//- 4)
            picItem(i + Dir1.ListCount).Print Left(picItem(i + Dir1.ListCount).Tag, Len(picItem(i + Dir1.ListCount).Tag) - 4)
            lTemp = picItem(i + Dir1.ListCount).TextWidth(picItem(i + Dir1.ListCount).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next
        '//--Take the Program Dir away from the startmenupath-------------------------------------------
       
        Q = 0
        For i = 0 To Dir2.ListCount - 1
            If Trim$(LCase(ExtractFileName(Dir2.List(i)))) <> "program" Then
                DrawIcon Dir2.List(i), i + Dir1.ListCount + File1.ListCount - Q
                picItem(i + Dir1.ListCount + File1.ListCount - Q).CurrentY = CurY
                picItem(i + Dir1.ListCount + File1.ListCount - Q).Tag = "        " & ExtractFileName(Dir2.List(i))
                picItem(i + Dir1.ListCount + File1.ListCount - Q).Print picItem(i + Dir1.ListCount + File1.ListCount - Q).Tag
                lTemp = picItem(i + Dir1.ListCount + File1.ListCount - Q).TextWidth(picItem(i + Dir1.ListCount + File1.ListCount - Q).Tag)
                If lTemp > MaxLen Then MaxLen = lTemp
            Else
                Q = 1
            End If
        Next
        '//---------------------------------------------------------------------------------
        For i = 0 To File2.ListCount - 1
            picTemp.BackColor = vbButtonFace
            DrawIcon Dir2.path & "\" & File2.List(i), i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q
            picItem(i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).CurrentY = CurY
            picItem(i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).Tag = "        " & Left(File2.List(i), Len(File2.List(i))) '//- 4)
            picItem(i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).Print Left(picItem(i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).Tag, Len(picItem(i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).Tag) - 4)
            lTemp = picItem(Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).TextWidth(picItem(i + Dir1.ListCount + File1.ListCount + Dir2.ListCount - Q).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next

    End If
'//--This is for all other Menu´s------------------------------------------------------------------------------------------
Else
  
  If File1.ListCount + Dir1.ListCount = 0 Then
      picItem(0).CurrentY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
      picItem(0).Print Space(5) & "[ Empty ]"
      MaxLen = picItem(0).TextWidth("[ Empty ]")
  Else
      For i = 1 To Dir1.ListCount + File1.ListCount - 1
          Load picItem(i)
          picItem(i).Visible = True
          picItem(i).Top = picItem(0).Height * (i + 1)
      Next
        CurYIcon = ((picItem(0).Height) - 20) / 2
        CurY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
        For i = 0 To Dir1.ListCount - 1
            DrawIcon Dir1.List(i), i
            picItem(i).CurrentY = CurY
            picItem(i).Tag = "        " & ExtractFileName(Dir1.List(i))
            picItem(i).Print picItem(i).Tag
            lTemp = picItem(i).TextWidth(picItem(i).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next
        For i = 0 To File1.ListCount - 1
            picTemp.BackColor = vbButtonFace
            DrawIcon Dir1.path & "\" & File1.List(i), i + Dir1.ListCount
            picItem(i + Dir1.ListCount).CurrentY = CurY
            picItem(i + Dir1.ListCount).Tag = "        " & Left(File1.List(i), Len(File1.List(i))) '//- 4)
            picItem(i + Dir1.ListCount).Print Left(picItem(i + Dir1.ListCount).Tag, Len(picItem(i + Dir1.ListCount).Tag) - 4)
            lTemp = picItem(i + Dir1.ListCount).TextWidth(picItem(i + Dir1.ListCount).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next
  End If
  
End If
'//--Set the Value for Head and Publicity-------------------------------------------------------------------------
    picHead.Tag = Acaption
    lTemp = picHead.TextWidth(picHead.Tag)
    If lTemp > MaxLen Then MaxLen = lTemp
    
    picPublicity.Tag = PublicityCaption ' "®" & Space(1) & "Softworld" & Space(1) & "™"
    lTemp = picPublicity.TextWidth(picPublicity.Tag)
    If lTemp > MaxLen Then MaxLen = lTemp
      
    Me.Width = MaxLen + 500
    
    If Me.Width > Screen.Width / 3 Then Me.Width = Screen.Width / 2.5
    
    Me.Height = ((picItem.Count + 2) * picItem(0).Height) * Screen.TwipsPerPixelY '+ 10
    
    picHead.CurrentY = CurY
    picHead.CurrentX = 10
    picHead.Print picHead.Tag
    picHead.Top = 0
    picHead.Visible = True
    picHead.Width = Me.ScaleWidth
    
    picPublicity.CurrentY = CurY
    picPublicity.CurrentX = ((picHead.Width * Screen.TwipsPerPixelX) / 2) - 650
    picPublicity.Print picPublicity.Tag
    picPublicity.Top = (picItem(picItem.Count - 1).Top) + (picItem(picItem.Count - 1).Height) 'Me.Height / Screen.TwipsPerPixelY - picPublicity.Height
    picPublicity.Visible = True
    picPublicity.Width = Me.ScaleWidth


'//--Seperator lines on the Menu-------------------------------------------------------------------------------
    '//--Line Upper...........
    picHead.ScaleMode = 3
    picHead.ForeColor = vb3DDKShadow
    picHead.Line (1, picHead.Top + (picHead.Height - 2))-(picPublicity.Width - 1, picHead.Top + (picHead.Height - 2))
    picHead.ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    picHead.Line (1, picHead.Top + (picHead.Height - 1))-(picPublicity.Width - 1, picHead.Top + (picHead.Height - 1))
    picHead.ScaleMode = 1
    '//--Line Down...........
    picPublicity.ScaleMode = 3
    picPublicity.ForeColor = vb3DDKShadow
    picPublicity.Line (1, 0)-(picPublicity.Width - 1, 0)
    picPublicity.ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    picPublicity.Line (1, 1)-(picPublicity.Width - 1, 1)
    picPublicity.ScaleMode = 1
'//-------------------------------------------------------------------------------------------
  SetWindowPos Me.hWnd, -1, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.ScaleWidth, Me.ScaleHeight + 10, SWP_NOREPOSITION
  For i = 0 To picItem.Count - 1
      picItem(i).Width = Me.ScaleWidth
  Next
  For i = 0 To Dir1.ListCount - 1
      BitBlt picItem(i).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
  Next
  If DoublePath = True Then
    For i = Dir1.ListCount + File1.ListCount + 1 To Dir1.ListCount + File1.ListCount + Dir2.ListCount - 1
        BitBlt picItem(i - Q).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
    Next
  End If
  If Me.Left + Me.Width > Screen.Width Then
        Me.Left = Aleft - Me.Width + 35
        If Me.Left < 0 Then Me.Left = 0
        Aleft = 0
  End If
  '//--If the menu is below the taskbar set it over.....
  If Me.Top + Me.Height > Screen.Height Then
        Me.Top = Me.Top - ((Me.Top + Me.Height) - (Screen.Height - Abottom))
  End If
  
  '//--If the form is to high to fit on the screen....
  If Me.Top < 0 Then
    With Me
        .Top = 0
        .Height = Screen.Height - Abottom
    End With
    With c1
        .Visible = True
        .Top = (Screen.Height / Screen.TwipsPerPixelY) - (33 + c1.Height)
        .Width = Me.ScaleWidth
        .Left = picItem(0).Left
        .Caption = "Up"
    End With
  End If
  
  '//--Show the form with Effect´s...
  If DoublePath = True Then
    '//--If the QuickButton is pressed...
    Call MakeFormEffect(FormFade, Me, 0, 200)
  Else
    If picItem.Count > 5 Then
        Call MakeFormEffect(FormFade, Me, 0, 1) 'picItem.Count)
    Else
        Call MakeFormEffect(FormFade, Me, 0, 1)
    End If
  End If
  
  DoublePath = False

End Sub

Sub DrawIcon(path, Index, Optional blt = True)
  
  Dim hImgLarge&
  
  hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
  BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
  picTemp.Cls '
  If blt Then
      ImageList_Draw hImgLarge&, shinfo.iIcon, picTemp.hdc, 2, 2, ILD_TRANSPARENT
      BitBlt picItem(Index).hdc, 0, CurYIcon, 20, 20, picTemp.hdc, 0, 0, vbSrcCopy
  Else
      ImageList_Draw hImgLarge&, shinfo.iIcon, picTemp.hdc, 2, 2, ILD_TRANSPARENT
  End If

End Sub

Private Sub Timer1_Timer()
  
  Timer1.Interval = 0
  If CurIndex <> Aindex Then
      If Not (fChild Is Nothing) Then fChild.Hide: Unload fChild

        '//--Show new child menu --------------------------------------------------------------------------------------------------------------
        FormFade = "side"
        If OldIndex < Dir1.ListCount Then
          Set fChild = New frmMenuSystem
          fChild.Top = Me.Top + (picItem(OldIndex).Top - picItem(OldIndex).Height) * Screen.TwipsPerPixelX
          fChild.Left = Me.Left + Me.Width - 70
          Aleft = Me.Left
          Acaption = Space(2) & ExtractFileName(Dir1.List(Aindex)) & "..."
          fChild.GetMenu Dir1.path & "\" & Right(picItem(OldIndex).Tag, Len(picItem(OldIndex).Tag) - 8), Me
          Call s_Playsound("open")
        '//--Just for the (Double) Quick Menu-----------------------------------------------------------------------------------------------------
        ElseIf OldIndex < Dir1.ListCount + File1.ListCount + Dir2.ListCount - 1 And OldIndex >= Dir1.ListCount + File1.ListCount Then
          Set fChild = New frmMenuSystem
          fChild.Top = Me.Top + (picItem(OldIndex).Top - picItem(OldIndex).Height) * Screen.TwipsPerPixelX
          fChild.Left = Me.Left + Me.Width - 70
          Aleft = Me.Left
          Acaption = Space(2) & ExtractFileName(Dir2.List(Aindex - (Dir1.ListCount + File1.ListCount - 1))) & "..."
          fChild.GetMenu Dir2.path & "\" & Right(picItem(OldIndex).Tag, Len(picItem(OldIndex).Tag) - 8), Me
          Call s_Playsound("open")
        End If
        '//---------------------------------------------------------------------------------------------------------------------------------------------
     
       
  End If
  CurIndex = Aindex

End Sub





