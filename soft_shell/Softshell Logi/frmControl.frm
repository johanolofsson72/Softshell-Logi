VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Softshell Logi Control Settings..."
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdbControl 
      Left            =   1560
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   0
      Left            =   480
      ScaleHeight     =   4695
      ScaleWidth      =   4935
      TabIndex        =   3
      Top             =   480
      Width           =   4935
      Begin VB.CheckBox chIni 
         Caption         =   "Have this Unchecked if you don´t want Softshell to start alone (without Explorer)"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   360
         TabIndex        =   29
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox txtPwidth 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Text            =   "2500"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtExevisareDivisionNumber 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Text            =   "75"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtStartbuttonCaption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Text            =   "Start"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "General Setting´s"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   2
      Left            =   480
      ScaleHeight     =   4695
      ScaleWidth      =   4935
      TabIndex        =   8
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton cmdCommonHighLightColor 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   28
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdCommonShadowColor 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   27
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdStartbuttonShadow 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdTaskbarForeColor 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdTaskbarBackcolor 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape shCommonHighLightColor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   375
         Left            =   3720
         Shape           =   2  'Oval
         Top             =   3600
         Width           =   615
      End
      Begin VB.Shape shCommonShadowColor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   375
         Left            =   3720
         Shape           =   2  'Oval
         Top             =   2880
         Width           =   615
      End
      Begin VB.Shape shStartbuttonShadow 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   375
         Left            =   3720
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   615
      End
      Begin VB.Shape shTaskbarForeColor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   375
         Left            =   3720
         Shape           =   2  'Oval
         Top             =   1440
         Width           =   615
      End
      Begin VB.Shape shTaskbarBackcolor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   375
         Left            =   3720
         Shape           =   2  'Oval
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Color Setting´s"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   1
      Left            =   480
      ScaleHeight     =   4695
      ScaleWidth      =   4935
      TabIndex        =   14
      Top             =   480
      Width           =   4935
      Begin VB.TextBox txtShutdown 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Text            =   "Text11"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtRecycled 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Text            =   "Text10"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtControlDevices 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Text            =   "Text9"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtRun 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   23
         Text            =   "Text8"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtFind 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Text            =   "Text7"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtMyDocuments 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Text            =   "Text6"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtDesktop 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtFavorites 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPrograms 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtPublicityCaption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtHeadCaption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   3480
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Startmenu Setting´s"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   4935
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControl.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControl.frx":4F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControl.frx":54F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox dirSkins 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   -480
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip tsControl 
      Height          =   5595
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9869
      MultiRow        =   -1  'True
      TabFixedWidth   =   2999
      TabStyle        =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General Setting´s"
            Key             =   "Generalsettings"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Startmenu Setting´s"
            Key             =   "Startmenusettings"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color Setting´s"
            Key             =   "Colorsettings"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    '//--Kill Quick menu----------------------
    If QuickMenuShow = True Then
        frmTaskbar.QuickButton.BorderStyle = 1
        Call frmMenuSystem.UnloadAll
        QuickMenuShow = False
    End If
    '//--Kill Start menu------------------------------
    If SubMenuShow = True Then
        Unload frmSubmenu
        frmTaskbar.StartButton.BorderStyle = 1
        Call frmMenuSystem.UnloadAll
        SubMenuShow = False
    End If

    txtStartbuttonCaption.Text = GetSetting("Softshell Logi", "General Settings", "StartbuttonCaption", "Start  ")
    txtExevisareDivisionNumber.Text = GetSetting("Softshell Logi", "General Settings", "ExeVisareDivisionNumber", 75)
    txtPwidth.Text = GetSetting("Softshell Logi", "General Settings", "ExeVisareWidth", 2500)
    
    shTaskbarBackcolor.BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarBackcolor", &HC0C0C0)
    shTaskbarForeColor.BackColor = GetSetting("Softshell Logi", "Color Settings", "TaskbarForeColor", vbButtonText)
    shStartbuttonShadow.BackColor = GetSetting("Softshell Logi", "Color Settings", "StartbuttonCaptionShadowColor", &HE0E0E0)
    shCommonShadowColor.BackColor = GetSetting("Softshell Logi", "Color Setting", "CommonShadowColor", vb3DShadow)
    shCommonHighLightColor.BackColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
    
    txtHeadCaption.Text = GetSetting("Softshell Logi", "Startmenu Settings", "SubHeadCaption", "Softshell Logi")
    txtPublicityCaption.Text = GetSetting("Softshell Logi", "Startmenu Settings", "SubPublicityCaption", "®" & Space(1) & "Softworld" & Space(1) & "™")
    txtPrograms.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Programs", "Programs")
    txtFavorites.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Favorites", "Favorites")
    txtDesktop.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Desktop", "Desktop")
    txtMyDocuments.Text = GetSetting("Softshell Logi", "Startmenu Settings", "My Documents", "My Documents")
    txtFind.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Find", "Find...")
    txtRun.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Run", "Run...")
    txtControlDevices.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Control Devices", "Control Devices...")
    txtRecycled.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Recycled", "Recycled...")
    txtShutdown.Text = GetSetting("Softshell Logi", "Startmenu Settings", "Shut-Down", "Shut-Down...")
    
    If GetSetting("Softshell Logi", "General Settings", "SystemIni", 0) = 0 Then
        chIni.Value = 0
    Else
        chIni.Value = 1
    End If
    
    
End Sub

Private Sub Form_Paint()

 With picSettings(0)
        .Cls
        .ForeColor = vbHighlight
        .FontName = "Comic Sans Ms"
        .FontBold = True
        .FontSize = 10
    End With
        picSettings(0).CurrentX = 500
        picSettings(0).CurrentY = 750
        picSettings(0).Print "The Text on the Startbutton:"
        picSettings(0).CurrentX = 500
        picSettings(0).CurrentY = 1470
        picSettings(0).Print "ExeVisareDivisionNumber:"
        picSettings(0).CurrentX = 500
        picSettings(0).CurrentY = 2190
        picSettings(0).Print "ExeVisare.Width  (" & Space(10) & "):"
        picSettings(0).ForeColor = vbRed
        picSettings(0).CurrentX = 2500
        picSettings(0).CurrentY = 2190
        picSettings(0).Print "Caution"

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

Me.Hide

    SaveSetting "Softshell Logi", "General Settings", "StartbuttonCaption", txtStartbuttonCaption.Text
    SaveSetting "Softshell Logi", "General Settings", "ExeVisareDivisionNumber", txtExevisareDivisionNumber.Text
    SaveSetting "Softshell Logi", "General Settings", "ExeVisareWidth", txtPwidth.Text
    
    SaveSetting "Softshell Logi", "Color Settings", "TaskbarBackcolor", shTaskbarBackcolor.BackColor
    SaveSetting "Softshell Logi", "Color Settings", "TaskbarForeColor", shTaskbarForeColor.BackColor
    SaveSetting "Softshell Logi", "Color Settings", "StartbuttonCaptionShadowColor", shStartbuttonShadow.BackColor
    SaveSetting "Softshell Logi", "Color Settings", "CommonShadowColor", shCommonShadowColor.BackColor
    SaveSetting "Softshell Logi", "Color Settings", "CommonHighLightColor", shCommonHighLightColor.BackColor
    
    SaveSetting "Softshell Logi", "Startmenu Settings", "SubHeadCaption", txtHeadCaption.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "SubPublicityCaption", txtPublicityCaption.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Programs", txtPrograms.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Favorites", txtFavorites.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Desktop", txtDesktop.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "My Documents", txtMyDocuments.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Find", txtFind.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Run", txtRun.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Control Devices", txtControlDevices.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Recycled", txtRecycled.Text
    SaveSetting "Softshell Logi", "Startmenu Settings", "Shut-Down", txtShutdown.Text

    
    SaveSetting "Softshell Logi", "General Settings", "SystemIni", chIni.Value
    If chIni.Value = 0 Then
        '//--Write to System.ini and change to default..
        Va = INIWrite("boot", "oldshell", "c:\windows\skrivbord\Lo\Logi.exe", "c:\windows\system.ini")
        Va = INIWrite("boot", "shell", "Explorer.exe", "c:\windows\system.ini")
    Else
        '//--Write to System.ini and change to Softshell Logi start..
        Va = INIWrite("boot", "shell", "c:\windows\skrivbord\Lo\Logi.exe", "c:\windows\system.ini")
        Va = INIWrite("boot", "oldshell", "Explorer.exe", "c:\windows\system.ini")
    End If

    Call frmTaskbar.KontrollList
    
Unload Me

End Sub

Private Sub tsControl_Click()

    For i = 0 To picSettings.Count - 1
        picSettings(i).Visible = False
    Next
    
    picSettings(tsControl.SelectedItem.Index - 1).Visible = True
    
    Select Case (tsControl.SelectedItem.Index - 1)
        
        Case 0 '//--General Settings
            With picSettings(0)
                .Cls
                .ForeColor = vbHighlight
                .FontName = "Comic Sans Ms"
                .FontBold = True
                .FontSize = 10
            End With
                picSettings(0).CurrentX = 500
                picSettings(0).CurrentY = 750
                picSettings(0).Print "The Text on the Startbutton:"
                picSettings(0).CurrentX = 500
                picSettings(0).CurrentY = 1470
                picSettings(0).Print "ExeVisareDivisionNumber:"
                picSettings(0).CurrentX = 500
                picSettings(0).CurrentY = 2190
                picSettings(0).Print "ExeVisare.Width  (" & Space(10) & "):"
                picSettings(0).ForeColor = vbRed
                picSettings(0).CurrentX = 2500
                picSettings(0).CurrentY = 2190
                picSettings(0).Print "Caution"
                

        Case 1 '//--Startmenu Settings
            With picSettings(1)
                .Cls
                .ForeColor = vbHighlight
                .FontName = "Comic Sans Ms"
                .FontBold = True
                .FontSize = 10
            End With
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 630
                picSettings(1).Print "Sub Head Caption:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 990
                picSettings(1).Print "All Publicity Caption:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 1470
                picSettings(1).Print "Sub Programs Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 1830
                picSettings(1).Print "Sub Favorites Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 2190
                picSettings(1).Print "Sub Desktop Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 2550
                picSettings(1).Print "Sub My Documents:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 2910
                picSettings(1).Print "Sub Find Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 3270
                picSettings(1).Print "Sub Run Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 3630
                picSettings(1).Print "Sub Control_Devices Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 3990
                picSettings(1).Print "Sub Recycled Text:"
                picSettings(1).CurrentX = 500
                picSettings(1).CurrentY = 4350
                picSettings(1).Print "Sub Shut-Down Text:"
                picSettings(1).ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonShadowColor", vb3DShadow)
                picSettings(1).Line (480, 1370)-(4920, 1370)
                picSettings(1).ForeColor = GetSetting("Softshell Logi", "Color Settings", "CommonHighLightColor", vb3DHighlight)
                picSettings(1).Line (480, 1380)-(4920, 1380)
                
        Case 2 '//--Color Settings
            With picSettings(2)
                .Cls
                .ForeColor = vbHighlight
                .FontName = "Comic Sans Ms"
                .FontBold = True
                .FontSize = 10
            End With
                picSettings(2).CurrentX = 500
                picSettings(2).CurrentY = 750
                picSettings(2).Print "Taskbar BackColor:"
                picSettings(2).CurrentX = 500
                picSettings(2).CurrentY = 1470
                picSettings(2).Print "Taskbar ForeColor:"
                picSettings(2).CurrentX = 500
                picSettings(2).CurrentY = 2190
                picSettings(2).Print "Taskbar ShadowColor:"
                picSettings(2).CurrentX = 500
                picSettings(2).CurrentY = 2910
                picSettings(2).Print "CommonShadowColor:"
                picSettings(2).CurrentX = 500
                picSettings(2).CurrentY = 3630
                picSettings(2).Print "CommonHighLightColor:"

    End Select

End Sub

Private Sub cmdTaskbarForeColor_Click()
    cdbControl.CancelError = True
    On Error GoTo Errfix
   
    cdbControl.flags = cdlCCFullOpen + cdlCCRGBInit
    cdbControl.Color = shTaskbarForeColor.BackColor
    cdbControl.ShowColor
    shTaskbarForeColor.BackColor = cdbControl.Color
    Exit Sub
    
Errfix:
    Exit Sub
End Sub

Private Sub cmdStartbuttonShadow_Click()
    cdbControl.CancelError = True
    On Error GoTo Errfix
    
    cdbControl.flags = cdlCCFullOpen + cdlCCRGBInit
    cdbControl.Color = shStartbuttonShadow.BackColor
    cdbControl.ShowColor
    shStartbuttonShadow.BackColor = cdbControl.Color
    Exit Sub
    
Errfix:
    Exit Sub
End Sub

Private Sub cmdTaskbarBackcolor_Click()
    cdbControl.CancelError = True
    On Error GoTo Errfix
    
    cdbControl.flags = cdlCCFullOpen + cdlCCRGBInit
    cdbControl.Color = shTaskbarBackcolor.BackColor
    cdbControl.ShowColor
    shTaskbarBackcolor.BackColor = cdbControl.Color
    Exit Sub
    
Errfix:
    Exit Sub
End Sub

Private Sub cmdCommonShadowColor_Click()
    cdbControl.CancelError = True
    On Error GoTo Errfix
    
    cdbControl.flags = cdlCCFullOpen + cdlCCRGBInit
    cdbControl.Color = shCommonShadowColor.BackColor
    cdbControl.ShowColor
    shCommonShadowColor.BackColor = cdbControl.Color
    Exit Sub
    
Errfix:
    Exit Sub
End Sub

Private Sub cmdCommonHighLightcolor_Click()
    cdbControl.CancelError = True
    On Error GoTo Errfix
    
    cdbControl.flags = cdlCCFullOpen + cdlCCRGBInit
    cdbControl.Color = shCommonHighLightColor.BackColor
    cdbControl.ShowColor
    shCommonHighLightColor.BackColor = cdbControl.Color
    Exit Sub
    
Errfix:
    Exit Sub
End Sub

