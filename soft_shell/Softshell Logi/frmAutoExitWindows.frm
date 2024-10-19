VERSION 5.00
Begin VB.Form frmAutoExitWindows 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3720
   ClientLeft      =   4710
   ClientTop       =   2415
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   5760
      Top             =   1800
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmAutoExitWindows.frx":0000
      Left            =   1440
      List            =   "frmAutoExitWindows.frx":0002
      TabIndex        =   4
      Text            =   "Stänga av datorn."
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Avbryt"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Off"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   975
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BETA  1.2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Softshell Logi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Softworld"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   480
      Picture         =   "frmAutoExitWindows.frx":0004
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   360
      Picture         =   "frmAutoExitWindows.frx":B1C6
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Vad vill du göra..?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmAutoExitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4

Dim ExitVal(4) As String

Private Sub Command4_Click() 'avbryt
Command3.Enabled = True
Command4.Enabled = False
Timer1.Enabled = False
Text1.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()
If Combo1.Text = "Aktivera Timer." Then _
    Timer1.Enabled = True: Label2.Caption = _
    "Windows kommer att avslutas när klockan är: " & Text1.Text: Exit Sub
If Combo1.Text = "Deaktivera Timer." Then _
    Timer1.Enabled = False: Label2.Caption = "": Exit Sub
If Combo1.Text = "Stänga av datorn." Then _
    Shutdown ExitVal(0), 0: Timer1.Enabled = False: Exit Sub: End
If Combo1.Text = "Starta om datorn." Then _
    Shutdown ExitVal(1), 0: Timer1.Enabled = False: Exit Sub: End
If Combo1.Text = "Logga ut." Then _
    Shutdown ExitVal(2), 0: Timer1.Enabled = False: Exit Sub: End
End Sub

Private Sub Form_Load()

Combo1.AddItem "Stänga av datorn."
Combo1.AddItem "Starta om datorn."
Combo1.AddItem "Logga ut."
Combo1.AddItem ""
Combo1.AddItem "Aktivera Timer."
Combo1.AddItem "Deaktivera Timer."
  
ExitVal(0) = EWX_SHUTDOWN
ExitVal(1) = EWX_REBOOT
ExitVal(2) = EWX_LOGOFF
    
Text1.Text = Time
    
End Sub

Private Sub Timer1_Timer()
Dim CurrentTime
 CurrentTime = Format(Time, "hh:mm")
 If CurrentTime >= Format(Text1.Text, "hh:mm") Then
  Shutdown ExitVal(0), 0
  Timer1.Enabled = False
  End
 End If
End Sub

