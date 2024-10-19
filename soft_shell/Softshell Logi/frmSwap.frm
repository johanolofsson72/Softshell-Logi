VERSION 5.00
Begin VB.Form frmSwap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Softshell Swap..........."
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Explorer"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.Label Label2 
         Caption         =   "If you run Softshell alone (without Explorer) then hit the [Swap-Button]..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "If you running Softshell with Explorer in the background hit       [Show Explorer-Button]...     "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frmSwap.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdShow_Click()
Call frmTaskbar.UnloadLogi
End Sub

Private Sub cmdSwap_Click()
'//--Unload Softshell Logi..
Call frmTaskbar.UnloadLogi
'//--Shutdown and restart the computer..
Shutdown 2, 0
End Sub

Private Sub Form_Load()
'//--Check system..
If GetSetting("Softshell Logi", "General Settings", "SystemIni", 0) = 0 Then
    cmdSwap.Enabled = False
    cmdShow.Enabled = True
Else
    cmdSwap.Enabled = True
    cmdShow.Enabled = False
End If

End Sub
