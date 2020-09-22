VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About MyApp"
   ClientHeight    =   3435
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370.898
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   379.26
      ScaleMode       =   0  'User
      ScaleWidth      =   379.26
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3000
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5408.938
      Y1              =   1687.582
      Y2              =   1687.582
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAbout.frx":0F72
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Archangel"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 1.1"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAbout.frx":1062
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   255
      TabIndex        =   3
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
'unload the about window
  Unload Me
End Sub

Private Sub Form_Load()
'set the window title
    Me.Caption = "About " & App.Title
End Sub
