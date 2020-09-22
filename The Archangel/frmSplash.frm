VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4440
      Top             =   2160
   End
   Begin VB.Label cover 
      BackColor       =   &H80000014&
      Height          =   510
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4530
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   120
      Picture         =   "frmSplash.frx":0BC2
      Top             =   1200
      Width           =   4530
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   360
      Picture         =   "frmSplash.frx":128D
      Top             =   120
      Width           =   3840
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
cover.Left = cover.Left + 10
If cover.Left = 5200 Then
Timer1.Enabled = False
frmoptions.Show
Unload Me
End If
End Sub
