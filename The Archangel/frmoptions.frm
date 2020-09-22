VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmoptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdata 
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Other Things"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   6240
      Width           =   3375
      Begin VB.TextBox txtnickname 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nickname"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Security Options"
      Height          =   3015
      Left            =   10
      TabIndex        =   12
      Top             =   3120
      Width           =   3375
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Twofish"
         Height          =   375
         Index           =   6
         Left            =   1680
         TabIndex        =   25
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TEA"
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Skipjack"
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rijndael"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gost"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DES"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Blowfish"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtpassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Choose the encryption algorithm"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter the encryption code"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Server Connection Options"
      Height          =   1095
      Left            =   10
      TabIndex        =   9
      Top             =   720
      Width           =   3375
      Begin VB.TextBox srvport 
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter the port's number to listen to"
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Client Connextion Options"
      Height          =   1095
      Left            =   10
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
      Begin VB.TextBox clientip 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox clientport 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter the server's IP"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter the server's port"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Work as . . ."
      Height          =   615
      Left            =   10
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton optclient 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Client"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optserver 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It!!!"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   3720
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare everything needed!
'
'{wizard}


Private Sub Form_Load()
    
    w1.Protocol = sckTCPProtocol
    
End Sub

Private Sub Command1_Click()

'check if no working mode has been chosen
If optserver.Value = False And optclient.Value = False Then
MsgBox "You have to choose between two working modes. Client or Server.", vbOKOnly + vbExclamation, "Ooooopssss..."
Exit Sub
End If

'check for any sort of mistakes like empty password or nickname box.
If txtpassword.Text = "" Then
MsgBox "Please fill the Encryption Code box!!!", vbInformation + vbOKOnly, "Ooooopssss..."
Exit Sub
End If
If txtnickname.Text = "" Then
MsgBox "Please fill the Nickname box!!!", vbInformation + vbOKOnly, "Ooooopssss..."
Exit Sub
End If


'lets move deaper


'1st check
'if Server working mode has been chosen then check for more errors
'and if no errors occur listen to the specific port and listen for any connection
If optserver.Value = True Then
    If srvport.Text = "" Then
    MsgBox "Please fill the Server Port box!!!", vbInformation + vbOKOnly, "Ooooopssss..."
    Exit Sub
    End If
    
    w1.LocalPort = srvport.Text
    w1.Listen
End If

'2nd check
'if Client working mode has been chosen then check for more errors
'and if no errors occur connect to the specific IP and port
If optclient.Value = True Then
    If clientip.Text = "" Then
    MsgBox "Please fill the Client IP box!!!", vbInformation + vbOKOnly, "Ooooopssss..."
    Exit Sub
    End If

    If clientport.Text = "" Then
    MsgBox "Please fill the Client Port box!!!", vbInformation + vbOKOnly, "Ooooopssss..."
    Exit Sub
    End If
    
    w1.RemoteHost = clientip.Text
    w1.RemotePort = clientport.Text
    w1.Connect

End If

'then show the main window
frmchat.Show
'and hide the options window
Me.Hide

End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
     CryptoEngine1.CryptAlgorithm = Index
End Sub

Private Sub w1_Connect()
Dim welcomess

'on connect set the welcome message
welcomemess = "Welcome to The Archangel v1.1. Created from {wizard}"
ConMessage = "Connection established with " + clientip.Text
frmchat.List1.AddItem welcomemess
frmchat.List1.AddItem ConMessage
End Sub

Private Sub w1_ConnectionRequest(ByVal requestID As Long)
Dim welcomemess

'if a connection is requested then accept the connection and set the welcome message
w1.Close
w1.Accept requestID
welcomemess = "Welcome to The Archangel v1.1. Created from {wizard}"
ConMessage = "Connection established with " + w1.RemoteHostIP
frmchat.List1.AddItem welcomemess
frmchat.List1.AddItem ConMessage
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String

'when the data arrive get them, decrypt them and paste them on the chat window
w1.GetData strdata
txtdata.Text = strdata

txtdata = frmchat.CryptoEngine1.DecryptString(txtdata, txtpassword, True)

frmchat.List1.AddItem txtdata
End Sub

Private Sub optclient_Click()
'check wich options will be available
If optclient.Value = True Then
srvport.Enabled = False
clientip.Enabled = True
clientport.Enabled = True
End If
End Sub

Private Sub optserver_Click()
'check wich options will be available
If optserver.Value = True Then
clientip.Enabled = False
clientport.Enabled = False
srvport.Enabled = True
End If
End Sub
