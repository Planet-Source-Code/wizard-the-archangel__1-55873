VERSION 5.00
Begin VB.Form frmchat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Archangel"
   ClientHeight    =   4275
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdata 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   6855
   End
   Begin TheArchangel.CryptoEngine CryptoEngine1 
      Left            =   7200
      Top             =   120
      _extentx        =   1614
      _extenty        =   1614
   End
   Begin VB.TextBox txtmessage 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   6855
   End
   Begin VB.ListBox List1 
      Height          =   3765
      ItemData        =   "frmchat.frx":0000
      Left            =   0
      List            =   "frmchat.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnusave 
         Caption         =   "Save Conversation"
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnumore 
      Caption         =   "&More"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare everything needed!


Private Sub Form_Unload(Cancel As Integer)
'if the form unloads then end the program
End
End Sub

Private Sub mnuabout_Click()
'if the about menu is clicked then show the about window
frmAbout.Show
End Sub

Private Sub mnuexit_Click()
'show a question before exiting,if yes is the reply exit,otherwise do nothing
reply = MsgBox("Are you sure you want to exit???", vbYesNo + vbExclamation, "Exit???")
If reply = vbYes Then
End
Else
'and set the focus to the message box
txtmessage.SetFocus
End If
End Sub


Private Sub mnusave_Click()

'show an input box and save the conversation with the given filename
        FlName = InputBox("Enter the file's name. eg. conversation.txt", "Save As...")
        
        Open FlName For Output As #1
        Print #1, List1.Text
        Close #1
     
End Sub

Private Sub txtmessage_KeyDown(KeyCode As Integer, Shift As Integer)
'on error do what ErrHandler sayzzz
On Error GoTo errhandler

'get the message to a string like strdata, encrypt it and send it

If KeyCode = 13 Then
txtdata = frmoptions.txtnickname.Text + "> " + txtmessage.Text

List1.AddItem txtdata

txtdata = CryptoEngine1.EncryptString(txtdata, frmoptions.txtpassword, True)

frmoptions.w1.SendData txtdata

txtmessage.Text = ""

txtdata.Text = ""

Exit Sub
End If

errhandler:
    'if the error is 40006 that means that a connection has not been established
    'so show the proper message without closing the program
    If Err.Number = 40006 Then
    MsgBox "Connection has not been established yet...", vbInformation + vbOKOnly, "Error"
    txtmessage.Text = ""
    End If
        
End Sub
