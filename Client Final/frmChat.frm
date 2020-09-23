VERSION 5.00
Begin VB.Form frmChat 
   Caption         =   "Chat Window"
   ClientHeight    =   4290
   ClientLeft      =   2715
   ClientTop       =   2760
   ClientWidth     =   5805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   5805
   Begin VB.Timer listUpd 
      Interval        =   5000
      Left            =   3240
      Top             =   3660
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   3300
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1200
      Top             =   3480
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   4620
      TabIndex        =   3
      Top             =   2940
      Width           =   855
   End
   Begin VB.TextBox txtSend 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2940
      Width           =   4335
   End
   Begin VB.ListBox lstUsers 
      Height          =   2685
      IntegralHeight  =   0   'False
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   1395
   End
   Begin VB.TextBox txtChat 
      Height          =   2655
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   3795
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ChatName As String
Public Topic As String

Private Sub cmdQuit_Click()
frmMain.W.SendData "quit" & ComSep & ChatName
End Sub

Private Sub cmdSend_Click()
If Trim(txtSend.Text) <> "" Then
    frmMain.W.SendData "send" & ComSep & ChatName & ComSep & txtSend.Text
    txtSend.Text = ""
End If
End Sub

Private Sub Form_Load()
If Trim(Topic) = "" Then
    Topic = "No Topic Specified"
End If
txtChat.Text = "Topic: [" & Topic & "]" & vbCrLf & "Joined " & ChatName & " - " & DateTime.Now & vbCrLf & vbCrLf
End Sub

Private Sub Form_Resize()
txtChat.Move 0, 0, Me.ScaleWidth - lstUsers.Width, Me.ScaleHeight - cmdSend.Height
lstUsers.Move Me.ScaleWidth - lstUsers.Width, 0, lstUsers.Width, txtChat.Height
txtSend.Move cmdQuit.Width, Me.ScaleHeight - txtSend.Height, Me.ScaleWidth - (cmdSend.Width + cmdQuit.Width)
cmdSend.Move Me.ScaleWidth - cmdSend.Width, Me.ScaleHeight - cmdSend.Height
cmdQuit.Move 0, Me.ScaleHeight - cmdQuit.Height

End Sub

Private Sub mnuQuit_Click()

End Sub

Private Sub listUpd_Timer()
If frmMain.W.State = 7 Then
frmMain.W.SendData "list" & ComSep & ChatName
End If
End Sub

Private Sub Timer1_Timer()
Me.Caption = "Chatting: [" & ChatName & "]"
End Sub

Private Sub txtChat_Change()
txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub
