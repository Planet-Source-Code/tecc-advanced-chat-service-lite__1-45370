VERSION 5.00
Begin VB.Form frmChat 
   Caption         =   "Chat Window"
   ClientHeight    =   4290
   ClientLeft      =   3945
   ClientTop       =   2400
   ClientWidth     =   5805
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
   ScaleHeight     =   4290
   ScaleWidth      =   5805
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   3360
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
   Begin VB.Menu mnuChannel 
      Caption         =   "&Channel"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Resize()
txtChat.Move 0, 0, Me.ScaleWidth - lstUsers.Width, Me.ScaleHeight - cmdSend.Height
lstUsers.Move Me.ScaleWidth - lstUsers.Width, 0, lstUsers.Width, txtChat.Height
txtSend.Move 0, Me.ScaleHeight - txtSend.Height, Me.ScaleWidth - cmdSend.Width
cmdSend.Move Me.ScaleWidth - cmdSend.Width, Me.ScaleHeight - cmdSend.Height
End Sub

Private Sub Timer1_Timer()
Me.Caption = "Chat window [" & FChannels(Me.Tag).cName & "]"
End Sub
