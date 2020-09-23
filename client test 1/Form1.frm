VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   1965
   ClientTop       =   1755
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6690
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Text            =   "Lobby"
      Top             =   3420
      Width           =   3735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Send"
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      Top             =   3060
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   2400
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      Caption         =   "List"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1980
      Width           =   1155
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   3420
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "J2"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   1560
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Join"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1140
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   420
      TabIndex        =   4
      Text            =   "Teccrc"
      Top             =   2640
      Width           =   2955
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   2
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   435
      Left            =   5400
      TabIndex        =   1
      Top             =   180
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   2355
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   3315
   End
   Begin MSWinsockLib.Winsock w 
      Left            =   2640
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   1001
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const ComSep As String = ""
Private Const USep As String = "Ö«"

Private Sub Command1_Click()
w.Connect

End Sub

Private Sub Command2_Click()
w.Close

End Sub

Private Sub Command3_Click()
w.SendData "join" & ComSep & Text4.Text
End Sub

Private Sub Command4_Click()
w.SendData "join" & ComSep & "ass"
End Sub

Private Sub Command5_Click()
w.SendData "list" & ComSep & Text4.Text
End Sub

Private Sub Command6_Click()
w.SendData "quiT" & ComSep & Text4.Text
End Sub

Private Sub Command7_Click()
w.SendData "send" & ComSep & Text4.Text & ComSep & Text3.Text
End Sub

Private Sub w_DataArrival(ByVal bytesTotal As Long)
Dim d As String
Dim ds() As String
Dim ul() As String
w.GetData d
Text1.Text = Text1.Text & d & vbCrLf
Select Case Left(d, 7)
Case "res_002"
    w.SendData "user" & ComSep & Text2.Text
Case "res_001"
    Me.Caption = "Connected!"
Case "res_004"
    ds = Split(d, ComSep)
    ul = Split(ds(2), USep)
    On Error Resume Next
    List1.Clear
    For i = 0 To UBound(ul)
        
        List1.AddItem ul(i)
    Next
Case "res_003"
    w.SendData "list" & ComSep & Text4.Text
Case Else
End Select
End Sub

