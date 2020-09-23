VERSION 5.00
Begin VB.Form frmAddChannel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Channel"
   ClientHeight    =   2580
   ClientLeft      =   4005
   ClientTop       =   3360
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddChannel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkaa 
      Caption         =   "Add Another"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2100
      TabIndex        =   8
      Top             =   2100
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3180
      TabIndex        =   7
      Top             =   2100
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Misc"
      Height          =   1215
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   3975
      Begin VB.TextBox txtUL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   2475
      End
      Begin VB.TextBox txtTopic 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label Label2 
         Caption         =   "User Limit:"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Topic:"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Channel Name"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3975
      Begin VB.TextBox txtChan 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmAddChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim UTC As Long

If Trim(txtUL.Text) <> "" Then
    If FindChannel(txtUL.Text) = -255 Then
        
        For i = 0 To UBound(FChannels)
            With FChannels(i)
                If Trim(.cName) = "" Then
                    UTC = i
                    GoTo formNew:
                End If
            End With
        Next
    
    UTC = UBound(FChannels) + 1
    ReDim Preserve FChannels(UTC)
    
formNew:

    With FChannels(UTC)
        .cName = txtChan.Text
        .cMaxUser = Val(txtUL.Text)
        .cTopic = txtTopic.Text
    End With
    
        GoTo Finish:
    
    
    Else
        MsgBox "Channel exists! choose a unique name", vbCritical, "Error"
        GoTo NONAME:
    End If
Else
    MsgBox "Channel name is invalid!", vbCritical, "Error"
    GoTo NONAME:
End If

Finish:
SaveChannelList
If chkaa.Value = 0 Then
    Unload Me
Else
    txtUL.Text = 0
    txtTopic.Text = ""
    txtChan.Text = ""
    txtChan.SetFocus
End If

Exit Sub
NONAME:
txtChan.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
