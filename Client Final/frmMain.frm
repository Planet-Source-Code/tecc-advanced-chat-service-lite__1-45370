VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ACS Lite Client"
   ClientHeight    =   5475
   ClientLeft      =   2910
   ClientTop       =   2550
   ClientWidth     =   6585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "ADS"
   NegotiateToolbars=   0   'False
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5220
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11086
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   3660
      Top             =   1620
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   1980
      Top             =   2340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuClient 
      Caption         =   "&Client"
      Begin VB.Menu mnucon 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDis 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTls 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCFG 
         Caption         =   "&Configuration"
      End
   End
   Begin VB.Menu mnuChan 
      Caption         =   "&Channel"
      Begin VB.Menu mnuJC 
         Caption         =   "&Join channel..."
      End
      Begin VB.Menu mnuLC 
         Caption         =   "&List channels"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
LoadSettings
ReDim ChatWindows(1)

End Sub

Private Sub mnuCFG_Click()
frmCFG.Show
End Sub

Private Sub mnucon_Click()
If Trim(MainP.nServer) <> "" Then
    If W.State <> sckClosed Then W.Close
    W.Connect MainP.nServer, MainP.lPort
End If
End Sub

Private Sub mnuDis_Click()
DoEvents
If W.State <> sckClosed Then W.Close
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuJC_Click()
Dim inp As String
inp = InputBox("Enter the channel name you wish to join", "Join channel", "#Lobby")
If Trim(inp) <> "" Then
    W.SendData "join" & ComSep & inp
End If
End Sub

Private Sub mnuLC_Click()
W.SendData "clst" & ComSep
End Sub

Private Sub Timer1_Timer()
Select Case W.State
    Case sckConnected, sckConnecting
        Me.Caption = "ACS Lite Client [Connected]"
    Case sckClosed
        Me.Caption = "ACS Lite Client"
    Case Else
        W.Close
End Select
End Sub

Private Sub W_DataArrival(ByVal bytesTotal As Long)
Dim dat As String
Dim spl() As String

W.GetData dat
spl = Split(dat, ComSep)

Select Case Trim(LCase(spl(0)))
    Case "res_001"
    DoEvents
        Stat.Panels(1).Text = "Username Accepted"
        If MainP.sLobbyJoin = 1 Then
            W.SendData "join" & ComSep & "#lobby"
        End If
        
    Case "res_002"
        Stat.Panels(1).Text = "Connected, Send username"
        W.SendData "user" & ComSep & MainP.sUser
    Case "res_003"
        Stat.Panels(1).Text = "Channel Joined ok: " & spl(1)
        For i = 0 To UBound(ChatWindows)
            With ChatWindows(i)
                If .ChatName = "" Then
                    .Show
                    .Topic = spl(2)
                    .ChatName = Trim(spl(1))
                    .txtChat.Text = "Topic: " & spl(2) & vbCrLf & "Joined " & spl(1) & " - " & DateTime.Now & vbCrLf & vbCrLf
                    GoTo done1:
                End If
            End With
        Next
        ReDim Preserve ChatWindows(UBound(ChatWindows) + 1)
        
        With ChatWindows(UBound(ChatWindows))
            .Show
            .ChatName = Trim(spl(1))
        End With
        
done1:
        
        DoEvents
        W.SendData "list" & ComSep & Trim(spl(1))
        
    Case "res_004"
        Dim USPL() As String
        Stat.Panels(1).Text = "User list recieved: " & spl(1)
        For i = 0 To UBound(ChatWindows)
            With ChatWindows(i)
                If .ChatName = spl(1) Then
                    .lstUsers.Clear
                    USPL = Split(spl(2), USep)
                    
                    For ii = 0 To UBound(USPL)
                        If Trim(USPL(ii)) <> "" Then
                        .lstUsers.AddItem USPL(ii)
                        End If
                    Next
                    
                End If
            End With
        Next
    Case "res_005"
        Stat.Panels(1).Text = "Quit from channel: " & spl(1)
        For i = 0 To UBound(ChatWindows)
            With ChatWindows(i)
                If .ChatName = spl(1) Then
                    ChatWindows(i).ChatName = ""
                    Unload ChatWindows(i)
                End If
            End With
        Next
    Case "err_001"
        Stat.Panels(1).Text = "Username exists on server!"
    Case "err_002"
        Stat.Panels(1).Text = "Username invalid"
    Case "err_003"
        Stat.Panels(1).Text = "Invalid Command  " & spl(1)
        
    Case "err_004"
        Stat.Panels(1).Text = "Invalid Channel"
    Case "err_005"
        Stat.Panels(1).Text = "Channel does not exist!"
        MsgBox "No channel exists with that name!", vbCritical, "Error"
    Case "err_006"
        Stat.Panels(1).Text = "You are already joined to this channel!"
        MsgBox "You are already in this channel!", vbCritical, "Error"
    Case "err_007"
        Stat.Panels(1).Text = "No permission!"
    Case "err_008"
        Stat.Panels(1).Text = "You are not a member of that channel!"
        MsgBox "You must join that channel first!", vbCritical, "Error"
    Case "err_009"
        Stat.Panels(1).Text = "Invalid message, not sent"
    Case "res_007"
        Dim spl2() As String
        spl2 = Split(spl(1), USep)
        frmList.Show
        frmList.lstChans.Clear
        For i = 0 To UBound(spl2)
            If Trim(spl2(i)) <> "" Then
                frmList.lstChans.AddItem spl2(i)
            End If
        Next
        Stat.Panels(1).Text = "Got channel list"
    Case "msg_chn"
        For i = 0 To UBound(ChatWindows)
            With ChatWindows(i)
                If .ChatName = spl(1) Then
                    .txtChat.Text = .txtChat.Text & spl(2) & ": " & spl(3) & vbCrLf
                End If
            End With
        Next
End Select
End Sub

