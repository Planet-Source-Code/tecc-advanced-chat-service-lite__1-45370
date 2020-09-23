VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ACS Server LITE"
   ClientHeight    =   5490
   ClientLeft      =   2115
   ClientTop       =   1590
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   4500
   End
   Begin VB.Frame Frame4 
      Caption         =   "Socket Status and Client list"
      Height          =   5115
      Left            =   4860
      TabIndex        =   9
      Top             =   60
      Width           =   3255
      Begin VB.Frame Frame5 
         Caption         =   "Control"
         Height          =   1275
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   3015
         Begin VB.Timer Timer2 
            Interval        =   200
            Left            =   240
            Top             =   720
         End
         Begin MSWinsockLib.Winsock D 
            Index           =   0
            Left            =   1620
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSWinsockLib.Winsock L 
            Left            =   1140
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
            LocalPort       =   1001
         End
         Begin VB.CommandButton cmdSTOP 
            Caption         =   "Stop Server"
            Height          =   315
            Left            =   1140
            TabIndex        =   14
            Top             =   300
            Width           =   1035
         End
         Begin VB.CommandButton cmdSTRT 
            Caption         =   "Start Server"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   1035
         End
      End
      Begin MSComctlLib.ListView lstSocks 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sock/User"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstUL 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   1980
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Channels"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Log"
      Height          =   1215
      Left            =   60
      TabIndex        =   7
      Top             =   3960
      Width           =   4875
      Begin VB.TextBox txtLog 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   2460
      TabIndex        =   5
      Top             =   60
      Width           =   2475
      Begin VB.ListBox lstUser 
         Height          =   3615
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSComctlLib.StatusBar Help1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5235
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1500
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":071E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Channels"
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2475
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   450
         Left            =   120
         TabIndex        =   3
         Top             =   3420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Remove"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Add"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstChan 
         Height          =   2535
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblChan 
         Caption         =   "Select a channel to view active users in that channel"
         ForeColor       =   &H8000000C&
         Height          =   555
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSTOP_Click()
StopServer
End Sub

Private Sub cmdSTRT_Click()
StartServer
End Sub

Private Sub D_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim dat As String
Dim ULL As String
Dim Dspl() As String
Dim ReQ As String
Dim TC As Long
Dim CHID As Long
Dim CHENT As Long

dat = ""

D(Index).GetData dat
If Len(dat) > 1 Then
Dspl = Split(dat, ComSep)
ReQ = LCase(Dspl(0))
Else
    Exit Sub
End If

DoEvents

TC = D(Index).Tag
With Clients(TC)
Select Case ReQ
    Case "user" 'user name sent by client
        If Trim(Dspl(1)) <> "" Then
            If FindClientByName(Dspl(1)) = -255 Then
            .sName = Dspl(1)
            .gotName = True
            txtLog.Text = txtLog.Text & "User sent name [" & .sName & "] | " & Index & " | " & .ID & " :" & vbCrLf
            D(Index).SendData "res_001" & ComSep
            
            Else
                'client name exists!
                D(Index).SendData "err_001" & ComSep
            End If
        Else
            .gotName = False
            'name invalid
            D(Index).SendData "err_002" & ComSep
        End If
        
    Case "join" 'user join channel
        If Not (.gotName) Then
            GoTo NoPermission:
        End If
        
        If Trim(Dspl(1)) <> "" Then
            CHID = FindChannel(Dspl(1))
            If CHID <> -255 Then
                'channel exists, join it
                For i = 0 To UBound(.Chans)
                    If Trim(LCase(.Chans(i))) _
                    = LCase(Trim(FChannels(CHID).cName)) Then
                        'user is already in this channel
                        D(Index).SendData "err_006" & ComSep & FChannels(CHID).cName
                        GoTo NoJoin:
                    Else
                        'user is not in this channel, join it
                        UserJoinChannel TC, FChannels(CHID).cName
                        D(Index).SendData "res_003" & ComSep & FChannels(CHID).cName & ComSep & FChannels(CHID).cTopic
                        GoTo Joined:
                    End If
                Next
Joined:
NoJoin:
            Else
                'channel does not exist
                D(Index).SendData "err_005" & ComSep
            End If
        Else
            D(Index).SendData "err_004" & ComSep
        End If
        
        
    Case "list"
        If Not (.gotName) Then
            GoTo NoPermission:
        End If
        
        If Trim(Dspl(1)) <> "" Then
            CHID = FindChannel(Dspl(1))
                If CHID <> -255 Then
                    ULL = FormUserList(FChannels(CHID).cName)
                    D(Index).SendData "res_004" & ComSep & FChannels(CHID).cName & ComSep & ULL
                Else
                    D(Index).SendData "err_005" & ComSep
                End If
        Else
            D(Index).SendData "err_004" & ComSep
        End If
        
    Case "quit" 'user quit channel
        If Trim(Dspl(1)) <> "" Then
            CHID = FindChannel(Dspl(1))
                If CHID <> -255 Then
                    CHENT = findChannelEntry(TC, FChannels(CHID).cName)
                    If CHENT <> -255 Then
                        .Chans(CHENT) = ""
                        D(Index).SendData "res_005" & ComSep & FChannels(CHID).cName
                    Else
                        D(Index).SendData "err_008" & ComSep & FChannels(CHID).cName
                    End If
                Else
                    D(Index).SendData "err_005" & ComSep
                End If
        Else
            D(Index).SendData "err_004" & ComSep
        End If
    
    Case "send" 'user send to channel
        If Not (.gotName) Then
            GoTo NoPermission:
        End If
        
        If Trim(Dspl(1)) <> "" Then
            CHID = FindChannel(Dspl(1))
                If CHID <> -255 Then
                    CHENT = findChannelEntry(TC, FChannels(CHID).cName)
                    If CHENT <> -255 Then
                        'data to be sent to channel
                        If Trim(Dspl(2)) <> "" Then
                        DoEvents
                        SendToChannel FChannels(CHID).cName, Dspl(2), .sName
                        Else
                            'invalid message
                            D(Index).SendData "err_009" & ComSep & FChannels(CHID).cName
                        End If
                    Else
                        D(Index).SendData "err_008" & ComSep & FChannels(CHID).cName
                    End If
                Else
                    D(Index).SendData "err_005" & ComSep
                End If
        Else
            D(Index).SendData "err_004" & ComSep
        End If
    
    Case "clst" 'user channel listing
        If Not (.gotName) Then
            GoTo NoPermission:
        End If
        
        Dim CL As String
        For i = 0 To UBound(FChannels)
            With FChannels(i)
                If Trim(.cName) <> "" Then
                    CL = CL & .cName & USep
                End If
            End With
        Next
        DoEvents
        D(Index).SendData "res_007" & ComSep & CL
        DoEvents
        CL = ""
        
    Case Else
        D(Index).SendData "err_003" & ComSep & dat
        dat = ""
End Select
End With

Erase Dspl
dat = ""

Exit Sub
NoPermission:
D(Index).SendData "err_007" & ComSep
End Sub

Private Sub Form_Load()
Clipboard.Clear
Clipboard.SetText Chr(6) & Chr(7)

lstSocks.ColumnHeaders(1).Width = (lstSocks.Width / 3)
lstSocks.ColumnHeaders(2).Width = lstSocks.ColumnHeaders(1).Width
lstSocks.ColumnHeaders(3).Width = lstSocks.ColumnHeaders(1).Width - 64

lstUL.ColumnHeaders(1).Width = (lstUL.Width / 3)
lstUL.ColumnHeaders(2).Width = lstUL.ColumnHeaders(1).Width
lstUL.ColumnHeaders(3).Width = lstUL.ColumnHeaders(1).Width - 64

ReDim Clients(1)
ReDim FChannels(1)

LoadChannelList
End Sub

Private Sub L_ConnectionRequest(ByVal requestID As Long)
Dim os As Long
Dim NewTag As Long
os = findOpenSocket
D(os).Accept requestID
D(os).Tag = GenID
NewTag = NewClient(D(os).Tag)
D(os).Tag = NewTag
'new client
txtLog.Text = txtLog.Text & "New Client " & D(os).Tag & " | " & Clients(D(os).Tag).ID & ":" & vbCrLf
End Sub

Private Sub L_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
StopServer
End Sub

Private Sub lstChan_Click()
Dim UL2 As String
Dim SPL1() As String
If Trim(lstChan.Text) <> "" Then
    UL2 = FormUserList(lstChan.Text)
Else
    Exit Sub
End If
lstUser.Clear
SPL1 = Split(UL2, USep)
For i = 0 To UBound(SPL1)
    If Trim(SPL1(i)) <> "" Then
        lstUser.AddItem SPL1(i)
    End If
Next
End Sub

Private Sub lstChan_DblClick()

Dim CHID As Long

If Trim(lstChan.Text) <> "" Then
    CHID = FindChannel(lstChan.Text)
    If CHID <> -255 Then
        Dim aa As New frmChat
        aa.Show
        aa.Tag = CHID
    End If
Else
    Exit Sub
End If

End Sub

Private Sub Timer1_Timer()
Select Case L.State
    Case sckListening
        cmdSTRT.Enabled = False
        cmdSTOP.Enabled = True
    Case sckClosed
        cmdSTRT.Enabled = True
        cmdSTOP.Enabled = False
    Case Else
        L.Close
End Select

For i = 0 To D.ubound
    With D(i)
        Select Case .State
            Case sckConnected
                With Clients(.Tag)
                    If Not (.WelcomeSent) Then
                        D(i).SendData "res_002" & ComSep
                        .WelcomeSent = True
                    End If
                End With
            Case sckConnecting
            Case sckClosed
            Case Else
                With Clients(.Tag)
                    txtLog.Text = txtLog.Text & "Client disconnected " & i & " | " & .ID & " | " & .sName & " :" & vbCrLf
                    Erase .Chans
                    .sName = ""
                    .ID = 0
                    .gotName = False
                    .WelcomeSent = False
                    
                End With
                .Close
                
        End Select
    End With
Next


End Sub

Private Sub StartServer()
L.Listen
txtLog.Text = txtLog.Text & "Server Started on port " & L.LocalPort & vbCrLf
End Sub

Private Sub StopServer()
If L.State <> sckClosed Then L.Close
For i = 0 To D.ubound
    If D(i).State <> sckClosed Then D(i).Close
    If i <> 0 Then
        Unload D(i)
    End If
Next
D(0).Tag = 0
ReDim Clients(1)
txtLog.Text = txtLog.Text & "Server Stopped" & vbCrLf

End Sub

Private Function findOpenSocket() As Long
Dim NS As Long
For i = 0 To D.ubound
    Select Case D(i).State
        Case sckConnected, sckConnecting
        Case Else
            If D(i).State <> sckClosed Then D(i).Close
            D(i).Tag = 0
            findOpenSocket = i
            Exit Function
    End Select
Next
NS = D.ubound + 1
Load D(NS)
D(NS).Tag = 0
findOpenSocket = NS
End Function

Private Function GenID() As Long
Randomize
GenID = Int(Rnd * 1000000000)
End Function

Private Sub Timer2_Timer()
On Error Resume Next
Dim lastSel As Long

Dim li As ListItem

lstSocks.ListItems.Clear
For i = 0 To D.ubound
    If D(i).State = sckConnected Then
        Set li = lstSocks.ListItems.Add(, , i & "/" & Clients(D(i).Tag).sName)
            li.SubItems(1) = D(i).State
            li.SubItems(2) = D(i).Tag
    End If
Next

Dim cs1 As String
Dim xx As String

lstUL.ListItems.Clear
For i = 0 To UBound(Clients)
    With Clients(i)
    cs1 = ""
    If .ID <> 0 Then
        If .sName <> "" Then
            xx = .sName
        Else
            xx = "Waiting..."
        End If
        
        Set li = lstUL.ListItems.Add(, , xx)
            For ii = 0 To UBound(.Chans)
                If Trim(.Chans(ii)) <> "" Then
                    cs1 = cs1 & .Chans(ii) & ","
                End If
            Next
            li.SubItems(1) = cs1
            li.SubItems(2) = .ID
            
    End If
    End With
Next

lastSel = lstChan.ListIndex
lstChan.Clear
For i = 0 To UBound(FChannels)
    With FChannels(i)
        If Trim(.cName) <> "" Then
            lstChan.AddItem .cName
        End If
    End With
Next
lstChan.ListIndex = lastSel
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim CHID As Long
Select Case Button.Index
    Case 1 'edit
        
    Case 2 'del
        If Trim(lstChan.Text) <> "" Then
            CHID = FindChannel(lstChan.Text)
            With FChannels(CHID)
                .cName = ""
                .cMaxUser = 0
                .cTopic = ""
            End With
            SaveChannelList
        End If
    Case 3 'add
        frmAddChannel.Show
End Select
End Sub

Private Sub txtLog_Change()
txtLog.SelStart = Len(txtLog.Text)
End Sub
