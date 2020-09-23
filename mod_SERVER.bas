Attribute VB_Name = "mod_SERVER"
Public Const ComSep As String = ""
Public Const USep As String = "Ö«"

Public Type ClientN
    sName As String
    ID As Long
    Chans() As String
    gotName As Boolean
    WelcomeSent As Boolean
End Type

Public Type ChannelN
    cName As String
    cMaxUser As Long
    cTopic As String
End Type

Public FChannels() As ChannelN
Public Clients() As ClientN


Public Function NewClient(ByVal uID As Long) As Long
Dim NC As Long
For i = 0 To UBound(Clients)
    If Clients(i).ID = 0 Then
        With Clients(i)
            .ID = uID
            ReDim .Chans(1)
            .sName = ""
            .gotName = False
            .WelcomeSent = False
            NewClient = i
            Exit Function
        End With
    End If
Next
NC = UBound(Clients) + 1
ReDim Preserve Clients(NC)
With Clients(NC)
    .ID = uID
    .gotName = False
    .WelcomeSent = False
    ReDim .Chans(1)
    .sName = ""
End With
NewClient = NC
End Function

Public Function FindClient(ByVal uID As Long) As Long
For i = 0 To UBound(Clients)
    With Clients(i)
        If .ID = uID Then
            FindClient = i
            Exit Function
        End If
    End With
Next
FindClient = -255
End Function

Public Function FindClientByName(strName As String) As Long
For i = 0 To UBound(Clients)
    With Clients(i)
        If Trim(LCase(.sName)) = Trim(LCase(strName)) Then
            FindClientByName = i
            Exit Function
        End If
    End With
Next
FindClientByName = -255
End Function

Public Function FindChannel(strName As String) As Long
For i = 0 To UBound(FChannels)
    With FChannels(i)
        If Trim(LCase(.cName)) = Trim(LCase(strName)) Then
            FindChannel = i
            Exit Function
        End If
    End With
Next
FindChannel = -255
End Function

Public Sub UserJoinChannel(ByVal UserIndex As Long, channelName As String)
Dim TNC As Long
With Clients(UserIndex)
    For i = 0 To UBound(.Chans)
        If Trim(.Chans(i)) = "" Then
            TNC = i
            GoTo ADDIT:
        End If
    Next
    
    TNC = UBound(.Chans) + 1
    ReDim Preserve .Chans(TNC)
    
ADDIT:

.Chans(TNC) = Trim(channelName)

End With
End Sub

Public Function FormUserList(strChannel As String) As String
On Error Resume Next
Dim Ulist As String
For i = 0 To UBound(Clients)
    With Clients(i)
        'MsgBox UBound(.Chans)
        For ii = 0 To UBound(.Chans)
            If Trim(LCase(.Chans(ii))) = _
            Trim(LCase(strChannel)) Then
            'user is in this channel, add them
            'to the list
            If Trim(.sName) <> "" Then
                Ulist = Ulist & .sName & USep
            End If
            
            GoTo NextClient:
            End If
        Next
NextClient:
    End With
Next
FormUserList = Ulist
End Function

Public Sub SendToChannel(strChannel As String, strText As String, strUser As String)
Dim sockID As Long
On Error Resume Next
For i = 0 To UBound(Clients)
    With Clients(i)
        'MsgBox UBound(.Chans)
        For ii = 0 To UBound(.Chans)
            If Trim(LCase(.Chans(ii))) = _
            Trim(LCase(strChannel)) Then
            'user is in this channel, send them
            'the data
            sockID = FindSockID(i)
            If sockID <> -255 Then
                DoEvents
                frmMain.D(sockID).SendData "msg_chn" & ComSep & strChannel & ComSep & strUser & ComSep & strText
            End If
            
            GoTo NextClient:
            End If
        Next
NextClient:
    End With
Next
End Sub

Public Function findChannelEntry(ByVal uID As Long, strChannelName As String) As Long
With Clients(uID)
    For i = 0 To UBound(.Chans)
        If Trim(LCase(.Chans(i))) = Trim(LCase(strChannelName)) Then
            'found channel
            findChannelEntry = i
            Exit Function
        End If
    Next
End With
findChannelEntry = -255
End Function

Public Function FindSockID(ByVal uIndex As Long) As Long
For i = 0 To frmMain.D.UBound
    If frmMain.D(i).Tag = uIndex Then
        FindSockID = i
        Exit Function
    End If
Next
'not found
FindSockID = -255
End Function

Public Sub SaveChannelList()
Dim p As String
Dim os As String

If Right(App.Path, 1) = "\" Then
    p = App.Path
Else
    p = App.Path & "\"
End If

For i = 0 To UBound(FChannels)
    With FChannels(i)
        If Trim(.cName) <> "" Then
        os = os & .cName & USep & _
                .cTopic & USep & _
                .cMaxUser & vbCrLf
        End If
    End With
Next

Open p & "channels.lst" For Output As #1
    Print #1, os
Close #1
End Sub

Public Sub LoadChannelList()
On Error GoTo noList:
Dim bDAT As String

Dim SPL1() As String
Dim SPL2() As String

Dim p As String

If Right(App.Path, 1) = "\" Then
    p = App.Path
Else
    p = App.Path & "\"
End If

bDAT = Space$(FileLen(p & "channels.lst"))

Open p & "channels.lst" For Binary Access Read As #1
    Get #1, , bDAT
Close #1

SPL1 = Split(bDAT, vbCrLf)
ReDim FChannels(UBound(SPL1))

For i = 0 To UBound(SPL1)
    SPL2 = Split(SPL1(i), USep)
    If Len(SPL1(i)) > 1 Then
    With FChannels(i)
    If Trim(SPL2(0)) <> "" Then
        .cName = SPL2(0)
        .cTopic = SPL2(1)
        .cMaxUser = Val(SPL2(2))
    End If
    End With
    End If
Next

bDAT = ""

Exit Sub
noList:

With FChannels(0)
    .cName = "#Lobby"
End With
With FChannels(1)
    .cName = "#Admin"
End With

End Sub
