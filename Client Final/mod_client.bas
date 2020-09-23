Attribute VB_Name = "mod_client"
Public Const ComSep As String = ""
Public Const USep As String = "Ö«"

Public Type nClient
    sUser As String
    nServer As String
    lPort As Long
    sLobbyJoin As Byte
End Type

Public ChatWindows() As New frmChat
Public MainP As nClient

Public Sub SaveSettings()
With MainP
    SaveSetting App.EXEName, "Settings", "User", .sUser
    SaveSetting App.EXEName, "Settings", "Port", .lPort
    SaveSetting App.EXEName, "Settings", "Server", .nServer
    SaveSetting App.EXEName, "Settings", "Ajoin", .sLobbyJoin
End With
End Sub
Public Sub LoadSettings()
With MainP
    Randomize
    .sUser = GetSetting(App.EXEName, "Settings", "User", "Guest-" & Int(Rnd * 10000))
    .lPort = Val(GetSetting(App.EXEName, "Settings", "Port", 1001))
    .nServer = GetSetting(App.EXEName, "Settings", "Server", "pscchat.mine.nu")
    .sLobbyJoin = Val(GetSetting(App.EXEName, "Settings", "Ajoin", 1))
End With
End Sub


