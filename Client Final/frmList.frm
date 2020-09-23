VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Channel List"
   ClientHeight    =   4080
   ClientLeft      =   2445
   ClientTop       =   1860
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstChans 
      Height          =   3960
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4095
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lstChans_DblClick()
frmMain.W.SendData "join" & ComSep & lstChans.Text
End Sub
