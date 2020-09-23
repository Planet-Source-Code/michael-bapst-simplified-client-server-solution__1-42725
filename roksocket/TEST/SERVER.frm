VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   2052
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents RokSock As RokSockets
Attribute RokSock.VB_VarHelpID = -1

Private Sub Form_Load()
Set RokSock = New RokSockets
With RokSock
    .KillSocketOnClose = True
    Debug.Print .CreateServer(1001)
End With
Label1.Caption = "Server Created..." & vbCrLf & "IP Address: " & RokSock.ServerLocalIP

End Sub


Private Sub Label1_DblClick()
RokSock.CloseConnections
RokSock.CloseServer
Set RokSock = Nothing
End
End Sub


Private Sub RokSock_ServerConnectionRequest(ByVal RequestId As Long, ByVal RemoteIP As String)

RokSock.SocketAcceptRequest RequestId

End Sub


Private Sub RokSock_SocketDataArrival(ByVal bytesTotal As Long, Socket As RokIPSocket.RokSocket)
Dim Tmp As String
Socket.GetData Tmp
MsgBox Tmp, , App.Title


End Sub


Private Sub Timer1_Timer()
RokSock.BroadcastAll Format(Now, "ttttt")
End Sub


