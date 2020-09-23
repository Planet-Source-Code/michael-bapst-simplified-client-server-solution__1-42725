Ragnarok Solutions Winsock Replacement
Michael Bapst
ragnaroksolutions@msn.com

Use:

'Declare the server obj 

Option Explicit

Private WithEvents RokServer as RokSockets


Private Sub Form_Load()
Set RokServer = New RokSockets
With RokServer
    'I never had this the original winsock control's
    'SendProgress event actually fire, but I put it
    'in here in case they get it to work. But sending
    'but firing progress events (and running whatever code you
    'add in the event) when there are alot of connections
    'is going to slow everything down. So we disable sending
    'progress events

    .DoProgressEvents = False
    
    'Now create the server, using TCP/IP and listen to port 999

    .CreateServer 999, sckTCPProtocol
    
    'the server is now up and listening
End With

End Sub

Private Sub Form_Unload(Cancel as integer)

'first close the server and disconnect all sockets

RokServer.CloseServer

'destroy the object
Set RokServer = Nothing

end sub


Private Sub RokServer_ServerClose()

'the server closed, so do any clean here

Debug.Print "Server Close"

End Sub


Private Sub RokServer_ServerConnectionRequest(ByVal RequestId As Long, ByVal RemoteIP As String)

'someone wants to connect, if we want, we could
'check the ip address against an banned list, and
'accept or deny accordingly

Debug.Print RemoteIP
Debug.Print RokServer.ServerAcceptRequest(RequestId)

End Sub



Private Sub RokServer_ServerSocketClosed(Key As String)

'A client discconected

Debug.Print "Socket Closed: " & CStr(Key)

End Sub


Private Sub RokServer_ServerSocketConnect(Socket As RokIPSocket.RokSocket)

'a client connected, we can alter the connection here
'ex: change the key froma guid to a client name or ip address

Debug.Print "Socket Connect"

End Sub



Private Sub RokServer_ServerSocketDataArrival(ByVal bytesTotal As Long, Socket As RokIPSocket.RokSocket)

'this is the heart of the server
'in this example the server simply rebroadcasts
'any data it recieves to all connections 
'this is your basic 'chat' program

Dim TmpStr As String

Socket.GetData TmpStr
Debug.Print RokServer.BroadcastAll(TmpStr)

End Sub



Private Sub RokServer_ServerError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)

'if the server encounters an error, this event will fire
'if possible

End Sub


