VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RokSockets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RokSocket"
Attribute VB_Ext_KEY = "Member0" ,"RokSocket"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Private WithEvents Server As Winsock 'MSWinsockLib.Winsock
Attribute Server.VB_VarHelpID = -1
Public DoProgressEvents As Boolean
Attribute DoProgressEvents.VB_VarDescription = "When true, each socket will fire a progress event, default is False (If you have many sockets, firing a progress event for each of them will slow the server down)"

Public Event ServerConnectionRequest(ByVal RequestId As Long, ByVal RemoteIP As String)
Attribute ServerConnectionRequest.VB_Description = "Fires when a client attempts to connect to the server"
Public Event SocketDataArrival(ByVal bytesTotal As Long, Socket As RokSocket)
Public Event SocketClosed(Key As String)
Attribute SocketClosed.VB_Description = "Fires when a socket in the collection closes a connection fom remote computer"
Public Event SocketError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, Socket As RokSocket)
Attribute SocketError.VB_Description = "Fires when an error occurs with in a socket"
Public Event SocketConnectionRequest(RequestId As Long, Socket As RokSocket)
Attribute SocketConnectionRequest.VB_Description = "Fires when a socket in the collection receives a connection request from a remote computer"
Public Event SocketConnect(Socket As RokSocket)
Attribute SocketConnect.VB_Description = "Fires when a socket in the collection connects to a remote computer"
Public Event SocketSendComplete(Socket As RokSocket)
Attribute SocketSendComplete.VB_Description = "Fires when a socket in the collection finishes sending data"
Public Event SocketSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long, ByVal Key As String)
Attribute SocketSendProgress.VB_Description = "fires while a socket in the collection is sending data to a remote computer"
Public Event ServerClose()
Attribute ServerClose.VB_Description = "Fires when a server closes"
Public Event ServerError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
Attribute ServerError.VB_Description = "Fires when a server encounters an error"

'local variable(s) to hold property value(s)
Private m_Sending As Boolean
Private m_KillSocket As Boolean
Private mvarSuspendServer As Boolean 'local copy
Public Property Let SuspendServer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SuspendServer = 5
    mvarSuspendServer = vData
End Property


Public Property Get SuspendServer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SuspendServer
    SuspendServer = mvarSuspendServer
End Property



Public Function BroadcastAll(Data As Variant, Optional SkipSocketKey As String = "") As Boolean
Attribute BroadcastAll.VB_Description = "Broadcasts data to every connected socket, optionally skipping a particular socket (most likely the socket who received the data)"
Dim TmpSck As RokSocket
On Error GoTo ErrHandle

If mvarSuspendServer Then
    Err.Raise vbObjectError + 1, "RokSockets.BroadcastAll", "The server is currently suspended."
    Exit Function
End If

BroadcastAll = True

For Each TmpSck In mCol
    If TmpSck.Key <> SkipSocketKey Then
        TmpSck.SendData Data
        Do While Sending
            DoEvents
        Loop
    End If
Next
Exit Function

ErrHandle:
BroadcastAll = False

End Function

Public Sub CloseConnections()
Attribute CloseConnections.VB_Description = "Closes all the connections in the collection"
Dim TmpSocket As RokSocket
For Each TmpSocket In mCol
    TmpSocket.CloseSocket
    mCol.Remove TmpSocket.Key
Next
Set TmpSocket = Nothing

End Sub


Public Function CloseServer() As Boolean
Attribute CloseServer.VB_Description = "Closes all the connections in the collection and then closes the server"

CloseConnections
Server.Close

End Function

Friend Sub ColSocketCnctRequest(RequestId As Long, Key As String)
RaiseEvent SocketConnectionRequest(RequestId, Item(Key))

End Sub

Friend Sub ColSocketConnect(Key As String)
    RaiseEvent SocketConnect(Item(Key))
    
End Sub


Friend Sub ColSocketError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal Key As String)
RaiseEvent SocketError(Number, Description, Scode, Source, Item(Key))
End Sub

Friend Sub ColSocketClosed(Key As String)
Me(Key).CloseSocket
If m_KillSocket Then
    Remove Key
End If

RaiseEvent SocketClosed(Key)
End Sub

Friend Sub ColSocketSendComplete(Key As String)
RaiseEvent SocketSendComplete(Item(Key))

End Sub

Friend Sub ColSocketSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long, ByVal Key As String)
RaiseEvent SocketSendProgress(bytesSent, bytesRemaining, Key)

End Sub

Public Property Get Connections() As Long
Attribute Connections.VB_Description = "Returns the number of connections currently connected to the server"
Connections = mCol.Count
End Property

Public Function CreateServer(ByVal ListenPort As Long, Optional ByVal Protocol As ProtocolConstants = sckTCPProtocol, Optional ByVal LocalIP As String = "", Optional ByVal RemotePort As Long = 0, Optional ByVal RemoteIP As String = "") As Boolean
Attribute CreateServer.VB_Description = "Creates a server and starts it listening to a port if using TCP, or binds it to a Port if using UDP"
On Error GoTo ErrHandle

CreateServer = True

If Not (Server Is Nothing) Then
    CloseServer
    Set Server = Nothing
End If

Set Server = New Winsock
With Server
    Select Case Protocol
        Case sckUDPProtocol
            If RemotePort = 0 Or RemoteIP = "" Then
                CreateServer = False
                Exit Function
            End If
            If .State <> sckClosed Then
                .Close
            End If
            .RemoteHost = RemoteIP
            .RemotePort = RemotePort
            If LocalIP <> "" Then
                .Bind ListenPort, LocalIP
            Else
                .Bind ListenPort
            End If
        Case sckTCPProtocol
            If .State <> sckClosed Then
                .Close
            End If
            .LocalPort = ListenPort
            .Listen
            
    End Select
End With
Exit Function

ErrHandle:

'Resume Next
CreateServer = False

End Function

Public Function Add(sKey As String) As RokSocket
Attribute Add.VB_Description = "Adds a socket to the collection"
On Error GoTo ErrHandle

    'create a new object
    Dim objNewMember As RokSocket
    Set objNewMember = New RokSocket

    objNewMember.Key = sKey
    objNewMember.Incollection = True
    'Set objNewMember.ParentCollection = Me
    Set m_SocketCol = Me
    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
Exit Function

ErrHandle:
Debug.Print CStr(Err.Number) & Err.Description & vbCrLf & Err.Source


End Function

Public Function DebugServer() As String
Attribute DebugServer.VB_Description = "Returns a string containing debug information about the server."
Attribute DebugServer.VB_MemberFlags = "40"
Dim TmpStr As String
Dim i As Long

With Server
    TmpStr = "Server Local Host Name: " & .LocalHostName & vbCrLf
    TmpStr = TmpStr & "Server Local IP: " & .LocalIP & vbCrLf
    TmpStr = TmpStr & "Server Local Port: " & CStr(.LocalPort) & vbCrLf
    TmpStr = TmpStr & "Server Remote Host Name: " & .RemoteHost & vbCrLf
    TmpStr = TmpStr & "Server Remote IP: " & .RemoteHostIP & vbCrLf
    TmpStr = TmpStr & "Server Remote Port: " & CStr(.RemotePort) & vbCrLf
    TmpStr = TmpStr & "Server Protocol: " & IIf(CBool(.Protocol), "UDP", "TCP") & vbCrLf
    TmpStr = TmpStr & "Server State: " & CStr(.State) & vbCrLf
    TmpStr = TmpStr & "Server Bytes Received: " & CStr(.BytesReceived) & vbCrLf
    TmpStr = TmpStr & "Server Connections: " & CStr(mCol.Count) & vbCrLf
    
    For i = 1 To mCol.Count
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Local Host Name: " & .LocalHostName & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Local IP: " & .LocalIP & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Local Port: " & CStr(.LocalPort) & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Remote Host Name: " & .RemoteHost & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Remote IP: " & .RemoteHostIP & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Remote Port: " & CStr(.RemotePort) & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Protocol: " & IIf(CBool(.Protocol), "UDP", "TCP") & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": State: " & CStr(.State) & vbCrLf
        TmpStr = TmpStr & vbTab & "Connection " & CStr(i) & ": Bytes Received: " & CStr(.BytesReceived) & vbCrLf
        TmpStr = TmpStr & vbCrLf
    Next i
End With
DebugServer = TmpStr
End Function

Public Property Let KillSocketOnClose(NewValue As Boolean)
Attribute KillSocketOnClose.VB_Description = "If True, Sockets are removed automatically from the collection when  the connection is closed from either client or server."
m_KillSocket = NewValue
End Property

Public Property Get KillSocketOnClose() As Boolean
KillSocketOnClose = m_KillSocket

End Property

Friend Property Let Sending(NewValue As Boolean)
m_Sending = NewValue
End Property

Public Property Get Sending() As Boolean
Attribute Sending.VB_Description = "If true, then one of the sockets is currently sending data, Used mainly for internal functions"
Sending = m_Sending

End Property

Public Property Get ServerHandle() As Long
Attribute ServerHandle.VB_Description = "Returns the ocket handle of the server, used for making windows API calls"
ServerHandle = Server.SocketHandle
End Property
Public Property Get ServerProtocol() As ProtocolConstants
Attribute ServerProtocol.VB_Description = "Return or sets the protocol that the server is using for connections, either TCP (Default) or UDP"
ServerProtocol = Server.Protocol
End Property
Public Property Get ServerBytesRecieved() As Long
Attribute ServerBytesRecieved.VB_Description = "Returns the number of bytes that the server has received"
ServerBytesRecieved = Server.BytesReceived
End Property
Public Property Get ServerLocalHostName() As String
Attribute ServerLocalHostName.VB_Description = "Local Host Name of the server"
    LocalHostName = Server.LocalHostName
End Property
Public Function ResetServer() As Boolean
Attribute ResetServer.VB_Description = "resets and restarts a previously created server, useful when the server gets an error"
On Error GoTo ErrHandle
Dim OldLPort As Long
Dim OldRPort As Long
Dim OldLIP As Long
Dim OldRIP As Long
Dim Protocol As Long

ResetServer = True

If (Server Is Nothing) Then
    ResetServer = False
    Exit Function
End If
'Get Old Server Values
OldLPort = Server.LocalPort
OldRPort = Server.RemotePort
OldLIP = Server.LocalIP
OldRIP = Server.RemoteHost
Protocol = Server.Protocol

CloseConnections
CloseServer
Set Server = Nothing
Set Server = New Winsock
With Server
    Select Case Protocol
        Case sckUDPProtocol
            If OldRPort = 0 Or OldRIP = "" Then
                ResetServer = False
                Exit Function
            End If
            If .State <> sckClosed Then
                .Close
            End If
            .RemoteHost = OldRIP
            .RemotePort = OldRPort
            If OldLIP <> "" Then
                .Bind OldLPort, OldLIP
            Else
                .Bind OldLPort
            End If
        Case sckTCPProtocol
            If .State <> sckClosed Then
                .Close
            End If
            .LocalPort = OldLPort
            .Listen
    End Select
End With
Exit Function

ErrHandle:
ResetServer = False

End Function

Public Property Get ServerLocalIP() As String
Attribute ServerLocalIP.VB_Description = "Local IP Address of the server"
ServerLocalIP = Server.LocalIP
End Property

Public Property Get ServerLocalPort() As Long
Attribute ServerLocalPort.VB_Description = "Returns or sets the port that the server is listening to"
ServerLocalPort = Server.LocalPort
End Property

Public Property Get ServerRemoteHost() As String
Attribute ServerRemoteHost.VB_Description = "Returns or sets the remote host address or name that server is connected to"
ServerRemoteHost = Server.RemoteHost
End Property
Public Property Get ServerRemoteHostIP() As String
Attribute ServerRemoteHostIP.VB_Description = "Returns the remote host address that server is connected to"
    ServerRemoteHostIP = Server.RemoteHostIP
End Property

Public Property Get ServerRemotePort() As Long
Attribute ServerRemotePort.VB_Description = "Returns the remote host port that server is connected to"
ServerRemotePort = Server.RemotePort

End Property

Public Property Get ServerState() As Integer
Attribute ServerState.VB_Description = "Returns the current state of the server"
ServerState = Server.State
End Property

Friend Sub SocketDataArrival(bytesTotal As Long, Key As String)
Attribute SocketDataArrival.VB_Description = "Fires whan a socket in the collection receives data from a remote computer"
If mvarSuspendServer Then
    Exit Sub
End If

RaiseEvent SocketDataArrival(bytesTotal, Item(Key))
End Sub


Public Property Get Item(vntIndexKey As Variant) As RokSocket
Attribute Item.VB_Description = "Returns a RokSocket Object based on index or key"
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of sockets in the socket collection (the connections proerty return the same result)"
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "revomes a socket from the collection"
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    If Me(vntIndexKey).State = sckConnected Or Me(vntIndexKey).State = sckOpen Then
        Me(vntIndexKey).CloseSocket
    End If

    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Function SocketAcceptRequest(ByVal RequestId As Long, Optional ByVal Key As String = "") As RokSocket
Attribute SocketAcceptRequest.VB_Description = "Passing the RequestID from the ServerConnectionRequest event will create a new socket, add it to the collection and have it accept the connection"
On Error GoTo ErrHandle

If mvarSuspendServer Then
    Err.Raise vbObjectError + 1, "RokSockets.SocketAcceptRequest", "The server is currently suspended."
    Exit Function
End If

Randomize Timer
ServerAcceptRequest = True

Dim TmpSocket As RokSocket

'Set TmpSocket = Add(GetGUID)
If Len(Key) = 0 Then
    Set TmpSocket = Add(GetGUID)
Else
    Set TmpSocket = Add(Key)
End If

With TmpSocket
    .LocalPort = 0
    .Accept RequestId
End With
Set SocketAcceptRequest = TmpSocket
Set TmpSocket = Nothing
Exit Function

ErrHandle:
ServerAcceptRequest = False

End Function

Public Function SocketSendData(Data As Variant, SocketKey As String) As Boolean
Attribute SocketSendData.VB_Description = "Allows you to send adat through a particular socket"
On Error GoTo ErrHandle

If mvarSuspendServer Then
    Err.Raise vbObjectError + 1, "RokSockets.SocketSendData", "The server is currently suspended."
    Exit Function
End If

SocketSendData = True

If Me(SocketKey).State = sckConnected Or Me(SocketKey).State = sckOpen Then
    Me(SocketKey).SendData Data
Else
    SocketSendData = False
End If

Exit Function

ErrHandle:
SocketSendData = False

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    DoProgressEvents = False
    m_Sending = False
    m_KillSocket = True
    mvarSuspendServer = False
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Private Sub Server_Close()
'fires when the server closes
RaiseEvent ServerClose

End Sub

Private Sub Server_Connect()
'no need to fire an event here, the server socket only only listens and then
'creates a new socket for each request. The server socket never connects

End Sub


Private Sub Server_ConnectionRequest(ByVal RequestId As Long)

If mvarSuspendServer Then
    Exit Sub
End If

RaiseEvent ServerConnectionRequest(RequestId, Server.RemoteHostIP)

End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
'no need to fire an event here, the server socket only only listens and then
'creates a new socket for each request. The server socket ignores all data
'sent to it, only acknowledges connection requests

End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent ServerError(Number, Description, Scode, Source)
CancelDisplay = True

End Sub


Private Sub Server_SendComplete()
'no need to fire an event here, the server socket only only listens and then
'creates a new socket for each request. The server socket sends data itself,
'that is delegated to the Sockets Collection

End Sub


Private Sub Server_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'no need to fire an event here, the server socket only only listens and then
'creates a new socket for each request. The server socket sends data itself,
'that is delegated to the Sockets Collection
End Sub


