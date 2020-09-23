VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RokSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'
'           Ragnarok Solutions Simple Server
'           Michael Bapst
'           Ragnarok Solutions
'           "Computer solutions for the common man..."
'           ragnaroksolutions@msn.com
'
' The code herein needs to be commented more clearly, but I just
' didn't have time to do it.
'
' The reason I created this was because The windsock control, while
' simplicity in design, had a few major bugs and needed to be indexed
' for mutliple connections.
'
' Winsock control Bugs: Only the last connected socket would
' be sent any information while using indexed controls (I have been
' imformed by BlackFire Interactive that this bug has been fixed
' in SP4, thanx for the Info!).
' Winsock CTL Limitations: Indexing the Control. When you need to remove a control
' from the control array, you leave huge gaps that you must either:
' 1) ReAssign the the index numbers in the array
' 2) Detect the the gaps and skip over them when needed
' Both are very time consuming, not to mention memory consuming
'
' Ragnarok Simple server addresses all of these concerns
' it has all the functionallity of the winsock control with some added extras
' It automatically adds and removes connections, all connections are stored
' in a collection for max speed, and can be enumerated with a for each...next loop
' if needed. The sockets and the server are early bound to the winsock control
' for max speed, but it also means that the server needs to recompiled if
' the winsock control's name (MSWinsockLib.Winsock) changes in future versions
' you can use late binding if you want, but I find it a serious hassle
'
' another thing you could try, if speed in you program is your main
' conecrn, is to compile this to an ActiveX EXE instead of a DLL, then
' the server would run in its own process, allowing your program to
' uses its threads for for what it needs to do. The only problem with this
' is making sure the server shuts down and unloads when your program does,
' otherwise the server will sit there and continue to accept connections
' even if your program isn't running.
'
' I was hoping to sell this server component, but microsoft turned
' around and made directplay available to VB and while more difficult
' to mess with, it seems to have rock solid performance and features
' that I cannot implement (such as port sharing, modem and serial
' connections, etc.) so I am releasing this code as FREEWARE. All I ask
' is mention in the about box, or in the version information comments
' (App.Comments)
'
' Modify as needed, mail me with questions or comments, Enjoy!
'********************************************************************************************


Private WithEvents m_WSock As Winsock 'MSWinsockLib.Winsock
Attribute m_WSock.VB_VarHelpID = -1
Public Enum SocketDataTypeConstants
    sck_Byte = vbByte
    sck_Integer = vbInteger
    sck_Long = vbLong
    sck_Single = vbSingle
    sck_Double = vbDouble
    sck_Currency = vbCurrency
    sck_Date = vbDate
    sck_Boolean = vbBoolean
    sck_SCODE = vbError
    sck_String = vbString
    sck_ByteArray = vbArray + vbByte
End Enum
Public Enum ProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
End Enum
Public Enum ErrorConstants
    sckOutOfMemory = 7
    sckInvalidPropertyValue = 380
    sckGetNotSupported = 394
    sckSetNotSupported = 383
    sckBadState = 40006
    sckInvalidArg = 40014
    sckSuccess = 40017
    sckUnsupported = 40018
    sckInvalidOp = 40020
    sckOutOfRange = 40021
    sckWrongProtocol = 40026
    sckOpCanceled = 1004
    sckInvalidArgument = 10014
    sckWouldBlock = 10035
    sckInProgress = 10036
    sckAlreadyComplete = 10037
    sckNotSocket = 10038
    sckMsgTooBig = 10040
    sckPortNotSupported = 10043
    sckAddressInUse = 10048
    sckAddressNotAvailable = 10049
    sckNetworkSubsystemFailed = 10050
    sckNetworkUnreachable = 10051
    sckNetReset = 10052
    sckConnectAborted = 11053
    sckConnectionReset = 10054
    sckNoBufferSpace = 10055
    sckAlreadyConnected = 10056
    sckNotConnected = 10057
    sckSocketShutdown = 10058
    sckTimedout = 10060
    sckConnectionRefused = 10061
    sckNotInitialized = 10093
    sckHostNotFound = 11001
    sckHostNotFoundTryAgain = 11002
    sckNonRecoverableError = 11003
    sckNoData = 11004
End Enum

Public Enum StateConstants
    sckClosed = 0
    sckOpen = 1
    sckListening = 2
    sckConnectionPending = 3
    sckResolvingHost = 4
    sckHostResolved = 5
    sckConnecting = 6
    sckConnected = 7
    sckClosing = 8
    sckError = 9
End Enum

'local variable(s) to hold property value(s)
Private m_InCol As Boolean
Private m_Key As String

'Events
'RaiseEvent SocketClose[(arg1, arg2, ... , argn)]
Public Event SocketClose()
Attribute SocketClose.VB_Description = "Fires when the remote computer and the socket closes"
Public Event Connect()
Public Event ConnectionRequest(ByVal RequestId As Long)
Attribute ConnectionRequest.VB_Description = "Fires when a remote computer requests a connection with the socket"
Public Event DataArrival(ByVal bytesTotal As Long)
Attribute DataArrival.VB_Description = "Fires when the socket receives data from a remote computer"
Public Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, CancelDisplay As Boolean)
Attribute Error.VB_Description = "Fires when a socket encounters an error"
Public Event SendComplete()
Attribute SendComplete.VB_Description = "Fires when a socket finishes sending data to a renmote computer"
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Attribute SendProgress.VB_Description = "Fires while a socket is sending data to a remote computer"
'Private Sending As Boolean







Friend Property Let Incollection(NewValue As Boolean)
m_InCol = NewValue
End Property

Friend Property Get Incollection() As Boolean
 Incollection = m_InCol
 
End Property
Public Property Get Key() As String
Attribute Key.VB_Description = "Return or sets the key associated with the socket"
Key = m_Key
End Property

Public Property Let Key(NewValue As String)
m_Key = NewValue
End Property
Friend Property Set ParentCollection(NewValue As RokSocket)
    Set m_SocketCol = NewValue
End Property

Public Sub SendData(Data As Variant)
Attribute SendData.VB_Description = "Sends data to the remote computer"
On Error GoTo ErrHandle
If m_WSock.State = sckConnected Or m_WSock.State = sckOpen Then
    m_WSock.SendData Data
    If m_InCol = True Then
        m_SocketCol.Sending = True
    End If
End If
Exit Sub
ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub

Public Sub PeekData(Data As Variant, Optional ByVal DataType As Long = sck_String, Optional ByVal MaxLen As Long = 0)
Attribute PeekData.VB_Description = "Gets data from the sockets data buffer without clearing the buffer"
On Error GoTo ErrHandle

If MaxLen = 0 Then
    m_WSock.PeekData Data, DataType
Else
    m_WSock.PeekData Data, DataType, MaxLen
End If

Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub

Public Sub Listen()
Attribute Listen.VB_Description = "Listens to a port for connection requests from a remote computer"
On Error GoTo ErrHandle

m_WSock.Listen

Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub

Public Sub GetData(Data As Variant, Optional ByVal DataType As SocketDataTypeConstants = sck_String, Optional ByVal MaxLen As Long = 0)
Attribute GetData.VB_Description = "Gets data from the sockets data buffer and clears the data buffer"
On Error GoTo ErrHandle

If MaxLen = 0 Then
    m_WSock.GetData Data, DataType
Else
    m_WSock.GetData Data, DataType, MaxLen
End If

Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub

Public Sub Connect(Optional ByVal RemoteHost As String = "", Optional ByVal RemotePort As Long = 0)
Attribute Connect.VB_Description = "Requesta connection from a remote computer"
On Error GoTo ErrHandle

If RemoteHost = "" And RemotePort = 0 Then
    m_WSock.Connect
    Exit Sub
End If

If RemoteHost <> "" And RemotePort = 0 Then
    m_WSock.Connect RemoteHost
    Exit Sub
End If

If RemoteHost = "" And RemotePort <> 0 Then
    m_WSock.Connect , RemotePort
    Exit Sub
End If

If RemoteHost <> "" And RemotePort <> 0 Then
    m_WSock.Connect RemoteHost, RemotePort
    Exit Sub
End If

Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub
Public Sub CloseSocket()
Attribute CloseSocket.VB_Description = "closes the connection between the socket and the remote computer"
On Error GoTo ErrHandle

m_WSock.Close

Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub


Public Sub Bind(Optional ByVal LocalPort As Long = 0, Optional ByVal LocalIP As String = "")
Attribute Bind.VB_Description = "Binds a socket to a port for use with the UDP protocol"
On Error GoTo ErrHandle

If LocalPort = 0 And LocalIP = "" Then
    m_WSock.Bind
    Exit Sub
End If

If LocalPort <> 0 And LocalIP = "" Then
    m_WSock.Bind LocalPort
    Exit Sub
End If

If LocalPort = 0 And LocalIP <> "" Then
    m_WSock.Bind , LocalIP
    Exit Sub
End If

If LocalPort <> 0 And LocalIP <> "" Then
    m_WSock.Bind LocalPort, LocalIP
    Exit Sub
End If


Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub

Public Sub Accept(ByVal RequestId As Long)
Attribute Accept.VB_Description = "Passing the RequestID, will open the connection between the socket and the remote computer"
On Error GoTo ErrHandle
If m_WSock.State <> sckClosed Then
    m_WSock.Close
End If

m_WSock.Accept RequestId

Exit Sub

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next
End Sub



Public Property Get State() As Integer
Attribute State.VB_Description = "Returns the current state of the socket"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
'    State = mvarState
On Error GoTo ErrHandle

State = m_WSock.State

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property





Public Property Get SocketHandle() As Long
Attribute SocketHandle.VB_Description = "Returns the socket handle, for use with windows API calls"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SocketHandle
'    SocketHandle = mvarSocketHandle
On Error GoTo ErrHandle

SocketHandle = m_WSock.SocketHandle

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property



Public Property Let RemotePort(ByVal vData As Long)
Attribute RemotePort.VB_Description = "Returns or sets the remote computer port number"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemotePort = 5
'    mvarRemotePort = vData
On Error GoTo ErrHandle

m_WSock.RemotePort = vData

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property


Public Property Get RemotePort() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemotePort
'    RemotePort = mvarRemotePort
On Error GoTo ErrHandle

RemotePort = m_WSock.RemotePort

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property





Public Property Get RemoteHostIP() As String
Attribute RemoteHostIP.VB_Description = "Returns the IP Address of the remote computer"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteHostIP
'    RemoteHostIP = mvarRemoteHostIP
On Error GoTo ErrHandle

RemoteHostIP = m_WSock.RemoteHostIP

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property



Public Property Let RemoteHost(ByVal vData As String)
Attribute RemoteHost.VB_Description = "Return or sets the Remote host name or address"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoteHost = 5
'    mvarRemoteHost = vData
On Error GoTo ErrHandle

m_WSock.RemoteHost = vData

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property


Public Property Get RemoteHost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteHost
'    RemoteHost = mvarRemoteHost
On Error GoTo ErrHandle

RemoteHost = m_WSock.RemoteHost

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property



Public Property Let Protocol(ByVal vData As ProtocolConstants)
Attribute Protocol.VB_Description = "Returns or sets the protocol the socket will use"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Protocol = 5
'    mvarProtocol = vData
On Error GoTo ErrHandle

m_WSock.Protocol = sckTCPProtocol

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property


Public Property Get Protocol() As ProtocolConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Protocol
'    Protocol = mvarProtocol
On Error GoTo ErrHandle

Protocol = m_WSock.Protocol

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property



Public Property Let LocalPort(ByVal vData As Long)
Attribute LocalPort.VB_Description = "Returns or sets the local port"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalPort = 5
'    mvarLocalPort = vData
On Error GoTo ErrHandle

m_WSock.LocalPort = vData

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property


Public Property Get LocalPort() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalPort
'    LocalPort = mvarLocalPort
On Error GoTo ErrHandle

LocalPort = m_WSock.LocalPort

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property




Public Property Get LocalIP() As String
Attribute LocalIP.VB_Description = "Returns the Local Host IP Address"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalIP
'    LocalIP = mvarLocalIP
On Error GoTo ErrHandle

LocalIP = m_WSock.LocalIP

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property





Public Property Get LocalHostName() As String
Attribute LocalHostName.VB_Description = "Returns the Local host name"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalHostName
'    LocalHostName = mvarLocalHostName
On Error GoTo ErrHandle

LocalHostName = m_WSock.LocalHostName

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property




Public Property Get BytesReceived() As Long
Attribute BytesReceived.VB_Description = "Returns the number of bytes recieved by the socket"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BytesReceived
'    BytesReceived = mvarBytesReceived
On Error GoTo ErrHandle

BytesReceived = m_WSock.BytesReceived

Exit Property

ErrHandle:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description, , App.Title
Resume Next

End Property






Private Sub Class_Initialize()
Set m_WSock = New Winsock 'MSWinsockLib.Winsock
m_InCol = False
'Sending = False
End Sub

Private Sub Class_Terminate()
Set m_WSock = Nothing

End Sub


Private Sub m_WSock_Close()
RaiseEvent SocketClose
If m_InCol Then
    m_SocketCol.ColSocketClosed m_Key
End If

End Sub

Private Sub m_WSock_Connect()
RaiseEvent Connect
If m_InCol Then
    m_SocketCol.ColSocketConnect m_Key
End If

End Sub


Private Sub m_WSock_ConnectionRequest(ByVal RequestId As Long)
RaiseEvent ConnectionRequest(RequestId)
If m_InCol Then
    m_SocketCol.ColSocketCnctRequest RequestId, m_Key
    
End If

End Sub


Private Sub m_WSock_DataArrival(ByVal bytesTotal As Long)
RaiseEvent DataArrival(bytesTotal)
If m_InCol Then
    m_SocketCol.SocketDataArrival bytesTotal, m_Key
End If

End Sub


Private Sub m_WSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, Scode, Source, CancelDisplay)
If m_InCol Then
    m_SocketCol.ColSocketError Number, Description, Scode, Source, m_Key
End If

End Sub


Private Sub m_WSock_SendComplete()
RaiseEvent SendComplete
If m_InCol Then
    m_SocketCol.ColSocketSendComplete m_Key
    If m_SocketCol.Sending = True Then
        m_SocketCol.Sending = False
    End If
End If
Sending = False
End Sub


Private Sub m_WSock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
RaiseEvent SendProgress(bytesSent, bytesRemaining)
If m_InCol Then
    If m_SocketCol.DoProgressEvents = True Then
        m_SocketCol.ColSocketSendProgress bytesSent, bytesRemaining, m_Key
    End If
End If
'Do While Sending
'    DoEvents
'Loop

End Sub


