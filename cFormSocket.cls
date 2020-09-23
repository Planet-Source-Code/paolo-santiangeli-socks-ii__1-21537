VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Const TIME_OUT = 60
Private Const RCV_BUFFER = 1024

Public Enum sckWaitedDataType
        sckBinary = 0
        sckASCII = 1
End Enum

Public Enum sckState
            
            sckDisconnected = 0
            sckConnecting = 1
            sckConnected = 2
            sckWaitingAnswer = 3
            sckWaitingBinary = 8
            sckCommandAnswered = 4
            sckCommandNotAnswered = 5
            sckDataReceived = 6
            sckListening = 7
            sckError = -1

End Enum
'************* EVENTS **************************************
Event sckData(pCMD As String, pDATA As String, hWnd As Long)
Event sckBinaryData(pDATA As Variant)
Event sckConnected()
Event sckConnectionAccepted()
Event sckDisconnectedByPeer()
Event sckLog(pLogEntry As String)
Event sckError(pErrorDescription As String)
'************* VARS ****************
Private mEvents As frmTesting
Private mSocket As Long
Private mhWnd As Long
Private mlpPrevWndProc As Long

'***************************************
Private mState As sckState
Private mWaitedDataType As sckWaitedDataType
Private mAnswer As String
Private mTotaleBytes As Long


Private mWaitForData As Boolean

Private mDataBuffer As Variant
Private mAnswerBuffer As String
Private mAnswerTimeOut As Long

Private mWaitedData As String

Private mCommandOK As Boolean

Private DataInBuffer As Boolean
Property Get LastCommandOK() As Boolean

    LastCommandOK = mCommandOK

End Property
Property Get LastAnswer() As String

    LastAnswer = mAnswer

End Property
Property Get State() As sckState

    State = mState

End Property
Property Let WaitedDataType(pDT As sckWaitedDataType)

    mWaitedDataType = pDT

End Property
Public Sub sckListen(pPort As Long)
Dim a As Long
    Open App.Path & "\tmp.tmp" For Binary As 1
    a = ListenForConnect(pPort, mhWnd)
    mState = sckListening
End Sub
Public Sub sckConnect(m_host As String, m_port As String, Optional mMode As String)
    
    Dim temp As Variant
    
    mState = sckConnecting
    
    'make sure the port is closed!
    If mSocket <> 0 Then Call closesocket(mSocket)
    
    temp = ConnectSock(m_host, m_port, 0, mhWnd, True)
    
    RaiseEvent sckLog("Connect socket return value :" & temp)
    RaiseEvent sckLog("Connecting to :" & m_host & " at port " & m_port)
        
    'Winsock Error
    If temp < 0 Then
        LogWSAError WSAGetLastError(), "sckConnect"
        mState = sckError
        Exit Sub
    End If
    
    'Wait for Error or FD_CONNECT Event
    Do While mState = sckConnecting And (Not mState = sckError)
        DoEvents
    Loop
    
End Sub
Function sckSendCommand(pServerCMD As String, pWaitFor As String)
'Send data to the socket and wait for data in Remote Host Answer
    
    
    'Wait for data flag
    'mWaitForData = True
    mCommandOK = False
    'mWaitedDataType = sckCommandAnswer
    mWaitedData = pWaitFor
    
    'Send Data
    Call SendData(mSocket, pServerCMD & vbCrLf)
    RaiseEvent sckLog("Command Sent: " & pServerCMD)
    
    mState = sckWaitingAnswer
    
    RaiseEvent sckLog("Waiting for data '" & pWaitFor & "'")

    mAnswerTimeOut = Timer + 360
    
    Do While mState = sckWaitingAnswer And (Not mState = sckError)
        DoEvents
        
        If Timer > mAnswerTimeOut Then
            Exit Do
        End If
        
    Loop
    
    If mCommandOK Then
        
        RaiseEvent sckLog("Waited data '" & pWaitFor & "': OK")
    
    Else
        
        RaiseEvent sckLog("Waited data '" & pWaitFor & "'): ERROR!")
    
    End If
    
    sckSendCommand = ""

End Function
Function sckSendData(pServerCMD As String) As String
'Send DATA without waiting Answer
    
    RaiseEvent sckLog("Command Sent: " & pServerCMD)
    
    Call SendData(mSocket, pServerCMD & vbCrLf)
    
    sckSendData = ""
    
End Function
Public Sub sckHook(ByVal pHwnd As Long)
    
    
     mhWnd = pHwnd
    
    mlpPrevWndProc = SetWindowLong(mhWnd, GWL_WNDPROC, AddressOf SubWndProc)
    
    Call SetWindowLong(mhWnd, GWL_USERDATA, ObjPtr(Me))
    
End Sub
Public Sub sckUnHook()
    'if we dont un-subclass before we shutdown
    'the program, we get an illigal procedure error.
    'fun.
    
    Call SetWindowLong(mhWnd, GWL_WNDPROC, mlpPrevWndProc)
    

    
End Sub
Sub sckClose()

    Call closesocket(mSocket)
    Close
End Sub

Private Sub DataArrival(ByVal pDATA As Variant, pDataLenght As Long)
Dim pAnswerBuffer As String
Dim S As String
    Select Case mState
    
        Case sckWaitingAnswer
            
            pAnswerBuffer = BytesToString(pDATA, pDataLenght - 1)
            
            If mWaitedData = Left(pAnswerBuffer, 3) And (Mid(pAnswerBuffer, 4, 1) = " ") Then
          
                mAnswerBuffer = ""
                mCommandOK = True
                mAnswer = pAnswerBuffer
                mState = sckCommandAnswered
                        
            End If
         
         Case sckWaitingBinary
            
            mTotaleBytes = mTotaleBytes + pDataLenght
            S = Left(StrConv(pDATA, vbUnicode), pDataLenght)
            
            RaiseEvent sckLog("Bytes to write :" & CStr(Len(S)))
            
            Put #1, , S
         
         Case Else
            
    End Select
    
    RaiseEvent sckLog("Bytes:" & CStr(pDataLenght))
    RaiseEvent sckBinaryData(StrConv(pDATA, vbUnicode))


End Sub

Private Sub DataArrivalOLD(ByVal pDATA As Variant, pDataLenght As Long)
Dim pDataStr As String
Dim tmpAnswer As String
Dim tmpStartLine As Long
Dim tmpEndLine As Long
Dim tmpLines As Variant
Dim pDataBuffer As Variant
Dim pTimeOUT As Long
       
    RaiseEvent sckLog("Bytes Received=" & CStr(pDataLenght))
        
    mDataBuffer = ""
        
    'Copy bytes to pbuffer
    pDataBuffer = CopyBytes(pDATA, pDataLenght)
    
    'Action waiting for response
    
    If mWaitForData = True Then
        
            
            Select Case mWaitedDataType
               
               Case sckBinary
                
               
               
               Case sckASCII
                
               
               Case sckCommandAnswer
                    
                    pDataStr = BytesToString(pDataBuffer, pDataLenght - 1)

                    tmpAnswer = Left$(pDataStr, 3)
                                          
                        If mWaitedData = tmpAnswer Then
                
                                mCommandOK = True
                                
                                mAnswer = tmpAnswer
                                RaiseEvent sckLog("Answer:" & tmpAnswer)
                
                        Else
                        
                                
                        
                        End If
                
                Case Else: mAnswer = StrConv(mbuffer, vbUnicode)

              
              End Select
            
            mWaitForData = False
    
        End If
    
    
    mState = sckDataReceived
    RaiseEvent sckBinaryData(mDataBuffer)

End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Sock As sockaddr
Dim newSock As Long
 
Dim x As Long
Dim temp As Variant

Dim ReadBuffer(RCV_BUFFER) As Byte
Dim recvBuffer As String

Dim WSAEvent As Long
Dim WSAError As Long


    If uMsg <> 1025 Then
        WindowProc = CallWindowProc(mlpPrevWndProc, hw, uMsg, wParam, ByVal lParam)
        Exit Function
    End If

    WSAEvent = WSAGetSelectEvent(lParam)
    WSAError = WSAGetAsyncError(lParam)
    
    If WSAError <> 0 Then
        RaiseEvent sckLog("WINPROC Error: " & GetWSAErrorString(WSAError))
        mState = sckError
        Exit Function
    End If
    
    Select Case WSAEvent
            
                Case FD_READ: 'Lets check for data
                    
                    x = recv(mSocket, ReadBuffer(0), RCV_BUFFER, 0)  'try to get some
                    
                    If x > 0 Then
                       DataArrival ReadBuffer, x
                    End If
                
                Case FD_CONNECT: 'Did we connect?
                      
                      mState = sckConnected
                      
                      RaiseEvent sckConnected
                      RaiseEvent sckLog("Connection Established..." & CStr(wParam))
                      
                      mSocket = wParam 'yep, we did! yayay
                
                Case FD_ACCEPT:
                    
                    newSock = accept(wParam, Sock, sockaddr_size)
                    
                    mState = sckWaitingBinary
                    
                    RaiseEvent sckConnectionAccepted
                    RaiseEvent sckLog("Connection Accepted... :" & lParam)
                    RaiseEvent sckLog("Remote Host:" & GetPeerAddress(newSock))
                    
                    mSocket = newSock

                Case FD_OOB:
                                                
                        
                Case FD_CLOSE: 'uh oh. they closed the connection
                    
                    Do
                        
                        x = recv(mSocket, ReadBuffer(0), RCV_BUFFER, 0)  'try to get some
                        
                        If x > 0 Then
                           DataArrival ReadBuffer, x
                        End If
                    
                    Loop While x <> 0
                    
                    mState = sckDisconnected
                    RaiseEvent sckLog(Str(mTotaleBytes))
                    
                    'RaiseEvent sckDisconnectedByPeer
                    RaiseEvent sckLog("Connection Closed By Peer")
                    
                    Call closesocket(mSocket)   'so we need to close
            
            End Select

End Function
Private Function ListenForConnect(ByVal Port As Long, ByVal pHwnd As Long) As Long
    
    Dim S As Long, dummy As Long
    Dim SelectOps As Long
    Dim saZero As sockaddr
    Dim sockin As sockaddr
    Dim errors As Long

    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    
    If sockin.sin_port = INVALID_SOCKET Then
        errors = WSAGetLastError()
        RaiseEvent sckLog("ERROR in sockin.sin_port -->" & GetWSAErrorString(errors))
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        errors = WSAGetLastError()
        RaiseEvent sckLog("ERROR in htonl" & GetWSAErrorString(errors))
        ListenForConnect = errors
        Exit Function
    End If
    'Create a socket TCP
    S = Socket(PF_INET, SOCK_STREAM, 0)
    If S < 0 Then
        errors = WSAGetLastError()
        gLog.Log "ERROR in Socket() --> " & GetWSAErrorString(errors)
        ListenForConnect = errors
        Exit Function
    End If
    'Bound the (unbound) socket to the address
    If bind(S, sockin, sockaddr_size) Then
        errors = WSAGetLastError()
        RaiseEvent sckLog("ERROR in bind()" & " --> " & GetWSAErrorString(errors))
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = errors
        Exit Function
    End If
    
    'Events to generate
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    'Set the message notification for the socket
    If WSAAsyncSelect(S, pHwnd, ByVal 1025, ByVal SelectOps) Then
        errors = WSAGetLastError()
        RaiseEvent sckLog("ERROR in WSAAsnycSelect --> " & GetWSAErrorString(errors))
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = WSAGetLastError()
        Exit Function
    End If
    
    'Listen the connection requests
    If listen(S, 1) Then
        errors = WSAGetLastError()
        RaiseEvent sckLog("ERROR in listen() --> " & GetWSAErrorString(errors))
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = errors
        Exit Function
    End If
    
    RaiseEvent sckLog("Listening...")

End Function
Sub sckStart()
Set mEvents = New frmTesting
mhWnd = mEvents.hWnd
mlpPrevWndProc = SetWindowLong(mhWnd, GWL_WNDPROC, AddressOf SubWndProc)
Call SetWindowLong(mhWnd, GWL_USERDATA, ObjPtr(Me))
mEvents.Show
End Sub
Sub sckStop()
Call closesocket(mSocket)
Call SetWindowLong(mhWnd, GWL_WNDPROC, mlpPrevWndProc)
Unload mEvents
Set mEvents = Nothing
End Sub
Private Function BytesToString(pDATA As Variant, BytesTotal As Long) As String
Dim tmpSTr As String
Dim i As Long
    
For i = 0 To BytesTotal
        
        tmpSTr = tmpSTr & Chr(pDATA(i))

Next

BytesToString = tmpSTr

End Function
Function sckGetBuffer() As Variant

    sckGetBuffer = mDataBuffer

End Function
Private Sub LogWSAError(pError As Long, Optional pProcedure As String = "NA")
Dim e As Long
     
     RaiseEvent sckLog("WSA Error: " & GetWSAErrorString(pError) & " (" & pProcedure & ")")

End Sub
Private Function CopyBytes(pDATA As Variant, BytesTotal As Long) As Variant
Dim i As Long
Dim pBuffer As Variant

For i = 0 To BytesTotal
    
    pBuffer = pBuffer & Chr(pDATA(i))

Next
    
    CopyBytes = pBuffer
                
                
End Function

