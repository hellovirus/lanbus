Attribute VB_Name = "Module1"
'******************定义以下全局变量，方便各模块调用
Public GBIP As String  '广播IP地址
Public MyNickN As String  '我的昵称
Public MyIP As String  '我的IP地址
Public MyFace As Integer  '我的头像号
'Public JianJie As String  '我的简介
Public MyInfo As String  '我的详细情况，包括"昵称|IP|头像号"

Private Const SOCKET_ERROR As Long = -1
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&

Private Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Private Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)


Function GetIPAddress(Optional sHost As String) As String
'返回给定机器名的Ip地址，机器名为空时返回本机Ip地址
    Dim sHostName   As String * 256
    Dim lpHost      As Long
    Dim HOST        As HOSTENT
    Dim dwIPAddr    As Long
    Dim tmpIPAddr() As Byte
    Dim i           As Integer
    Dim sIPAddr     As String
    Dim werr        As Long

    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    If sHost = "" Then
        If gethostname(sHostName, 256) = SOCKET_ERROR Then
            werr = WSAGetLastError()
            GetIPAddress = ""
            SocketsCleanup
            Exit Function
        End If

        sHostName = Trim$(sHostName)
        'TxtCmpName.Text = sHostName '如果机器名为空，则查询本机用户名并显示！
    Else
        sHostName = Trim$(sHost) & Chr$(0)
    End If
    
    lpHost = gethostbyname(sHostName)

    If lpHost = 0 Then
        werr = WSAGetLastError()
        GetIPAddress = ""
        SocketsCleanup
        Exit Function
    End If

    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4

    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen

    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next

    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function
Private Function SocketsInitialize(Optional sErr As String) As Boolean
    Dim WSAD As WSADATA, sLoByte As String, sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        sErr = "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If

    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        sErr = "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."

        SocketsInitialize = False
        Exit Function
    End If


    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
            (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
            HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then

        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))

        sErr = "Sockets version " & sLoByte & "." & sHiByte & _
                " is not supported by 32-bit Windows Sockets."

        SocketsInitialize = False
        Exit Function
    End If
    SocketsInitialize = True
End Function

Private Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        App.LogEvent "Socket error occurred in Cleanup.", vbLogEventTypeError
    End If
End Sub

Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H1 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Function FileExists(FileName As String) As Boolean
    On Error Resume Next
    FileExists = Dir$(FileName) <> ""
    If Err.Number <> 0 Then
        FileExists = False
    End If
    On Error GoTo 0
End Function
