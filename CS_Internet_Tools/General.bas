Attribute VB_Name = "General"
Option Explicit
Public lngNextPort As Long

Global PortDone As Integer
Global OnPort As Long

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type Inet_address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
    End Type


Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
    End Type


Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
    End Type

Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    FLAGS As Byte
    OptionsSize As Long
    OptionsData As String * 128
    End Type


Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    Data As Long
    Options As IP_OPTION_INFORMATION
    End Type
    
    Public pIPe As IP_ECHO_REPLY
    Public pIPe2 As IP_ECHO_REPLY
    Public pIPe3 As IP_ECHO_REPLY
    Public pIPo As IP_OPTION_INFORMATION
    Public pIPo2 As IP_OPTION_INFORMATION
    Public pIPo3 As IP_OPTION_INFORMATION
    Public IPLong As Inet_address
    Public IPLong2 As Inet_address
    Public IPLong3 As Inet_address
    Public IPLong4 As Inet_address
    Public IPLong5 As Inet_address
    Public IPLong6 As Inet_address
    Public IPLong7 As Inet_address
    
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, _
    ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, _
    ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal timeout As Long) As Boolean

Declare Function gethostname Lib "wsock32.dll" (ByVal hostname$, HostLen&) As Long
Declare Function gethostbyname& Lib "wsock32.dll" (ByVal hostname$)
Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Declare Function WSACleanup Lib "wsock32.dll" () As Long
Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function IcmpCreateFile Lib "ICMP.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "ICMP.dll" (ByVal HANDLE As Long) As Boolean

Function ScanPort(thePort As Long, ws1 As Winsock) As Boolean
ScanPort = False
On Error GoTo gotport
ws1.Close
ws1.LocalPort = thePort
ws1.Listen
pause 0.1
ws1.Close
Exit Function
gotport:
If Err.Number = 10048 Then
    ScanPort = True
End If
End Function


Sub pause(interval)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

