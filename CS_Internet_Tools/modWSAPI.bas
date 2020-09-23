Attribute VB_Name = "modWSAPI"
Option Explicit

'Win32 API declarations
' this allows us to copy a string by the strings pointer
' sometimes we can not extract a string into the VB string data type
' so we have to get the pointer to that string this Win32 API function
' allows us to do just that
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

' this function gives us the length of the string from the pointer
' we need to know this before we can copy the entire string to
' lpstring2 above
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long


' winsock API declarations

Public Const INADDR_NONE = &HFFFF
Public Const SOCKET_ERROR = -1
Public Const WSABASEERR = 10000
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEINPROGRESS = (WSABASEERR + 50)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004



Public Type servent
    s_name    As Long
    s_aliases As Long
    s_port    As Integer
    s_proto   As Long
End Type

Public Type protoent
    p_name    As String
    p_aliases As Long
    p_proto   As Long
End Type


' used to convert IP Address to long
Public Declare Function inet_addr _
    Lib "ws2_32.dll" (ByVal cp As String) As Long

Public Declare Function inet_ntoa _
    Lib "ws2_32.dll" (ByVal inn As Long) As Long
    
' these functions of the winsock API are used to convert
' byte ordering - to learn more about why this is important
' msdn library - look under byte ordering

Public Declare Function htons _
    Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

Public Declare Function htonl _
    Lib "ws2_32.dll" (ByVal hostlong As Long) As Long

Public Declare Function ntohl _
    Lib "ws2_32.dll" (ByVal netlong As Long) As Long

Public Declare Function ntohs _
    Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
'end winsock byte ordering declarations

Public Declare Sub RtlMoveMemory _
    Lib "kernel32" (hpvDest As Any, _
                    ByVal hpvSource As Long, _
                    ByVal cbCopy As Long)

' end winsock API declarations

' lets start here
' I am not going to go into real detail here about what is going on
' I will however try to explain a little bit and give you links to what helped me
' the winsock api uses UnSigned Data Types (unsigned short and long (integer and long))
' which is used in C - C++ - visual basic uses signed data types (signed short and long (integer and long))
' well if you look at the microsoft knowledge base there is an article with
' the code I will use below to convert between the 2 types (unsigned and signed)
' Article ID: Q189323 - it can be found online at microsofts web site or
' in the MSDN CD's that contain the VB help files
' there you will find an explination of unsigned and signed data types
' and a link on how to pass unsigned values to a dll from VB
' got it good :)


' code for unsigned to signed data types starts here

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Public Function UnsignedToLong(Value As Double) As Long
    '
    'The function takes a Double containing a value in the
    'range of an unsigned Long and returns a Long that you
    'can pass to an API that requires an unsigned Long
    '
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
    '
End Function

Public Function LongToUnsigned(Value As Long) As Double
    '
    'The function takes an unsigned Long from an API and
    'converts it to a Double for display or arithmetic purposes
    '
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
    '
End Function

Public Function UnsignedToInteger(Value As Long) As Integer
    '
    'The function takes a Long containing a value in the range
    'of an unsigned Integer and returns an Integer that you
    'can pass to an API that requires an unsigned Integer
    '
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
    '
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
    '
    'The function takes an unsigned Integer from and API and
    'converts it to a Long for display or arithmetic purposes
    '
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
    '
End Function

' End unsigned to signed data types


Public Function GetString(ByVal lPoint As Long) As String

    Dim sTemp As String
    Dim lPointStringVal As Long

' get the sTemp String ready to accept the pointer value
    sTemp = String$(lstrlen(ByVal lPoint), 0)

' copy the string pointer into the sTemp string
    lPointStringVal = lstrcpy(ByVal sTemp, ByVal lPoint)


    If lPointStringVal Then GetString = sTemp

End Function
 


Public Function WSErrHandle(lErr As Long)
' this function will display the error message for the
' winsock error - provided we have any - lets hope not
' it is better to create a function like this
' otherwise in every sub or function we need error handling
' we will have to re-type all this code
' this reduces the amount of code we need
' all the full error explinations can be found again in the MSDN Library
' under winsock api or Platform SDK: Windows Sockets
' I just copied the short descripton here to give you an idea
' of what is going on

    Select Case lErr
        Case WSANOTINITIALISED ' Successful WSAStartup not yet performed
            MsgBox "A successful WSAStartup call must occur before using this function."
        Case WSAENETDOWN ' Network is down
            MsgBox "The network subsystem has failed.", vbOKOnly + vbCritical, "Winsock Error"
        Case WSAHOST_NOT_FOUND ' Host not found
            MsgBox "Authoritative answer host not found.", vbOKOnly + vbCritical, "Winsock Error"
        Case WSATRY_AGAIN ' Nonauthoritative host not found
            MsgBox "Nonauthoritative host not found, or server failure.", vbOKOnly + vbCritical, "Winsock Error"
        Case WSANO_RECOVERY ' This is a nonrecoverable error
            MsgBox "A nonrecoverable error occurred.", vbOKOnly + vbCritical, "Winsock Error"
        Case WSANO_DATA ' Valid name, no data record of requested type
            MsgBox "Valid name, no data record of requested type.", vbOKOnly + vbCritical, "Winsock Error"
        Case WSAEINPROGRESS ' Operation now in progress
            MsgBox "A blocking Windows Sockets 1.1 call is in progress, or the service provider is still processing a callback function.", vbOKOnly + vbCritical, "Winsock Error"
        Case WSAEFAULT ' Bad address
            MsgBox "The name parameter is not a valid part of the user address space.", vbOKOnly + vbCritical, "Winsock Error"
        'Case WSAEINTR ' Interrupted function call
        '    MsgBox "A blocking Windows Socket 1.1 call was canceled through WSACancelBlockingCall.", vbOKOnly + vbCritical, "Winsock Error"
        Case Else ' any other error - not likely thougth but just incase
            MsgBox "Unkown Error Occured!", vbOKOnly + vbCritical, "Winsock Error"
    End Select
End Function

