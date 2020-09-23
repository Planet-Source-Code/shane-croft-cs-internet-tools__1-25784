VERSION 5.00
Begin VB.Form FrmResolve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolve Host or IP"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmResolve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   3705
   Begin VB.Frame Frame2 
      Caption         =   "Get Host By Name"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Type Host Name Here"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get"
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2880
         Picture         =   "FrmResolve.frx":1D2A
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Get Host By IP Address"
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Type IP Address Here"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get Host Name"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "Host Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "FrmResolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type
'
Private Type WSADATA
    wversion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NO_DATA = 232&
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&
'
Private Const MIB_TCP_STATE_CLOSED = 1
Private Const MIB_TCP_STATE_LISTEN = 2
Private Const MIB_TCP_STATE_SYN_SENT = 3
Private Const MIB_TCP_STATE_SYN_RCVD = 4
Private Const MIB_TCP_STATE_ESTAB = 5
Private Const MIB_TCP_STATE_FIN_WAIT1 = 6
Private Const MIB_TCP_STATE_FIN_WAIT2 = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT = 8
Private Const MIB_TCP_STATE_CLOSING = 9
Private Const MIB_TCP_STATE_LAST_ACK = 10
Private Const MIB_TCP_STATE_TIME_WAIT = 11
Private Const MIB_TCP_STATE_DELETE_TCB = 12
'
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPROW) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
'
Private aTcpTblRow() As MIB_TCPROW

Private Declare Function gethostbyaddr _
    Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, _
                      ByVal addr_type As Long) As Long
Private Declare Function WSAStartup _
    Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Private mWSData As WSADATA ' this will hold the wsadata we need

Private Function WSAPIFun1(icType As Integer, tText As TextBox, tlist As TextBox)
' we are seeting up a function here to do some of the
' WS api calls - again we set up functions so there isnt much code being repeated
' ictype returns 1 = get name and IP address of local system
' ictype returns 2 = get remote host by name


' Pointer to host
    Dim lPointtoHost As Long
' stores all the host info
    Dim mHost As HOSTENT
' pointer to the IP address list - there may be several IP address for 1 host
    Dim lPointtoIP As Long
' array that holds elemets of an IP address
    Dim aIPAdd() As Byte
' IP address to add into the ListBox
    Dim sIPAdd As String


' here we are checking to see what type of call we need
' if we want the host by name then we do not need the following code
' else if we want local ip address and name then we do
If icType = 1 Then
    Dim sHostN As String * 256
    Dim lV As Long

    lV = gethostname(sHostN, 256)

    If lV = SOCKET_ERROR Then
        'WSErrHandle (Err.LastDllError)
        Text2.Text = "Unable To Resolve"
        Exit Function
    End If

    tText.Text = Left(sHostN, InStr(1, sHostN, Chr(0)) - 1)
End If

' Call the gethostbyname Winsock API function
    lPointtoHost = gethostbyname(Trim$(tText.Text))

' Check to see if the lPointtoHost value has returned anything
' if we get a 0 then that means there was an error getting the host info
' here is where we saved time typeing and we call the error function
' we created for the winsock api
    If lPointtoHost = 0 Then
        'WSErrHandle (Err.LastDllError)
        Text2.Text = "Unable To Resolve"
    Else
' Copy data to mHost structure
        RtlMoveMemory mHost, lPointtoHost, LenB(mHost)

        RtlMoveMemory lPointtoIP, mHost.hAddrList, 4

        Do Until lPointtoIP = 0
            
            ReDim aIPAdd(1 To mHost.hLength)

            RtlMoveMemory aIPAdd(1), lPointtoIP, mHost.hLength

            For i = 1 To mHost.hLength
                sIPAdd = sIPAdd & aIPAdd(i) & "."
            Next

            sIPAdd = Left$(sIPAdd, Len(sIPAdd) - 1)

' Add the IP address to the listbox
            tlist.Text = sIPAdd

            sIPAdd = ""

            mHost.hAddrList = mHost.hAddrList + LenB(mHost.hAddrList)
            RtlMoveMemory lPointtoIP, mHost.hAddrList, 4

         Loop
    End If

End Function
Private Sub Command1_Click()
If Text1.Text <> "Type Host Name Here" Then
Screen.MousePointer = vbHourglass
WSAPIFun1 2, Text1, Text2
Screen.MousePointer = vbNormal
End If
End Sub

Private Sub Command3_Click()
If Text3.Text = "Type IP Address Here" Then Exit Sub
Screen.MousePointer = vbHourglass
' The inet_addr function returns a long value
    Dim lInteAdd As Long
' pointer to the HOSTENT
    Dim lPointtoHost As Long
' host name we are looking for
    Dim sHost As String
' Hostent
    Dim mHost As HOSTENT
' IP Address
    Dim sIP As String

    sIP = Trim$(Text3.Text)

' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        'WSErrHandle (Err.LastDllError)
        Text4.Text = "Unable To Resolve"

    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            'WSErrHandle (Err.LastDllError)
            Text4.Text = "Unable To Resolve"

        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.hName, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            Text4.Text = sHost

        End If

    End If
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
Dim mWSD As WSADATA
Me.Top = 0
Me.Left = 0
lV = WSAStartup(&H202, mWSD)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Text1_LostFocus()
On Error Resume Next
Text1.Text = Replace(Text1.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command3_Click
 DoEvents
 End If
End Sub

Private Sub Text3_LostFocus()
On Error Resume Next
Text3.Text = Replace(Text3.Text, " ", "", 1, , vbTextCompare)
End Sub
