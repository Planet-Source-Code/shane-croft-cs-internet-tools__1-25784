VERSION 5.00
Begin VB.Form FrmPing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "FrmPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   5985
   Begin VB.TextBox lblPacketSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2880
      TabIndex        =   2
      Text            =   "32"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox lblPingTimes 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2880
      TabIndex        =   1
      Text            =   "4"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5775
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Host 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   330
      Width           =   2055
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4920
      Picture         =   "FrmPing.frx":1A7A
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblIpHost 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Ip/Host:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   570
   End
   Begin VB.Label lblPings 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Ping(s):"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblPacket 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Packet:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   525
   End
End
Attribute VB_Name = "FrmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PingTimes As Integer
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Integer
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
Dim ExitTheFor As Integer
' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0


Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdPing_Click()
On Error Resume Next
If Host.Text = "" Then
MsgBox "Please Enter A Ip/Host To Ping.", vbInformation
Exit Sub
End If
If gethostbyname(Host.Text) = 0 Then
txtStatus.Text = "Unable To Resolve Host"
Exit Sub
End If
    Speed = 0
    PingTimes = 0
    cmdPing.Enabled = False
    txtStatus = ""
    szBuffer = Space(Val(lblPacketSize))
    vbWSAStartup
    If Len(Host.Text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
    vbIcmpCreateFile
    pIPo2.TTL = Trim$(255)
    '
    For Times = 1 To lblPingTimes
    If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
    vbIcmpSendEcho
    Next
    vbIcmpCloseHandle
    vbWSACleanup
    cmdPing.Enabled = True
    On Error GoTo skipit
    Speed = Speed / PingTimes
    txtStatus = txtStatus & vbCrLf & " Average Speed: " & Speed & "."
    txtStatus.SelStart = Len(txtStatus)
    Exit Sub
skipit:
End Sub

Public Sub GetRCode()
RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Destination Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Destination Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Destination Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Requested Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = "Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"

    DoEvents

        If RCode <> "" Then
            If RCode = "Success" Then
                Speed = Speed + Val(Trim$(CStr(pIPe2.RoundTripTime)))
                txtStatus.Text = txtStatus.Text + " Reply from " + RespondingHost + ": Bytes = " + Trim$(CStr(pIPe2.DataSize)) + " RTT = " + Trim$(CStr(pIPe2.RoundTripTime)) + "ms TTL = " + Trim$(CStr(pIPe2.Options.TTL)) + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            Exit Sub
            End If
            KeepGoing = 1
            txtStatus.Text = txtStatus.Text & RCode
        Else
            KeepGoing = 1
            txtStatus.Text = txtStatus.Text & RCode
        End If
        txtStatus.SelStart = Len(txtStatus)
    End Sub


Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HOSTENT.h_name, ByVal _
        PointerToPointer, Len(HOSTENT) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HOSTENT.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong2, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong2.Byte4)) + "." + CStr(Asc(IPLong2.Byte3)) _
        + "." + CStr(Asc(IPLong2.Byte2)) + "." + CStr(Asc(IPLong2.Byte1)))
    End If
End Sub


Public Sub vbGetHostName()
    
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.Text = Host
    End If
End Sub


Public Sub vbIcmpSendEcho()
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                GetRCode
            Else
                txtStatus.Text = txtStatus.Text + " Request Timeout" + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            End If
        Next NbrOfPkts
    End Sub


Sub vbWSAStartup()
Dim wsdaata As WSADATA
    iReturn = WSAStartup(&H101, WSADATA)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        txtStatus = "WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(WSADATA.wversion) < WS_VERSION_MAJOR Or (LoByte(WSADATA.wversion) = WS_VERSION_MAJOR And HiByte(WSADATA.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSADATA.wversion)))
        sLowByte = Trim$(Str$(LoByte(WSADATA.wversion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        txtStatus = sMsg
        ExitTheFor = 1
        End
    End If


    If WSADATA.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            txtStatus = sMsg
            ExitTheFor = 1
        End
    End If
    
    MaxSockets = WSADATA.iMaxSockets


    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = WSADATA.iMaxUdpDg


    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Description = ""


    For i = 0 To WSADESCRIPTION_LEN
        If WSADATA.szDescription(i) = 0 Then Exit For
        Description = Description + Chr$(WSADATA.szDescription(i))
    Next i
    Status = ""


    For i = 0 To WSASYS_STATUS_LEN
        If WSADATA.szSystemStatus(i) = 0 Then Exit For
        Status = Status + Chr$(WSADATA.szSystemStatus(i))
    Next i
End Sub


Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function


Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function


Public Sub vbWSACleanup()
    iReturn = WSACleanup()
End Sub


Public Sub vbIcmpCloseHandle()
    bReturn = IcmpCloseHandle(hIP)
End Sub


Public Sub vbIcmpCreateFile()
    hIP = IcmpCreateFile()
End Sub


Private Sub Form_Load()
Dim mWSD As WSADATA
Me.Top = 0
Me.Left = 0
lV = WSAStartup(&H202, mWSD)
vbWSAStartup
vbWSACleanup
End Sub


Private Sub Host_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdPing_Click
 DoEvents
 End If
End Sub

Private Sub Host_LostFocus()
On Error Resume Next
Host.Text = Replace(Host.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub lblPacketSize_LostFocus()
On Error Resume Next
lblPacketSize.Text = Replace(lblPacketSize.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub lblPingTimes_LostFocus()
On Error Resume Next
lblPingTimes.Text = Replace(lblPingTimes.Text, " ", "", 1, , vbTextCompare)
End Sub
