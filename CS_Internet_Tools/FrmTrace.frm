VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTrace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trace Route"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTrace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   7545
   Begin VB.CommandButton Command2 
      Caption         =   "Save To File"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resolve"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Resolve Ip To Host When Finished."
      Height          =   210
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Double Click To Port Scan Selected IP"
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hop"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Host"
         Object.Width           =   7938
      EndProperty
   End
   Begin VB.CommandButton TraceRT2 
      Caption         =   "Trace Route"
      Default         =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   255
      Left            =   4785
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Host 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   825
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6480
      Picture         =   "FrmTrace.frx":1D12
      Top             =   240
      Width           =   480
   End
   Begin VB.Label IP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   825
      TabIndex        =   8
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   585
      TabIndex        =   7
      Top             =   480
      Width           =   210
   End
   Begin VB.Label lblIPHost 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "IP/Host:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "FrmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalNum As Long
Dim KeepGoing As Integer
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
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
Private Declare Function gethostbyaddr _
    Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, _
                      ByVal addr_type As Long) As Long

Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0

Public Sub GetRCode()
RCode = ""
DoEvents
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
'    If pIPe.Status = 11013 Then RCode = "TTL Exprd In Transit"
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
'    RCode = RCode + " (" + CStr(pIPe.Status) + ")"


    DoEvents

        If RCode <> "" Then
            If RCode = "Reqested Timed Out" Then
                vbWSACleanup
                If TotalNum < 10 Then
                Set Item = ListView1.ListItems.Add(, , " # 0" & TotalNum)
                Item.SubItems(1) = RCode
                Else
                Set Item = ListView1.ListItems.Add(, , " # " & TotalNum)
                Item.SubItems(1) = RCode
                End If
            Exit Sub
            End If
            If RCode = "Success" Then
                vbWSACleanup
                If TotalNum < 10 Then
                Set Item = ListView1.ListItems.Add(, , " # 0" & TotalNum)
                Item.SubItems(1) = IP
                Else
                Set Item = ListView1.ListItems.Add(, , " # " & TotalNum)
                Item.SubItems(1) = IP
                End If
            Exit Sub
            End If
            KeepGoing = 1
            Set Item = ListView1.ListItems.Add(, , RCode)
        Else
            If TTL - 1 < 10 Then
            Set Item = ListView1.ListItems.Add(, , " # 0" & TotalNum)
            Item.SubItems(1) = RespondingHost
            Else
            Set Item = ListView1.ListItems.Add(, , " # " & TotalNum)
            Item.SubItems(1) = RespondingHost
            End If
        End If
    End Sub


Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HOSTENT.h_name, ByVal _
        PointerToPointer, Len(HOSTENT) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HOSTENT.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory addr, ByVal ListAddr, 4
        IP.Caption = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
        + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
    End If
End Sub


Public Sub vbGetHostName()
    
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.Text = Host
    End If
End Sub


Public Sub vbIcmpSendEcho()
    vbWSACleanup
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)


        DoEvents
        vbWSACleanup
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo, pIPe, Len(pIPe) + 8, 2700)
            If bReturn Then
                TotalNum = TotalNum + 1
                RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))
                GetRCode
            Else
                TotalNum = TotalNum + 1
                    GetRCode
                    TTL = TTL + 1
            End If
        Next NbrOfPkts
    End Sub


Sub vbWSAStartup()
Dim wsdaata As WSADATA
    iReturn = WSAStartup(&H101, WSADATA)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        MsgBox "WSock32.dll is Not responding!", 0, ""
    End If


    If LoByte(WSADATA.wversion) < WS_VERSION_MAJOR Or (LoByte(WSADATA.wversion) = WS_VERSION_MAJOR And HiByte(WSADATA.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSADATA.wversion)))
        sLowByte = Trim$(Str$(LoByte(WSADATA.wversion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        MsgBox sMsg
        End
    End If


    If WSADATA.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            MsgBox sMsg
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

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Command2.Enabled = False
Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count

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

    sIP = Trim$(ListView1.SelectedItem.SubItems(1))
Label1.Caption = "Resolving " & ListView1.SelectedItem.SubItems(1) & " To Host"
DoEvents
' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.h_name, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            ListView1.SelectedItem.SubItems(2) = sHost
            DoEvents

        End If

    End If
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
Loop

    sIP = Trim$(ListView1.SelectedItem.SubItems(1))
Label1.Caption = "Resolving " & ListView1.SelectedItem.SubItems(1) & " To Host"
DoEvents
' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.h_name, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            ListView1.SelectedItem.SubItems(2) = sHost
            DoEvents

        End If

    End If
Label1.Caption = "Resolving IP To Host Is Complete."
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim X As Long
X = Me.ListView1.SelectedItem.Index - 1
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - X)
DoEvents
FrmSaveTrace.Show
DoEvents
FrmSaveTrace.List1.Clear
DoEvents
FrmSaveTrace.List1.AddItem "Address Traced: " & FrmTrace.Host.Text
FrmSaveTrace.List1.AddItem ""
FrmSaveTrace.List1.AddItem "Total Hops: " & FrmTrace.ListView1.ListItems.Count
FrmSaveTrace.List1.AddItem ""
DoEvents
Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count
FrmSaveTrace.List1.AddItem "Hop: " & FrmTrace.ListView1.SelectedItem.Text
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  IP: " & FrmTrace.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  Host: " & FrmTrace.ListView1.SelectedItem.SubItems(2)
DoEvents
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
DoEvents
Loop
FrmSaveTrace.List1.AddItem "Hop: " & FrmTrace.ListView1.SelectedItem.Text
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  IP: " & FrmTrace.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  Host: " & FrmTrace.ListView1.SelectedItem.SubItems(2)
DoEvents
End Sub

Private Sub Form_Load()
Dim mWSD As WSADATA
Me.Top = 0
Me.Left = 0
lV = WSAStartup(&H202, mWSD)
vbWSAStartup
vbWSACleanup
End Sub

Private Sub Form_Unload(Cancel As Integer)
KeepGoing = 1

End Sub

Private Sub Host_LostFocus()
On Error Resume Next
Host.Text = Replace(Host.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
If ListView1.ListItems.Count = 0 Then
Exit Sub
End If

FrmPortScanner.Show
FrmPortScanner.cmdStop_Click
DoEvents
FrmPortScanner.cmdClearList_Click
DoEvents
FrmPortScanner.txtIP = Me.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmPortScanner.SetFocus
End Sub

Private Sub Timer1_Timer()
If Host.Text = "" Then
TraceRT2.Enabled = False
Else
TraceRT2.Enabled = True
End If
End Sub

Private Sub TraceRT2_Click()
Command1.Enabled = False
Command2.Enabled = False
TotalNum = 0
    szBuffer = Space(32)
    ListView1.ListItems.Clear
    vbWSAStartup


    If Len(Host.Text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
    vbIcmpCreateFile
    ' The following determines the TTL of th
    '     e ICMPEcho for TRACE function
    TraceRT = True
    Label1.Caption = "Tracing Route To " + IP.Caption

    For TTL = 2 To 255
        If KeepGoing = 1 Then
        KeepGoing = 0
        Exit For
        End If
        pIPo.TTL = TTL
        DoEvents
        vbIcmpSendEcho


        DoEvents

            If RespondingHost = IP.Caption Then
                Label1.Caption = "Trace Route has Completed"
                Exit For ' Stop TraceRT
            End If
        Next TTL
        TraceRT = False
        vbIcmpCloseHandle
        vbWSACleanup
DoEvents
DoEvents
Command1.Enabled = True
Command2.Enabled = True
If Check1.Value = 1 Then
Command1_Click
End If
End Sub

