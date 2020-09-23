VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmOnline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Online/Offline Checker"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   7170
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6720
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check All"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit List"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
      Begin VB.Label Label1 
         Caption         =   "Total Online:"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Total Offline:"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Total In List:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5640
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5650
      _ExtentX        =   9975
      _ExtentY        =   3201
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   270
      Left            =   5520
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmOnline.frx":1D2A
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "FrmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Long
Dim xx As Long
Dim PingTimes As Long
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Long
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Long
Dim Description As String, Status As String
Dim ExitTheFor As Long
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

Public Sub Check_Status()
If gethostbyname(ListView1.SelectedItem.Text) = 0 Then
ListView1.SelectedItem.SubItems(1) = "Offline"
Label2(1).Caption = Label2(1).Caption + 1
Exit Sub
End If
    Speed = 0
    PingTimes = 0
    ListView1.SelectedItem.SubItems(1) = ""
    szBuffer = Space(Val("32"))
    DoEvents
    vbWSAStartup
    DoEvents
    If Len(ListView1.SelectedItem) = 0 Then
        vbGetHostName
    End If
    DoEvents
    vbGetHostByName
    vbIcmpCreateFile
    DoEvents
    pIPo2.TTL = Trim$(255)
    '
    For Times = 1 To "1"
    If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
    vbIcmpSendEcho
    DoEvents
    Next
    DoEvents
    vbIcmpCloseHandle
    vbWSACleanup
    On Error GoTo skipit
    'Speed = Speed / PingTimes
    Exit Sub
skipit:
End Sub
Public Sub List_Add(list As listbox, txt As String)
On Error Resume Next
Set Item = ListView1.ListItems.Add(, , txt)
    'List1.AddItem txt
End Sub

Public Sub List_Load(thelist As listbox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add(List1, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub

Private Sub Command1_Click()

DoEvents
ListView1.Enabled = False
DoEvents
Call Check_Status
DoEvents
vbWSACleanup
DoEvents
ListView1.Enabled = True
DoEvents
End Sub

Private Sub Command2_Click()
On Error Resume Next

Label2(0).Caption = "0"
Label2(1).Caption = "0"
DoEvents
Call Form_Load
DoEvents
ProgressBar1.Value = 0
DoEvents
ListView1.Enabled = False
DoEvents
Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count
DoEvents
Call Check_Status
DoEvents
DoEvents
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
ProgressBar1.Value = ProgressBar1.Value + 1
Loop
DoEvents
Call Check_Status
ProgressBar1.Value = ProgressBar1.Value + 1
DoEvents
ListView1.Enabled = True
DoEvents
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
FrmList.Show
End Sub

Public Sub Form_Load()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Dim mWSD As WSADATA
lV = WSAStartup(&H202, mWSD)
ListView1.ListItems.Clear
DoEvents
Call List_Load(List1, App.Path & "\List.ini")
DoEvents
ProgressBar1.Min = 0
ProgressBar1.Max = ListView1.ListItems.Count
Label2(2).Caption = ListView1.ListItems.Count
Label2(0).Caption = "0"
Label2(1).Caption = "0"
DoEvents
vbWSAStartup
vbWSACleanup
DoEvents
End Sub
Public Sub GetRCode()
RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Reqested Timed Out"
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
DoEvents
        If RCode <> "" Then
        DoEvents
            If RCode = "Success" Then
                'Speed = Speed + Val(Trim$(CStr(pIPe2.RoundTripTime)))
                DoEvents
                ListView1.SelectedItem.SubItems(1) = "Online"
                Label2(0).Caption = Label2(0).Caption + 1
            Exit Sub
            End If
            DoEvents
            KeepGoing = 1
            ListView1.SelectedItem.SubItems(1) = RCode
            DoEvents
        Else
        DoEvents
            KeepGoing = 1
            ListView1.SelectedItem.SubItems(1) = RCode
            DoEvents
        End If
    End Sub


Public Sub vbGetHostByName()
    Dim szString As String
    
    Host = Trim$(ListView1.SelectedItem.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        ListView1.SelectedItem.SubItems(1) = sMsg
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
        ListView1.SelectedItem.SubItems(1) = sMsg
        ExitTheFor = 1
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        ListView1.SelectedItem.Text = Host
    End If
End Sub


Public Sub vbIcmpSendEcho()
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
           DoEvents
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                DoEvents
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                GetRCode
            Else
                ListView1.SelectedItem.SubItems(1) = "Offline"
                Label2(1).Caption = Label2(1).Caption + 1
            End If
        Next NbrOfPkts
    End Sub


Sub vbWSAStartup()
Dim wsdaata As WSADATA
    iReturn = WSAStartup(&H101, WSADATA)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        ListView1.SelectedItem.SubItems(1) = "WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(WSADATA.wversion) < WS_VERSION_MAJOR Or (LoByte(WSADATA.wversion) = WS_VERSION_MAJOR And HiByte(WSADATA.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSADATA.wversion)))
        sLowByte = Trim$(Str$(LoByte(WSADATA.wversion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        ListView1.SelectedItem.SubItems(1) = sMsg
        ExitTheFor = 1
        End
    End If


    If WSADATA.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            ListView1.SelectedItem.SubItems(1) = sMsg
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

Private Sub Timer1_Timer()
If ListView1.ListItems.Count = 0 Then
Command1.Enabled = False
Command2.Enabled = False
Else
Command1.Enabled = True
Command2.Enabled = True
End If
End Sub
