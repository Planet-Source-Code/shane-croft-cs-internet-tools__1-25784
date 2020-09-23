VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmNetstat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Netstat"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmNetstat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   7620
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save List"
      Height          =   375
      Left            =   6323
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update List"
      Height          =   375
      Left            =   5003
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   83
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Local IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Local Computer"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Remote IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Remote Computer"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4440
      Picture         =   "FrmNetstat.frx":1D12
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   3000
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "Total In List:"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
End
Attribute VB_Name = "FrmNetstat"
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
Private Function GetIp(lIPAdd As Long) As String

    GetIp = GetString(inet_ntoa(lIPAdd))

End Function
 

Private Function GetPort(lPort As Long) As Long

    GetPort = IntegerToUnsigned(ntohs(UnsignedToInteger(lPort)))

End Function
Private Function GetState(lngState As Long) As String

    Select Case lngState
        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case MIB_TCP_STATE_ESTAB: GetState = "ESTAB"
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select

End Function
Private Function HostNameFromLong(lngInetAdr As Long) As String

    Dim lPointtoHost As Long
    
    Dim lPointtoHostName As Long
    
    Dim sHName As String
    
    Dim mHost As HOSTENT

' Get the pointer to the Host
    lPointtoHost = gethostbyaddr(lngInetAdr, 4, 1)

' put data into the Host
    RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

    sHName = String(256, 0)

' Copy the host name
    RtlMoveMemory ByVal sHName, ByVal mHost.hName, 256

    sHName = Left(sHName, InStr(1, sHName, Chr(0)) - 1)

    HostNameFromLong = sHName

End Function
Private Function WSAPIFun1(icType As Integer, tText As TextBox, tlist As listbox)
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

    tlist.Clear
' here we are checking to see what type of call we need
' if we want the host by name then we do not need the following code
' else if we want local ip address and name then we do
If icType = 1 Then
    Dim sHostN As String * 256
    Dim lV As Long

    lV = gethostname(sHostN, 256)

    If lV = SOCKET_ERROR Then
        WSErrHandle (Err.LastDllError)
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
        WSErrHandle (Err.LastDllError)
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
            tlist.AddItem sIPAdd

            sIPAdd = ""

            mHost.hAddrList = mHost.hAddrList + LenB(mHost.hAddrList)
            RtlMoveMemory lPointtoIP, mHost.hAddrList, 4

         Loop
    End If

End Function
Private Function UpdateList()
Command1.Enabled = False
Command2.Enabled = False
DoEvents
    Dim aBuf() As Byte
    Dim lSize As Long
    Dim lV As Long
    Dim lR As Long
    Dim i As Long
    Dim TCPtr As MIB_TCPROW

    ListView1.ListItems.Clear

    Me.MousePointer = vbHourglass

    lSize = 0

' Call the GetTcpTable just to get the buffer size into the lSize variable
    lV = GetTcpTable(ByVal 0&, lSize, 0)

    If lV = ERROR_NOT_SUPPORTED Then
' API is not supported
        MsgBox "not supported by this system.", vbOKOnly + vbInformation, "Error"
        Exit Function
    End If
    
    ReDim aBuf(0 To lSize - 1) As Byte

    lV = GetTcpTable(aBuf(0), lSize, 0)

    If lV = ERROR_SUCCESS Then

        CopyMemory lR, aBuf(0), 4

        ReDim aTcpTblRow(1 To lR)

            Dim lcIP As String
            Dim lcHst As String
            Dim lcPrt As String
            Dim rmIP As String
            Dim rmHst As String
            Dim rmPrt As String
            Dim tStat As String

        For i = 1 To lR

            DoEvents

' Copy the table row data to the TCPtr structure
            CopyMemory TCPtr, aBuf(4 + (i - 1) * Len(TCPtr)), Len(TCPtr)

' Add data to the listbox
           
                With TCPtr
                
                 
                 lcIP = GetIp(.dwLocalAddr)
                 lcHst = HostNameFromLong(.dwLocalAddr)
                 lcPrt = GetPort(.dwLocalPort)
                 rmIP = GetIp(.dwRemoteAddr)
                 rmHst = HostNameFromLong(.dwRemoteAddr)
                 rmPrt = GetPort(.dwRemotePort)
                 tStat = GetState(.dwState)
                
                ' just a check to see the type of data returned
                ' in the list box we need to add an extra tab space if
                ' localhost is returned - this may only be the case on my network
                ' because of the name sceme on the network - you may see the data displayed
                ' all messed up - in my case this make sthe data display nice.
                ' there are better ways to display this data but I just threw this
                ' together quick and didnt want to get into the listview control
                ' or a datagrid control - you should be able to get the idea and
                ' change this part as you see fit to display this data
                
                If lcHst = "localhost" Then
                lcHst = lcHst
                End If
                
                If rmHst = "localhost" Then
                rmHst = rmHst
                End If
                
                    Set Item = ListView1.ListItems.Add(, , lcIP)
                   Item.SubItems(1) = lcHst
                   Item.SubItems(2) = lcPrt
                   Item.SubItems(3) = rmIP
                   Item.SubItems(4) = rmHst
                   Item.SubItems(5) = rmPrt
                   Item.SubItems(6) = tStat
                   DoEvents
                End With
            aTcpTblRow(i) = TCPtr

        Next i

    End If

    Me.MousePointer = vbNormal
    Label2.Caption = "Netstat status as of: " & Date & " " & Time
    Command1.Enabled = True
    Command2.Enabled = True
End Function

Private Sub Command1_Click()
Label2.Caption = "Please Wait Loading List..."
DoEvents
UpdateList
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim X As Long
If Me.ListView1.ListItems.Count = 0 Then
MsgBox "There is nothing in the list to save."
Exit Sub
End If
X = Me.ListView1.SelectedItem.Index - 1
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - X)
DoEvents
FrmNetstatSave.Show
DoEvents
FrmNetstatSave.List1.Clear
DoEvents
FrmNetstatSave.List1.AddItem Label2.Caption
FrmNetstatSave.List1.AddItem ""
FrmNetstatSave.List1.AddItem "Total Connections: " & Me.ListView1.ListItems.Count
FrmNetstatSave.List1.AddItem ""
DoEvents
Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count
FrmNetstatSave.List1.AddItem "Connection: " & Format(Me.ListView1.SelectedItem.Index, "00") & vbTab & "Status: " & Me.ListView1.SelectedItem.SubItems(6)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Local IP: " & Me.ListView1.SelectedItem.Text
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Local Computer: " & Me.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Local Port: " & Me.ListView1.SelectedItem.SubItems(2)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Remote IP: " & Me.ListView1.SelectedItem.SubItems(3)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Remote Computer: " & Me.ListView1.SelectedItem.SubItems(4)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Remote Port: " & Me.ListView1.SelectedItem.SubItems(5)
DoEvents
FrmNetstatSave.List1.AddItem ""
DoEvents
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
DoEvents
Loop
FrmNetstatSave.List1.AddItem "Connection: " & Format(Me.ListView1.SelectedItem.Index, "00") & vbTab & "Status: " & Me.ListView1.SelectedItem.SubItems(6)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Local IP: " & Me.ListView1.SelectedItem.Text
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Local Computer: " & Me.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Local Port: " & Me.ListView1.SelectedItem.SubItems(2)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Remote IP: " & Me.ListView1.SelectedItem.SubItems(3)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Remote Computer: " & Me.ListView1.SelectedItem.SubItems(4)
DoEvents
FrmNetstatSave.List1.AddItem vbTab & "  Remote Port: " & Me.ListView1.SelectedItem.SubItems(5)
DoEvents
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
DoEvents
FrmNetstatSave.List1.AddItem ""
DoEvents
End Sub

Private Sub Form_Load()
Dim lV As Long
Dim mWSD As WSADATA
Me.Top = 0
Me.Left = 0
' start the winsock service
' we need to load this before we can do any type of winsocking :)

    lV = WSAStartup(&H101, mWSD)
    DoEvents

End Sub

Private Sub Timer1_Timer()
Label1.Caption = "Total In List: " & ListView1.ListItems.Count
End Sub
