VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmPortListen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Listener"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPortListen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   4530
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000007&
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox port3 
      Height          =   285
      Left            =   600
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "139"
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton optUDP 
      BackColor       =   &H8000000B&
      Caption         =   "UDP"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.OptionButton optTCP 
      BackColor       =   &H8000000B&
      Caption         =   "TCP/IP"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Listen"
      Default         =   -1  'True
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   2520
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3120
      Picture         =   "FrmPortListen.frx":1D12
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Protocol:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "FrmPortListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdConnect_Click()
If port3.Text = "" Then
MsgBox "Please Enter A Port Number", vbCritical
Exit Sub
End If
cmdConnect.Enabled = False
port3.Enabled = False
cmdDisconnect.Enabled = True
txtStatus = ""
If optTCP = True Then
    ws1.Protocol = sckTCPProtocol
End If
If optUDP = True Then
    ws1.Protocol = sckUDPProtocol
End If
On Error GoTo PortIsOpen
ws1.Close
ws1.LocalPort = port3.Text
ws1.Listen
Exit Sub
PortIsOpen:
ws1.Close
If Err.Number = 10048 Then
    txtStatus = "The port " & port3.Text & " is already open."
Else
    txtStatus = "Error: " & Err.Number & vbCrLf & "   " & Err.Description
End If
cmdDisconnect.Enabled = False
port3.Enabled = True
cmdConnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
ws1.Close
cmdDisconnect.Enabled = False
port3.Enabled = True
cmdConnect.Enabled = True
End Sub


Private Sub Form_Load()
Dim mWSD As WSADATA
Me.Top = 0
Me.Left = 0
lV = WSAStartup(&H202, mWSD)
optTCP = True
End Sub


Private Sub port3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdConnect_Click
 DoEvents
 End If
End Sub

Private Sub port3_LostFocus()
On Error Resume Next
port3.Text = Replace(port3.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
 'If ws1.State <> sckClosed Then ws1.Close
 'ws1.Accept (requestID)
 txtStatus.Text = txtStatus.Text & vbCrLf & "Connection" & " - " & Date & " " & Time
End Sub

Private Sub ws1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
ws1.GetData strData
txtStatus.Text = txtStatus.Text & vbCrLf & " - " & strData & " - " & Date & " " & Time
End Sub

Private Sub ws1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtStatus = txtStatus.Text & vbCrLf & "Winsock Error: " & Number & vbCrLf & "   " & descriptoin & " - " & Date & " " & Time
End Sub
