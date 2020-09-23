VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmPortScanner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPortScanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   6330
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   210
      Left            =   75
      TabIndex        =   16
      Text            =   "Ports Scanned:"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Portn 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   210
      Left            =   75
      TabIndex        =   15
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Text            =   "Open Ports:"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   195
      Left            =   2520
      TabIndex        =   13
      Text            =   "Ports to scan:"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Text            =   "IP Address:"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Text            =   "Max Connections:"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   120
   End
   Begin VB.ListBox lstOpenPorts 
      Height          =   2160
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   4695
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtMaxConnections 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "25"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtLowerBound 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtUpperBound 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "65535"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save To File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2880
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock wskSocket 
      Index           =   0
      Left            =   5520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblTo 
      BackColor       =   &H80000013&
      Caption         =   "To"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmPortScanner.frx":1CFA
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FrmPortScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngNextPort As Long

Public Sub cmdClearList_Click()
   Me.lstOpenPorts.Clear
   Label4.Caption = ""
   Label3.Caption = ""
   Label2.Caption = ""
   Command1.Enabled = False
End Sub

Private Sub cmdScan_Click()
On Error Resume Next
If txtIP.Text = "" Then
MsgBox "Please Enter A IP", vbCritical
Exit Sub
End If
If txtLowerBound.Text = "" Then
MsgBox "Please Enter A Begining Port", vbCritical
Exit Sub
End If
If txtUpperBound.Text = "" Then
MsgBox "Please Enter A Ending Port", vbCritical
Exit Sub
End If
If txtMaxConnections.Text = "" Then
MsgBox "Please Enter A Max Connection", vbCritical
Exit Sub
End If
   Dim intI As Integer
   Command1.Enabled = False
   cmdClearList.Enabled = False
   Me.txtIP.Enabled = False
   Me.txtLowerBound.Enabled = False
   Me.txtUpperBound.Enabled = False
   Me.txtMaxConnections.Enabled = False
   Label4.Caption = ""
   Label2.Caption = "Scan Started at " & Date & " " & Time
   Label3.Caption = ""
   lstOpenPorts.Clear
   lngNextPort = Val(Me.txtLowerBound)
   PB1.Max = txtUpperBound.Text
   PB1.Min = txtLowerBound.Text
   For intI = 1 To Val(Me.txtMaxConnections)
   
      Load Me.wskSocket(intI)
     
      lngNextPort = lngNextPort + 1
      
      Me.wskSocket(intI).Connect Me.txtIP, lngNextPort
   
   Next intI
timTimer.Enabled = True
 cmdStop.Enabled = True
cmdScan.Enabled = False
End Sub

Public Sub cmdStop_Click()
On Error Resume Next
   Dim intI As Integer
   timTimer.Enabled = False
   For intI = 1 To Val(Me.txtMaxConnections)
   
      Me.wskSocket(intI).Close
 
      Unload Me.wskSocket(intI)
   
   Next intI
   
cmdStop.Enabled = False
cmdScan.Enabled = True
Command1.Enabled = True
cmdClearList.Enabled = True
Label1.Caption = "0%"
Portn.Text = "0"
PB1.Value = txtLowerBound.Text
Label4.Caption = "Scan Stopped By User!"
Label3.Caption = "Scan Stopped at " & Date & " " & Time
FrmPortScanner.Caption = "Port Scanner"
   Me.txtIP.Enabled = True
   Me.txtLowerBound.Enabled = True
   Me.txtUpperBound.Enabled = True
   Me.txtMaxConnections.Enabled = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
FrmReport.Show
DoEvents
FrmReport.List1.Clear
DoEvents
FrmReport.List1.AddItem "Address Scanned: " & FrmPortScanner.txtIP.Text
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem FrmPortScanner.Label2.Caption
FrmReport.List1.AddItem FrmPortScanner.Label3.Caption
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem "Ports Scanned: " & FrmPortScanner.txtLowerBound.Text & " To " & FrmPortScanner.txtUpperBound.Text
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem "Total Ports Found Open: " & FrmPortScanner.lstOpenPorts.ListCount
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem "Current Ports Found Open:"
FrmPortScanner.lstOpenPorts.ListIndex = 0
Do Until FrmPortScanner.lstOpenPorts.ListIndex = FrmPortScanner.lstOpenPorts.ListCount - 1
FrmReport.List1.AddItem FrmPortScanner.lstOpenPorts.Text
FrmPortScanner.lstOpenPorts.ListIndex = FrmPortScanner.lstOpenPorts.ListIndex + 1
Loop
FrmReport.List1.AddItem FrmPortScanner.lstOpenPorts.Text
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub timTimer_Timer()
On Error Resume Next
   Me.Portn.Text = Str(lngNextPort)
   PB1.Value = Str(lngNextPort)
   Label1.Caption = Int((lngNextPort - Me.txtLowerBound.Text) / (Me.txtUpperBound.Text - Me.txtLowerBound.Text) * 100) & " %" '
   FrmPortScanner.Caption = Label1.Caption & " Port Scanner"
End Sub

Private Sub txtIP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If
End Sub

Private Sub txtIP_LostFocus()
On Error Resume Next
txtIP.Text = Replace(txtIP.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub txtLowerBound_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If
End Sub

Private Sub txtLowerBound_LostFocus()
On Error Resume Next
txtLowerBound.Text = Replace(txtLowerBound.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub txtMaxConnections_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If
End Sub

Private Sub txtMaxConnections_LostFocus()
On Error Resume Next
txtMaxConnections.Text = Replace(txtMaxConnections.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub txtUpperBound_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If
End Sub

Private Sub txtUpperBound_LostFocus()
On Error Resume Next
txtUpperBound.Text = Replace(txtUpperBound.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub wskSocket_Connect(Index As Integer)

   Me.lstOpenPorts.AddItem "Port: " & Format(Me.wskSocket(Index).RemotePort, "00000")
  
   Try_Next_Port (Index)

End Sub

Private Sub wskSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

   Try_Next_Port (Index)

End Sub

Private Sub Try_Next_Port(Index As Integer)
On Error Resume Next
   Me.wskSocket(Index).Close

   If lngNextPort <= Val(Me.txtUpperBound) Then
      
      Me.wskSocket(Index).Connect , lngNextPort
      
      lngNextPort = lngNextPort + 1

   Else

      Unload Me.wskSocket(Index)
      Me.cmdScan.Enabled = True
      Me.cmdStop.Enabled = False
      Command1.Enabled = True
      Me.timTimer.Enabled = False
      cmdClearList.Enabled = True
      Label4.Caption = "Scan Finished!"
      Label1.Caption = "0%"
      Portn.Text = "0"
      PB1.Value = txtLowerBound.Text
      Label3.Caption = "Scan Finished at " & Date & " " & Time
      FrmPortScanner.Caption = "Port Scanner"
   Me.txtIP.Enabled = True
   Me.txtLowerBound.Enabled = True
   Me.txtUpperBound.Enabled = True
   Me.txtMaxConnections.Enabled = True

   End If

End Sub

Private Sub Close_Click()
Unload Me
End Sub
