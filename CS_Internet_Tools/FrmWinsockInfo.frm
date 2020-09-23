VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmWinsockInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock & Internet Connection Information"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmWinsockInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   7365
   Begin VB.Frame Frame2 
      Caption         =   "Internet Connection Information"
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   7095
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   22
         Text            =   "?"
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Text            =   "?"
         Top             =   960
         Width           =   525
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6480
         TabIndex        =   20
         Text            =   "?"
         Top             =   600
         Width           =   525
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Text            =   "?"
         Top             =   600
         Width           =   525
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6480
         TabIndex        =   18
         Text            =   "?"
         Top             =   240
         Width           =   525
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Check Internet Connection"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   3285
      End
      Begin VB.ListBox List1 
         Height          =   1320
         Left            =   3720
         TabIndex        =   16
         Top             =   960
         Width           =   3255
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3600
         Top             =   1920
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   10
      End
      Begin VB.Label LanConnection 
         Caption         =   "Check For Lan Connection"
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2850
      End
      Begin VB.Label Label9 
         Caption         =   "Check For RAS Installed"
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   2850
      End
      Begin VB.Label Label10 
         Caption         =   "Check if Connected to the Internet"
         Height          =   270
         Left            =   3600
         TabIndex        =   27
         Top             =   600
         Width           =   2850
      End
      Begin VB.Label Label11 
         Caption         =   "Check For connection by Proxy"
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2850
      End
      Begin VB.Label Label1 
         Caption         =   "Check For Modem Connection"
         Height          =   270
         Index           =   7
         Left            =   3600
         TabIndex        =   25
         Top             =   240
         Width           =   2850
      End
      Begin VB.Label ProgressLabel 
         Caption         =   "Progress..."
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Winsock Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6480
         Picture         =   "FrmWinsockInfo.frx":1D2A
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Version:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "High Version:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "System Status:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Max Sockets:"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Max UdpDg:"
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Vendor Info:"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmWinsockInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type WSAData2
    wversion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Private Declare Function WSAStartup _
    Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData2) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private mWSData As WSAData2 ' this will hold the wsadata we need

Private Sub Command1_Click()
On Error Resume Next
List1.Clear
ProgressBar1.Value = 0
DoEvents
    ProgressLabel.Caption = "Checking for Lan Connection..."
    ProgressBar1 = 1
    DoEvents
    Text1 = IsLanConnection()
    ProgressLabel.Caption = "Checking for Modem Connection..."
    ProgressBar1 = ProgressBar1 + 1
    DoEvents
    Text2 = IsModemConnection()
    ProgressLabel.Caption = "Checking for Connection Via Proxy..."
    ProgressBar1 = ProgressBar1 + 1
    DoEvents
    Text3 = IsProxyConnection()
    ProgressLabel.Caption = "Checking for Any Internet Connection..."
    ProgressBar1 = ProgressBar1 + 1
    DoEvents
    Text4 = IsConnected()
    ProgressLabel.Caption = "Checking if RAS is installed..."
    ProgressBar1 = ProgressBar1 + 1
    DoEvents
    Text5 = IsRasInstalled()
    ProgressLabel.Caption = "Getting connection type..."
    ProgressBar1 = ProgressBar1 + 1
    DoEvents
    Call ConnectionTypeMsg(List1)
    ProgressBar1 = 10
    ProgressLabel.Caption = "Done."

End Sub

Private Sub Form_Load()
' I would go into more detail here but most of this information can be found in the MSDN
' Library that came with VB when you bought it
' Otherwise the knowledge base on the microsoft web site has
' almost all of the information needed if not all of it
' for this version we are using winsock version 1.1
' if you want to use winsock version 2.2 then change
' lV = WSAStartup(&H101, mWSD) to
' lV = WSAStartup(&H202, mWSD)
Me.Top = 0
Me.Left = 0
Dim lV As Long
Dim mWSD As WSAData2

' start the winsock service
' we need to load this before we can do any type of winsocking :)

    lV = WSAStartup(&H101, mWSD)

' this is to check and make sure the winsock service has started
' before we proceed any further

    If lV <> 0 Then

    Select Case lV
        Case WSASYSNOTREADY ' winsock error system not ready
            MsgBox "The system is not ready!", vbOKOnly + vbInformation, "Winsock Error"
        Case WSAVERNOTSUPPORTED ' winsock error API not supported
            MsgBox "The version of Windows Sockets API is not supported!", vbOKOnly + vbInformation, "Winsock Error"
        Case WSAEINVAL ' winsock error the socket version is not supported
            MsgBox "The Windows Sockets version is not supported!", vbOKOnly + vbInformation, "Winsock Error"
        Case Else
            MsgBox "An unknown error has occured!", vbOKOnly + vbInformation, "Winsock Error"
        End Select

    End If
    
mWSData = mWSD 'set our declaration to the wsadata

' set up our labels on our form with the winsock information
Label2.Caption = mWSData.wversion \ 256 & "." & mWSData.wversion Mod 256

Label3.Caption = mWSData.wHighVersion \ 256 & "." & mWSData.wHighVersion Mod 256
                  
Label4.Caption = mWSData.szDescription

Label5.Caption = mWSData.szSystemStatus

Label6.Caption = IntegerToUnsigned(mWSData.iMaxSockets)

Label7.Caption = IntegerToUnsigned(mWSData.iMaxUdpDg)

Label8.Caption = mWSData.lpVendorInfo


End Sub

Private Sub Timer1_Timer()
On Error Resume Next
List1.ToolTipText = List1.Text
End Sub
