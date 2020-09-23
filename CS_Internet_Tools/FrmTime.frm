VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Sync"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3225
   Begin VB.CommandButton Command1 
      Caption         =   "Synch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2385
      TabIndex        =   1
      Top             =   0
      Width           =   780
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   2310
   End
   Begin MSWinsockLib.Winsock StinkySock 
      Left            =   2040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2640
      Picture         =   "FrmTime.frx":1D12
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "FrmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetSystemTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME) As Long
   
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Dim sNTP As String ' the 32bit time stamp returned by the server
Dim TimeDelay As Single 'the time between the acknowledgement of
                        'the connection and the data received.
                        'we compensate by adding half of the round
                        'trip latency

Private Sub Command1_Click()
Label1.Caption = "Please Wait..."
DoEvents
StinkySock.Close
sNTP = Empty
StinkySock.RemoteHost = Combo1.Text
StinkySock.RemotePort = 37 'NTP servers port
StinkySock.Connect
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Combo1.AddItem "time.ien.it"
Combo1.AddItem "ntp.cs.mu.oz.au"
Combo1.AddItem "tock.usno.navy.mil"
Combo1.AddItem "tick.usno.navy.mil"
Combo1.AddItem "swisstime.ethz.ch"
Combo1.AddItem "ntp-cup.external.hp.com"
Combo1.AddItem "ntp1.fau.de"
Combo1.AddItem "ntps1-0.cs.tu-berlin.de"
Combo1.AddItem "ntps1-1.rz.Uni-Osnabrueck.DE"
Combo1.AddItem "tempo.cstv.to.cnr.it"
Combo1.ListIndex = 0
End Sub

Private Sub StinkySock_DataArrival(ByVal bytesTotal As Long)
Dim Data As String

StinkySock.GetData Data, vbString
sNTP = sNTP & Data
End Sub

Private Sub StinkySock_Connect()
TimeDelay = Timer
End Sub

Private Sub StinkySock_Close()
On Error Resume Next
Do Until StinkySock.State = sckClosed
 StinkySock.Close
 DoEvents
Loop
TimeDelay = ((Timer - TimeDelay) / 2)
Call SyncClock(sNTP)
End Sub

Private Sub SyncClock(tStr As String)
Dim NTPTime As Double
Dim UTCDATE As Date
Dim LngTimeFrom1990 As Long
Dim ST As SYSTEMTIME
     
tStr = Trim(tStr)
If Len(tStr) <> 4 Then
 Label1.Caption = "NTP Server returned an invalid response."
 Exit Sub
End If

NTPTime = Asc(Left$(tStr, 1)) * 256 ^ 3 + Asc(Mid$(tStr, 2, 1)) * 256 ^ 2 + Asc(Mid$(tStr, 3, 1)) * 256 ^ 1 + Asc(Right$(tStr, 1))
      
LngTimeFrom1990 = NTPTime - 2840140800#

UTCDATE = DateAdd("s", CDbl(LngTimeFrom1990 + CLng(TimeDelay)), #1/1/1990#)

ST.wYear = Year(UTCDATE)
ST.wMonth = Month(UTCDATE)
ST.wDay = Day(UTCDATE)
ST.wHour = Hour(UTCDATE)
ST.wMinute = Minute(UTCDATE)
ST.wSecond = Second(UTCDATE)

Call SetSystemTime(ST)
Label1.Caption = "Clock synchronised succesfully."

End Sub

