VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMassEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Email Sender"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMassEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   5265
   Begin VB.CommandButton Command4 
      Caption         =   "-"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      ToolTipText     =   "Remove Email From List"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      ToolTipText     =   "Add Email To List"
      Top             =   960
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Height          =   690
      Left            =   2040
      TabIndex        =   14
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtEmailBodyOfMessage 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox txtEmailSubject 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtEmailServer 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtFromEmailAddress 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtFromName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Status:"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   5055
      Begin VB.Label StatusTxt 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2520
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Total Emails In List:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4440
      Picture         =   "FrmMassEmail.frx":1D2A
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      Caption         =   "SMTP Email Server:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "Send To:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Your Email Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "FrmMassEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ClipCursor Lib "user32" _
    (lpRect As Any) As Long

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096

Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim start As Single, Tmr As Single
Private FilePathName As String



Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.Connect ' Start connection
    
    WaitFor ("220")
    
    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    StatusTxt.Caption = "Connected"
    StatusTxt.Refresh

    Winsock1.SendData (first)

    StatusTxt.Caption = "Sending Message"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    StatusTxt.Caption = "Disconnecting"
    StatusTxt.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub


Private Sub Command1_Click()
If List1.ListCount = 0 Then
MsgBox "There Must Be At Least 1 Email In The List To Continue.", vbCritical
Exit Sub
End If

PB1.Max = List1.ListCount
PB1.Min = 0
PB1.Value = 0
List1.ListIndex = 0
Do Until List1.ListIndex = List1.ListCount - 1
    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, List1.Text, List1.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
    'MsgBox ("Mail Sent")
    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    DoEvents
    PB1.Value = PB1.Value + 1
    DoEvents
    'Beep
    List1.ListIndex = List1.ListIndex + 1
    Close
Loop
    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, List1.Text, List1.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
    'MsgBox ("Mail Sent")
    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    DoEvents
    PB1.Value = PB1.Value + 1
    DoEvents
    'Beep
    Close
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()
FrmMassEmailAdd.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim AppDir As String
Dim name As String
Me.Top = 0
Me.Left = 0
AppDir = App.Path

Call List_Load(List1, App.Path & "\EmailList.ini")
DoEvents

FilePathName = AppDir + "\MassEmail.inf"
EmailAddress = GetPrivateProfileString("settings", "emailaddress", "", FilePathName)
name = GetPrivateProfileString("settings", "name", "", FilePathName)
smtp = GetPrivateProfileString("settings", "smtp", "", FilePathName)
subject = GetPrivateProfileString("settings", "subject", "", FilePathName)

DoEvents
txtFromEmailAddress.Text = EmailAddress
txtFromName.Text = name
txtEmailServer.Text = smtp
txtEmailSubject.Text = subject
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim fFile As Integer
fFile = FreeFile
Winsock1.Close
Call List_Save(List1, App.Path & "\EmailList.ini")
DoEvents

Open App.Path & "\MassEmail.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "emailaddress=" & txtFromEmailAddress.Text
Print #fFile, "name=" & txtFromName.Text
Print #fFile, "smtp=" & txtEmailServer.Text
Print #fFile, "subject=" & txtEmailSubject.Text
Close fFile

End Sub

Private Sub Timer1_Timer()
Label4.Caption = "Total Emails In List: " & List1.ListCount
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*

End Sub
Public Sub List_Add(list As listbox, txt As String)
On Error Resume Next
    List1.AddItem txt
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

Public Sub List_Save(thelist As listbox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.list(Save)
    Next Save
    Close fFile
End Sub
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

End Function
