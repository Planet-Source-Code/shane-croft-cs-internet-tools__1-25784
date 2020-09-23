VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIFrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "CS Internet Tools"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9615
   Icon            =   "MDIFrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   840
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":578A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":74C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":9202
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":AC8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":C9B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":E6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":103FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1211E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":13E5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   6510
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   11483
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Domain Name Lookup"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mass Email Sender"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "NetStat"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Online/Offline Checker"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ping"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Port Listener"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Port Scanner"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Resolve Host or IP"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Time Sync"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trace Route"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Winsock && Internet Connection Information"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6510
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10689
            MinWidth        =   9948
            Text            =   "© Crofts Software - Networking Software & More"
            TextSave        =   "© Crofts Software - Networking Software & More"
            Object.ToolTipText     =   "© Crofts Software - Networking Software && More"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   609
            MinWidth        =   617
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":15B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":15CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":15E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":15F92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuTools 
      Caption         =   "&Tools"
      Begin VB.Menu MenuDomain 
         Caption         =   "Domain Name Lookup"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuMass 
         Caption         =   "Mass Email Sender"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuNetstat 
         Caption         =   "NetStat"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuOOChecker 
         Caption         =   "Online/Offline Checker"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuPing 
         Caption         =   "Ping"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuPortListen 
         Caption         =   "Port Listener"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuPortScan 
         Caption         =   "Port Scanner"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuResolve 
         Caption         =   "Resolve Host or IP"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenuTime 
         Caption         =   "Time Sync"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MenuTrace 
         Caption         =   "Trace Route"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MenuWinsock 
         Caption         =   "Winsock && Internet Connection Information"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu Menusettings 
      Caption         =   "&Settings"
      Begin VB.Menu MenuSound 
         Caption         =   "Sounds"
         Begin VB.Menu MenuStartupSound 
            Caption         =   "Startup"
            Checked         =   -1  'True
         End
         Begin VB.Menu MenuShutDownSound 
            Caption         =   "Shutdown"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu menuBand 
         Caption         =   "Bandwidth Monitor"
         Begin VB.Menu MenBandwith 
            Caption         =   "Enable"
            Checked         =   -1  'True
         End
         Begin VB.Menu M 
            Caption         =   "View"
         End
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuWeb 
         Caption         =   "Web Page"
      End
      Begin VB.Menu MenuBug 
         Caption         =   "Bug Report"
      End
   End
End
Attribute VB_Name = "MDIFrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private FilePathName As String
Private m_objIpHelper As CIpHelper
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub UpdateInterfaceInfo()
On Error Resume Next
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)

Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
If blnIsRecv And blnIsSent Then
StatusBar1.Panels(4).Picture = ImageList2.ListImages(4).Picture
ElseIf (Not blnIsRecv) And blnIsSent Then
StatusBar1.Panels(4).Picture = ImageList2.ListImages(3).Picture
ElseIf blnIsRecv And (Not blnIsSent) Then
StatusBar1.Panels(4).Picture = ImageList2.ListImages(2).Picture
ElseIf Not (blnIsRecv And blnIsSent) Then
StatusBar1.Panels(4).Picture = ImageList2.ListImages(1).Picture
End If
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
StatusBar1.Panels(4).ToolTipText = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & "  Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
End Sub

Private Sub M_Click()
FrmBandwidth.Show
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
Set m_objIpHelper = New CIpHelper
AppDir = App.Path

FilePathName = AppDir + "\Settings.inf"
startupsound = GetPrivateProfileString("settings", "startupsound", "", FilePathName)
shutdownsound = GetPrivateProfileString("settings", "shutdownsound", "", FilePathName)
bandwidth = GetPrivateProfileString("settings", "bandwidth", "", FilePathName)

DoEvents
Me.MenuStartupSound.Checked = startupsound
Me.MenuShutDownSound.Checked = shutdownsound
Me.MenBandwith.Checked = bandwidth

MDIFrmMain.Caption = "CS Internet Tools v" & App.Major & "." & App.Minor & "." & App.Revision
StatusBar1.Panels(2).Text = Me.Winsock1.LocalHostName
StatusBar1.Panels(3).Text = Me.Winsock1.LocalIP
StatusBar1.Panels(2).ToolTipText = "Current Local Computer Name"
StatusBar1.Panels(3).ToolTipText = "Current Local IP Address"
StatusBar1.Panels(4).Picture = ImageList2.ListImages(1).Picture
DoEvents
FrmResolve.Show
FrmResolve.Hide
DoEvents
DoEvents
If Me.MenBandwith.Checked = True Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
DoEvents
 Dim wavSetup As String
 If Me.MenuStartupSound.Checked = True Then
 wavSetup = NoiseGet(App.Path & "\Sounds\" & "startup.wav")
 NoisePlay wavSetup, SND_SYNC
 End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Timer1.Enabled = False
Dim fFile As Integer
Dim wavSetup As String
fFile = FreeFile
 
Open App.Path & "\Settings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "startupsound=" & Me.MenuStartupSound.Checked
Print #fFile, "shutdownsound=" & Me.MenuShutDownSound.Checked
Print #fFile, "bandwidth=" & Me.MenBandwith.Checked
Close fFile
DoEvents
DoEvents
 If Me.MenuShutDownSound.Checked = True Then
 wavSetup = NoiseGet(App.Path & "\Sounds\" & "shutdown.wav")
 NoisePlay wavSetup, SND_SYNC
  DoEvents
  DoEvents
End
Else
End
End If
End Sub

Private Sub MenBandwith_Click()
If Me.MenBandwith.Checked = True Then
Me.MenBandwith.Checked = False
Me.Timer1.Enabled = False
StatusBar1.Panels(4).Picture = ImageList2.ListImages(1).Picture
Else
Me.MenBandwith.Checked = True
StatusBar1.Panels(4).Picture = ImageList2.ListImages(1).Picture
Me.Timer1.Enabled = True
End If
End Sub

Private Sub MenuAbout_Click()
frmAbout.Show
End Sub

Private Sub MenuBug_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "mailto:webmaster@croftssoftware.com", "", App.Path, 1)
End Sub

Private Sub MenuDomain_Click()
frminternetdomain.Show
End Sub

Private Sub MenuExit_Click()
On Error Resume Next
Timer1.Enabled = False
Dim fFile As Integer
Dim wavSetup As String
fFile = FreeFile
 
Open App.Path & "\Settings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "startupsound=" & Me.MenuStartupSound.Checked
Print #fFile, "shutdownsound=" & Me.MenuShutDownSound.Checked
Print #fFile, "bandwidth=" & Me.MenBandwith.Checked
Close fFile
DoEvents
DoEvents
 If Me.MenuShutDownSound.Checked = True Then
 wavSetup = NoiseGet(App.Path & "\Sounds\" & "shutdown.wav")
 NoisePlay wavSetup, SND_SYNC
  DoEvents
  DoEvents
End
Else
End
End If
End Sub

Private Sub MenuMass_Click()
FrmMassEmail.Show
End Sub

Private Sub MenuNetstat_Click()
FrmNetstat.Show
End Sub

Private Sub MenuOOChecker_Click()
FrmOnline.Show
End Sub

Private Sub MenuPing_Click()
FrmPing.Show
End Sub

Private Sub MenuPortListen_Click()
FrmPortListen.Show
End Sub

Private Sub MenuPortScan_Click()
FrmPortScanner.Show
End Sub

Private Sub MenuResolve_Click()
FrmResolve.Show
End Sub

Private Sub MenuShutDownSound_Click()
If Me.MenuShutDownSound.Checked = True Then
Me.MenuShutDownSound.Checked = False
Else
Me.MenuShutDownSound.Checked = True
End If
End Sub

Private Sub MenuStartupSound_Click()
If Me.MenuStartupSound.Checked = True Then
Me.MenuStartupSound.Checked = False
Else
Me.MenuStartupSound.Checked = True
End If
End Sub

Private Sub MenuTime_Click()
FrmTime.Show
End Sub

Private Sub MenuTrace_Click()
FrmTrace.Show
End Sub

Private Sub MenuWeb_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)
End Sub

Private Sub MenuWinsock_Click()
FrmWinsockInfo.Show
End Sub
Private Sub Timer1_Timer()
If Me.MenBandwith.Checked = True Then
Call UpdateInterfaceInfo
Else
Timer1.Enabled = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Index

Case 1
frminternetdomain.Show
Case 2
FrmMassEmail.Show
Case 3
FrmNetstat.Show
Case 4
FrmOnline.Show
Case 5
FrmPing.Show
Case 6
FrmPortListen.Show
Case 7
FrmPortScanner.Show
Case 8
FrmResolve.Show
Case 9
FrmTime.Show
Case 10
FrmTrace.Show
Case 11
FrmWinsockInfo.Show
End Select
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
