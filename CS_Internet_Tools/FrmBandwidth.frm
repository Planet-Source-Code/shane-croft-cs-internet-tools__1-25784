VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmBandwidth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bandwidth Monitor"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   8865
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2295
      Left            =   3840
      OleObjectBlob   =   "FrmBandwidth.frx":0000
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   1575
      Left            =   15
      OleObjectBlob   =   "FrmBandwidth.frx":2024
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   4320
   End
   Begin MSChart20Lib.MSChart MSChart3 
      Height          =   1575
      Left            =   4440
      OleObjectBlob   =   "FrmBandwidth.frx":3ADA
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chart Options"
      Height          =   615
      Left            =   7200
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "2D"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "3D"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   0
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Bandwidth Monitor by Crofts Software"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5055
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6405
      TabIndex        =   33
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5055
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6405
      TabIndex        =   31
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   1815
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   1815
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4995
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4995
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "0 KB"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   13
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0 KB"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   12
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Estimated Upload Speed"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   11
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Estimated Download Speed"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   10
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   915
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   915
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   75
      X2              =   2115
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   75
      X2              =   2115
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      X1              =   2115
      X2              =   2115
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   2115
      X2              =   3795
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      X1              =   3795
      X2              =   3795
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line Line7 
      X1              =   75
      X2              =   3795
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      X1              =   75
      X2              =   3795
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   915
      X2              =   915
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3315
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3315
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "FrmBandwidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_objIpHelper As CIpHelper
Private TransferRate                    As Single
Private TransferRate2                   As Single
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Set m_objIpHelper = New CIpHelper
DoEvents
MDIFrmMain.Timer1.Enabled = False
Call UpdateInterfaceInfo
DoEvents
Me.Timer1.Enabled = True
Me.Icon = MDIFrmMain.ImageList2.ListImages(1).Picture
MDIFrmMain.StatusBar1.Panels(4).Picture = MDIFrmMain.ImageList2.ListImages(1).Picture
DoEvents
Label5.Caption = Me.lblRecv.Caption
Label6.Caption = Me.lblSent.Caption
DoEvents
Timer2.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Timer1.Enabled = False
MDIFrmMain.Timer1.Enabled = True
End Sub

Private Sub Option1_Click()
If Me.Option1.Value = True Then
MSChart1.chartType = VtChChartType2dBar
MSChart2.chartType = VtChChartType2dArea
MSChart3.chartType = VtChChartType2dArea
End If
If Me.Option2.Value = True Then
MSChart1.chartType = VtChChartType3dBar
MSChart2.chartType = VtChChartType3dArea
MSChart3.chartType = VtChChartType3dArea
End If
DoEvents
End Sub

Private Sub Option2_Click()
If Me.Option1.Value = True Then
MSChart1.chartType = VtChChartType2dBar
MSChart2.chartType = VtChChartType2dArea
MSChart3.chartType = VtChChartType2dArea
End If
If Me.Option2.Value = True Then
MSChart1.chartType = VtChChartType3dBar
MSChart2.chartType = VtChChartType3dArea
MSChart3.chartType = VtChChartType3dArea
End If
DoEvents
End Sub

Private Sub Timer1_Timer()
Call UpdateInterfaceInfo
End Sub
Private Sub UpdateInterfaceInfo()
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
If blnIsRecv And blnIsSent Then
Me.Icon = MDIFrmMain.ImageList2.ListImages(4).Picture
ElseIf (Not blnIsRecv) And blnIsSent Then
Me.Icon = MDIFrmMain.ImageList2.ListImages(3).Picture
ElseIf blnIsRecv And (Not blnIsSent) Then
Me.Icon = MDIFrmMain.ImageList2.ListImages(2).Picture
ElseIf Not (blnIsRecv And blnIsSent) Then
Me.Icon = MDIFrmMain.ImageList2.ListImages(1).Picture
End If
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
MSChart1.Visible = True
MSChart2.Visible = True
MSChart3.Visible = True
Me.Frame1.Visible = True
Option1.Visible = True
Option2.Visible = True
DoEvents
Dim XX As Long
Dim YY As Long
Dim XXX As Long
Dim YYY As Long
YYY = Label6.Caption
YY = Label5.Caption
DoEvents
XX = Me.lblRecv.Caption - YY
XXX = Me.lblSent.Caption - YYY
DoEvents
TransferRate = Format(Int(XX) / 1024, "####.00")
DoEvents
TransferRate2 = Format(Int(XXX) / 1024, "####.00")
DoEvents
 With MSChart1
         'To choose the type of the Chart

         'Specifying the title location

         'Determining the number of columns and rows
         
            .RowCount = 1
            .ColumnCount = 3
         'One record is one data series.
         ' Each data series is a collection of columns
         'Each data is assigned to a particular row and column with .data
         
            For Row = 1 To 1
            For Column = 1 To 3
           
           .Column = Column
           .Row = Row
           If Column = 1 Then
                .ColumnLabel = "Sent - " & XXX
                Label10.Caption = TransferRate2 & " KB"
                DoEvents
                .Data = XXX
            End If
           If Column = 2 Then
                .ColumnLabel = "Received - " & XX
                Label9.Caption = TransferRate & " KB"
                DoEvents
                .Data = XX
            End If
           If Column = 3 Then
                .Data = 100000
            End If
            Next Column
            Next Row
        
 
        End With

 With MSChart2
         
            .RowCount = 8
            .ColumnCount = 1
        
                Label11(0).Caption = Label12(0).Caption
                DoEvents
                .Row = 1
                .Data = Label11(0).Caption
                
                Label12(0).Caption = Label13(0).Caption
                DoEvents
                .Row = 2
                .Data = Label12(0).Caption

                Label13(0).Caption = Label14(0).Caption
                DoEvents
                 .Row = 3
                .Data = Label13(0).Caption

                Label14(0).Caption = Label11(1).Caption
                DoEvents
                .Row = 4
                .Data = Label14(0).Caption
                DoEvents
                
                Label11(1).Caption = Label12(1).Caption
                DoEvents
                .Row = 5
                .Data = Label11(1).Caption
                
                Label12(1).Caption = Label13(1).Caption
                DoEvents
                .Row = 6
                .Data = Label12(1).Caption

                Label13(1).Caption = Label14(1).Caption
                DoEvents
                 .Row = 7
                .Data = Label13(1).Caption

                Label14(1).Caption = XXX
                DoEvents
                .Row = 8
                .Data = Label14(1).Caption
                DoEvents
 
        End With

 With MSChart3
         
            .RowCount = 8
            .ColumnCount = 1
        
                Label15(0).Caption = Label16(0).Caption
                DoEvents
                .Row = 1
                .Data = Label15(0).Caption
                
                Label16(0).Caption = Label17(0).Caption
                DoEvents
                .Row = 2
                .Data = Label16(0).Caption

                Label17(0).Caption = Label18(0).Caption
                DoEvents
                 .Row = 3
                .Data = Label17(0).Caption

                Label18(0).Caption = Label15(1).Caption
                DoEvents
                .Row = 4
                .Data = Label18(0).Caption
                DoEvents
 
                Label15(1).Caption = Label16(1).Caption
                DoEvents
                .Row = 5
                .Data = Label15(1).Caption
                
                Label16(1).Caption = Label17(1).Caption
                DoEvents
                .Row = 6
                .Data = Label16(1).Caption

                Label17(1).Caption = Label18(1).Caption
                DoEvents
                 .Row = 7
                .Data = Label17(1).Caption

                Label18(1).Caption = XXX
                DoEvents
                .Row = 8
                .Data = Label18(1).Caption
                DoEvents
 
        End With
        
    DoEvents
    Label5.Caption = Me.lblRecv.Caption
    Label6.Caption = Me.lblSent.Caption
    DoEvents

End Sub

