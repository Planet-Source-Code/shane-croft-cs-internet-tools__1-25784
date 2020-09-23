VERSION 5.00
Begin VB.Form FrmMassEmailAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Email Address To List"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMassEmailAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Email Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMassEmailAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter A Valid Email Address", vbCritical
Exit Sub
End If

FrmMassEmail.List1.AddItem Me.Text1.Text
DoEvents
Call xListKillDupes(FrmMassEmail.List1) 'calls sub from module
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub
