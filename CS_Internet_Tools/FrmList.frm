VERSION 5.00
Begin VB.Form FrmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "FrmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4500
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   1440
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save && Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Selected"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3120
      Picture         =   "FrmList.frx":1D2A
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Total Entries In List:"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub List_Add(List As listbox, txt As String)
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
        Print #fFile, List1.List(Save)
    Next Save
    Close fFile
End Sub
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please enter a Computer Name,Web Site Address, or a IP"
FrmList.Text1.SetFocus
Exit Sub
End If
List1.AddItem Text1.Text
DoEvents
Text1.Text = ""
FrmList.Text1.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
Call List_Save(List1, App.Path & "\List.ini")
DoEvents
FrmOnline.Form_Load
DoEvents
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Call List_Load(List1, App.Path & "\List.ini")
DoEvents
FrmList.Text1.SetFocus
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Text1_LostFocus()
On Error Resume Next
Text1.Text = Replace(Text1.Text, " ", "", 1, , vbTextCompare)
End Sub

Private Sub Timer1_Timer()
Label2.Caption = List1.ListCount
End Sub
