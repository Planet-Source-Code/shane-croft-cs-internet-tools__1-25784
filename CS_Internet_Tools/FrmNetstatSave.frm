VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmNetstatSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save NetStat List"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "FrmNetstatSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7410
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save To File"
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6240
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmNetstatSave.frx":1D12
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FrmNetstatSave"
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
        Call List_Add(List1, TheContents$)
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
On Error GoTo exitme
Dim FileName As String
CD1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
CD1.DefaultExt = "txt"
CD1.DialogTitle = "Select the destination file"
CD1.FileName = "NetStat.txt"
CD1.CancelError = True
CD1.ShowSave
FileName = CD1.FileName

Call List_Save(List1, FileName)
exitme:

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

