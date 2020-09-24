VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmanswers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Answers"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmanswers.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   4320
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "frmanswers.frx":0442
      Left            =   0
      List            =   "frmanswers.frx":0444
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   5055
   End
End
Attribute VB_Name = "frmanswers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim x As String
x = MsgBox("Would you like to save changes?", vbYesNo + vbExclamation)
If x = vbYes Then
    cd.Filter = "List Files (*.LST)|*.lst|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    cd.ShowSave
    List_Save List1, cd.FileName
End If
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
cd.Filter = "List Files (*.LST)|*.lst|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
cd.ShowSave
List_Save List1, cd.FileName
End Sub

Private Sub Command3_Click()
On Error Resume Next
List_Print List1
End Sub

Public Sub List_Load(TheList As ListBox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile


    Do
        Line Input #fFile, TheContents$
        'Call List_Add(TheList, TheContents$)
        TheList.AddItem TheContents$
    Loop Until EOF(fFile)
    Close fFile
End Sub


Public Sub List_Save(TheList As ListBox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile


    For Save = 0 To TheList.ListCount - 1
        Print #fFile, TheList.List(Save)
    Next Save
    Close fFile
End Sub

Public Sub List_Print(TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Printer.FontSize = 12


    For SaveList& = 0 To TheList.ListCount - 1
        Printer.Print TheList.List(SaveList&)
    Next SaveList&
    Printer.EndDoc
End Sub

