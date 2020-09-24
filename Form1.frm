VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pythagorean Triples"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   0
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Load ..."
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Process!"
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&About ..."
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variables"
      Height          =   2415
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3135
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "C Minimum:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "C Maximum:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "B Maximum:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "B Minimum:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "A Maximum:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "A Minimum:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Left            =   960
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "c"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "b"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "a ² + b ² = c ²"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This utility will find all the pythagorean triples."
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   1080
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   1080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   240
      Y1              =   240
      Y2              =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim amin As Integer
Dim amax As Integer
Dim bmin As Integer
Dim bmax As Integer
Dim cmin As Integer
Dim cmax As Integer

Private Sub Command1_Click()
MsgBox "PTF (Pythagorean Triples Finder) created and programmed by Yusuf Simonson.", vbInformation, "About"
End Sub

Private Sub Command2_Click()
Dim x As String
x = MsgBox("Are you sure you want to quit?")
If x = vbYes Then End
End Sub

Private Sub Command3_Click()
On Error GoTo err
amin = Text1.Text
amax = Text2.Text
bmin = Text3.Text
bmax = Text4.Text
cmin = Text5.Text
cmax = Text6.Text
process

Exit Sub
err:
    MsgBox "Either there are no variables or you have put an illegal character in one or more of the variables (numbers only accepted)", vbCritical
   
End Sub

Private Sub process()
frmanswers.List1.Clear
Dim x As Long
Dim y As Long
Dim z As Long
Dim results As Long
results = 0
x = amin
y = bmin
z = cmin
Do Until z = cmax + 1
    If x ^ 2 + y ^ 2 = z ^ 2 Then
        If x <= y Then
            frmanswers.List1.AddItem x & " ²" & " + " & y & " ²" & " = " & z & " ²"
            results = results + 1
            frmanswers.Label1.Caption = results & " results found."
        End If
    End If
    x = x + 1
    If x = amax + 1 Then
        x = amin
        y = y + 1
    End If
    If y = bmax + 1 Then
        y = bmin
        z = z + 1
    End If
Loop

frmanswers.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
With frmanswers
.cd.Filter = "List Files (*.LST)|*.lst|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
.cd.ShowOpen
.List1.Clear
.List_Load .List1, .cd.FileName
End With
frmanswers.Show
End Sub

