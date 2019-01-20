VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5940
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1215
      Left            =   3120
      TabIndex        =   16
      Top             =   1560
      Width           =   2535
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   9
         Left            =   2040
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "余"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Question
End Sub

Private Sub Form_Load()
    Dim i  As Integer
    For Each cmd In Command1
    cmd.Caption = i
    i = i + 1
    Next
End Sub


Private Sub Question()
    Dim FirstOperand As Integer
    Dim SecondOperand As Integer
    Dim Operator_N As Integer
    Dim Operator_S As String
    Dim QBody As String
    Dim Answer As Integer
    Dim Remainder As Integer
    Dim blank As String
    
    blank = String(2, Chr(32))
    
    Select Case Operator_N
        Case 0
            Operator_S = "┼"
        Case 1
            Operator_S = "—"
        Case 2
            Operator_S = "×"
        Case 3
            Operator_S = "÷"
    End Select
    
10: FirstOperand = Int(Rnd * 200)
20: SecondOperand = Int(Rnd * 200) + 1
    Operator_N = Int(Rnd * 4)
    


    If (Operator_N = 1 And SecondOperand < FirstOperand) Then GoTo 10 '小数减大数
    If (Operator_N = 3 And secondeoperand = 0) Then GoTo 20 '除数为零
    
    QBody = FirstOperand & blank & Operator_S & blank & SecondOperand
    
End Sub
