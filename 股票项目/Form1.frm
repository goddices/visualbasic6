VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "模拟炒股"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "现价修改完毕"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox TxFld 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Text            =   "在这里输入"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6 股票资本"
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5 现价"
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4 买入股数 "
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3 买入价"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 购买日期 "
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 买入股票"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label LbFld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 总资本"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FldStr() As String
Private i As Integer
Private FileTxt As String

Private Sub Command1_Click()
If IsNumeric(TxFld.Text) Then
FldStr(5) = "[现价]" + TxFld.Text
LbFld(5) = FldStr(5)
Close #1
Open App.Path + "\模拟炒股.txt" For Output As #1
Print #1, LbFld(5)
Close #1
End If
End Sub

Private Sub Command2_Click()
 x = Split(FileTxt, vbCrLf)
 MsgBox UBound(x)
End Sub

Private Sub Form_Activate()
Do While Not EOF(1)
i = i + 1
ReDim Preserve FldStr(i)
Line Input #1, FldStr(i)
'start = InStr(1, FldStr(i), "]")
LbFld(i).Caption = FldStr(i)
Loop
End Sub

Private Sub Form_Load()
Open App.Path + "\模拟炒股.txt" For Input As #1
Get #1, , FileTxt
i = -1
'
 
End Sub

Private Sub LblFld_Click(Index As Integer)

End Sub

Private Sub Label2_Click()

End Sub
