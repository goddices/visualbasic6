VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00404000&
   Caption         =   "帐户管理"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   4845
   Begin VB.TextBox Text3 
      DataField       =   "qx"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text3"
      ToolTipText     =   """1""超级用户""2""管理员""3""收费员"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   465
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   465
      Left            =   2040
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改"
      Height          =   465
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   " 帐户管理"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\user.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "user"
      Top             =   2280
      Width           =   2220
   End
   Begin VB.TextBox Text2 
      DataField       =   "password"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "user"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "权  限："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1020
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "新增" Then
   Command1.Caption = "确定"
   Data1.Recordset.AddNew
   Text2.SetFocus
   Command2.Enabled = False
   Command3.Enabled = False
Else
   Command1.Caption = "新增"
   Data1.Recordset.Update
   Data1.Recordset.MoveLast
   Command2.Enabled = True
   Command3.Enabled = True
End If
'Download by http://down.liehuo.net
End Sub

Private Sub Command2_Click()
Data1.Recordset.Delete
Data1.Recordset.MovePrevious
End Sub

Private Sub Command3_Click()
Data1.Recordset.Edit
Data1.Recordset.Update
End Sub

Private Sub Form_Load()
Form12.Width = 4965
Form12.Height = 4335
Form12.Move (MDIForm1.Width - Form12.Width) / 2, (MDIForm1.Height - Form12.Height) / 4

End Sub
