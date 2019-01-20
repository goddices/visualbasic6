VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H00404000&
   Caption         =   "缴纳水费"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   7230
   Begin VB.TextBox Text8 
      DataField       =   "应缴月份"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "应缴月份"
      DataSource      =   "Data2"
      Height          =   270
      Left            =   960
      TabIndex        =   28
      Text            =   "Text7"
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text6 
      DataField       =   "总户号"
      DataSource      =   "Data3"
      Height          =   270
      Left            =   6120
      TabIndex        =   27
      Text            =   "Text6"
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "总户号"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   4440
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查  找"
      Height          =   375
      Left            =   4320
      TabIndex        =   25
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "总费用"
      DataSource      =   "Data3"
      Height          =   375
      Left            =   6120
      TabIndex        =   24
      Text            =   "Text5"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "地址"
      DataSource      =   "Data3"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "户名"
      DataSource      =   "Data3"
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\water.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "用户管理"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      DataField       =   "jg"
      DataSource      =   "Data2"
      Height          =   270
      Left            =   120
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\water.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "当前价格"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      DataField       =   "缴费日期"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   2400
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "当月水费"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2400
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "当前单价"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   2400
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "用水量"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2400
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "地址"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2400
      TabIndex        =   14
      Top             =   2468
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "户名"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   1834
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加入库"
      Height          =   345
      Left            =   4920
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   " 浏览"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\water.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "水费管理"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "吨"
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
      Index           =   5
      Left            =   4200
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "元/吨"
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
      Index           =   3
      Left            =   4320
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5400
      TabIndex        =   12
      Top             =   480
      Width           =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   1440
      X2              =   5160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   1440
      X2              =   5160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缴  纳  水  费"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1500
      TabIndex        =   11
      Top             =   60
      Width           =   2910
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缴  纳  水  费"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   2910
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   5175
      Left            =   360
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缴费日期："
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
      Index           =   9
      Left            =   1080
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当月费用："
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
      Index           =   8
      Left            =   1080
      TabIndex        =   6
      Top             =   5004
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前水价："
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
      Index           =   7
      Left            =   1080
      TabIndex        =   5
      Top             =   4370
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用水量："
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
      Index           =   6
      Left            =   1320
      TabIndex        =   4
      Top             =   3735
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应缴月份："
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
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   3102
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地  址："
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
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   2468
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "户  名："
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
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   1834
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总户号："
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
      Index           =   0
      Left            =   1335
      TabIndex        =   0
      Top             =   1200
      Width           =   1020
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
MsgBox "已入库！", , "提示"
Text5.Text = Val(Text5.Text) + Val(Text1(5).Text)
Data1.UpdateRecord
Text1(0).SetFocus
Text1(0).Text = ""
   For i = 1 To 9
     Label1(i).Visible = False
   Next i
   For i = 1 To 6
     Text1(i).Visible = False
   Next i
   Text8.Visible = False

End Sub


Private Sub Command2_Click()
SQL = "select * from 用户管理 where 总户号='" & Trim(Text1(0).Text) & "'"
Data3.RecordSource = SQL
Data3.Refresh
If Data3.Recordset.EOF Then
   MsgBox "没有此总户号！请重新输入[总户号]！", , "提示"
   Text1(0).Text = ""
   Text1(0).SetFocus
Else
   For i = 1 To 9
     Label1(i).Visible = True
   Next i
   For i = 1 To 6
     Text1(i).Visible = True
   Next i
   Text8.Visible = True
   Data1.Recordset.AddNew
   Text1(7) = Text6
   Text1(1) = Text3
   Text1(2) = Text4
   Text1(4) = Text2
   Text1(6).Text = Date$
   Text1(3).SetFocus
   Text8 = Text7
End If
End Sub



Private Sub Form_Load()
 Form21.Width = 7350
 Form21.Height = 7395
Form21.Move (MDIForm1.Width - Form21.Width) / 2, (MDIForm1.Height - Form21.Height) / 4
Label5.Caption = Date
End Sub

Private Sub Picture1_Click()
Picture1.Picture = Clipboard.GetData
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 3 Or Index = 4 Then
  Text1(5).Text = Val(Text1(3).Text) * Val(Text1(4).Text)
End If
End Sub
