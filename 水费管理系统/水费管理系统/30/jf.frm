VERSION 5.00
Begin VB.Form Form51 
   BackColor       =   &H00404000&
   Caption         =   "环境设置"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   6135
   Begin VB.Data Data1 
      Caption         =   "   价格表"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\water.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "当前价格"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "应缴月份"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "格式：yyyymm"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "jg"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label3 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
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
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   1275
   End
End
Attribute VB_Name = "Form51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Form51.Width = 6255
Form51.Height = 3360
Form51.Move (MDIForm1.Width - Form51.Width) / 2, (MDIForm1.Height - Form51.Height) / 4
End Sub


