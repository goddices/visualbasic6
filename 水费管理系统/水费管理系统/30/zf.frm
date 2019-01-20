VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form78 
   Caption         =   "转房管理"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   7560
   Begin VB.CommandButton Command3 
      Caption         =   "放弃转房"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "主治医师"
      DataSource      =   "Data1"
      Height          =   270
      Left            =   1680
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "病房号"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   3570
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查找"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "zf.frx":0000
      Height          =   330
      Left            =   4080
      TabIndex        =   18
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "病房号"
      Text            =   "请选择"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6360
      Top             =   2400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\医院管理系统\doctor.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\医院管理系统\doctor.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "主治医师情况"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "zf.frx":0015
      Height          =   330
      Left            =   4080
      TabIndex        =   17
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "姓名"
      Text            =   "请选择"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认转房"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   16
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   " 浏览"
      Connect         =   "Access"
      DatabaseName    =   "C:\医院管理系统\doctor.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "住院患者情况"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "出院时病情"
      DataSource      =   "Data1"
      Height          =   630
      Index           =   14
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2715
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "入院时病情"
      DataSource      =   "Data1"
      Height          =   630
      Index           =   13
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2025
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "入院时间"
      DataSource      =   "Data1"
      Height          =   270
      Index           =   11
      Left            =   2640
      TabIndex        =   13
      Top             =   1605
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "年龄"
      DataSource      =   "Data1"
      Height          =   270
      Index           =   3
      Left            =   2640
      TabIndex        =   12
      Top             =   1260
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "性别"
      DataSource      =   "Data1"
      Height          =   270
      Index           =   2
      Left            =   2640
      TabIndex        =   11
      Top             =   930
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "姓名"
      DataSource      =   "Data1"
      Height          =   270
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   3960
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主治医师："
      Height          =   180
      Index           =   17
      Left            =   600
      TabIndex        =   8
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "病房号："
      Height          =   180
      Index           =   16
      Left            =   780
      TabIndex        =   7
      Top             =   3615
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出院时病情况："
      Height          =   180
      Index           =   14
      Left            =   1200
      TabIndex        =   6
      Top             =   2715
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入院时病情况："
      Height          =   180
      Index           =   13
      Left            =   1200
      TabIndex        =   5
      Top             =   2025
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入院时间："
      Height          =   180
      Index           =   11
      Left            =   1560
      TabIndex        =   4
      Top             =   1605
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年  龄："
      Height          =   180
      Index           =   3
      Left            =   1740
      TabIndex        =   3
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性  别："
      Height          =   180
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓  名："
      Height          =   180
      Index           =   1
      Left            =   1740
      TabIndex        =   1
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "住院号："
      Height          =   180
      Index           =   0
      Left            =   1740
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Form78"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = DataCombo2.Text
Text3.Text = DataCombo1.Text
End Sub

Private Sub Command2_Click()
sql = "select * from 住院患者情况 where 住院号='" & Trim(Text1(0).Text) & "'"
Data1.RecordSource = sql
Data1.Refresh
If Data1.Recordset.EOF Then
   MsgBox "住院号错！"
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Form31.Width = 7680
 Form31.Height = 5220
End Sub
