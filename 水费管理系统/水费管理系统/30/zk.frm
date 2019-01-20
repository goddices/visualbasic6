VERSION 5.00
Begin VB.Form Form31 
   BackColor       =   &H00404000&
   Caption         =   "编辑用户"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   6975
   Begin VB.CommandButton Command1 
      Caption         =   "刷  新"
      Height          =   375
      Index           =   7
      Left            =   4320
      TabIndex        =   13
      Top             =   2880
      Width           =   1000
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\water.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "用户管理"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修  改"
      Height          =   375
      Index           =   6
      Left            =   3360
      TabIndex        =   12
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删  除"
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   11
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "增  加"
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   10
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "末  条"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   9
      Top             =   3720
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "前一条"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   8
      Top             =   3720
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一条"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   3720
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "首  条"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   6
      Top             =   3720
      Width           =   1000
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
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
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
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
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
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   1335
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
      Left            =   1575
      TabIndex        =   2
      Top             =   600
      Width           =   1020
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1230
      Width           =   1035
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
      Left            =   1560
      TabIndex        =   0
      Top             =   1875
      Width           =   1035
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
  Data1.Recordset.MoveFirst
  Command1(2).Enabled = False
  Command1(1).Enabled = True
End If
If Index = 1 Then
   Command1(2).Enabled = True
   Data1.Recordset.MoveNext
   If Data1.Recordset.EOF Then
      Data1.Recordset.MoveLast
      Command1(1).Enabled = False
   End If
End If
If Index = 2 Then
  Command1(1).Enabled = True
  Data1.Recordset.MovePrevious
  If Data1.Recordset.BOF Then
     Data1.Recordset.MoveFirst
     Command1(2).Enabled = False
  End If
End If
If Index = 3 Then
   Data1.Recordset.MoveLast
   Command1(1).Enabled = False
   Command1(2).Enabled = True
End If
If Index = 4 Then
   Data1.Recordset.AddNew
   Text1(0) = Data1.Recordset.RecordCount + 1
   Text1(1).SetFocus
  Command1(4).Enabled = False
  Command1(5).Enabled = False
  Command1(6).Enabled = False
End If
If Index = 5 Then
   Data1.Recordset.Delete
   Data1.Recordset.MoveNext
   If Data1.Recordset.EOF Then
      Data1.Recordset.MoveLast
      Command1(1).Enabled = False
   End If
End If
If Index = 6 Then
  Data1.Recordset.Edit
  Command1(4).Enabled = False
  Command1(5).Enabled = False
  Command1(6).Enabled = False
End If
If Index = 7 Then
  Data1.UpdateRecord
  Data1.Recordset.MoveLast
  Command1(1).Enabled = False
  Command1(2).Enabled = True
  Command1(4).Enabled = True
  Command1(5).Enabled = True
  Command1(6).Enabled = True
End If
End Sub
'Download by http://down.liehuo.net
Private Sub Form_Load()
 Form31.Width = 7095
 Form31.Height = 5340
Form31.Move (MDIForm1.Width - Form31.Width) / 2, (MDIForm1.Height - Form31.Height) / 4
End Sub
