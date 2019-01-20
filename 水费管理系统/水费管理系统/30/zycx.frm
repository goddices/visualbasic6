VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form22 
   BackColor       =   &H00404000&
   Caption         =   "查询缴费情况"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "选择查询项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404000&
         Caption         =   "总户号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404000&
         Caption         =   "户  名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404000&
         Caption         =   "缴费日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查  询"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21430273
      CurrentDate     =   38516
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "zycx.frx":0000
      Height          =   3495
      Left            =   360
      OleObjectBlob   =   "zycx.frx":0014
      TabIndex        =   0
      Top             =   2280
      Width           =   9615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\water.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "水费管理"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1(0).Value = True Then
   SQL = "select * from 水费管理 where 总户号='" & Trim(Text1.Text) & "'"
   Else
   If Option1(1).Value = True Then
      SQL = "select * from 水费管理 where 户名='" & Trim(Text1.Text) & "'"
      Else
      If Option1(2).Value = True Then
         SQL = "select * from 水费管理 where 缴费日期='" & Format(DTPicker1.Value, "yyyy-mm-dd") & "'"
      End If
   End If
End If
Data1.RecordSource = SQL
Data1.Refresh
If Data1.Recordset.EOF Then
   MsgBox "没有您要查询的缴纳水费情况！", , "提示"
End If
End Sub
'Download by http://down.liehuo.net
Private Sub Form_Load()
Form22.Width = 10320
Form22.Height = 6525
Form22.Move (MDIForm1.Width - Form22.Width) / 2, (MDIForm1.Height - Form22.Height) / 4
DTPicker1.Value = Date
End Sub

Private Sub Option1_Click(Index As Integer)
For i = 0 To 2
    If Option1(0).Value = True Or Option1(1).Value = True Then
        Text1.Visible = True
        DTPicker1.Visible = False
        Else
          If Option1(2).Value = True Then
             Text1.Visible = False
             DTPicker1.Visible = True
          Else
             MsgBox "请选择查询的项！", , "提示"
          End If
    End If
Next i
Command1.Visible = True
End Sub
