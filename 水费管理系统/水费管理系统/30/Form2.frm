VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00404000&
   Caption         =   "登录窗口"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   4830
   Begin VB.TextBox Text5 
      DataField       =   "qx"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "password"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "user"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\水费管理系统\user.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "user"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退  出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "进  入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
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
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1200
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.DL.Caption = "注销"
SQL = "select * from user where user ='" & Trim(Text1.Text) & "'"
Data1.RecordSource = SQL
Data1.Refresh
If Data1.Recordset.EOF Then
   MsgBox "没有此用户！", , "提示"
   Data1.RecordSource = "user"
   Data1.Refresh
   
Else
   MDIForm1.StatusBar1.Panels(1).Text = "用户名：" & Trim(Text1.Text)
   yfm = Trim(Text1.Text)
   If Trim(Text2.Text) = Trim(Text4.Text) Then
     qxqx = Text5.Text
    If qxqx = 2 Then
       MDIForm1.sfgl.Enabled = True
       MDIForm1.yfgl.Enabled = True
       MDIForm1.DYFW.Enabled = True
       MDIForm1.hjsz.Enabled = True
       Else
         If qxqx = 3 Then
            MDIForm1.sfgl.Enabled = True
            MDIForm1.DYFW.Enabled = True
            Else
               MDIForm1.mmxg.Enabled = True
               MDIForm1.sfgl.Enabled = True
               MDIForm1.yfgl.Enabled = True
               MDIForm1.DYFW.Enabled = True
               MDIForm1.hjsz.Enabled = True
        End If
    End If
    Unload Me
   Else
      MsgBox "密码错误！", , "提示"
   End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Form11.Width = 4950
Form11.Height = 3810
Form11.Move (MDIForm1.Width - Form11.Width) / 2, (MDIForm1.Height - Form11.Height) / 4
End Sub
