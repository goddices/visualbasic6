VERSION 5.00
Begin VB.Form LoginSys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "枫叶图书管理系统_管理员登录"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOkCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdOkCancel 
      Caption         =   "确定(&E)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox comPer 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "管理员"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   690
   End
End
Attribute VB_Name = "LoginSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Asc() As Integer

Dim pwd As String
Dim db As Database
Dim rst As Recordset
Dim Rec As Integer
Private Sub cmdOkCancel_Click(Index As Integer)
Dim i As Integer
i = 0
Select Case Index
    Case 0
        If txtPass.Text = "" Or comPer.Text = "" Then
            MsgBox "请选择用户名或输入密码!", 0 + 48, "提示"
            txtPass.SetFocus
            Exit Sub
        End If
        If Val(txtPass.Text) = pwd Then
            'MsgBox "正确"
            Mainfrm.Show
            Unload Me
        Else
            MsgBox "密码错误,请重试！", 0 + 16, "错误"
            txtPass.SetFocus
            Exit Sub
        End If
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
Dim i As Integer

Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Pass", dbOpenTable)

rst.MoveLast
Rec = rst.RecordCount
ReDim Asc(Rec - 1)
rst.MoveFirst
 
comPer.AddItem rst.Fields("名称")
'Asc(i - 1) = Val(rst.Fields("密码"))
pwd = CStr(rst.Fields("密码"))
rst.MoveNext
    
 
comPer.Text = ""
txtPass.Text = ""
End Sub

