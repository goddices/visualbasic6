VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SetPer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "管理员设置"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "SetPer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   2160
      TabIndex        =   12
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtOkPass 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtPass 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   270
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "管理员列表"
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
      Begin MSComctlLib.ListView Lv 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "管理员"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "=>单击右键显示菜单"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   2400
         TabIndex        =   16
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "=>双击列表可修改设置"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   2400
         TabIndex        =   15
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "=>输入各项之后按保存"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "添加设置管理员："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "重复"
      Height          =   180
      Index           =   2
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "口令"
      Height          =   180
      Index           =   1
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   360
   End
   Begin VB.Menu MainMnu 
      Caption         =   "MainMnu"
      Visible         =   0   'False
      Begin VB.Menu EditMnu 
         Caption         =   "修改"
      End
      Begin VB.Menu DeleteMnu 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "SetPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim Rec As Integer
Dim StrFlag As String
Dim i As Integer
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If StrFlag = "修改" Then
    rst.Seek "=", Lv.SelectedItem.Text
    If txtName.Text = "" Or txtPass.Text = "" Or txtOkPass = "" Then
    MsgBox "请将所有信息填写完整！", 0 + 16, "提示"
    Exit Sub
    End If
    If txtPass.Text <> txtOkPass.Text Then
    MsgBox "密码不相同！", 0 + 16, "密码"
    txtOkPass.SetFocus
    Exit Sub
    End If
    rst.Edit
    rst.Fields("名称") = txtName.Text
    rst.Fields("密码") = txtPass.Text
    rst.Update
    Disp
    StrFlag = ""
    MsgBox "修改成功！", 0 + 48, "提示"
Else

If txtName.Text = "" Or txtPass.Text = "" Or txtOkPass = "" Then
    MsgBox "请将所有信息填写完整！", 0 + 16, "提示"
    Exit Sub
End If
If txtPass.Text <> txtOkPass.Text Then
    MsgBox "密码不相同！", 0 + 16, "密码"
    txtOkPass.SetFocus
    Exit Sub
End If
rst.AddNew
rst.Fields("名称") = txtName.Text
rst.Fields("密码") = txtPass.Text
rst.Update
Disp
StrFlag = ""
MsgBox "添加成功！", 0 + 48, "提示"
End If
txtName.Text = ""
txtPass.Text = ""
txtOkPass.Text = ""

End Sub

Private Sub DeleteMnu_Click()
Dim Str As String
If Lv.SelectedItem.Text = "超级用户" Then
    MsgBox "超级用户不能删除!", 0 + 16, "错误"
    Exit Sub
End If
rst.Seek "=", Lv.SelectedItem.Text
Str = "确实要删除 " & Lv.SelectedItem.Text & "吗？"
If MsgBox(Str, 4 + 32, "删除") = vbYes Then
    rst.Delete
    Disp
End If
End Sub
Private Sub EditMnu_Click()
Lv_DblClick
End Sub

Private Sub Form_Load()
Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Pass", dbOpenTable)
rst.Index = "名称"
Disp
End Sub
Private Sub Disp()
Lv.ListItems.Clear
rst.MoveLast
Rec = rst.RecordCount
rst.MoveFirst
For i = 1 To Rec
    Lv.ListItems.Add i, , rst.Fields("名称")
    rst.MoveNext
    If rst.EOF Then Exit Sub
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
rst.Close
db.Close
End Sub

Private Sub Lv_DblClick()
If Lv.SelectedItem.Text = "超级用户" Then
    MsgBox "超级用户不能修改！", 0 + 16, "错误"
    Exit Sub
End If
StrFlag = "修改"
txtName.Text = Lv.SelectedItem.Text
End Sub

Private Sub Lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu MainMnu
End If
End Sub

Private Sub txtName_Change()
If txtName.Text <> "" Then
    cmdSave.Enabled = True
Else
    cmdSave.Enabled = False
End If
End Sub
