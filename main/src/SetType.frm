VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SetType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置图书类别和借出时间"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "SetType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton cmdSaveCancel 
         Caption         =   "取消(&C)"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveCancel 
         Caption         =   "保存(&S)"
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   12
         Top             =   2520
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   2535
      End
      Begin MSComCtl2.UpDown UpD 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   1560
         Width           =   225
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "comTime"
         BuddyDispid     =   196612
         OrigLeft        =   1920
         OrigTop         =   1440
         OrigRight       =   2145
         OrigBottom      =   1695
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox comTime 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "SetType.frx":0442
         Left            =   840
         List            =   "SetType.frx":045B
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtTypeName 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label labFlag 
         AutoSize        =   -1  'True
         Caption         =   "添加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1680
         TabIndex        =   16
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "规定借出时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "类别名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   3120
      ScaleHeight     =   3255
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "修改类别"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":047F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "关闭<=>返回"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":08C1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除旧类别"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":0D03
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加新类别"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":1145
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView Lv 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5530
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "图书类别"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "规定借出时间"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu MainMnu 
      Caption         =   "MainMnu"
      Visible         =   0   'False
      Begin VB.Menu AddMnu 
         Caption         =   "添加新类别(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditMnu 
         Caption         =   "编辑类别(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu DeleteMnu 
         Caption         =   "删除类别(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu ShowMnu 
         Caption         =   "显示所有类别(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "SetType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim Rec As Integer
Dim StrFlag As String
Dim Se As Integer

Private Sub AddMnu_Click()
cmdAdd_Click
End Sub

Private Sub cmdAdd_Click()
StrFlag = "添加"
labFlag.Caption = "添加状态"
txtTypeName = ""
comTime = ""
Lv.Visible = False
Picture2.Visible = True
cmdFlag (False)
End Sub
Private Sub cmdDelete_Click()
Dim St As String
rst.Seek "=", Lv.SelectedItem.Text
St = "确实要删除 " & Lv.SelectedItem.Text & " 类吗？"
If MsgBox(St, 4 + 32, "删除类别") = vbYes Then
    rst.Delete
    Disp
Else
    Exit Sub
End If
End Sub
Private Sub cmdEdit_Click()
StrFlag = "编辑"
labFlag.Caption = "修改状态"
Se = Lv.SelectedItem.Index
rst.Seek "=", Lv.SelectedItem.Text
txtTypeName.Text = rst.Fields("类别")
comTime.Text = rst.Fields("借出天数")
Picture2.Visible = True
Lv.Visible = False
cmdFlag (False)
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSaveCancel_Click(Index As Integer)
Select Case Index
    Case 0
    If StrFlag = "添加" Then
        If txtTypeName.Text = "" Or comTime.Text = "" Then
            MsgBox "请填写完整！", 0 + 48, "提示"
            Exit Sub
        End If
        rst.Seek "=", txtTypeName
        If rst.NoMatch = False Then
            MsgBox txtTypeName & " 类别已经存在，请填写其它类！", 0 + 48, "类别重复"
            txtTypeName.SetFocus
            Exit Sub
        End If
        rst.AddNew
        rst.Fields("类别") = txtTypeName.Text & vbNullString
        rst.Fields("借出天数") = comTime.Text & vbNullString
        rst.Update
        Picture2.Visible = False
        Lv.Visible = True
        Disp
        cmdFlag (True)
    ElseIf StrFlag = "编辑" Then
        If txtTypeName.Text = "" Or comTime.Text = "" Then
            MsgBox "请填写完整！", 0 + 48, "提示"
            Exit Sub
        End If
        rst.Edit
        rst.Fields("类别") = txtTypeName.Text & vbNullString
        rst.Fields("借出天数") = comTime.Text
        rst.Update
        Picture2.Visible = False
        Lv.Visible = True
        Disp
        cmdFlag (True)
    End If
    Case 1
        Picture2.Visible = False
        Lv.Visible = True
        cmdFlag (True)
End Select
End Sub

Private Sub DeleteMnu_Click()
cmdDelete_Click
End Sub

Private Sub EditMnu_Click()
cmdEdit_Click
End Sub

Private Sub ExitMnu_Click()
cmdExit_Click
End Sub

Private Sub Form_Load()
Lv.Visible = True
Picture2.Visible = False
Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Type", dbOpenTable)
rst.Index = "类别"
Disp
End Sub
Private Sub Disp()
Dim i As Integer
Lv.ListItems.Clear
rst.MoveLast
Rec = rst.RecordCount
rst.MoveFirst
For i = 1 To Rec
    Lv.ListItems.Add i, , rst.Fields("类别")
    Lv.ListItems(i).SubItems(1) = rst.Fields("借出天数")
    rst.MoveNext
    If rst.EOF Then Exit For
Next
End Sub
Private Sub cmdFlag(Bool As Boolean)
cmdAdd.Enabled = Bool
cmdDelete.Enabled = Bool
cmdExit.Enabled = Bool
cmdEdit.Enabled = Bool
End Sub

Private Sub Form_Unload(Cancel As Integer)
rst.Close
db.Close
End Sub

Private Sub Lv_DblClick()
cmdEdit_Click
End Sub

Private Sub Lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu MainMnu
End If
End Sub

Private Sub ShowMnu_Click()
Disp
End Sub
