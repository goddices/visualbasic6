VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EditBookId 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "编辑借书证"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   Icon            =   "EditBookId.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "edit "
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "delete"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "add"
      Height          =   255
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBookId.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBookId.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBookId.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBookId.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBookId.frx":1592
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   660
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1164
         ButtonWidth     =   1984
         ButtonHeight    =   1005
         Appearance      =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "添加"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "关闭"
               ImageIndex      =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin MSComctlLib.ListView mLv 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8493
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu PoMnu 
      Caption         =   "PoMnu"
      Visible         =   0   'False
      Begin VB.Menu AddMnu 
         Caption         =   "添加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditMnu 
         Caption         =   "修改(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu DeleteMnu 
         Caption         =   "删除(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu SearchMnu 
         Caption         =   "查找(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu ShowAllMnu 
         Caption         =   "显示所有人员"
         Shortcut        =   {F3}
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "EditBookId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim Rec As Integer

Private Sub AddMnu_Click()
cmdAdd_Click
End Sub

Private Sub cmdAdd_Click()
loop1:
AddBookId.Show (1)
If mSave Then
rst.AddNew
rst.Fields("借书证号") = mAddEditId & vbNullString
rst.Fields("姓名") = mAddEditIdName & vbNullString
rst.Fields("班级") = mAddEditIdClass & " "
rst.Fields("部门") = mAddEditIdDepart & " "
rst.Fields("职称") = mAddEditIdZhi & " "
rst.Update
DispId
mSave = False
If MsgBox("已成功添加，要继续添加按回车，否则按取消！", 4 + 32, "添加成功") = vbYes Then
    GoTo loop1
Else
    Exit Sub
End If
End If
End Sub
Private Sub cmdDelete_Click()
Dim St As String
St = "确实要删除 " & mLv.SelectedItem.Text & " " & mLv.SelectedItem.SubItems(1) & " 吗？"
If MsgBox(St, 4 + 32, "删除") = vbYes Then
    rst.Seek "=", mLv.SelectedItem.Text
    rst.Delete
    DispId
End If
End Sub
Private Sub cmdEdit_Click()
Dim i As Integer
i = mLv.SelectedItem.Index
rst.Seek "=", mLv.SelectedItem.Text
mAddEditId = rst.Fields("借书证号") & vbNullString
mAddEditIdName = rst.Fields("姓名") & vbNullString
mAddEditIdClass = rst.Fields("班级") & vbNullString
mAddEditIdDepart = rst.Fields("部门") & vbNullString
mAddEditIdZhi = rst.Fields("职称") & vbNullString
AEditBookId.Show (1)
If mSave Then
    rst.Edit
    rst.Fields("借书证号") = mAddEditId & vbNullString
    rst.Fields("姓名") = mAddEditIdName & vbNullString
    rst.Fields("班级") = mAddEditIdClass & " "
    rst.Fields("部门") = mAddEditIdDepart & " "
    rst.Fields("职称") = mAddEditIdZhi & " "
    rst.Update
    With mLv.ListItems(i)
        .SubItems(1) = rst.Fields("姓名")
        .SubItems(2) = rst.Fields("班级")
        .SubItems(3) = rst.Fields("部门")
        .SubItems(4) = rst.Fields("职称")
    End With
    'DispId
    mSave = False
End If
End Sub

Private Sub cmdSearch_Click()
SearchNum.Show (1)
If SearchFlag Then
    rst.Seek "=", BookBianHao
    If rst.NoMatch Then
        MsgBox "没有找到匹配记录！", 0 + 48, "查找失败"
        Exit Sub
    End If
    mLv.ListItems.Clear
    mLv.ListItems.Add , , rst.Fields("借书证号")
    With mLv.ListItems(1)
        .SubItems(1) = rst.Fields("姓名")
        .SubItems(2) = rst.Fields("班级")
        .SubItems(3) = rst.Fields("部门")
        .SubItems(4) = rst.Fields("职称")
    End With
    SearchFlag = False
End If
End Sub
Private Sub DeleteMnu_Click()
cmdDelete_Click
End Sub
Private Sub EditMnu_Click()
cmdEdit_Click
End Sub
Private Sub ExitMnu_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Personal", dbOpenTable)
rst.Index = "借书证号"

mLv.View = lvwReport
mLv.GridLines = True

mLv.ColumnHeaders.Add , , "借书证号"
mLv.ColumnHeaders.Add , , "姓名"
mLv.ColumnHeaders.Add , , "班级"
mLv.ColumnHeaders.Add , , "部门"
mLv.ColumnHeaders.Add , , "职称"
If rst.RecordCount <> 0 Then
DispId
End If
End Sub
Public Sub DispId()
Dim i As Integer
mLv.ListItems.Clear
rst.MoveLast
Rec = rst.RecordCount
rst.MoveFirst
For i = 1 To Rec
    mLv.ListItems.Add i, , rst.Fields("借书证号")
    With mLv.ListItems(i)
        .SubItems(1) = rst.Fields("姓名") & vbNullString
        .SubItems(2) = rst.Fields("班级") & vbNullString
        .SubItems(3) = rst.Fields("部门") & vbNullString
        .SubItems(4) = rst.Fields("职称") & vbNullString
    End With
    rst.MoveNext
    If rst.EOF Then Exit For
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
rst.Close
db.Close
End Sub

Private Sub mLv_DblClick()
cmdEdit_Click
End Sub

Private Sub mLv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu PoMnu
End If
End Sub

Private Sub SearchMnu_Click()
cmdSearch_Click
End Sub

Private Sub ShowAllMnu_Click()
DispId
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        cmdAdd_Click
    Case 2
        cmdEdit_Click
    Case 3
        cmdDelete_Click
    Case 4
        cmdSearch_Click
    Case 7
        Unload Me
    End Select
End Sub
