VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EditBookId 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�༭����֤"
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
   StartUpPosition =   2  '��Ļ����
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
               Caption         =   "���"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ر�"
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
         Name            =   "����"
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
         Caption         =   "���(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditMnu 
         Caption         =   "�޸�(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu DeleteMnu 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu SearchMnu 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu ShowAllMnu 
         Caption         =   "��ʾ������Ա"
         Shortcut        =   {F3}
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "�˳�(&X)"
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
rst.Fields("����֤��") = mAddEditId & vbNullString
rst.Fields("����") = mAddEditIdName & vbNullString
rst.Fields("�༶") = mAddEditIdClass & " "
rst.Fields("����") = mAddEditIdDepart & " "
rst.Fields("ְ��") = mAddEditIdZhi & " "
rst.Update
DispId
mSave = False
If MsgBox("�ѳɹ���ӣ�Ҫ������Ӱ��س�������ȡ����", 4 + 32, "��ӳɹ�") = vbYes Then
    GoTo loop1
Else
    Exit Sub
End If
End If
End Sub
Private Sub cmdDelete_Click()
Dim St As String
St = "ȷʵҪɾ�� " & mLv.SelectedItem.Text & " " & mLv.SelectedItem.SubItems(1) & " ��"
If MsgBox(St, 4 + 32, "ɾ��") = vbYes Then
    rst.Seek "=", mLv.SelectedItem.Text
    rst.Delete
    DispId
End If
End Sub
Private Sub cmdEdit_Click()
Dim i As Integer
i = mLv.SelectedItem.Index
rst.Seek "=", mLv.SelectedItem.Text
mAddEditId = rst.Fields("����֤��") & vbNullString
mAddEditIdName = rst.Fields("����") & vbNullString
mAddEditIdClass = rst.Fields("�༶") & vbNullString
mAddEditIdDepart = rst.Fields("����") & vbNullString
mAddEditIdZhi = rst.Fields("ְ��") & vbNullString
AEditBookId.Show (1)
If mSave Then
    rst.Edit
    rst.Fields("����֤��") = mAddEditId & vbNullString
    rst.Fields("����") = mAddEditIdName & vbNullString
    rst.Fields("�༶") = mAddEditIdClass & " "
    rst.Fields("����") = mAddEditIdDepart & " "
    rst.Fields("ְ��") = mAddEditIdZhi & " "
    rst.Update
    With mLv.ListItems(i)
        .SubItems(1) = rst.Fields("����")
        .SubItems(2) = rst.Fields("�༶")
        .SubItems(3) = rst.Fields("����")
        .SubItems(4) = rst.Fields("ְ��")
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
        MsgBox "û���ҵ�ƥ���¼��", 0 + 48, "����ʧ��"
        Exit Sub
    End If
    mLv.ListItems.Clear
    mLv.ListItems.Add , , rst.Fields("����֤��")
    With mLv.ListItems(1)
        .SubItems(1) = rst.Fields("����")
        .SubItems(2) = rst.Fields("�༶")
        .SubItems(3) = rst.Fields("����")
        .SubItems(4) = rst.Fields("ְ��")
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
rst.Index = "����֤��"

mLv.View = lvwReport
mLv.GridLines = True

mLv.ColumnHeaders.Add , , "����֤��"
mLv.ColumnHeaders.Add , , "����"
mLv.ColumnHeaders.Add , , "�༶"
mLv.ColumnHeaders.Add , , "����"
mLv.ColumnHeaders.Add , , "ְ��"
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
    mLv.ListItems.Add i, , rst.Fields("����֤��")
    With mLv.ListItems(i)
        .SubItems(1) = rst.Fields("����") & vbNullString
        .SubItems(2) = rst.Fields("�༶") & vbNullString
        .SubItems(3) = rst.Fields("����") & vbNullString
        .SubItems(4) = rst.Fields("ְ��") & vbNullString
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
