VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SetType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͼ�����ͽ��ʱ��"
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
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton cmdSaveCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveCancel 
         Caption         =   "����(&S)"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�涨���ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�޸����"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":047F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�ر�<=>����"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":08C1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ�������"
         Height          =   735
         Left            =   120
         Picture         =   "SetType.frx":0D03
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "��������"
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
         Text            =   "ͼ�����"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�涨���ʱ��"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu MainMnu 
      Caption         =   "MainMnu"
      Visible         =   0   'False
      Begin VB.Menu AddMnu 
         Caption         =   "��������(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditMnu 
         Caption         =   "�༭���(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu DeleteMnu 
         Caption         =   "ɾ�����(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu ShowMnu 
         Caption         =   "��ʾ�������(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "�˳�(&X)"
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
StrFlag = "���"
labFlag.Caption = "���״̬"
txtTypeName = ""
comTime = ""
Lv.Visible = False
Picture2.Visible = True
cmdFlag (False)
End Sub
Private Sub cmdDelete_Click()
Dim St As String
rst.Seek "=", Lv.SelectedItem.Text
St = "ȷʵҪɾ�� " & Lv.SelectedItem.Text & " ����"
If MsgBox(St, 4 + 32, "ɾ�����") = vbYes Then
    rst.Delete
    Disp
Else
    Exit Sub
End If
End Sub
Private Sub cmdEdit_Click()
StrFlag = "�༭"
labFlag.Caption = "�޸�״̬"
Se = Lv.SelectedItem.Index
rst.Seek "=", Lv.SelectedItem.Text
txtTypeName.Text = rst.Fields("���")
comTime.Text = rst.Fields("�������")
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
    If StrFlag = "���" Then
        If txtTypeName.Text = "" Or comTime.Text = "" Then
            MsgBox "����д������", 0 + 48, "��ʾ"
            Exit Sub
        End If
        rst.Seek "=", txtTypeName
        If rst.NoMatch = False Then
            MsgBox txtTypeName & " ����Ѿ����ڣ�����д�����࣡", 0 + 48, "����ظ�"
            txtTypeName.SetFocus
            Exit Sub
        End If
        rst.AddNew
        rst.Fields("���") = txtTypeName.Text & vbNullString
        rst.Fields("�������") = comTime.Text & vbNullString
        rst.Update
        Picture2.Visible = False
        Lv.Visible = True
        Disp
        cmdFlag (True)
    ElseIf StrFlag = "�༭" Then
        If txtTypeName.Text = "" Or comTime.Text = "" Then
            MsgBox "����д������", 0 + 48, "��ʾ"
            Exit Sub
        End If
        rst.Edit
        rst.Fields("���") = txtTypeName.Text & vbNullString
        rst.Fields("�������") = comTime.Text
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
rst.Index = "���"
Disp
End Sub
Private Sub Disp()
Dim i As Integer
Lv.ListItems.Clear
rst.MoveLast
Rec = rst.RecordCount
rst.MoveFirst
For i = 1 To Rec
    Lv.ListItems.Add i, , rst.Fields("���")
    Lv.ListItems(i).SubItems(1) = rst.Fields("�������")
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
