VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Findfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͼ��"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "Findfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2520
      ScaleHeight     =   735
      ScaleWidth      =   3135
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��ʵ�ֶ����¼�Ĳ���"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ʾ:������ѯ������*���������ַ�"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   3060
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4471
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   32768
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "��  ��(&C)"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "�رմ˶Ի���"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdKong 
      Caption         =   "ȫ�����(&L)"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "��������ı�"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBeginFind 
      Caption         =   "��ʼ����(&F)"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      ToolTipText     =   "��ʼ���ҷ��������ļ�¼"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtBookName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtBookBian 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   600
      ScaleHeight     =   2175
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   720
      Width           =   1815
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��   ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "Findfrm.frx":0442
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ͼ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   1035
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6600
      Picture         =   "Findfrm.frx":088C
      Top             =   1920
      Width           =   480
   End
End
Attribute VB_Name = "Findfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst1 As Recordset '�򿪱�Book
Dim rst2 As Recordset '�򿪱�BookFf
Dim rst As Recordset
Dim db1 As Database
Dim db2 As Database
Dim qry1 As QueryDef
Dim qry2 As QueryDef
Dim RecNum As Integer '���ҷ��������ܼ�¼��
Dim i As Integer
Dim FindStr As String  '����SQL���
Private Sub cmdBeginFind_Click()
If txtBookBian = "" And txtBookName = "" Then
    MsgBox "����д��ز�����Ϣ��", 0 + 48, "��ʾ"
    txtBookBian.SetFocus
    Exit Sub
End If
Lv.ListItems.Clear
Findfrm.MousePointer = 11
If txtBookBian <> "" And txtBookName = "" Then
    rst1.Seek "=", txtBookBian
    If rst1.NoMatch Then
        MsgBox "û���ҵ�ƥ���¼��", 0 + 48, "����ʧ��"
        Findfrm.MousePointer = 0
        Exit Sub
    End If
    If rst1.Fields("�Ƿ���") = True Then
        rst2.Seek "=", txtBookBian
        Lv.ListItems.Add , , rst1.Fields("ͼ����") & vbNullString
        With Lv.ListItems(1)
            .SubItems(1) = rst1.Fields("����") & vbNullString
            .SubItems(2) = rst1.Fields("���") & vbNullString
            .SubItems(3) = rst1.Fields("�۸�") & Empty
            .SubItems(4) = rst1.Fields("������") & vbNullString
            .SubItems(5) = rst1.Fields("�Ƿ���")
            .SubItems(6) = rst2.Fields("����֤��") & vbNullString
            .SubItems(7) = rst2.Fields("����") & vbNullString
            .SubItems(8) = rst2.Fields("�������")
        End With
    Else
        Lv.ListItems.Add , , rst1.Fields("ͼ����") & vbNullString
        With Lv.ListItems(1)
            .SubItems(1) = rst1.Fields("����") & vbNullString
            .SubItems(2) = rst1.Fields("���") & vbNullString
            .SubItems(3) = rst1.Fields("�۸�") & Empty
            .SubItems(4) = rst1.Fields("������") & vbNullString
            .SubItems(5) = rst1.Fields("�Ƿ���")
        End With
    End If
ElseIf txtBookBian = "" And txtBookName <> "" Then
    FindStr = "select * from Book where ���� like"
    FindStr = FindStr & "'" & txtBookName & "'"
    
    qry1.SQL = FindStr
    Set rst = qry1.OpenRecordset
    If rst.RecordCount = 0 Then
        MsgBox "û���ҵ�ƥ���¼��", 0 + 48, "����ʧ��"
        Findfrm.MousePointer = 0
        Exit Sub
    End If
    rst.MoveLast
    RecNum = rst.RecordCount
    rst.MoveFirst
    For i = 1 To RecNum
        If rst.Fields("�Ƿ���") = True Then
        rst2.Seek "=", rst.Fields("ͼ����")
        Lv.ListItems.Add i, , rst.Fields("ͼ����") & vbNullString
        With Lv.ListItems(i)
            .SubItems(1) = rst.Fields("����") & vbNullString
            .SubItems(2) = rst.Fields("���") & vbNullString
            .SubItems(3) = rst.Fields("�۸�") & Empty
            .SubItems(4) = rst.Fields("������") & vbNullString
            .SubItems(5) = rst.Fields("�Ƿ���")
            .SubItems(6) = rst2.Fields("����֤��") & vbNullString
            .SubItems(7) = rst2.Fields("����") & vbNullString
            .SubItems(8) = rst2.Fields("�������")
        End With
        Else
           Lv.ListItems.Add i, , rst.Fields("ͼ����") & vbNullString
        With Lv.ListItems(i)
            .SubItems(1) = rst.Fields("����") & vbNullString
            .SubItems(2) = rst.Fields("���") & vbNullString
            .SubItems(3) = rst.Fields("�۸�") & Empty
            .SubItems(4) = rst.Fields("������") & vbNullString
            .SubItems(5) = rst.Fields("�Ƿ���")
        End With
        End If
        rst.MoveNext
        If rst.EOF Then Exit For
    Next
Else
    MsgBox "��ѡ��һ����в���", 0 + 48, "��ʾ"
    txtBookBian = ""
    txtBookName = ""
    txtBookBian.SetFocus
    Findfrm.MousePointer = 0
    Exit Sub
End If
Findfrm.MousePointer = 0
End Sub
Private Sub cmdKong_Click()
txtBookBian = ""
txtBookName = ""
Lv.ListItems.Clear
txtBookBian.SetFocus
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set db1 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst1 = db1.OpenRecordset("Book", dbOpenTable)
Set qry1 = db1.CreateQueryDef("")
rst1.Index = "ͼ����"

Set db2 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst2 = db2.OpenRecordset("BookFf", dbOpenTable)
Set qry2 = db2.CreateQueryDef("")
rst2.Index = "ͼ����"

txtBookBian = ""
txtBookName = ""

Lv.View = lvwReport
Lv.GridLines = False
Lv.ColumnHeaders.Add , , "ͼ����"
Lv.ColumnHeaders.Add , , "����"
Lv.ColumnHeaders.Add , , "���"
Lv.ColumnHeaders.Add , , "�۸�"
Lv.ColumnHeaders.Add , , "������"
Lv.ColumnHeaders.Add , , "�Ƿ���"
Lv.ColumnHeaders.Add , , "����֤��"
Lv.ColumnHeaders.Add , , "����������"
Lv.ColumnHeaders.Add , , "��������"
End Sub
Private Sub txtBookBian_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBookName.Text = ""
    cmdBeginFind_Click
End If
End Sub
Private Sub txtBookName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBookBian.Text = ""
    cmdBeginFind_Click
End If
End Sub
