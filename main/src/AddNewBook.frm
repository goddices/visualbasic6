VERSION 5.00
Begin VB.Form AddNewBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ͼ��"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "AddNewBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOkCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   495
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "�رմ˶Ի���"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOkCancel 
      Caption         =   "���(&E)"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "��ͼ��������ݿ�"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "���������Ϣ"
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "AddNewBook.frx":0442
         Left            =   4800
         List            =   "AddNewBook.frx":0444
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtCost 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtBookChu 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtBookName 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1290
         Width           =   2535
      End
      Begin VB.TextBox txtBookNum 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "��Ҫ��ʾ������۸�ʱ��ʹ��Ӣ��Բ����š�.��"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4800
         Picture         =   "AddNewBook.frx":0446
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "������"
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
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   5
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   4
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ͼ����"
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
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1545
      End
   End
End
Attribute VB_Name = "AddNewBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim db1 As Database
Dim rst1 As Recordset
Private Sub cmdOkCancel_Click(Index As Integer)
Dim con As New ADODB.Connection
Set con = New ADODB.Connection
'Dim time As Integer
Dim re As New ADODB.Recordset
Select Case Index
    Case 0
        If txtBookNum = "" Or txtBookName = "" Or Combo1.Text = "" _
            Or txtCost = "" Or txtBookChu = "" Then
                MsgBox "�뽫������Ϣ��д������", 0 + 48, "��ʾ"
                Exit Sub
        End If
        'rst.Seek "=", Trim(txtBookNum.Text)
  con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source=" & App.Path & "\DataBase\data.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"

      re.Open "select * from book where ͼ����='" & txtBookNum.Text & "'", con, 3, 3
      
        If re.RecordCount <> 0 Then
            MsgBox "�˱���Ѿ����ڣ�����д������ţ�", 0 + 48, "��ʾ"
            'txtBookNum.SelText = txtBookNum.Text
            txtBookNum.SetFocus
            Exit Sub
        End If
        rst.AddNew
        rst.Fields("ͼ����") = txtBookNum.Text
        rst.Fields("����") = txtBookName.Text
        rst.Fields("���") = Combo1.Text
        rst.Fields("�۸�") = txtCost.Text
        rst.Fields("������") = txtBookChu.Text
        rst.Update
        MsgBox "��ӳɹ������س�����", 0 + 48, "�ɹ�"
        txtBookNum.Text = ""
        txtBookName = ""
        txtCost = ""
        Combo1.Text = ""
        txtBookChu = ""
        txtBookNum.SetFocus
    Case 1
        Unload Me
End Select
End Sub


Private Sub Form_Load()

Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Book", dbOpenTable)
rst.Index = "ͼ����"

Set db1 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst1 = db1.OpenRecordset("Type", dbOpenTable)

TypeAdd
txtBookNum.Text = ""
txtBookName = ""
txtCost = ""
Combo1.Text = ""
txtBookChu = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
rst.Close
rst1.Close
db1.Close
db.Close
End Sub
Private Sub TypeAdd()
Dim i As Integer
rst1.MoveLast
rst1.MoveFirst
For i = 1 To rst1.RecordCount
    Combo1.AddItem rst1.Fields("���")
    rst1.MoveNext
    If rst1.EOF Then Exit Sub
Next
End Sub
