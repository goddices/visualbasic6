VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "��Ժ"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   9285
   Begin VB.Frame Frame3 
      Caption         =   "��ѯ�����"
      Height          =   4815
      Left            =   600
      TabIndex        =   11
      Top             =   1320
      Width           =   8175
      Begin VB.CommandButton Command5 
         Caption         =   "ϵͳ����"
         Height          =   375
         Left            =   480
         TabIndex        =   36
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text12 
         DataField       =   "ѧ������"
         DataSource      =   "Data1"
         Height          =   1335
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   3360
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         DataField       =   "ѧ��"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         DataField       =   "����"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         DataField       =   "�Ա�"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text4 
         DataField       =   "רҵ"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   19
         Top             =   600
         Width           =   1332
      End
      Begin VB.TextBox Text5 
         DataField       =   "��������"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   18
         Top             =   960
         Width           =   1332
      End
      Begin VB.PictureBox Picture1 
         DataField       =   "��Ƭ"
         DataSource      =   "Data1"
         Height          =   2655
         Left            =   5760
         ScaleHeight     =   2595
         ScaleWidth      =   2115
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         DataField       =   "������ò"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   16
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text7 
         DataField       =   "��ͥסַ"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Text8 
         DataField       =   "��ͥ�绰"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         DataField       =   "С��ͨ"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   13
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         DataField       =   "�ֻ�"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   35
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ѧ    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   33
         Top             =   255
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   32
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��    ��: "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   31
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ר    ҵ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   30
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   29
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��Ƭ: "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5280
         TabIndex        =   28
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "������ò:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   27
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "��ͥסַ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   26
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "��ͥ�绰:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   25
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "С �� ͨ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   24
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   23
         Top             =   2760
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ѯ���ͣ�"
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   " רҵ"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   " ����"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   " ѧ��"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ѯ������"
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "|<"
         Height          =   372
         Left            =   3480
         TabIndex        =   10
         ToolTipText     =   "��һ��"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   372
         Left            =   3960
         TabIndex        =   9
         ToolTipText     =   "��һ��"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">>"
         Height          =   372
         Left            =   4440
         TabIndex        =   8
         ToolTipText     =   "��һ��"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">|"
         Height          =   372
         Left            =   4920
         TabIndex        =   7
         ToolTipText     =   "���һ��"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\��������ϵͳ\student.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "�������"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Data1.Recordset.MoveFirst
End Sub

Private Sub Command10_Click()
Form5.Hide
End Sub

Private Sub Command2_Click()
    Data1.Recordset.MovePrevious
    If Data1.Recordset.BOF Then
       MsgBox "�ѵ���ͷ��"
       Data1.Recordset.MoveFirst
    End If
End Sub

Private Sub Command3_Click()
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then
       MsgBox "�ѵ���β��"
       Data1.Recordset.MoveLast
    End If
End Sub

Private Sub Command4_Click()
    Data1.Recordset.MoveLast
End Sub

'Download by http://down.liehuo.net

Private Sub Command5_Click()
s1 = Text2.Text
Form1.Show
End Sub

Private Sub Command6_Click()
Dim sql As String
If Option1(0).Value = True Then
    sql = "select * from ������� where ѧ�� ='" & Trim(Text11.Text) & "'"

Else
   If Option1(1).Value = True Then
      sql = "select * from ������� where ���� ='" & Trim(Text11.Text) & "'"
   Else
      If Option1(2).Value = True Then
         sql = "select * from ������� where רҵ ='" & Trim(Text11.Text) & "'"
      End If
   End If
End If
Data1.RecordSource = sql
Data1.Refresh
If Data1.Recordset.EOF Then
   MsgBox "û������ѯ����Ϣ��", , "��ʾ"
   Data1.RecordSource = "�������"
   Data1.Refresh
End If
End Sub
Private Sub Command9_Click()
    Dim mzy As String
    zymc = InputBox$("������רҵ����:", "������ʾ����")
    Data1.RecordSource = "Select * From ������� Where רҵ = '" & zymc & "'"
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "���޴�רҵ!", , "��ʾ"
        Data1.RecordSource = "�������"
        Data1.Refresh
        End If
End Sub

Private Sub Form_Load()
Form5.Width = 9405
Form5.Height = 6765
End Sub

Private Sub Picture1_Click()
   Picture1.Picture = Clipboard.GetData
End Sub

