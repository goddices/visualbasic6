VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form32 
   BackColor       =   &H00404000&
   Caption         =   "�û����"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   7965
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "ѡ���ѯ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404000&
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404000&
         Caption         =   "�ܻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "bfhcx.frx":0000
      Height          =   3855
      Left            =   360
      OleObjectBlob   =   "bfhcx.frx":0014
      TabIndex        =   2
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ˮ�ѹ���ϵͳ\water.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "�û�����"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ѯ"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "���벡���Ż�ҽʦ������"
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
SQL = "select * from �û����� where �ܻ���='" & Trim(Text1.Text) & "'"
Data1.RecordSource = SQL
Data1.Refresh
If Data1.Recordset.EOF Then
   MsgBox "û�д��ܻ��ţ�", , "��ʾ"
End If
End If

If Option2.Value = True Then
SQL = "select * from �û����� where ����='" & Trim(Text1.Text) & "'"
Data1.RecordSource = SQL
Data1.Refresh
If Data1.Recordset.EOF Then
   MsgBox "û�д˻�����", , "��ʾ"
End If
End If

If Option1.Value = False And Option2.Value = False Then
   MsgBox "��ѡ���ѯ����Ŀ���ٽ��в�ѯ��", , "��ʾ"
End If
End Sub

Private Sub Form_Load()
  Form32.Width = 8085
  Form32.Height = 6300
  Form32.Move (MDIForm1.Width - Form32.Width) / 2, (MDIForm1.Height - Form32.Height) / 4
End Sub
