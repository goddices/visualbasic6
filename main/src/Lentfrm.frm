VERSION 5.00
Begin VB.Form Lentfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�黹ͼ��"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "Lentfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame5 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H0000C000&
         Cancel          =   -1  'True
         Caption         =   "ȡ��"
         DownPicture     =   "Lentfrm.frx":0442
         Height          =   495
         Left            =   4320
         Picture         =   "Lentfrm.frx":0884
         TabIndex        =   30
         ToolTipText     =   "�رմ˶Ի���"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtBookBian1 
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
         Left            =   3120
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   660
         Width           =   2055
      End
      Begin VB.Frame Frame6 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox txtFa 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   5880
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox txtChaoChu 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   3480
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtXianDing 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   5880
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtLentDay 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   1200
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtToday 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   3480
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtType1 
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
            Height          =   315
            Left            =   5520
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtBookhao1 
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
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtCost1 
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
            Height          =   315
            Left            =   3600
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtBookname1 
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
            Height          =   345
            Left            =   1200
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtChuban1 
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
            Height          =   345
            Left            =   1200
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox txtLentDate1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ա"
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
            Height          =   210
            Index           =   7
            Left            =   6840
            TabIndex        =   29
            Top             =   2760
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��"
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
            Height          =   210
            Index           =   6
            Left            =   6840
            TabIndex        =   28
            Top             =   2160
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "������"
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
            Height          =   210
            Index           =   5
            Left            =   4920
            TabIndex        =   26
            Top             =   2760
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��������"
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
            Height          =   210
            Index           =   4
            Left            =   2520
            TabIndex        =   24
            Top             =   2760
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�޶�����"
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
            Height          =   210
            Index           =   3
            Left            =   4920
            TabIndex        =   22
            Top             =   2160
            Width           =   840
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   5640
            Picture         =   "Lentfrm.frx":0B8E
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   2760
            Width           =   840
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "���"
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
            Height          =   210
            Index           =   0
            Left            =   4920
            TabIndex        =   16
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��������"
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
            Height          =   210
            Index           =   9
            Left            =   2520
            TabIndex        =   12
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "������"
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
            Height          =   210
            Index           =   10
            Left            =   240
            TabIndex        =   11
            Top             =   1560
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�۸�"
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
            Height          =   210
            Index           =   11
            Left            =   3000
            TabIndex        =   10
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��  ��"
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
            Height          =   210
            Index           =   12
            Left            =   240
            TabIndex        =   9
            Top             =   960
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ͼ����"
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
            Height          =   210
            Index           =   13
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdOkCancel 
         Caption         =   "�黹ͼ��(&C)"
         Height          =   495
         Index           =   1
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "�黹��ǰͼ��"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   5400
         Picture         =   "Lentfrm.frx":0FD0
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Lentfrm.frx":1412
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����Ҫ����ͼ����"
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
         TabIndex        =   15
         Top             =   720
         Width           =   2310
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Enter"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   3
         Left            =   5520
         TabIndex        =   14
         Top             =   960
         Width           =   450
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         Index           =   1
         X1              =   0
         X2              =   7440
         Y1              =   1320
         Y2              =   1320
      End
   End
End
Attribute VB_Name = "Lentfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst2 As Recordset '�򿪱�BookFlag
Dim rst3 As Recordset '�򿪱�Book
Dim rst1 As Recordset '�򿪱�personal
Dim db2 As Database
Dim db3 As Database
Dim db1 As Database
Dim db As Database
Dim rst As Recordset
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 1
        rst2.Seek "=", txtBookBian1.Text
        If rst2.NoMatch Then
            MsgBox "û�н���Ȿ�飡�ǲ��Ǳ�Ŵ��ˣ�", 0 + 48, "��ʾ"
            txtBookBian1.Text = ""
            txtBookBian1.SetFocus
            Frame6.Visible = False
            cmdOkCancel(1).Visible = False
            Exit Sub
        End If
        If rst3.Fields("�Ƿ���") = False Then
            MsgBox "���黹û�н��", 0 + 48, "��ʾ"
            Exit Sub
        End If
        rst1.Seek "=", rst2.Fields("����֤��")
        rst1.Edit
        '��������д�����ݿ�
        rst1.Fields("����") = Val(txtFa.Text) + rst1.Fields("����")
        rst1.Update
        If txtFa.Text > 0 Then
            MsgBox "�������Ѿ�д�����ݿ⣡", 0 + 48, "��ʾ"
        End If
        rst2.Delete
        rst3.Edit
        rst3.Fields("�Ƿ���") = False
        rst3.Fields("�������") = Empty
        rst3.Update
        txtBookBian1.Text = ""
        txtBookBian1.SetFocus
        Frame6.Visible = False
        cmdOkCancel(1).Visible = False
        MsgBox "����ɹ������س�����", 0 + 48, "���"
End Select
End Sub
Private Sub Form_Load()
Set db2 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst2 = db2.OpenRecordset("BookFf", dbOpenTable)
rst2.Index = "ͼ����"

Set db3 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst3 = db3.OpenRecordset("Book", dbOpenTable)
rst3.Index = "ͼ����"

Set db1 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst1 = db1.OpenRecordset("Personal", dbOpenTable)
rst1.Index = "����֤��"

Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Type", dbOpenTable)
rst.Index = "���"

txtBookBian1.Text = ""
txtBookhao1.Text = ""
txtBookname1.Text = ""
txtCost1 = ""
txtChuban1 = ""
txtLentDate1 = ""
txtToday = ""
txtType1 = ""
txtLentDay = ""
txtXianDing.Text = ""
txtChaoChu.Text = ""
txtFa.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
rst2.Close
rst3.Close
rst1.Close
db1.Close
db2.Close
db3.Close
End Sub

Private Sub txtBookBian1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rst3.Seek "=", txtBookBian1.Text
    If rst3.NoMatch Then
        MsgBox "û�д�ͼ���ţ���������д", 0 + 48, "��д����"
        txtBookBian1.Text = ""
        'txtBookBian1.SelLength
        txtBookBian1.SetFocus
        Exit Sub
    End If
    Frame6.Visible = True
    cmdOkCancel(1).Visible = True
    txtBookhao1.Text = txtBookBian1.Text
    txtBookname1.Text = rst3.Fields("����") & vbNullString
    txtChuban1.Text = rst3.Fields("������") & vbNullString
    txtCost1.Text = rst3.Fields("�۸�") & Empty
    txtLentDate1 = rst3.Fields("�������") & Empty
    txtToday.Text = rst3.Fields("�������") & vbNullString
    txtType1.Text = rst3.Fields("���") & vbNullString
    txtLentDay.Text = rst3.Fields("�������") - rst3.Fields("�������") & Empty
    rst.Seek "=", rst3.Fields("���")
    BookDay = rst.Fields("�������")
    txtXianDing.Text = BookDay  'BookDay Ϊ�޶����������
    If Val(txtLentDay.Text) - BookDay <= 0 Then  '�ж��Ƿ񳬳�������
        txtChaoChu.Text = "δ����"
        txtFa.Text = "0"
        Exit Sub
    Else
        txtChaoChu.Text = Val(txtLentDay.Text) - BookDay
    End If
    'txtFa.Text = Format(FaCost * Val(txtChaoChu.Text), "#.00") '�������
End If
End Sub

