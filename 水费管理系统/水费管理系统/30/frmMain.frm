VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѧ������������"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   7560
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   600
      TabIndex        =   33
      Text            =   "Text3"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   435
      Left            =   6120
      TabIndex        =   31
      Top             =   4800
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmMain.frx":0000
      Left            =   2760
      List            =   "frmMain.frx":0016
      TabIndex        =   29
      Text            =   "��ѡ�����"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Frame Frame4 
      Caption         =   "��"
      Height          =   795
      Left            =   480
      TabIndex        =   24
      Top             =   3480
      Width           =   1695
      Begin VB.OptionButton Option4 
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   28
         Top             =   480
         Width           =   555
      End
      Begin VB.OptionButton Option4 
         Caption         =   "һ��"
         Height          =   180
         Index           =   2
         Left            =   900
         TabIndex        =   27
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton Option4 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   795
      End
      Begin VB.OptionButton Option4 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��"
      Height          =   735
      Left            =   480
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton Option3 
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   22
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton Option3 
         Caption         =   "һ��"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "����"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��"
      Height          =   795
      Left            =   480
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton Option2 
         Caption         =   "����"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "һ��"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton Option2 
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   14
         Top             =   480
         Width           =   675
      End
   End
   Begin VB.TextBox Text2 
      Height          =   630
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "��"
      Height          =   795
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "��"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "һ��"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   435
      Left            =   4920
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2400
      Width           =   4275
   End
   Begin VB.Frame fraScore 
      Caption         =   "��"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton optScore 
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   4
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton optScore 
         Caption         =   "һ��"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optScore 
         Caption         =   "����"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optScore 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      Height          =   255
      Left            =   180
      TabIndex        =   32
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "����Զ�������"
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "��    ��    ��"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   1920
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'�������������ͱ���
Dim Mind As String, Score As String, PE As String
Dim Manner As String, Labor As String, jy As String

Private Sub Combo1_Click()
    jy = Combo1.Text
End Sub

'����
Private Sub Command1_Click()
    If Text3.Text = "" Then
        MsgBox "����������Ϊ��", vbOKOnly, "��ʾ"
        Text3.SetFocus
        Text1.Text = ""
    Else
        Text1 = "    " + Text3.Text + "ͬѧһѧ����" + Mind + Score + _
            PE + Manner + Labor + Text2.Text + jy
        s2 = Text1.Text
    End If
End Sub

'����
Private Sub Command2_Click()
    Dim sF As String
    sF = Text3.Text + ".txt"
    Open sF For Output As #1
    Print #1, Text1.Text
    Close #1
End Sub

Private Sub Form_Load()
Form1.Height = 6015
Form1.Width = 7650
Text3.Text = s1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Form5.Text12 = s2
End Sub

'��
Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Mind = "������Ӹߴ���Ҫ���Լ����Ͻ���ǿ�����ؼ��ɡ�"
        Case 1
           Mind = "˼����Ҫ��������������ѣ��ܽϺõ�����У�Ͱ�档"
        Case 2
            Mind = "��һ�����Ͻ��ģ�������������ѧУ�Ĺ����ƶȡ�"
        Case 3
            Mind = "˼����ɱ���һ�㣬������������У�Ͱ�棬������������ǿ��ż��" _
                + "��������Ϊ��"
    End Select
End Sub

'��
Private Sub optScore_Click(Index As Integer)
    Select Case Index
    Case 0
        Score = "ѧϰ�������ڣ��ڷ��ù����ɼ����죬�籣����ȥ��ǰ;���ޡ�"
    Case 1
        Score = "ѧϰ�����Ϻã����¹��򣬳ɼ��������нϺõķ�չǱ����"
    Case 2
        Score = "��һ������ѧ��������������Ŭ��ȡ����Ӧ�ĳɼ����ھ�Ǳ���ɹۡ�"
    Case 3
        Score = "ѧϰ�����ĳ̶Ȳ������ɼ���̫���룬�����ԽϺã���Ŭ����ܿ��г�" _
            + "������ġ�"
    End Select
    
End Sub

'��
Private Sub Option3_Click(Index As Integer)
    Select Case Index
        Case 0
            PE = "�Ȱ������������μ����������������������������彡����"
        Case 1
            PE = "���������������μ����������������������ã����彡����"
        Case 2
            PE = "�����ϰ�Ҫ��μ������������ɼ��ϸ�"
        Case 3
            PE = "�������治̫���룬���������д���ǿ��"
    End Select
End Sub

'��
Private Sub Option4_Click(Index As Integer)
    Select Case Index
        Case 0
            Manner = "Ϊ�˳�ʵ��������ʦ���������Ž�ͬѧ���������ˡ�"
        Case 1
            Manner = "��������ò������ʦ�����Ž�ͬѧ��"
        Case 2
            Manner = "�������ˣ���ͬѧ�ͺ��ദ��"
        Case 3
            Manner = "��ֹ�н�Ƿ�ܵ�������ǿ������"
    End Select
End Sub

'��
Private Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0
            Labor = "�Ȱ��Ͷ����гԿ����͵ľ���"
        Case 1
            Labor = "�Ͷ������ϸɣ��ܳԿ����͡�"
        Case 2
            Labor = "�Ͷ�����������񣬵��������д���ߡ�"
        Case 3
            Labor = "�Ͷ�������ʶ���㣬����һ�㡣"
    End Select
End Sub
