VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "学生评语生成器"
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
      Caption         =   "保存"
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
      Text            =   "请选择寄语"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Frame Frame4 
      Caption         =   "美"
      Height          =   795
      Left            =   480
      TabIndex        =   24
      Top             =   3480
      Width           =   1695
      Begin VB.OptionButton Option4 
         Caption         =   "差"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   28
         Top             =   480
         Width           =   555
      End
      Begin VB.OptionButton Option4 
         Caption         =   "一般"
         Height          =   180
         Index           =   2
         Left            =   900
         TabIndex        =   27
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton Option4 
         Caption         =   "良好"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   795
      End
      Begin VB.OptionButton Option4 
         Caption         =   "优秀"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "体"
      Height          =   735
      Left            =   480
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton Option3 
         Caption         =   "差"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   22
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton Option3 
         Caption         =   "一般"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "良好"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "优秀"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "劳"
      Height          =   795
      Left            =   480
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "优秀"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton Option2 
         Caption         =   "良好"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "一般"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton Option2 
         Caption         =   "差"
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
      Caption         =   "德"
      Height          =   795
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "差"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "一般"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "良好"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "优秀"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
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
      Caption         =   "智"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton optScore 
         Caption         =   "差"
         Height          =   180
         Index           =   3
         Left            =   900
         TabIndex        =   4
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton optScore 
         Caption         =   "一般"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optScore 
         Caption         =   "良好"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optScore 
         Caption         =   "优秀"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label Label2 
      Caption         =   "姓名"
      Height          =   255
      Left            =   180
      TabIndex        =   32
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "添加自定义评语"
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "评    语    栏"
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

'声明德智体美劳变量
Dim Mind As String, Score As String, PE As String
Dim Manner As String, Labor As String, jy As String

Private Sub Combo1_Click()
    jy = Combo1.Text
End Sub

'生成
Private Sub Command1_Click()
    If Text3.Text = "" Then
        MsgBox "姓名栏不能为空", vbOKOnly, "提示"
        Text3.SetFocus
        Text1.Text = ""
    Else
        Text1 = "    " + Text3.Text + "同学一学期来" + Mind + Score + _
            PE + Manner + Labor + Text2.Text + jy
        s2 = Text1.Text
    End If
End Sub

'保存
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

'德
Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Mind = "各方面从高从严要求自己，上进心强，遵守纪律。"
        Case 1
           Mind = "思想上要求进步，严于律已，能较好地遵守校纪班规。"
        Case 2
            Mind = "有一定的上进心，基本上能遵守学校的规章制度。"
        Case 3
            Mind = "思想纪律表现一般，基本上能遵守校纪班规，但自律能力不强，偶尔" _
                + "有意外行为。"
    End Select
End Sub

'智
Private Sub optScore_Click(Index As Integer)
    Select Case Index
    Case 0
        Score = "学习能力出众，勤奋用功，成绩优异，如保持下去，前途无限。"
    Case 1
        Score = "学习能力较好，肯下功夫，成绩优良，有较好的发展潜力。"
    Case 2
        Score = "有一定的自学能力，所作出的努力取得相应的成绩，挖掘潜力可观。"
    Case 3
        Score = "学习上用心程度不够，成绩不太理想，但悟性较好，经努力会很快有长" _
            + "足进步的。"
    End Select
    
End Sub

'体
Private Sub Option3_Click(Index As Integer)
    Select Case Index
        Case 0
            PE = "热爱体育，积极参加体育锻炼，体能素质优良，身体健康。"
        Case 1
            PE = "爱好体育，主动参加体育锻炼，体能素质良好，身体健康。"
        Case 2
            PE = "基本上按要求参加体育锻炼，成绩合格。"
        Case 3
            PE = "体育方面不太理想，身体素质有待加强。"
    End Select
End Sub

'美
Private Sub Option4_Click(Index As Integer)
    Select Case Index
        Case 0
            Manner = "为人诚实，尊重老师、长辈，团结同学，乐于助人。"
        Case 1
            Manner = "待人有礼貌，尊重师长、团结同学。"
        Case 2
            Manner = "和气待人，与同学和好相处。"
        Case 3
            Manner = "举止有进欠周到，望加强修养。"
    End Select
End Sub

'劳
Private Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0
            Labor = "热爱劳动，有吃苦耐劳的精神。"
        Case 1
            Labor = "劳动积极肯干，能吃苦耐劳。"
        Case 2
            Labor = "劳动课能完成任务，但主动性有待提高。"
        Case 3
            Labor = "劳动方面认识不足，表现一般。"
    End Select
End Sub
