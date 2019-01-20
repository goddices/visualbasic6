VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "打字游戏 小游戏而已"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5610
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Menu MenuGame 
      Caption         =   "数量 (&N)"
      Index           =   0
   End
   Begin VB.Menu MenuGame 
      Caption         =   "速度 (&P)"
      Index           =   1
   End
   Begin VB.Menu MenuGame 
      Caption         =   "重置 (&R)"
      Index           =   2
   End
   Begin VB.Menu MenuGame 
      Caption         =   "开始 (&S)"
      Index           =   3
   End
   Begin VB.Menu MenuGame 
      Caption         =   "输赢 (&W)"
      Index           =   4
   End
   Begin VB.Menu MenuGame 
      Caption         =   "帮助 (&H)"
      Index           =   5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartPause As Boolean
Dim n As Integer
Dim Speed As Integer
Dim Down As Integer, Hit As Integer
Dim DownLost As Integer, HitWin As Integer
 
Rem 自定义函数效率不高啊。。。
 
 
Private Sub Form_Initialize()
Speed = 10
DownLost = 100
HitWin = 100
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Randomize
For Index = 0 To n
    If Chr(KeyCode) = Label1(Index).Caption Then
         With Label1(Index)
            .Top = Me.ScaleTop
            .Caption = Chr(Int(Rnd * 26) + 65)
            .Left = Rnd * (Me.ScaleWidth - Label1(Index).Width)
            .ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        End With
    Hit = Hit + 1
    Me.Caption = "打字游戏   " & "掉落: " & Down & "  命中: " & Hit

    End If
Next Index
End Sub

Private Sub Form_Load()
On Error Resume Next

Timer1.Interval = 10
Timer1.Enabled = False
 Randomize
With Label1(0)
    .Top = Me.ScaleTop
    .Caption = Chr(Int(Rnd * 26) + 65)
    .ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    .Left = Rnd * (Me.ScaleWidth - Label1(0).Width)
    .FontSize = 30
    .BackStyle = 0
End With

For Index = 1 To n
    Load Label1(Index)
        With Label1(Index)
            .Visible = True
            .FontSize = 30
            .BackStyle = 0
            .Top = Me.ScaleTop
            .Caption = Chr(Int(Rnd * 26) + 65)
            .Left = Rnd * (Me.ScaleWidth - Label1(Index).Width)
            .ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        End With
Next Index
End Sub

Private Sub MenuGame_Click(Index As Integer)
On Error Resume Next

Select Case Index
    Case 0
        n = Int(InputBox("输入数量,建议输入1至5,如果输入0或者按取消，将会去缺省值1", "输入数量") - 1)
        Form_Load
        StartPause = False: MenuGame(3).Caption = "开始 (&S)"
    Case 1
        Speed = Int(Val(InputBox("输入速度参数，建议5-20，如果输入0或按取消，将会取缺省值0，就是不会移动", "输入速度参数")))
        Timer1.Enabled = False
        StartPause = False: MenuGame(3).Caption = "开始 (&S)"
    Case 2
        Hit = 0
        Down = 0
        Form_Load
        StartPause = False: MenuGame(3).Caption = "开始 (&S)"
    Case 3
         StartPause = Not StartPause
            If StartPause = True Then
                MenuGame(3).Caption = "暂停 (&P)"
                Timer1.Enabled = True
            ElseIf StartPause = False Then
                MenuGame(3).Caption = "开始 (&S)"
                Timer1.Enabled = False
            End If
    Case 4
        HitWin = Int(InputBox("输入数字，当命中数等于该数时即为胜利。", "输入数字"))
        DownLost = Int(InputBox("输入数字，当掉落数等于该数时即为胜利。", "输入数字"))
    Case 5
        MsgBox "目前没有编辑帮助"
 End Select
End Sub

Private Sub Timer1_Timer()
Randomize
For Index = 0 To n
    Label1(Index).Top = Label1(Index).Top + Speed
        If Label1(Index).Top >= Me.ScaleHeight Then
            With Label1(Index)
                .Top = Me.ScaleTop
                .Caption = Chr(Int(Rnd * 26) + 65)
                .Left = Rnd * (Me.ScaleWidth - Label1(Index).Width)
                .ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

            End With
        Down = Down + 1
        Me.Caption = "打字游戏   " & "掉落 " & Down & "  命中 " & Hit

        End If
Next Index

If Down >= DownLost Then MsgBox "你输啦！", vbOKOnly, "You lost": End
If Hit >= HitWin Then MsgBox "你赢啦！", vbOKOnly, "You Win": End
End Sub

 
 
