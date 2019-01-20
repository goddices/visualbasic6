VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   9375
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdrestart 
      Caption         =   "Command2"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blackturn As Boolean '轮到黑方下子
Dim whiteturn As Boolean '轮到白方下子
Dim table(0 To 15, 0 To 15) As Integer '用此二维数组表示棋盘
Dim inti As Integer '数组元素……
Dim intj As Integer
Dim boolstatus As Boolean '表示棋局状态：进行/结束

Private Sub cmdclose_Click() '关闭窗口
Unload Me
Set frmmain = Nothing
End Sub

Private Sub cmdrestart_Click() '重新开始
'窗口清除
Me.Cls

'数组清零
For inti = 0 To 15
 For intj = 0 To 15
 table(inti, intj) = 0
 Next
Next

'重画棋盘
Form_Load
End Sub

Private Sub Form_Load()
'画棋盘
Form_Paint
blackturn = True '黑方先下
boolstatus = True '开始
Label1.Caption = "黑方先下"
End Sub

'下子
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intx As Integer '落子横向位置
Dim inty As Integer '落子竖向位置

'确定棋局是否在进行中，否，跳出
If boolstatus = False Then
 Label1.Caption = "结束"
 Exit Sub
End If

'确定落子的确切位置
'如果鼠标点击位置不在棋盘中，则跳出
If x < 10 Or x > 310 Or y < 10 Or y > 310 Then
 Exit Sub
End If
'如果鼠标点击位置在棋盘中，则转化为相应棋盘落子点的坐标
If (x - 10) Mod 20 < 10 Then
 intx = x - (x - 10) Mod 20
Else
 intx = x + 20 - (x - 10) Mod 20
End If
If (y - 10) Mod 20 < 10 Then
 inty = y - (y - 10) Mod 20
Else
 inty = y + 20 - (y - 10) Mod 20
End If

'把坐标转换成数组中的相应元素
inti = (intx - 10) / 20
intj = (inty - 10) / 20

'如果该数组元素不为零，即表示棋盘中相应点已有棋子，则跳出
If table(inti, intj) <> 0 Then
 Exit Sub
End If

'画子（圆）
If blackturn = True Then
 '黑色
 Me.FillColor = RGB(0, 0, 0)
 table(inti, intj) = 1 '黑子赋1
 Label1.Caption = "白方"
Else
 '白色
 Me.FillColor = RGB(255, 255, 255)
 table(inti, intj) = 2 '白子赋2
 Label1.Caption = "黑方"
End If
Me.FillStyle = 0 '不可缺
Me.Circle (intx, inty), 8

'判断是否有五子连线
Call judgeman

'轮流
blackturn = Not blackturn '取反


End Sub
Private Sub judgeman() '判断是否有五子连线

Dim strwho As String '下子方名称

If table(inti, intj) = 1 Then '表示黑方下的子
 strwho = "黑方"
Else
 strwho = "白方"
End If

'分别判断横竖，对角线是否有五子，此段代码比较复杂，可能那以理解，但其执行效率极高
'非常适合与棋盘格子很多的情况

If samelinenums(1, 0) >= 5 Or samelinenums(0, 1) >= 5 Or samelinenums(1, 1) >= 5 Or samelinenums(-1, 1) >= 5 Then
 MsgBox strwho & "胜！"
 boolstatus = False '棋局结束
End If
End Sub

Function samelinenums(changei As Integer, changej As Integer) '判断同一直线上的棋子数
Dim i As Integer
Dim j As Integer
Dim num As Integer '同一线上相同颜色棋子数

'计算落子一边同颜色的棋子数
i = inti: j = intj
Do
 If table(i, j) <> table(inti, intj) Then
 num = max(Abs(inti - i), Abs(intj - j))
 Exit Do
 End If
 i = i + changei: j = j + changej
Loop Until i < 0 Or i > 15 Or j < 0 Or j > 15

'计算落子另一边同颜色的棋子数
i = inti: j = intj
Do
 If table(i, j) <> table(inti, intj) Then
 num = num - 1 + max(Abs(inti - i), Abs(intj - j))
 Exit Do
 End If
 i = i - changei: j = j - changej
Loop Until i < 0 Or i > 15 Or j < 0 Or j > 15
'MsgBox num
samelinenums = num
End Function

'求较大值
Function max(inta As Integer, intb As Integer)
 max = inta
 If max < intb Then max = intb
End Function

Private Sub Form_Paint() '以(10,10)为左上角坐标画一个16*16,每格边长为20象素的棋盘
Cls '清除
Dim i As Integer
ScaleMode = 3 '设定窗体画布的单位为象素
For i = 10 To 330 Step 20
 Me.Line (10, i)-(330, i)
 Me.Line (i, 10)-(i, 330)
Next
End Sub

