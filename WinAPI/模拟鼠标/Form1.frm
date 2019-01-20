VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  '-------------------------------------------
  '               模拟鼠标的左键单击和右键单击
  '-------------------------------------------
  '                       洪恩在线   求知无限
  '-------------------------------------------
  '程序说明：
  '本例采用API函数实现模拟的鼠标事件，程序运行后会
  '产生十分有趣的效果。也来试一试。
  '本例中只使用了相对鼠标坐标，我们也可以使用绝对
  '鼠标坐标来试一试。
  '-------------------------------------------
    
  '【VB声明】
  '     Private   Declare   Sub   mouse_event   Lib   "user32"   (ByVal   dwFlags   As   Long,   ByVal   dx   As   Long,   ByVal   dy   As   Long,   ByVal   cButtons   As   Long,   ByVal   dwExtraInfo   As   Long)
    
  '【说明】
  '     模拟一次鼠标事件
    
  '【备注】
  '     进行相对运动的时候，由SystemParametersInfo函数规定的系统鼠标轨迹速度会应用于鼠标运行的速度
    
  '【参数表】
  '     dwFlags   --------     Long，下述标志的一个组合
  '     MOUSEEVENTF_ABSOLUTE
  '     dx和dy指定鼠标坐标系统中的一个绝对位置。在鼠标坐标系统中，屏幕在水平和垂直方向上均匀分割成65535×65535个单元   -
  '     MOUSEEVENTF_MOVE                   移动鼠标
  '     MOUSEEVENTF_LEFTDOWN           模拟鼠标左键按下
  '     MOUSEEVENTF_LEFTUP               模拟鼠标左键抬起
  '     MOUSEEVENTF_RIGHTDOWN         模拟鼠标右键按下
  '     MOUSEEVENTF_RIGHTUP             模拟鼠标右键抬起
  '     MOUSEEVENTF_MIDDLEDOWN       模拟鼠标中键按下
  '     MOUSEEVENTF_MIDDLEUP           模拟鼠标中键抬起
  '     dx   -------------     Long，根据是否指定了MOUSEEVENTF_ABSOLUTE标志，指定水平方向的绝对位置或相对运动'
    
  '     dy   -------------     Long，根据是否指定了MOUSEEVENTF_ABSOLUTE标志，指定垂直方向的绝对位置或相对运动
    
  '     cButtons   -------     Long，未使用
    
  '     dwExtraInfo   ----     Long，通常未用的一个值。用GetMessageExtraInfo函数可取得这个值。可用的值取决于特定的驱动程序
  Option Explicit
          Private Declare Sub mouse_event Lib "user32" _
          ( _
          ByVal dwFlags As Long, _
          ByVal dx As Long, _
          ByVal dy As Long, _
          ByVal cButtons As Long, _
          ByVal dwExtraInfo As Long _
          )
    
  'Option_Tag标示选择了哪一种模拟事件
  Dim Option_Tag     As Integer
  'OnTest标示是否处于模拟状态，以便我们停止模拟
  Dim OnTest     As Boolean
  '对API变量的定义
  Const MOUSEEVENTF_LEFTDOWN = &H2
  Const MOUSEEVENTF_LEFTUP = &H4
  Const MOUSEEVENTF_MIDDLEDOWN = &H20
  Const MOUSEEVENTF_MIDDLEUP = &H40
  Const MOUSEEVENTF_MOVE = &H1
  Const MOUSEEVENTF_ABSOLUTE = &H8000
  Const MOUSEEVENTF_RIGHTDOWN = &H8
  Const MOUSEEVENTF_RIGHTUP = &H10
    
  '控制   模拟的开始与结束
  Private Sub Command1_Click()
    
  '如果不处于模拟状态
  If OnTest = False Then
  Command1.Caption = "快停下来吧"
  Timer1.Enabled = True
  OnTest = True
  '如果处于模拟状态
  Else
  Command1.Caption = "试一试"
  Timer1.Enabled = False
  OnTest = False
  End If
  End Sub
    
Private Sub Command2_Click()
Print "sb"
End Sub

  '窗体加载时一些变量需要设置
  Private Sub Form_Load()
  Option_Tag = 1
  Timer1.Enabled = False
  OnTest = False
  End Sub
    
 
    
  '每隔一秒中模拟一次鼠标事件
  Private Sub Timer1_Timer()
  If Option_Tag = 1 Then
          '调用了mouse_event函数，其参数的设置见前面说明
          '如果同时要模拟两个鼠标事件，可以用   Or   将两个参数连接
          '这里是   鼠标左键按下   和松开两个事件的组合即一次单击
          mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  ElseIf Option_Tag = 2 Then
          '模拟鼠标右键单击事件
          mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
  Else
          '两次连续的鼠标左键单击事件   构成一次鼠标双击事件
          mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
          mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  End If
  End Sub

