VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu mnuTray 
      Caption         =   "SS"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuTrayMaximize 
         Caption         =   "Maximize"
      End
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuTrayMove 
         Caption         =   "Move"
      End
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTraySize 
         Caption         =   "Size"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------
'           使用系统托盘程序演示
'---------------------------------------------
'           洪恩在线 求知无限
'---------------------------------------------
'程序说明：
'   这是一个比较完整的使用系统托盘的程序实例，包括
'了：添加托盘图标，删除托盘图标，动态改变托盘图标，
'为托盘图标添加浮动提示信息，实现托盘图标的鼠标右键
'菜单等内容。
'-------名称-------------------作用------------
'       Form1                   主窗体
'       mnuFile,mnuFileExit     文件菜单，菜单项
'       mnuTray,mnuTrayClose... 托盘区右键菜单，菜单项
'---------------------------------------------

Option Explicit

'LastState变量的作用是标示主窗体原有状态
Public LastState As Integer

'【VB声明】
'  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'【说明】
'  调用一个窗口的窗口函数，将一条消息发给那个窗口。除非消息处理完毕，否则该函数不会返回。SendMessageBynum，
'  SendMessageByString是该函数的“类型安全”声明形式

'【返回值】
'  Long，由具体的消息决定

'【参数表】
'  hwnd -----------  Long，要接收消息的那个窗口的句柄

'  wMsg -----------  Long，消息的标识符

'  wParam ---------  Long，具体取决于消息

'  lParam ---------  Any，具体取决于消息
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'表示发送的是系统命令
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

'当主窗体加载时
Private Sub Form_Load()
    
    '窗体的WindowState属性，返回或设置一个值，该值用来指定在运行时窗体窗口的可视状态
    'vbNormal    0   （缺省值）正常 。
    'VbMinimized 1   最小化（最小化为一个图标）
    'VbMaximized 2   最大化（扩大到最大尺寸）
    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If
    
    '将图标添加到托盘的函数，参见模块中的解释
    '注意了这是从主程序到模块的入口，本例中并没有直接调用Shell_NotifyIcon函数
    AddToTray Me, mnuTray
    
    SetTrayTip "托盘图标演示，点击右键弹出菜单"
End Sub

'在主窗体Form1大小改变时，相应改变右键菜单mnuTray的菜单项的可用属性Enabled
Private Sub Form_Resize()
    Select Case WindowState
        
        '如果窗体最小化了，把菜单项“最大化”“恢复”设为可用，
        '而把“最小化”“移动”“大小”三项设为不可用.
        '如果这时在托盘图标上点击鼠标右键，会发现不可用项变为灰色
        Case vbMinimized
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = False
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        
        '窗体最大化时
        Case vbMaximized
            mnuTrayMaximize.Enabled = False
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        
        '一般状态下
        Case vbNormal
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = True
            mnuTrayRestore.Enabled = False
            mnuTraySize.Enabled = True
    End Select

    If WindowState <> vbMinimized Then LastState = WindowState
End Sub

'保证在程序退出时删除托盘图标
Private Sub Form_Unload(Cancel As Integer)
    RemoveFromTray
End Sub

'“文件”菜单的“退出”项被点击时
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'托盘图标右键菜单上的“退出”项被点击时
Private Sub mnuTrayClose_Click()
    Unload Me
End Sub

'托盘图标右键菜单上的“最大化”项被点击时
Private Sub mnuTrayMaximize_Click()
    WindowState = vbMaximized
End Sub

'托盘图标右键菜单上的“最小化”项被点击时
Private Sub mnuTrayMinimize_Click()
    WindowState = vbMinimized
End Sub

'托盘图标右键菜单上的“移动”项被点击时
Private Sub mnuTrayMove_Click()
    SendMessage HWnd, WM_SYSCOMMAND, _
        SC_MOVE, 0&
End Sub

'托盘图标右键菜单上的“恢复”项被点击时
Private Sub mnuTrayRestore_Click()
    SendMessage HWnd, WM_SYSCOMMAND, _
        SC_RESTORE, 0&
End Sub

'托盘图标右键菜单上的“退出”项被点击时
Private Sub mnuTraySize_Click()
    SendMessage HWnd, WM_SYSCOMMAND, _
        SC_SIZE, 0&
End Sub

 
