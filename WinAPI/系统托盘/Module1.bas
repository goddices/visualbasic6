Attribute VB_Name = "Module1"
'-----------------------------------------
'以下为模块中的代码：
'-----------------------------------------
Option Explicit

Public OldWindowProc As Long
Public TheForm As Form
Public TheMenu As Menu
'【VB声明】
'Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'【说明】
'  此函数发送消息到一个窗口过程

'【返回值】
'  Long，依据发送的消息不同而变化

'【参数表】
' lpPrevWndFunc----- Long，原来的窗口过程地址

' HWnd-------------- Long，窗口句柄

' Msg -------------- Long，发送的消息

' wParam ----------- Long，消息类型，参考wParam参数表

' lParam ----------- Long，依据wParam参数的不同而不同

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'【VB声明】
'  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'【说明】
'  在窗口结构中为指定的窗口设置信息

'【返回值】
'  Long，指定数据的前一个值

'【参数表】
'  hwnd -----------  Long，欲为其取得信息的窗口的句柄

'  nIndex ---------  Long，请参考GetWindowLong函数的nIndex参数的说明

'  dwNewLong ------  Long，由nIndex指定的窗口信息的新值
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'【VB声明】
'Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'【说明】

'【参数表】
'参数dwMessage ---- 为消息设置值，它可以是以下的几个常数值：0、1、2

'NIM_ADD = 0        加入图标到系统状态栏中
'NIM_MODIFY = 1     修改系统状态栏中的图标
'NIM_DELETE = 2     删除系统状态栏中的图标

'参数LpData ---- 用以传入NOTIFYICONDATA数据结构变量，我们也需要在"模块"中定义其结构如下：

'Type NOTIFYICONDATA
'       cbSize As Long              需填入NOTIFYICONDATA数据结构的长度
'       HWnd As Long                设置成窗口的句柄
'       Uid As Long                 为图标所设置的ID值
'       UFlags As Long              用来设置以下三个参数uCallbackMessage、hIcon、szTip是否有效
'       UCallbackMessage As Long    消息编号
'       HIcon As Long               显示在状态栏上的图标
'       SzTip As String * 64        提示信息
'End Type

'---- 其中参数uCallbackMessage、hIcon、szTip也应在模块中声明为以下的常量：
'Public Const NIF_MESSAGE = 1
'Public Const NIF_ICON = 2
'Public Const NIF_TIP = 4

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long


Public Const WM_USER = &H400
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

'记录 设置托盘图标的数据 的数据类型NOTIFYICONDATA
Public Type NOTIFYICONDATA
    cbSize As Long
    HWnd As Long
    Uid As Long
    UFlags As Long
    UCallbackMessage As Long
    HIcon As Long
    SzTip As String * 64
End Type

'TheData变量记录设置托盘图标的数据
Private TheData As NOTIFYICONDATA
' *********************************************
' 新的窗口过程--主程序中采用SetWindowLong函数改变了窗口函数的地址，消息转向由NewWindowProc处理
' *********************************************
Public Function NewWindowProc(ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    '如果用户点击了托盘中的图标，则进行判断是点击了左键还是右键
    If Msg = TRAY_CALLBACK Then
        '如果点击了左键
        If lParam = WM_LBUTTONUP Then
            '而这时窗体的状态是最小化时
            If TheForm.WindowState = vbMinimized Then _
                '恢复到最小化前的窗体状态
                TheForm.WindowState = TheForm.LastState
            TheForm.SetFocus
            Exit Function
            End If
        End If
        '如果点击了右键
        If lParam = WM_RBUTTONUP Then
            '则弹出右键菜单
            TheForm.PopupMenu TheMenu
            Exit Function
        End If
    End If
    
    '如果是其他类型的消息则传递给原有默认的窗口函数
    NewWindowProc = CallWindowProc(OldWindowProc, HWnd, Msg, wParam, lParam)
End Function
' *********************************************
' 把主窗体的图标（Form1.icon属性可改变）添加到托盘中
' *********************************************
Public Sub AddToTray(frm As Form, mnu As Menu)

    '保存当前窗体和菜单信息
    Set TheForm = frm
    Set TheMenu = mnu
    
    'GWL_WNDPROC获得该窗口的窗口函数的地址
    OldWindowProc = SetWindowLong(frm.HWnd, GWL_WNDPROC, AddressOf NewWindowProc)
    
    '知识点滴：HWnd属性
    '返回窗体或控件的句柄。语法: object.HWnd
    '说明:Microsoft Windows 运行环境，通过给应用程序中的每个窗体和控件
    '分配一个句柄（或 hWnd）来标识它们。hWnd 属性用于Windows API调用。

    '将主窗体图标添加在托盘中
    With TheData
        .Uid = 0    '忘了吗？参考一下前面内容,Uid图标的序号，做动画图标有用
        .HWnd = frm.HWnd
        .cbSize = Len(TheData)
        .HIcon = frm.Icon.Handle
        .UFlags = NIF_ICON                  '指明要对图标进行设置
        .UCallbackMessage = TRAY_CALLBACK
        .UFlags = .UFlags Or NIF_MESSAGE    '指明要设置图标或返回信息给主窗体，此句不能省去
        .cbSize = Len(TheData)              '为什么呢？我们需要在添加图标的同时，让其返回信息
    End With                                '给主窗体，Or的意思是同时进行设置和返回消息
    Shell_NotifyIcon NIM_ADD, TheData       '根据前面定义NIM_ADD，设置为“添加模式”
End Sub
' *********************************************
' 删除系统托盘中的图标
' *********************************************
Public Sub RemoveFromTray()
    '删除托盘中的图标
    With TheData
        .UFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData   '根据前面定义NIM_DELETE，设置为“删除模式”
    
    '恢复原有的设置
    SetWindowLong TheForm.HWnd, GWL_WNDPROC, OldWindowProc
End Sub
' *********************************************
' 为托盘中的图标加上浮动提示（也就是鼠标移上去时出现的提示字条）
' *********************************************
Public Sub SetTrayTip(tip As String)
    With TheData
        .SzTip = tip & vbNullChar
        .UFlags = NIF_TIP   '指明要对浮动提示进行设置
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData    '根据前面定义NIM_MODIFY，设置为“修改模式”
End Sub
' *********************************************
' 设置托盘的图标（在本例中没有用到，如果要动态改变托盘内显示的图标，它非常有用）
' 例如：1、显示动画图标（方法你一定猜到了，对！使用Timer控件，不断调用此过程，注意把动画放在pic数组中）
'       2、程序处于不同状态时，显示不同的图标，方法是类似的
' 有兴趣的话试一试吧。
' *********************************************
Public Sub SetTrayIcon(pic As Picture)
    '判断一下pic中存放的是不是图标
    If pic.Type <> vbPicTypeIcon Then Exit Sub

    '更换图标为pic中存放的图标
    With TheData
        .HIcon = pic.Handle
        .UFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub




