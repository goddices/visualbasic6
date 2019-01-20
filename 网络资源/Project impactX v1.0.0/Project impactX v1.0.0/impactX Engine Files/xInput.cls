VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'impactX Game Engine
'本类模块用于处理鼠标键盘和手柄的输入
'使用本类模块必须遵守:
'你可以免费使用本引擎及代码
'使用本引擎后的责任由使用者承担
'你可以任意拷贝本引擎代码，但必须保证其完整性
'希望我能得到你使用本引擎制作出的程序
'Davy.xu sunicdavy@sina.com qq:20998333

Option Explicit
Dim di As DirectInput8
Dim DIDevice(0 To 4) As DirectInputDevice8 'DX输入设备
Dim diState As DIKEYBOARDSTATE '键盘按钮状态
Dim KeyState(255) As Integer
Dim JoyPadState(31) As Integer
Dim MouseState(3) As Integer
Dim m_hWnd As Long
''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\鼠标\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function showCursor Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Enum ENUM_XG_MOUSEBUTTON
    xgL_BUTTON = 1
    xgR_BUTTON = 2
    xgM_BUTTON = 3
End Enum

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\手柄\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Enum ENUM_XG_PSBUTTON 'PS改制手柄的键盘码，其他手柄的可能会有些出入
    xgPadUP = 13
    xgPadDOWN = 15
    xgPadLEFT = 14
    xgPadRIGHT = 16
    xgPadBTN1 = 1
    xgPadBTN2 = 2
    xgPadBTN3 = 3
    xgPadBTN4 = 4
    xgPadL1 = 7
    xgPadL2 = 8
    xgPadR1 = 5
    xgPadR2 = 6
    xgPadSTART = 9
    xgPadSELECT = 10
End Enum
'DirectInput设备枚举,列出手柄及其他输入设备
Dim diDevEnum As DirectInputEnumDevices8
'手柄状态,可以获取Axis的参数
Dim JoyCaps(4) As DIDEVCAPS
'可用手柄的数量
Dim m_JoyPadNum As Integer



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\DirectInput基础函数\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'功能:初始化DirectInput
'参数:hWnd为窗体的句柄,若某窗口名称为Main，则可以获得Main.hWnd
Public Function InitDXInput(hWnd As Long) As Boolean
    On Error GoTo ErrH
    m_hWnd = hWnd
    Dim dx As New DirectX8
    Set di = dx.DirectInputCreate()
    If Err.Number <> 0 Then
        InitDXInput = False
        Debug.Print "Err [InitdxInput] DirectInput创建错误!"
        Exit Function
    End If
    '初始化键盘
    Set DIDevice(0) = di.CreateDevice("GUID_SysKeyboard") '创建键盘
    DIDevice(0).SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice(0).SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DIDevice(0).Acquire
    '初始化手柄
    Set diDevEnum = di.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
    m_JoyPadNum = CInt(diDevEnum.GetCount)
'    If diDevEnum.GetCount = 0 Then
'        Debug.Print "Warning [InitdxInput] 没有连接手柄"
'    End If
    Dim n As Integer
    If m_JoyPadNum > 4 Then m_JoyPadNum = 4
        For n = 1 To m_JoyPadNum
            Set DIDevice(n) = di.CreateDevice(diDevEnum.GetItem(n).GetGuidInstance)
            DIDevice(n).SetCommonDataFormat DIFORMAT_JOYSTICK
            DIDevice(n).SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
            DIDevice(n).GetCapabilities JoyCaps(n)
'            Debug.Print "Pad:" & n
'            Debug.Print JoyCaps(n).lButtons
            DIDevice(n).SetEventNotification 0
            DIDevice(n).Acquire
        Next
    InitDXInput = True
    Exit Function
    
ErrH:
    InitDXInput = False
    Debug.Print "Err [InitdxInput] 初始化输入设备错误!"
End Function
'功能:卸载DirectInput
Public Sub UnloadDXInput()
    Dim i As Integer
    For i = 0 To 4
        If Not (DIDevice(i) Is Nothing) Then
            DIDevice(i).Unacquire
        End If
    Next i
    Set di = Nothing
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\键盘相关函数\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'功能:指定的键盘按键是否按下
'CONST_DIKEYFLAGS请查看DXSDK或者本引擎的说明
Public Function KeyInput(ByVal KeyCode As CONST_DIKEYFLAGS, Optional ByVal Once As Boolean = False) As Boolean
    If KeyCode < 0 Or KeyCode > 255 Then
        Debug.Print "Err [GetKeyInput] 输入键盘检测码不在范围内!"
        Exit Function
    End If
    DIDevice(0).GetDeviceStateKeyboard diState
    KeyInput = IIf(diState.Key(KeyCode) = 0, False, True)
    
    If KeyInput Then
        If KeyState(KeyCode) > 0 And Once Then
            KeyInput = False
        End If
        If KeyState(KeyCode) > 10000 Then KeyState(KeyCode) = 1
        KeyState(KeyCode) = KeyState(KeyCode) + 1
    Else
        KeyState(KeyCode) = 0
    End If
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\手柄相关函数\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'功能:获得可用的手柄个数
Public Function GetJoyPadNum() As Integer
    GetJoyPadNum = m_JoyPadNum
End Function
'功能:获得可用手柄按钮个数
'参数:手柄号(例如JoyPadNum=1 为1号手柄)
Public Function GetBtnNum(JoyPadNum As Integer) As Integer
    If JoyPadNum < 0 Or JoyPadNum > m_JoyPadNum Then Exit Function
    GetBtnNum = JoyCaps(JoyPadNum).lButtons
End Function

'功能:指定的按键码是否按下
'参数:手柄号(例如JoyPadNum=1 为1号手柄)
'       按钮:1~16
'注意:在Win2000以上可以调节手柄的Axis模式和Button模式
'       对于PS改制手柄 无论在Axis模式还是Button下本函数都会自动识别方向键

Public Function JoyInput(ByVal JoyPadNum As Integer, ByVal Button As ENUM_XG_PSBUTTON, Optional Once As Boolean) As Boolean
    Dim JoyState As DIJOYSTATE
    If m_JoyPadNum = 0 Then
        JoyInput = False
        'Debug.Print "Err:[Joyinput] 没有安装手柄"
        Exit Function
    End If
    If Button = 0 Then JoyInput = False: Exit Function
    Button = Button - 1 '纠正到WINDOWS里的按键码
    
    DIDevice(JoyPadNum).Poll
    DIDevice(JoyPadNum).GetDeviceStateJoystick JoyState
    
        If JoyState.Buttons(Button) = 0 Then
            JoyInput = False
        Else
            JoyInput = True
        End If
    'Axis模式下的号码对应
    Select Case Button
        Case 12
            If JoyState.Y < 15000 Then JoyInput = True
        Case 14
            If JoyState.Y > 50000 Then JoyInput = True
        Case 13
            If JoyState.X < 15000 Then JoyInput = True
        Case 15
            If JoyState.X > 50000 Then JoyInput = True
    End Select
    
    If JoyInput Then
        If JoyPadState(Button) > 0 And Once Then
            JoyInput = False
        End If
        JoyPadState(Button) = JoyPadState(Button) + 1
    Else
        JoyPadState(Button) = 0
    End If
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\鼠标相关函数\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'功能:返回鼠标X坐标
Public Function MouseX() As Long
    Dim t As PointAPI
    Dim client As RECT
    GetCursorPos t
    GetClientRect m_hWnd, client
    ScreenToClient m_hWnd, t
    MouseX = t.X
    If t.X < client.Left Then MouseX = 0
    If t.X > client.Right Then MouseX = client.Right
    
End Function
'功能:返回鼠标Y坐标
Public Function MouseY() As Long
    Dim t As PointAPI
    Dim client As RECT
    GetCursorPos t
    GetClientRect m_hWnd, client
    ScreenToClient m_hWnd, t
    MouseY = t.Y
    If t.Y < client.Top Then MouseY = 0
    If t.Y > client.Bottom Then MouseY = client.Bottom
End Function
'功能:隐藏鼠标
Public Sub HideMouse()
    Do: Loop Until showCursor(0) < 0
End Sub
'功能:显示鼠标
Public Sub ShowMouse()
    Do: Loop Until showCursor(1) > 0
End Sub
'功能:指定的鼠标按钮是否按下
'参数:由ENUM_XG_MOUSEBUTTON给出常用的鼠标按钮定义
Public Function MouseKey(ByVal KeyCode As ENUM_XG_MOUSEBUTTON, Optional ByVal Once As Boolean) As Boolean
    MouseKey = False
    Select Case KeyCode
        Case xgL_BUTTON
            If (GetKeyState(vbKeyLButton) And &H8000) Then
                MouseKey = True
            Else
                MouseKey = False
            End If
        Case xgR_BUTTON
            If (GetKeyState(vbKeyRButton) And &H8000) Then
                MouseKey = True
            Else
                MouseKey = False
            End If
        Case xgM_BUTTON
            If (GetKeyState(vbKeyMButton) And &H8000) Then
               MouseKey = True
            Else
                MouseKey = False
            End If
    End Select
    
    If MouseKey Then
        If MouseState(KeyCode) > 0 And Once Then
            MouseKey = False
        End If
        MouseState(KeyCode) = MouseState(KeyCode) + 1
    Else
        MouseState(KeyCode) = 0
    End If
    If MouseState(KeyCode) > 10000 Then MouseState(KeyCode) = 1
End Function





