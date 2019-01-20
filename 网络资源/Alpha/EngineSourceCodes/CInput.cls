VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const INPUT_SPEED = 10                  '每10ms刷新一次数据
Private Const DCLICK_DELAY = 500                '处理双击事件

Private m_DI As DirectInput
Private m_MouseDev As DirectInputDevice
Private m_KeyboardDev As DirectInputDevice
Private mLastDownKey As Integer
Private m_MouseX As Integer
Private m_MouseY As Integer
Private m_MouseButton As Integer
Private m_MouseK As Single                      '灵敏度
Private m_LastMouseButton As Integer

Public Mouse_DX As Integer                      '鼠标的移动Step
Public Mouse_DY As Integer
Public Mouse_DZ As Integer

Public Event InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)

Public Sub Destroy()
    Set m_DI = Nothing
    Set m_MouseDev = Nothing
    Set m_KeyboardDev = Nothing
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Property Get MouseX() As Integer
    MouseX = m_MouseX
End Property

Public Property Get MouseY() As Integer
    MouseY = m_MouseY
End Property

Public Property Get MouseButton() As Integer
    MouseButton = m_MouseButton
End Property

Public Function ScanKeys() As Integer
    Dim i As Integer
    Dim Keyboard_State As DIKEYBOARDSTATE
    '扫描按键，返回键值，0为无键
    m_KeyboardDev.GetDeviceStateKeyboard Keyboard_State
    For i = 0 To 255
        If Keyboard_State.Key(i) <> 0 Then ScanKeys = i: Exit Function
    Next i
    ScanKeys = 0
End Function

Public Function KeyDown(ByVal KeyCode As Byte) As Boolean
    Dim Keyboard_State As DIKEYBOARDSTATE
    If KeyCode < 0 Or KeyCode > 255 Then Exit Function
    m_KeyboardDev.GetDeviceStateKeyboard Keyboard_State
    KeyDown = Keyboard_State.Key(KeyCode)
End Function

Public Function KeyUp(ByVal KeyCode As Byte) As Boolean
    Dim NewDownKey As Integer
    NewDownKey = ScanKeys
    If Not KeyDown(mLastDownKey) And mLastDownKey > 0 Then
        If NewDownKey <> mLastDownKey And KeyCode = mLastDownKey Then KeyUp = True
        'Msg KeyUP
        Exit Function
    End If
    mLastDownKey = NewDownKey
End Function

Public Function Create(Optional ShowSysMouse As Boolean = False) As Boolean
    Create = True
    Set m_DI = g_DX7.DirectInputCreate()
    Set m_KeyboardDev = m_DI.CreateDevice("GUID_SysKeyboard")
    Set m_MouseDev = m_DI.CreateDevice("GUID_SysMouse")
    
    m_KeyboardDev.SetCommonDataFormat DIFORMAT_KEYBOARD
    m_MouseDev.SetCommonDataFormat DIFORMAT_MOUSE
    
    If ShowSysMouse Then
        m_MouseDev.SetCooperativeLevel g_Hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND  ' DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Else
        m_MouseDev.SetCooperativeLevel g_Hwnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    End If
    'm_MouseDev.Acquire
    
    
    m_KeyboardDev.SetCooperativeLevel g_Hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    m_KeyboardDev.Acquire
    
    If m_MouseDev Is Nothing Or m_KeyboardDev Is Nothing Then Create = False
    'Init Data
    m_MouseK = 1
    m_MouseX = g_ViewRect.Right / 2
    m_MouseY = g_ViewRect.Bottom / 2
End Function

Public Sub Refresh()
    Dim Mouse_State As DIMOUSESTATE
    Dim NewDownKey As Integer
    Dim Msg As InputMessage
    Static oldTime As Long

    '处理当前鼠标
    On Error GoTo errhand:
    
    '开启刷新率控制
    'If g_DX7.TickCount - oldTime < INPUT_SPEED Then Exit Sub

    m_MouseDev.GetDeviceStateMouse Mouse_State
    Mouse_DX = m_MouseK * Mouse_State.x
    Mouse_DY = m_MouseK * Mouse_State.y
    Mouse_DZ = Mouse_State.z
    
    m_MouseX = m_MouseX + Mouse_DX
    m_MouseY = m_MouseY + Mouse_DY
    
    If m_MouseX > g_ViewRect.Right - 1 Then m_MouseX = g_ViewRect.Right - 1
    If m_MouseY > g_ViewRect.Bottom - 1 Then m_MouseY = g_ViewRect.Bottom - 1
    If m_MouseX < 0 Then m_MouseX = 0
    If m_MouseY < 0 Then m_MouseY = 0
    
    m_MouseButton = 0
    If Mouse_State.buttons(0) Then m_MouseButton = 1
    If Mouse_State.buttons(1) Then m_MouseButton = 2
    If Mouse_State.buttons(2) Then m_MouseButton = 3
    
    If m_LastMouseButton > 0 And m_MouseButton = 0 Then
        'Msg Mouse UP,Click
        RaiseEvent InputMsg(MSG_MouseUp, m_LastMouseButton, m_MouseX, m_MouseY)
        RaiseEvent InputMsg(MSG_MouseClick, m_LastMouseButton, m_MouseX, m_MouseY)
        
        '处理鼠标的双击事件
        Static lastClickTime As Long
        Static lastClickMouseButton As Integer
        If g_DX7.TickCount - lastClickTime < DCLICK_DELAY And m_LastMouseButton = lastClickMouseButton Then
            RaiseEvent InputMsg(MSG_MouseDblClick, m_LastMouseButton, m_MouseX, m_MouseY)
        End If
        lastClickTime = g_DX7.TickCount
        lastClickMouseButton = m_LastMouseButton
    ElseIf m_MouseButton > 0 And m_MouseButton <> m_LastMouseButton Then
        'Msg Mouse Down
        RaiseEvent InputMsg(MSG_MouseDown, m_MouseButton, m_MouseX, m_MouseY)
    End If
    If Mouse_DX <> 0 Or Mouse_DY <> 0 Then
        RaiseEvent InputMsg(MSG_MouseMove, m_MouseButton, m_MouseX, m_MouseY)
    End If
    If Mouse_DZ <> 0 Then
        RaiseEvent InputMsg(MSG_MouseScroll, Mouse_DZ, m_MouseX, m_MouseY)
    End If
    m_LastMouseButton = m_MouseButton
    
    '处理键盘
    NewDownKey = ScanKeys
    If NewDownKey > 0 Then
        'Key Down/May be A Bug Here!
        RaiseEvent InputMsg(MSG_Keydown, NewDownKey, 0, 0)
    End If
    If Not KeyDown(mLastDownKey) And mLastDownKey > 0 Then
        'Msg KeyUP
        If NewDownKey <> mLastDownKey Then RaiseEvent InputMsg(MSG_KeyUp, mLastDownKey, 0, 0)
    End If
    If NewDownKey <> mLastDownKey And NewDownKey > 0 Then
        'Msg KeyPress
        RaiseEvent InputMsg(MSG_KeyPress, NewDownKey, 0, 0)
    End If
    mLastDownKey = NewDownKey
    'oldTime = g_DX7.TickCount
    Exit Sub

errhand:
    AcquireMouse
End Sub

Public Sub SetMouseK(ByVal k As Single)
    m_MouseK = k
End Sub

Private Sub AcquireMouse()
    On Error GoTo CANNOTACQUIRE
    m_MouseDev.Acquire
    Exit Sub
    
CANNOTACQUIRE:

End Sub
