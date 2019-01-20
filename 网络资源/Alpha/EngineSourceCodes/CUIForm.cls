VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public Actived As Boolean
Public Release As Boolean               '释放输入

Private m_Controls As Collection
Private m_WindowRect As RECT

Private m_WindowState As Long
Private m_w As Long
Private m_h As Long
Private m_ActivedControl As Long

Private Sub Class_Terminate()
    Set m_Controls = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_Controls = New Collection
End Sub

Public Function RegControl(UIControl As CUIControl) As Long
    RegControl = ID + 255 + m_Controls.Count    '0-255保留
    m_Controls.Add UIControl, Hex(RegControl)
    m_ActivedControl = m_Controls.Count
End Function

Public Sub Move(ByVal x As Integer, ByVal y As Integer)
    m_WindowRect.Left = x
    m_WindowRect.Top = y
    m_WindowRect.Right = x + m_w
    m_WindowRect.Bottom = y + m_h
End Sub

Public Function GetWindowRect() As RECT
    GetWindowRect = m_WindowRect
    GetWindowRect.Top = GetWindowRect.Top + 20
End Function

Public Sub Create(x As Integer, y As Integer, w As Integer, h As Integer)
    Debug.Assert w > 0 And h > 0
    m_w = w
    m_h = h
    Move x, y
    
    '注册窗体
    Release = False
End Sub

Public Sub SendInputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    '首先处理消息
    '状态机的实现
    Dim inRequest As Long, i As Long
    inRequest = GetMsgRequest(MsgType, Value, x, y)
    m_WindowState = FSM(m_WindowState, inRequest)
    
    Select Case m_WindowState
        Case 0  '正常窗体
            SendChildMsg MsgType, Value, x, y
        Case 1  '开始拖放
            Move m_WindowRect.Left + g_Inputs.Mouse_DX, m_WindowRect.Top + g_Inputs.Mouse_DY
    End Select
End Sub

Private Sub SendChildMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpControl As CUIControl, i As Integer
    '键盘发送到激活控件
    Select Case MsgType
        Case MSG_MouseDown, MSG_MouseMove, MSG_MouseUp, MSG_MouseClick
            For i = m_Controls.Count To 1 Step -1
                 Set tmpControl = m_Controls.Item(i)
                 If tmpControl.IsInRect(x, y) Then
                    tmpControl.SendInputMsg MsgType, Value, x, y
                    If MsgType = MSG_MouseDown Then
                        'SetFouce
                        If m_ActivedControl > 0 Then
                            m_Controls.Item(m_ActivedControl).Actived = False
                        End If
                        m_ActivedControl = i
                        tmpControl.Actived = True
                    End If
                    Exit Sub
                 End If
            Next i
            
            If MsgType = MSG_MouseDown And m_ActivedControl > 0 Then
                '切换焦点
                m_Controls.Item(m_ActivedControl).Actived = False
                m_ActivedControl = 0
            End If
        Case MSG_KeyPress, MSG_Keydown, MSG_KeyUp
            '支持全局的热键
            If m_ActivedControl > 0 Then
                Set tmpControl = m_Controls.Item(m_ActivedControl)
                tmpControl.SendInputMsg MsgType, Value, x, y
            End If
    End Select
    '发送
    For i = 1 To m_Controls.Count
         Set tmpControl = m_Controls.Item(i)
         If tmpControl.Release Then
            tmpControl.SendInputMsg MsgType, Value, x, y
            Exit Sub
         End If
    Next i
    
    Set tmpControl = Nothing
End Sub

Public Sub Render()
    'First Default Draw
    With g_MainSurface.DD_Surface
        .SetFillStyle 0
        .setDrawWidth 1
        .SetFillColor RGB(100, 100, 100)
        .DrawRoundedBox m_WindowRect.Left, _
                        m_WindowRect.Top, _
                        m_WindowRect.Right, _
                        m_WindowRect.Bottom, _
                        10, _
                        10
        .DrawText m_WindowRect.Left + 5, m_WindowRect.Top + 3, Hex(ID), False
        .DrawLine m_WindowRect.Left, m_WindowRect.Top + 20, m_WindowRect.Right, m_WindowRect.Top + 20
    End With
    
    Dim i As Integer
    For i = 1 To m_Controls.Count
        m_Controls.Item(i).Render
    Next i
End Sub

Public Function IsInRect(ByVal x As Integer, ByVal y As Integer) As Boolean
    IsInRect = CheckRect(m_WindowRect, x, y)
End Function

Public Function IsInBarRect(ByVal x As Integer, ByVal y As Integer) As Boolean
    '标题栏
    Dim tmpRect As RECT
    tmpRect = m_WindowRect
    tmpRect.Bottom = tmpRect.Top + 20
    
    IsInBarRect = CheckRect(tmpRect, x, y)
End Function

Private Function GetMsgRequest(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long) As Long
    '转换条件
    If MsgType = MSG_MouseDown And IsInBarRect(x, y) Then GetMsgRequest = 1
    If MsgType = MSG_MouseUp Then GetMsgRequest = 2
End Function

Private Function FSM(oldState As Long, ByVal inRequest As Long) As Long
    '输入 oldState,条件 inRequest
    FSM = oldState
    Select Case oldState
        Case 0  '正常窗体
            If inRequest = 1 Then FSM = 1: Release = True
        Case 1  '开始拖放
            If inRequest = 2 Then FSM = 0: Release = False
    End Select
End Function
