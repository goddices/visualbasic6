VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHotRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Caption As String
Public Style As Integer
Public Enable As Boolean
Public Visiable As Boolean
Public index As Integer
Public Actived As Boolean
Public ChildNum As Integer
Public Parent As CHotRect
Public ActivedRect As CHotRect
Public SelectIndex As Integer


'Public Width As Integer
'Public Height As Integer
Private m_Static As DefMsgType
Private m_SendPaintMessage As Boolean
Private m_SendInputMessage As Boolean
Private m_Childen() As CHotRect
Private m_Rect As RECT
Private m_RectStack() As CHotRect           '读取文件时
Private m_RectNumStack() As Integer

Public Event InputRequire(ByVal Name As String, ByVal MsgStyle As DefMsgType, Default As Boolean)
Public Event PaintRequire(ByVal Name As String, ByVal MsgStyle As DefMsgType, Default As Boolean)

Private Sub Class_Terminate()
    Dim i As Integer

    For i = 0 To ChildNum - 1
        Set m_Childen(i) = Nothing
    Next i
    
    Set Parent = Nothing
    Set ActivedRect = Nothing
    ChildNum = 0
    ReDim m_Childen(ChildNum)
    ReDim m_RectNumStack(0)
    ReDim m_RectStack(0)
    Set m_RectStack(0) = Nothing
End Sub


Public Property Let SendPaintMsg(ByVal vNewValue As Boolean)
    Dim i As Integer
    For i = 0 To ChildNum - 1
        m_Childen(i).SendPaintMsg = vNewValue
    Next i
    m_SendPaintMessage = vNewValue
End Property

Public Property Let SendInputMsg(ByVal vNewValue As Boolean)
    Dim i As Integer
    For i = 0 To ChildNum - 1
        m_Childen(i).SendInputMsg = vNewValue
    Next i
    m_SendInputMessage = vNewValue
End Property

Public Property Get Left() As Integer
    Left = m_Rect.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    m_Rect.Left = vNewValue
End Property

Public Property Get Top() As Integer
    Top = m_Rect.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    m_Rect.Top = vNewValue
End Property

Public Property Get Right() As Integer
    Right = m_Rect.Right
End Property

Public Property Let Right(ByVal vNewValue As Integer)
    m_Rect.Right = vNewValue
End Property

Public Property Get Bottom() As Integer
    Bottom = m_Rect.Bottom
End Property

Public Property Let Bottom(ByVal vNewValue As Integer)
    m_Rect.Bottom = vNewValue
End Property

Public Property Get width() As Integer
    width = m_Rect.Right - m_Rect.Left
End Property

Public Property Get height() As Integer
    height = m_Rect.Bottom - m_Rect.Top
End Property

Public Function GetViewRect() As RECT
    GetViewRect = m_Rect
End Function

Private Sub Class_Initialize()
    Style = 0
    Enable = True
    Visiable = True
    Actived = False
    ChildNum = 0
    m_SendPaintMessage = False
    m_SendInputMessage = False
    ReDim m_Childen(ChildNum)
End Sub

Public Sub Render()
    '绘图
    Dim i As Integer, PaintDefault As Boolean
    
    'On Error Resume Next
    For i = 0 To ChildNum - 1
        m_Childen(i).Render
    Next i

    'Draw A HotRect
    If Parent Is Nothing Then Exit Sub
    
    If Parent.Actived Then
        '请求区域重画
        If m_SendPaintMessage Then
            'g_MsgPost.PostMessage "Paint"
            'g_MsgPost.PostMessage Name
        End If
        
        PaintDefault = True
        Call Parent.ChildPaintRequire(Name, m_Static, PaintDefault)
        
        If PaintDefault Then
            g_MainSurface.DD_Surface.DrawBox m_Rect.Left, m_Rect.Top, m_Rect.Right, m_Rect.Bottom
            g_MainSurface.TextOut m_Rect.Left + 5, m_Rect.Top, Caption
            
            If Actived Then
                g_MainSurface.SetForeColor vbWhite
                g_MainSurface.DD_Surface.setDrawStyle 2
                g_MainSurface.DD_Surface.DrawBox _
                    m_Rect.Left + 1, m_Rect.Top + 1, m_Rect.Right - 1, m_Rect.Bottom - 1
                g_MainSurface.DD_Surface.setDrawStyle 0
            End If
        End If
    End If
End Sub

Public Function GetChildByName(sName As String) As CHotRect
    '绘图
    Dim i As Integer
    If Name = sName Then Set GetChildByName = Me: Exit Function
    For i = 0 To ChildNum - 1
        Set GetChildByName = m_Childen(i).GetChildByName(sName)
        If Not GetChildByName Is Nothing Then
            Exit Function
        End If
    Next i
End Function

Public Function Child(index As Integer) As CHotRect
    If ChildNum > 0 And index < ChildNum Then Set Child = m_Childen(index)
End Function

Public Sub Move(ByVal x As Integer, ByVal y As Integer)
    Dim dx As Integer, dy As Integer, i As Integer
    dx = x - m_Rect.Left
    dy = y - m_Rect.Top
    m_Rect.Left = x
    m_Rect.Top = y
    m_Rect.Right = m_Rect.Right + dx
    m_Rect.Bottom = m_Rect.Bottom + dy
    
    For i = 0 To ChildNum - 1
        m_Childen(i).Move dx + m_Childen(i).Left, dy + m_Childen(i).Top
    Next i
End Sub

Public Sub TreatMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Integer
    '如何实现按照层次的遍历？？？
    If Me.Actived And Me.ChildNum > 0 Then
        '遍历所有子
        For i = 0 To ChildNum - 1
            m_Childen(i).TreatMsg MsgType, Value, x, y
        Next i
    End If
    '处理输入的消息
    If Parent Is Nothing Then Exit Sub
    
    Dim Default As Boolean
    Default = True
    m_Static = MsgType
    
    Call Parent.ChildInputRequire(Name, MsgType, Default)
    If Default Then
        '默认输入处理,Space/Click选定,MouseMove的实现
        If MsgType = MSG_MouseClick And CheckRect(x, y) Then
            '单击选中事件
            Me.SetFouse
            If m_SendInputMessage Then
                'g_MsgPost.PostMessage "Click"
                'g_MsgPost.PostMessage Name
            End If
        End If
    End If
 End Sub

Public Sub TabChange(Optional IndexAdd As Boolean = True)
    '切换焦点
    If IndexAdd Then
        If SelectIndex >= ChildNum - 1 Then SelectIndex = 0 Else SelectIndex = SelectIndex + 1
    Else
        If SelectIndex <= 0 Then SelectIndex = ChildNum - 1 Else SelectIndex = SelectIndex - 1
    End If
    m_Childen(SelectIndex).SetFouse
End Sub

Public Function CheckRect(ByVal x As Integer, ByVal y As Integer) As Boolean
    '检测点是否在区域内
    CheckRect = (x > m_Rect.Left And x < m_Rect.Right) _
        And (y > m_Rect.Top And y < m_Rect.Bottom)
End Function

Public Sub SetFouse()
    '得到焦点,底层优先
    If Not Parent Is Nothing Then
        If Not Parent.ActivedRect Is Nothing Then Parent.ActivedRect.LostFouse
        Set Parent.ActivedRect = Me
        Parent.SelectIndex = index
        SelectIndex = index
    End If
    Actived = True
End Sub

Public Sub LostFouse()
    '失去焦点
    Actived = False
    If Not ActivedRect Is Nothing Then
        ActivedRect.LostFouse
        Set ActivedRect = Nothing
    End If
End Sub

Public Sub AddChild(Child As CHotRect)
    Set m_Childen(ChildNum) = Child
    Set Child.Parent = Me
    Child.index = ChildNum
    ChildNum = ChildNum + 1
    ReDim Preserve m_Childen(ChildNum)
End Sub

Public Sub Create(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer)
    With m_Rect
        .Top = Top
        .Left = Left
        .Right = Right
        .Bottom = Bottom
    End With
End Sub

Public Sub LoadRects(FileName As String)
    Dim NewRect As CHotRect
    Dim WorkRect As CHotRect
    Dim i As Integer, Name As String, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer
    Dim childNums As Integer
    ReDim m_RectNumStack(0)
    ReDim m_RectStack(0)
    Open FileName For Input As #1
    While Not EOF(1)
        Input #1, Name, Left, Top, Right, Bottom, childNums
        Set NewRect = New CHotRect
        NewRect.Create Left, Top, Right, Bottom
        'Add to WorkRect
        If Not WorkRect Is Nothing Then
            
            If Not SubStackTop Then Pop WorkRect, 0
            WorkRect.AddChild NewRect
        End If
        If childNums > 0 Then
            If WorkRect Is Nothing Then Set NewRect = Me
            Push NewRect, childNums
            Set WorkRect = NewRect
        End If
        NewRect.Name = Name
        NewRect.Caption = Name
    Wend
    Close #1
    Set NewRect = Nothing
    Set WorkRect = Nothing
End Sub

Private Sub Push(mRect As CHotRect, ChildNum As Integer)
    Dim StackLen As Integer
    StackLen = UBound(m_RectNumStack())
    m_RectNumStack(StackLen) = ChildNum
    Set m_RectStack(StackLen) = mRect
    ReDim Preserve m_RectNumStack(StackLen + 1)
    ReDim Preserve m_RectStack(StackLen + 1)
End Sub

Private Sub Pop(mRect As CHotRect, ChildNum As Integer)
    Dim StackLen As Integer
    StackLen = UBound(m_RectNumStack())
    ReDim Preserve m_RectNumStack(StackLen - 1)
    ReDim Preserve m_RectStack(StackLen - 1)
    Set mRect = m_RectStack(StackLen - 1)
    ChildNum = m_RectNumStack(StackLen - 1)
End Sub

Private Function SubStackTop() As Boolean
    Dim StackLen As Integer
    SubStackTop = True
    StackLen = UBound(m_RectNumStack()) - 1
    m_RectNumStack(StackLen) = m_RectNumStack(StackLen) - 1
    If m_RectNumStack(StackLen) <= 0 Then SubStackTop = False
End Function

Public Function GetBottomActiveRect() As CHotRect
    '得到最底层的激活区域
    If ChildNum > 0 Then
        If ActivedRect Is Nothing Then Set GetBottomActiveRect = Me: Exit Function
        Set GetBottomActiveRect = ActivedRect.GetBottomActiveRect
    Else
        Set GetBottomActiveRect = Me
    End If
End Function

Public Sub ChildPaintRequire(ByVal Name As String, ByVal MsgStyle As Integer, Default As Boolean)
    RaiseEvent PaintRequire(Name, MsgStyle, Default)
End Sub

Public Sub ChildInputRequire(ByVal Name As String, ByVal MsgStyle As DefMsgType, Default As Boolean)
    RaiseEvent InputRequire(Name, MsgStyle, Default)
End Sub
