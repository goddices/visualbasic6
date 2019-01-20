VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnimates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------动画播放--------------------------------
'.AddAnimates  Name,Surface,MinFrame,MaxFrame,DelayTime 加入源动画
'.Item Name     返回一个 New CAnimate
'.Play          置播放开始位置
'.Refresh       将动画绘到主缓冲区
'--------------------------------------------------------------------

Private m_Animates As New Collection                '资源

Private m_PlayingAnimates As New Collection

Public Sub Destroy()
    'Clear
    Set m_Animates = Nothing
    Set m_PlayingAnimates = Nothing
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Function Count() As Integer
    Count = m_Animates.Count
End Function

Public Function ItemByIndex(ByVal index As Integer) As CAnimate
    '没有检测输入---
    Set ItemByIndex = m_Animates.Item(index)
End Function

Public Function ItemByName(ByVal Name As String, Optional StartAddress As Integer = 1) As CAnimate
    Dim i As Integer
    For i = StartAddress To Count
        If m_Animates.Item(i).Name = Name Then
            Set ItemByName = m_Animates.Item(i)
            Exit Function
        End If
    Next i
End Function

Public Function Item(ByVal Name As String, Optional NewCreate As Boolean = False) As CAnimate
    Dim NewAnimate As New CAnimate
    On Error GoTo errhand:
    If Not NewCreate Then Set Item = m_Animates.Item(Name): Exit Function
    '若NewCreate=1则新建一个，否则返回一个共享
    With NewAnimate
        Set .Surface = m_Animates.Item(Name).Surface
        .DelayTime = m_Animates.Item(Name).DelayTime
        .MinFrame = m_Animates.Item(Name).MinFrame
        .MaxFrame = m_Animates.Item(Name).MaxFrame
    End With
    Set Item = NewAnimate
    Exit Function
errhand:
    Set Item = Nothing
End Function

Public Sub Play(ByVal Name As String, ByVal x As Integer, ByVal y As Integer, _
    Optional ByVal PlayTimes As Integer = 1, Optional ByVal LoopPlay As Boolean = False)
    Dim NewAnimate As CAnimate
    Set NewAnimate = Item(Name)
    If NewAnimate Is Nothing Then g_Err_Description = "未加载动画！": Exit Sub

    With NewAnimate
        .x = x
        .y = y
        .PlayTimes = PlayTimes
        .LoopPlay = LoopPlay
    End With
    NewAnimate.Play
    m_PlayingAnimates.Add NewAnimate
End Sub

Public Function AddAnimates(Name As String, SrcSurface As CSurface, _
    MinFrame As Integer, MaxFrame As Integer, ByVal DelayTime As Long) As Boolean
    '载入动画
    Dim NewAnimate As New CAnimate
    On Error GoTo errhand:
    If SrcSurface Is Nothing Then Exit Function
    
    With NewAnimate
        Set .Surface = SrcSurface
        .Name = Name
        .MinFrame = MinFrame
        .MaxFrame = MaxFrame
        .DelayTime = DelayTime
        .PlayTimes = 1
        .LoopPlay = False
    End With
    m_Animates.Add NewAnimate, Name
    AddAnimates = True
    Exit Function
errhand:
    
End Function

Public Sub Refresh(Optional x As Integer, Optional y As Integer)
    Dim i As Integer
    'x,y 为地图上可见区域的左上位置
    For i = m_PlayingAnimates.Count To 1 Step -1
        m_PlayingAnimates(i).Render x, y
        g_MainSurface.TextOut m_PlayingAnimates(i).x - x, m_PlayingAnimates(i).y - y, Str(i), vbRed
        If m_PlayingAnimates(i).Done Then
            'Kill it
            m_PlayingAnimates.Remove i
        End If
    Next i
End Sub

Public Sub Clear()
    Dim i As Integer
    For i = m_PlayingAnimates.Count To 1 Step -1
        m_PlayingAnimates.Remove i
    Next i
    For i = m_Animates.Count To 1 Step -1
        m_Animates.Remove i
    Next i
End Sub

