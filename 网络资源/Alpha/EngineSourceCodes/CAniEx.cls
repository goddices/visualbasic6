VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAniEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type PlayingFrame
    startedtime As Long
    ani As AniNode
End Type

Private m_TimeStick As Long                         '时间轴,当前播放桢数
Private m_PlaySpeed As Integer                      '播放速度
Private m_srcSurfaces() As CSurface                 '所有的图象资源
Private m_AllPlayingFrames() As PlayingFrame        '所有的正在显示的动画桢
Private m_PlayingFrameCount As Integer
Private m_Static As Byte                            '播放状态

Private i As Integer

Public Sub Play(ByVal x As Integer, ByVal y As Integer)
    m_TimeStick = 0
End Sub

Public Sub Render()
    Static oldpainttime As Long
    If GetTickCount - oldpainttime < m_PlaySpeed Then Exit Sub
    For i = 0 To m_PlayingFrameCount - 1
        If GetTickCount - m_AllPlayingFrames(i).startedtime > m_AllPlayingFrames(i).ani.delayTime Then
            'kill it
        Else
            'Show m_AllPlayingFrames(i).ani
        End If
    Next i
    m_TimeStick = m_TimeStick + 1
    oldpainttime = GetTickCount
End Sub

Private Sub AddNewFrame(ByVal frameindex As Integer, ByVal x As Integer, ByVal y As Integer, _
    ByVal delay As Integer, ByVal blt_mode As Byte)
    
End Sub

