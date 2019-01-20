VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Public Name As String
Public MinFrame As Integer
Public MaxFrame As Integer
Public DelayTime As Long
Public Surface As CSurface
Public x As Integer                 '地图上的实际坐标
Public y As Integer
Public Done As Boolean
Public PlayTimes As Integer
Public LoopPlay As Boolean
Public CurrentFrame As Integer
Public PlayedTimes As Integer

Private mOldPaintTime As Long
Private mAutoDestroy As Boolean
Private mFlash As Boolean           '闪烁标志
Private mFlashDelayTime As Long
Private mFlashStartTime As Long
Private mFlashFrameTime As Long
Private mFlashOldTime As Long

Public Sub Destroy()
    Set Surface = Nothing
End Sub

Private Sub Class_Initialize()
    mAutoDestroy = True
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Property Get BltReady() As Boolean
    BltReady = GetTickCount - mFlashOldTime > mFlashFrameTime
End Property

Public Sub Render(Optional x0 As Integer, Optional y0 As Integer, Optional Default As Boolean = True)
    '动画绘制
    Dim NewTime As Long
    If Done Then Exit Sub
    NewTime = GetTickCount
    If NewTime - mOldPaintTime > DelayTime Then
        CurrentFrame = CurrentFrame + 1
        If mAutoDestroy And CurrentFrame > MaxFrame Then
            '销毁它
            PlayedTimes = PlayedTimes + 1
            If PlayedTimes > 255 Then PlayedTimes = 0  '重新开始
            'Play?????
            CurrentFrame = MinFrame - 1
            If Not LoopPlay And PlayedTimes >= PlayTimes Then
                Done = True
                Exit Sub
            End If
            CurrentFrame = CurrentFrame + 1
            'Done = True
        End If
        mOldPaintTime = NewTime
    End If
    If CurrentFrame < MinFrame Then CurrentFrame = MinFrame         '防止越界情况
    If CurrentFrame > MaxFrame Then CurrentFrame = MaxFrame
    
    If Not mFlash And Default Then g_MainSurface.Blt Surface, x - x0, y - y0, CurrentFrame
    '支持闪烁
    If NewTime - mFlashStartTime > mFlashDelayTime Then mFlash = False: Exit Sub
    If NewTime - mFlashOldTime > mFlashFrameTime Then
        mFlashOldTime = NewTime
        mFlash = Not mFlash
    End If
End Sub

Public Sub Flash(FrameTime As Long, DelayTime As Long)
    mFlashStartTime = GetTickCount
    mFlashFrameTime = FrameTime
    mFlashDelayTime = DelayTime
    mFlash = True
End Sub

Public Sub Play()
    Done = False
    PlayedTimes = 0
    CurrentFrame = MinFrame - 1                  '先-1因为播放开始时，自动加一
End Sub
