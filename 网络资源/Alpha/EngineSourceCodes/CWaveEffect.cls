VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWaveEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_lpScreenBuffer As Long                        '首地址
Private m_lpScreenBufferPitch As Long
Private m_DataBuffer() As Long                        '波的能量区域
Private m_OldBuffer() As Long
Private m_lpDataBuffer As Long
Private m_lpOldBuffer As Long
Private m_DataBufferPitch As Long
Private m_with As Integer
Private m_height As Integer
Private m_DataReady As Boolean
Private m_test As Long
Private m_WaveRect As RECT

Private Const PI = 3.141592

Public Sub CreateWaveRect(ByVal lpscreen As Long, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, ByVal pitch As Long)
    m_lpScreenBuffer = lpscreen
    m_lpScreenBufferPitch = pitch
    '创建数据区域
    ReDim m_DataBuffer(w - 1, h - 1)
    ReDim m_OldBuffer(w - 1, h - 1)
    m_lpDataBuffer = VarPtr(m_DataBuffer(0, 0))
    m_lpOldBuffer = VarPtr(m_OldBuffer(0, 0))
    m_DataBufferPitch = w
    m_with = w
    m_height = h
    m_DataReady = True
    'Add Data
    m_test = m_lpDataBuffer
    'm_DataBuffer(200, 200) = 200
    m_WaveRect.Left = x
    m_WaveRect.Top = y
    m_WaveRect.Right = x + w
    m_WaveRect.Bottom = y + h
End Sub

Public Sub Render()
    renderipple ByVal g_MainSurface.DataAddress + m_WaveRect.Left + m_WaveRect.Top * g_MainSurface.DataPitch / 2, g_MainSurface.DataPitch / 2, _
        ByVal m_lpScreenBuffer + m_WaveRect.Left + m_WaveRect.Top * m_lpScreenBufferPitch / 2, m_lpScreenBufferPitch / 2, ByVal m_lpDataBuffer, _
        m_with, m_height
End Sub

Public Sub Blur()
    'blur_mmx ByVal g_MainSurface.DataAddress, g_MainSurface.DataPitch, _
        ByVal g_MainSurface.DataAddress, g_MainSurface.DataPitch, 640, 480
    blur_c ByVal g_MainSurface.DataAddress, g_MainSurface.DataPitch, _
        1, 1, 640 - 2, 480 - 2
End Sub

Public Sub SetPoint(ByVal x As Integer, ByVal y As Integer, ByVal a As Integer)
    Dim i As Integer, j As Integer
    x = x - m_WaveRect.Left
    y = y - m_WaveRect.Top
    On Error Resume Next
    For i = -2 To 2
        For j = -2 To 2
            m_DataBuffer(x + i, y + j) = m_DataBuffer(x + i, y + j) - a
        Next
    Next
End Sub

Public Sub SetPoint2(ByVal x As Integer, ByVal y As Integer, ByVal r As Integer, ByVal a As Integer, Optional n As Integer = 10, Optional dt As Single)
    Dim i As Integer, j As Integer
    On Error Resume Next
    x = x - m_WaveRect.Left
    y = y - m_WaveRect.Top
    For i = -r To r
        For j = -r To r
            If Sqr(i ^ 2 + j ^ 2) < r Then
                m_DataBuffer(x + i, y + j) = m_DataBuffer(x + i, y + j) + Cos(dt + n * Sqr(i ^ 2 + j ^ 2) / r) * a
            End If
        Next
    Next
End Sub

Public Sub SetPoint3()
    Dim i As Integer, j As Integer
    On Error Resume Next
    For j = 0 To m_height - 1
        For i = 0 To m_with - 1
            m_DataBuffer(i, j) = 100
            m_OldBuffer(i, j) = 0
        Next
    Next
End Sub
Private Sub LoadSinTable()

End Sub

Public Sub Spread()
    '移动水波
    'fastmemset ByVal m_lpDataBuffer, CLng(640) * CLng(480), Rnd * 2 + 1
    ripplespread ByVal m_lpDataBuffer, ByVal m_lpOldBuffer, m_with, m_height
    Dim i As Long
    i = m_lpDataBuffer                              '交换数据指针
    m_lpDataBuffer = m_lpOldBuffer
    m_lpOldBuffer = i
    '
End Sub
