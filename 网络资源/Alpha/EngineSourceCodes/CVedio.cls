VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVedioAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Used ActiveMovie Control Type Library in Quartz.Dll
Public m_objMediaControl As FilgraphManager

Private m_objVedioWindow As IVideoWindow
Private m_objBasicAudio  As IBasicAudio
Private m_objMediaPosition As IMediaPosition

Private Enum PlayType
    m_Vedio
    m_Audio
End Enum

Public Sub OpenMedia(path As String)
    Set m_objMediaControl = New FilgraphManager
    m_objMediaControl.RenderFile path
    Set m_objBasicAudio = m_objMediaControl
    Set m_objMediaPosition = m_objMediaControl
End Sub

Public Sub SetVedioRect(hWnd As Long, vedioRect As RECT)
    Set m_objVedioWindow = m_objMediaControl
    
    'm_objVedioWindow.FullScreenMode = Not g_Windowed
    m_objVedioWindow.WindowStyle = CLng(&H6000000)
    m_objVedioWindow.Top = vedioRect.Top
    m_objVedioWindow.Left = vedioRect.Left
    m_objVedioWindow.Width = vedioRect.Right - vedioRect.Left
    m_objVedioWindow.Height = vedioRect.Bottom - vedioRect.Top
    
    m_objVedioWindow.Owner = hWnd
End Sub

Public Sub PlayMedia(Optional mode As Boolean = True)
    m_objMediaControl.Run
    If mode Then
        While CurrentPosition < Round(m_objMediaPosition.Duration, 2)
            DoEvents
        Wend
    End If
    m_objMediaControl.Stop
End Sub

Public Sub StopMedia()
    m_objMediaControl.Stop
End Sub

Public Function CurrentPosition() As Single
    CurrentPosition = m_objMediaPosition.CurrentPosition
    Debug.Print CurrentPosition
End Function

Public Sub SetVolume(ByVal Vol As Long)
    m_objBasicAudio.Volume = Vol
End Sub

Public Sub Destroy()
    Set m_objBasicAudio = Nothing
    Set m_objVedioWindow = Nothing
    Set m_objMediaPosition = Nothing
    Set m_objMediaControl = Nothing
End Sub
