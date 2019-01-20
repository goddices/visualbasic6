VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSoundWav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Sound As DirectSound
Private m_SoundDeviceReady As Boolean
Private m_SoundBuffers As New Collection
Private m_DirectMusicLoader As DirectMusicLoader
Private m_MusicPerform As DirectMusicPerformance
'Private m_Musicseg As DirectMusicSegment

Public Sub Destroy()
    Set m_Sound = Nothing
    Set m_SoundBuffers = Nothing
    Set m_DirectMusicLoader = Nothing
    Set m_MusicPerform = Nothing
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Function Create(hWnd As Long) As Boolean
    On Error GoTo errhand
    Set m_Sound = g_DX7.DirectSoundCreate("")
    m_Sound.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    m_SoundDeviceReady = True
    
    '---------Midi
    'Set m_DirectMusicLoader = g_DX7.DirectMusicLoaderCreate
    'Set m_MusicPerform = g_DX7.DirectMusicPerformanceCreate
    'm_MusicPerform.Init Nothing, hwnd
    'm_MusicPerform.SetPort -1, 1
    
    Create = True
    Exit Function
errhand:
    m_SoundDeviceReady = False
End Function

Public Function LoadWav(ByVal filePath As String, Optional sName As String) As Boolean
    '‘ÿ»Î…˘“Ù◊ ‘¥
    'On Error GoTo errhand:
    Dim NewSound As DirectSoundBuffer
    Dim BufferDesc As DSBUFFERDESC
    Dim WaveForm As WAVEFORMATEX
    BufferDesc.lFlags = DSBCAPS_STATIC
    Set NewSound = m_Sound.CreateSoundBufferFromFile(filePath, BufferDesc, WaveForm)
    
    If NewSound Is Nothing Then GoTo errhand:
    If sName <> "" Then m_SoundBuffers.Add NewSound, sName Else _
        m_SoundBuffers.Add NewSound
    LoadWav = True
    'add-----
    'Dim tm As DirectMusicSegmentState
    'Set m_Musicseg = m_DirectMusicLoader.LoadSegment(filePath)
    'm_Musicseg.Download m_MusicPerform
    'Set tm = m_MusicPerform.PlaySegment(m_Musicseg, 0, 0)
    'm_MusicPerform.Stop m_Musicseg, tm, 0, 0
    
    Exit Function
errhand:
    
End Function

Public Sub PlayWav(ByVal index As Long, Optional ByVal Looping As Boolean = False, Optional ByVal startPoint As Long = 0)
    Dim t_Sound As DirectSoundBuffer
    If Not m_SoundDeviceReady Then Exit Sub
    Set t_Sound = m_SoundBuffers.Item(index)
    t_Sound.SetCurrentPosition startPoint
    If Looping = True Then
       t_Sound.Play DSBPLAY_LOOPING
    Else
       t_Sound.Play DSBPLAY_DEFAULT
    End If
End Sub

Public Sub Play(ByVal sName As String, Optional ByVal Looping As Boolean = False, Optional ByVal startPoint As Long = 0)
    Dim t_Sound As DirectSoundBuffer
    If Not m_SoundDeviceReady Then Exit Sub
    Set t_Sound = m_SoundBuffers.Item(sName)
    t_Sound.SetCurrentPosition startPoint
    If Looping = True Then
       t_Sound.Play DSBPLAY_LOOPING
    Else
       t_Sound.Play DSBPLAY_DEFAULT
    End If
End Sub

Public Sub StopPlay(ByVal index As Long, endPoint As Long)
    Dim position As DSCURSORS
    Dim t_Sound As DirectSoundBuffer
    If Not m_SoundDeviceReady Then Exit Sub
    Set t_Sound = m_SoundBuffers.Item(index)
    t_Sound.Stop
    t_Sound.GetCurrentPosition position
    endPoint = position.lPlay
End Sub

Public Sub Clear()
    Dim i As Integer
    For i = m_SoundBuffers.Count To 1 Step -1
        m_SoundBuffers.Remove i
    Next i
End Sub

Public Sub Remove()
    'ReDim m_SoundBuffers(0)
End Sub
