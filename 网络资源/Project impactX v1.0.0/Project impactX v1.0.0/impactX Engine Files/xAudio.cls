VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'impactX Game Engine
'本类模块用于对WAV,MIDI格式的声音进行回放和处理
'使用本类模块必须遵守:
'你可以免费使用本引擎及代码
'使用本引擎后的责任由使用者承担
'你可以任意拷贝本引擎代码，但必须保证其完整性
'希望我能得到你使用本引擎制作出的程序
'Davy.xu sunicdavy@sina.com qq:20998333
Option Explicit
Dim dx As New DirectX8
Dim DS As DirectSound8
Dim DMA As DMUS_AUDIOPARAMS
Dim DAperformance As DirectMusicPerformance8    '播放器
Dim DAloader As DirectMusicLoader8              '载入器
Dim dmPath As DirectMusicAudioPath8 '媒体路径,做调节音量用
Dim m_PausePos As Long '停止位置(待修正)
'功能:初始化DirectAudio
Public Function InitDXAudio(hWnd As Long) As Boolean
    On Error GoTo ErrH
    
    Set DAloader = dx.DirectMusicLoaderCreate
    Set DAperformance = dx.DirectMusicPerformanceCreate
    
    DAperformance.InitAudio hWnd, DMUS_AUDIOF_ALL, DMA, Nothing, DMUS_APATH_DYNAMIC_STEREO, 64
        Set dmPath = DAperformance.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_STEREO, 64, True)
        InitDXAudio = True
    Exit Function
ErrH:
    Debug.Print "Err:[InitDXAudio] 初始化错误"
    InitDXAudio = False
    
End Function
'功能:初始化DirectAudio的WAVE处理部分
Public Function InitDXSound(hWnd As Long) As Boolean
    InitDXSound = False
    Set DS = dx.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    InitDXSound = True
End Function
Public Function LoadWav(Pathname As String) As DirectSoundSecondaryBuffer8
    On Error GoTo ErrH
        Dim DSbufSC As DSBUFFERDESC
        Pathname = Trim(Pathname)
        If Len(Pathname) = 0 Then
            Debug.Print "Err [LoadWav] 路径为空"
            End
        End If
        
        If Mid(Pathname, 2, 1) <> ":" Then Pathname = App.Path & "\" & Pathname
        If LCase(Right(Pathname, 3)) <> "wav" And LCase(Right(Pathname, 3)) <> "mid" Then
             Debug.Print "Err [LoadWav] 载入格式不正确,只能载入wav文件"
             End
        End If
        If Len(Dir(Pathname)) = 0 Then
            Debug.Print "Err:[LoadWav] 文件不存在"
            Debug.Print Pathname
            End
    End If
        DSbufSC.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_CTRLPOSITIONNOTIFY
        Set LoadWav = DS.CreateSoundBufferFromFile(Pathname, DSbufSC)
        Exit Function
ErrH:
    Debug.Print "Err [LoadWav] 载入错误"
    Debug.Print Pathname
End Function

'功能:载入音乐文件
'参数:音乐缓冲索引,路径.没有盘符的路径自动识别为工作目录
Public Function LoadAudio(Pathname As String) As DirectMusicSegment8
    On Error GoTo ErrH
    Pathname = Trim(Pathname)
    If Len(Pathname) = 0 Then
        Debug.Print "Err [LoadAudio] 路径为空"
        End
    End If
    
    If Mid(Pathname, 2, 1) <> ":" Then Pathname = App.Path & "\" & Pathname
    If LCase(Right(Pathname, 3)) <> "wav" And LCase(Right(Pathname, 3)) <> "mid" Then
         Debug.Print "Err [LoadAudio] 载入格式不正确,只能载入wav和mid文件"
         End
    End If
    If Len(Dir(Pathname)) = 0 Then
        Debug.Print "Err:[LoadAudio] 文件不存在"
        Debug.Print Pathname
        End
    End If
    Set LoadAudio = DAloader.LoadSegment(Pathname)
    LoadAudio.Download dmPath
    Exit Function
ErrH:
    Debug.Print "Err [LoadAudio] 载入错误 "
    Debug.Print Pathname
    Debug.Print "在非NT系统中(如Win98),请不要在路径中带有中文"
End Function
'功能: 播放索引号对应音乐缓冲里的音乐
Public Sub PlayAudio(Buf As DirectMusicSegment8, Optional isRepeat As Boolean = False)
    On Error GoTo ErrH
    If isRepeat Then
        Buf.SetRepeats INFINITE
    End If
    DAperformance.PlaySegmentEx Buf, DMUS_SEGF_SECONDARY, 0, Nothing, dmPath
    Exit Sub
ErrH:
    Debug.Print "Err [PlayAudio] 播放时错误"
End Sub
'功能: 播放索引号对应音乐缓冲里的音乐
Public Sub PlayWav(Buf As DirectSoundSecondaryBuffer8, Optional isRepeat As Boolean = False)
    On Error GoTo ErrH
    Buf.SetCurrentPosition 0
    If isRepeat Then
        Buf.Play DSBPLAY_LOOPING
    Else
        Buf.Play DSBPLAY_DEFAULT
    End If
    Exit Sub
ErrH:
    If Buf Is Nothing Then
        Debug.Print "Err [PlayWav] 没有载入音乐，播放时错误"
    Else
        Debug.Print "Err [PlayWav] 播放时错误 "
    End If
    
End Sub
'功能:停止播放音乐
Public Sub StopWav(Buf As DirectSoundSecondaryBuffer8)
   On Error GoTo ErrH
    Buf.Stop
    Exit Sub
ErrH:
    Debug.Print "Err [StopWav] 停止时错误"
End Sub

'功能:停止播放音乐
Public Sub StopAudio(Buf As DirectMusicSegment8)
    On Error GoTo ErrH
    m_PausePos = Buf.GetStartPoint
    DAperformance.StopEx Buf, 0, 0
    Exit Sub
ErrH:
    Debug.Print "Err [StopAudio] 停止时错误 "
End Sub
'功能:设置Wav音乐音量
'参数:范围(0~100)
Public Sub SetWavVolume(Buf As DirectSoundSecondaryBuffer8, Volume As Integer)
    If Volume < 0 Or Volume > 100 Then Exit Sub
    Buf.SetVolume Volume * 30 - 3000
End Sub
'功能:设定声音左右平衡度
'参数:范围(左)-10~10(右)
Public Sub SetWavPan(Buf As DirectSoundSecondaryBuffer8, Lev As Integer)
    If Lev < -10 Or Lev > 10 Then Exit Sub
    Buf.SetPan ((Lev + 10) * 5 - 50) * 100
End Sub
'功能:设置音乐音量
'参数:范围(0~100)
Public Sub SetAudioVolume(Vol As Integer)
    If Vol < 0 Or Vol > 100 Then Exit Sub
    dmPath.SetVolume -(1 - Vol / 100) * 5000, 0
End Sub
'功能:音乐是否在播放
Public Function IsWavPlaying(Buf As DirectSoundSecondaryBuffer8) As Boolean
    IsWavPlaying = IIf(Buf.GetStatus = DSBSTATUS_PLAYING, True, False)
End Function

'功能:音乐是否在播放
Public Function IsAudioPlaying(Buf As DirectMusicSegment8) As Boolean
    IsAudioPlaying = DAperformance.isPlaying(Buf, Nothing)
End Function
'功能:设定声音左右平衡度
'参数:范围(左)-10~10(右)
Public Sub SetAudioBalance(Lev As Integer)
    If Lev < -10 Or Lev > 10 Then Exit Sub
    Dim DSbuf As DirectSound3DBuffer8
    Set DSbuf = dmPath.GetObjectinPath(DMUS_PCHANNEL_ALL, DMUS_PATH_BUFFER, 0, vbNullString, 0, "IID_IDirectSound3DBuffer")
    DSbuf.SetPosition Lev / 5, 0, 0, DS3D_IMMEDIATE
    Set DSbuf = Nothing
End Sub
'卸载DirectAudio
Public Sub UnloadDXAudio()
    On Error GoTo ErrH
        Dim i As Long
        DAperformance.CloseDown '关闭DirectMusicPerformance8
        Set DAperformance = Nothing
        Set DAloader = Nothing
    Set DS = Nothing
    Exit Sub
ErrH:
    Debug.Print "Err [UnloadDXAudio] 卸载错误"
End Sub
'卸载DirectAudio
Public Sub UnloadDXSound()
    Set DS = Nothing
End Sub

Public Sub ReleaseWav(Buf As DirectSoundSecondaryBuffer8)
    Set Buf = Nothing
End Sub
Public Sub ReleaseAudio(Buf As DirectMusicSegment8)
    Set Buf = Nothing
End Sub
