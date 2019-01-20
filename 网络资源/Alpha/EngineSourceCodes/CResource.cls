VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----提供图象资源、声音资源、动画资源
Private m_Sounds As CSoundWav       '初始化设置m_Sound为Engins->Sound
Private m_Animates As New CAnimates
Private m_Surfaces As New Collection

Private m_ResTableFileNum As Long

'处理动态加载资源

Public Sub Destroy()
    'Clear
    Set m_Animates = Nothing
    Set m_Surfaces = Nothing
    Set m_Sounds = Nothing
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Private Sub Class_Initialize()
    Set m_Sounds = g_Sounds
End Sub

Public Property Get PicNum() As Integer
    PicNum = m_Surfaces.Count
End Property

Public Property Get Sound() As CSoundWav
    Set Sound = m_Sounds
End Property

Public Property Let Sound(NewSound As CSoundWav)
    Set m_Sounds = NewSound         '初始化设置m_Sound为Engins->Sound
End Property

Public Function Animates() As CAnimates
    Set Animates = m_Animates
End Function

Public Function ItemPic(Name As String) As CSurface
    Set ItemPic = m_Surfaces.Item(Name)
End Function

Public Function Pics(ByVal index As Integer) As CSurface
    Set Pics = m_Surfaces.Item(index + 1)  '0
End Function

Public Function LoadWav(ByVal filePath As String, Optional sName As String) As Boolean
    '载入声音资源
    If m_Sounds Is Nothing Then Exit Function
    LoadWav = m_Sounds.LoadWav(filePath, sName)
    LoadWav = True
End Function

Public Function LoadPic(ByVal filePath As String, Optional sName As String) As Boolean
    Dim NewSurface As New CSurface
    '载入图象资源
    On Error GoTo errhand:
    NewSurface.LoadJPG filePath
    If sName <> "" Then
        '防止重复添加
        m_Surfaces.Add NewSurface, sName
    Else
        m_Surfaces.Add NewSurface
    End If
    LoadPic = True
    Set NewSurface = Nothing
    Exit Function
errhand:
    LoadPic = False
End Function

Public Function LoadAnimate(ByVal srcPic As String, ByVal Name As String, _
    ByVal MinFrame As Integer, ByVal MaxFrame As Integer, ByVal DelayTime As Long) As Boolean
    '载入图象资源
    On Error GoTo errhand:
    LoadAnimate = m_Animates.AddAnimates(Name, m_Surfaces.Item(srcPic), MinFrame, MaxFrame, DelayTime)
    Exit Function
errhand:
    LoadAnimate = False
End Function

Public Function LoadRes(ByVal FileName As String, Optional Appended As Boolean = False) As Boolean
    '文件格式
    On Error GoTo errhand:
    Dim PicNum As Integer, WavNum As Integer, AniNum As Integer
    Dim i As Integer, PicPath As String, PicName As String
    Dim WavPath As String, WavName As String, AniSrcPic As String, AniName As String
    Dim MinFrame As Integer, MaxFrame As Integer, DelayTime As Long
    Dim FileNumber As Long
    
    If Not Appended Then Clear              '清除旧的资源
    FileNumber = FreeFile
    
    Open FileName For Input As #FileNumber
        Input #FileNumber, PicNum, WavNum, AniNum
        '
        For i = 1 To PicNum
            Input #FileNumber, PicPath, PicName
            LoadPic App.Path + PicPath, PicName
        Next i
        For i = 1 To WavNum
            Input #FileNumber, WavPath, WavName
            LoadWav App.Path + WavPath, WavName
        Next i
        For i = 1 To AniNum
            Input #FileNumber, AniSrcPic, AniName, MinFrame, MaxFrame, DelayTime
            LoadAnimate AniSrcPic, AniName, MinFrame, MaxFrame, DelayTime
        Next i
    
    Close #FileNumber
    LoadRes = True
    Exit Function
errhand:
    LoadRes = False
End Function

Public Sub Clear()
    Dim i As Integer
    For i = m_Surfaces.Count To 1 Step -1
        m_Surfaces.Remove i
    Next i
    m_Sounds.Clear
    m_Animates.Clear
End Sub

'---------------New-----------

Public Function OpenResTable(ByVal FileName As String, Optional AutoLoad As Boolean = False) As Boolean
    Dim s As String, Parameters() As String, i As Long
    Dim Name As String, ID As Long, Style As Long, ParametersNum As Long
    On Error GoTo errhand:
    If m_ResTableFileNum > 0 Then Close #m_ResTableFileNum
    
    m_ResTableFileNum = FreeFile
    
    Open FileName For Binary As #m_ResTableFileNum
    If AutoLoad Then
        While Not EOF(m_ResTableFileNum)
            Input #m_ResTableFileNum, Name, ID, Style, ParametersNum
            'Add TableNode
            ReDim Parameters(ParametersNum - 1)
            'Read Parameters
            For i = 0 To ParametersNum - 1
                Input #m_ResTableFileNum, Parameters(i)
            Next i
            'Load Res
            Select Case Style
                Case 0
                    If Not LoadPic(App.Path + Parameters(0), Name) Then Debug.Print "PIC_LoadErr:" + Parameters(0)
                Case 1
                    If Not LoadWav(App.Path + Parameters(0), Name) Then Debug.Print "WAV_LoadErr:" + Parameters(0)
                Case 2
                    If Not Me.LoadAnimate(Parameters(0), Name, Parameters(1), Parameters(2), Parameters(3)) Then
                        Debug.Print "ANI_LoadErr:" + Name
                    End If
                Case 3
                    '保留
            End Select
        Wend
    End If
    OpenResTable = True
    '注意并未关闭文件
    
    Exit Function
    
errhand:
    OpenResTable = False
End Function
 
Private Function LocTableRes(ByVal sName As String, ByVal lID As Long) As Boolean
    '定位资源
    On Error GoTo errhand:
    Dim address As Long, s As String
    Dim Name As String, ID As Long, Style As Long, ParametersNum As Long
    If m_ResTableFileNum = 0 Then Exit Function
    
    Seek #m_ResTableFileNum, 1
    While Not EOF(m_ResTableFileNum)
        address = Loc(m_ResTableFileNum)
        Input #m_ResTableFileNum, Name, ID, Style, ParametersNum
        Line Input #m_ResTableFileNum, s
        If Name = sName Or ID = lID Then
            Seek #m_ResTableFileNum, address + 1
            LocTableRes = True
            Exit Function
        End If
    Wend
    
errhand:
    LocTableRes = False
End Function


Public Function LoadTableRes(ByVal sName As String, ByVal lID As Long) As Boolean
    On Error GoTo errhand:
    '首先定位
    Dim Parameters() As String, i As Long
    Dim Name As String, ID As Long, Style As Long, ParametersNum As Long
    sName = Trim(sName)
    If Not LocTableRes(sName, lID) Then Exit Function
    
    Input #m_ResTableFileNum, Name, ID, Style, ParametersNum
    'Loaded Res
    ReDim Parameters(ParametersNum - 1)
    'Read Parameters
    For i = 0 To ParametersNum - 1
        Input #m_ResTableFileNum, Parameters(i)
    Next i
    'Load Res
    Select Case Style
        Case 0
            If Not LoadPic(App.Path + Parameters(0), Name) Then Debug.Print "PIC_LoadErr:" + Parameters(0)
        Case 1
            If Not LoadWav(App.Path + Parameters(0), Name) Then Debug.Print "WAV_LoadErr:" + Parameters(0)
        Case 2
            If Not FindPic(Parameters(0)) Then LoadTableRes Parameters(0), 0
            If Not LoadAnimate(Parameters(0), Name, Parameters(1), Parameters(2), Parameters(3)) Then
                Debug.Print "ANI_LoadErr:" + Name
            End If
        Case 3
            '保留
    End Select
    LoadTableRes = True
    Exit Function
    
errhand:
    LoadTableRes = False
End Function

Private Function FindPic(ByVal Name As String) As Boolean
    On Error GoTo errhand:
    If Not m_Surfaces.Item(Name) Is Nothing Then FindPic = True
    Exit Function
errhand:
    FindPic = False
End Function
